// Avanti PPT Anonymizer
// Generaliserar presentationer med klassificering + platshållare

const SUPABASE_URL = 'https://vnjcwffdhywckwnjothu.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZuamN3ZmZkaHl3Y2t3bmpvdGh1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU4NjA4MTAsImV4cCI6MjA4MTQzNjgxMH0.ETCptr-BYt7wunTOXVAsBCsv9L9kICR30GGHoC5X3ZQ';

// ============================================
// CONSTANTS
// ============================================

const SECTION_LABELS = new Set([
    'syfte', 'mål', 'bakgrund', 'agenda', 'sammanfattning',
    'nästa steg', 'tidplan', 'organisation', 'risker',
    'budget', 'resurser', 'bilagor', 'innehåll', 'analys',
    'resultat', 'slutsats', 'rekommendationer', 'översikt',
    'introduktion', 'diskussion', 'metod', 'uppföljning'
]);

const PLACEHOLDER_MAP = {
    title:       '[Rubrik]',
    body:        '[Beskrivning]',
    name:        '[Namn]',
    number:      '[Värde]',
    initials:    '[Initialer]',
    email:       '[email]',
    phone:       '[telefon]',
    url:         '[URL]',
    date:        '[Period]',
};

// ============================================
// LOCAL CLASSIFICATION PATTERNS
// ============================================

const LOCAL_PATTERNS = [
    {
        category: 'email',
        test: (text) => /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(text.trim())
    },
    {
        category: 'phone',
        test: (text) => /^(?:\+46|0)[\s-]?(?:\d[\s-]?){8,11}$/.test(text.trim())
    },
    {
        category: 'url',
        test: (text) => /^https?:\/\/[^\s]+$/.test(text.trim())
    },
    {
        category: 'initials',
        test: (text) => /^[A-ZÅÄÖ]{2,3}(?:\s*[+&,]\s*[A-ZÅÄÖ]{2,3})+$/.test(text.trim())
    },
    {
        category: 'date',
        test: (text) => /^(?:Q[1-4]\s*\d{4}|\d{4}-\d{2}-\d{2}|\d{4}-\d{2}|\d{1,2}\/\d{1,2}[\/-]\d{2,4}|(?:jan|feb|mar|apr|maj|jun|jul|aug|sep|okt|nov|dec)\w*\s+\d{4})$/i.test(text.trim())
    },
    {
        category: 'number',
        test: (text) => /^\d+(?:[,.\s]\d{3})*(?:[,.\s]\d+)?\s*(?:kr|SEK|MSEK|TSEK|EUR|USD|miljoner|miljarder|st|%|kkr|mkr|mdr)$/i.test(text.trim())
    },
    {
        category: 'section_label',
        test: (text) => SECTION_LABELS.has(text.trim().toLowerCase())
    }
];

// ============================================
// KEY HELPER
// ============================================

function makeKey(slideIndex, shapeIndex, row, col) {
    if (row !== undefined) return `${slideIndex}-${shapeIndex}-${row}-${col}`;
    return `${slideIndex}-${shapeIndex}`;
}

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log('Avanti Anonymizer loaded');
        initializeUI();
    }
});

function initializeUI() {
    document.getElementById('scan-btn').addEventListener('click', anonymizePresentation);
    document.getElementById('new-scan-btn').addEventListener('click', resetAndAnonymize);
}

// ============================================
// MAIN FLOW
// ============================================

async function anonymizePresentation() {
    hideStatus();

    try {
        showLoading('Extraherar text...');
        const slideData = await extractAllText();

        if (slideData.length === 0) {
            hideLoading();
            showStatus('Ingen text hittades i presentationen.', 'info');
            return;
        }

        showLoading('Klassificerar text...');
        const { classified, unclassified } = classifyLocally(slideData);

        let aiClassifications = [];
        if (unclassified.length > 0) {
            try {
                showLoading('AI klassificerar text...');
                aiClassifications = await classifyWithAI(unclassified);
            } catch (error) {
                console.warn('AI classification unavailable, using fallback:', error);
                aiClassifications = fallbackClassify(unclassified);
            }
        }

        const allClassifications = [...classified, ...aiClassifications];

        showLoading('Ersätter text...');
        const rewrites = buildRewrites(slideData, allClassifications);

        if (rewrites.length === 0) {
            hideLoading();
            showStatus('Ingen text behövde generaliseras.', 'success');
            return;
        }

        const count = await replaceAllShapes(slideData, rewrites);

        hideLoading();
        showDoneSection(count);

    } catch (error) {
        hideLoading();
        console.error('Anonymization error:', error);
        showStatus('Ett fel uppstod: ' + error.message, 'error');
    }
}

// ============================================
// TEXT EXTRACTION
// ============================================

async function extractAllText() {
    const slides = [];

    await PowerPoint.run(async (context) => {
        const slideCollection = context.presentation.slides;
        slideCollection.load('items');
        await context.sync();

        const canCaptureFormatting = Office.context.requirements.isSetSupported('PowerPointApi', '1.4');

        for (let i = 0; i < slideCollection.items.length; i++) {
            const slide = slideCollection.items[i];
            const shapes = slide.shapes;
            shapes.load('items');
            await context.sync();

            const slideTexts = [];

            for (let j = 0; j < shapes.items.length; j++) {
                const shape = shapes.items[j];

                // Try as text shape first
                try {
                    const textRange = shape.textFrame.textRange;
                    textRange.load('text');
                    await context.sync();

                    const text = textRange.text;
                    if (text && text.trim()) {
                        let fontData = null;
                        if (canCaptureFormatting) {
                            fontData = await captureFormatting(shape.textFrame, context);
                        }
                        slideTexts.push({ shapeIndex: j, text, fontData });
                    }
                } catch (e) {
                    // Not a text shape — try as table
                    try {
                        await extractTableCells(shape, j, context, slideTexts);
                    } catch (e2) {
                        // Neither text nor table, skip
                        continue;
                    }
                }
            }

            if (slideTexts.length > 0) {
                slides.push({ slideIndex: i, texts: slideTexts });
            }
        }
    });

    return slides;
}

async function extractTableCells(shape, shapeIndex, context, slideTexts) {
    const table = shape.table;
    table.rows.load('items');
    await context.sync();

    // Load all row cells
    for (const row of table.rows.items) {
        row.cells.load('items');
    }
    await context.sync();

    // Load all cell text in one batch
    const cellRefs = [];
    for (let r = 0; r < table.rows.items.length; r++) {
        const row = table.rows.items[r];
        for (let c = 0; c < row.cells.items.length; c++) {
            const cell = row.cells.items[c];
            cell.body.textRange.load('text');
            cellRefs.push({ row: r, col: c, cell });
        }
    }
    await context.sync();

    // Collect cells with text
    for (const { row, col, cell } of cellRefs) {
        const text = cell.body.textRange.text;
        if (text && text.trim()) {
            slideTexts.push({ shapeIndex, text, fontData: null, row, col });
        }
    }
}

// ============================================
// FORMAT CAPTURE
// ============================================

async function captureFormatting(textFrame, context) {
    const fontProps = ['bold', 'italic', 'color', 'size', 'name', 'underline'];

    try {
        const textRange = textFrame.textRange;
        textRange.font.load(fontProps);
        await context.sync();

        // Check if formatting is uniform
        const wholeFont = {};
        let isUniform = true;
        for (const prop of fontProps) {
            const val = textRange.font[prop];
            if (val === null) { isUniform = false; break; }
            wholeFont[prop] = val;
        }

        if (isUniform) {
            return { type: 'uniform', font: wholeFont };
        }

        // Mixed formatting — capture per paragraph
        const text = textRange.text;
        const paragraphs = text.split('\r');
        let offset = 0;
        const paragraphFonts = [];

        for (const para of paragraphs) {
            if (para.length > 0) {
                try {
                    const subRange = textRange.getSubstring(offset, para.length);
                    subRange.font.load(fontProps);
                    await context.sync();

                    const paraFont = {};
                    for (const prop of fontProps) {
                        const val = subRange.font[prop];
                        if (val !== null && val !== undefined) paraFont[prop] = val;
                    }
                    paragraphFonts.push(paraFont);
                } catch (e) {
                    paragraphFonts.push({});
                }
            } else {
                paragraphFonts.push({});
            }
            offset += para.length + 1; // +1 for \r
        }

        return { type: 'perParagraph', fonts: paragraphFonts };

    } catch (e) {
        console.warn('Could not capture formatting:', e);
        return null;
    }
}

// ============================================
// LOCAL CLASSIFICATION
// ============================================

function classifyLocally(slideData) {
    const classified = [];
    const unclassified = [];

    for (const slide of slideData) {
        for (const t of slide.texts) {
            const text = t.text;

            // Table cells: classify entirely locally (no AI needed)
            if (t.row !== undefined) {
                if (t.row === 0) {
                    // First row = table header, keep as-is
                    classified.push({
                        slideIndex: slide.slideIndex,
                        shapeIndex: t.shapeIndex,
                        row: t.row,
                        col: t.col,
                        category: 'table_header'
                    });
                } else {
                    // Other rows: try local patterns, default to 'body'
                    let category = 'body';
                    for (const pattern of LOCAL_PATTERNS) {
                        if (pattern.test(text)) {
                            category = pattern.category;
                            break;
                        }
                    }
                    classified.push({
                        slideIndex: slide.slideIndex,
                        shapeIndex: t.shapeIndex,
                        row: t.row,
                        col: t.col,
                        category
                    });
                }
                continue;
            }

            // Regular shapes: try local patterns
            let matched = false;
            for (const pattern of LOCAL_PATTERNS) {
                if (pattern.test(text)) {
                    classified.push({
                        slideIndex: slide.slideIndex,
                        shapeIndex: t.shapeIndex,
                        category: pattern.category
                    });
                    matched = true;
                    break;
                }
            }

            if (!matched) {
                const paragraphs = text.split('\r');
                unclassified.push({
                    slideIndex: slide.slideIndex,
                    shapeIndex: t.shapeIndex,
                    text,
                    paragraphs: paragraphs.filter(p => p.trim())
                });
            }
        }
    }

    return { classified, unclassified };
}

// ============================================
// AI CLASSIFICATION
// ============================================

async function classifyWithAI(shapes) {
    const payload = shapes.map(s => {
        const item = { slideIndex: s.slideIndex, shapeIndex: s.shapeIndex };
        if (s.paragraphs && s.paragraphs.length > 1) {
            item.paragraphs = s.paragraphs;
        } else {
            item.text = s.text;
        }
        return item;
    });

    const response = await fetch(`${SUPABASE_URL}/functions/v1/analyze-presentation`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
        },
        body: JSON.stringify({ shapes: payload })
    });

    if (!response.ok) {
        throw new Error('AI classification failed');
    }

    const result = await response.json();
    return result.classifications || [];
}

// Fallback when AI is unavailable: classify all unclassified as 'body'
function fallbackClassify(unclassified) {
    return unclassified.map(s => ({
        slideIndex: s.slideIndex,
        shapeIndex: s.shapeIndex,
        category: 'body'
    }));
}

// ============================================
// BUILD REWRITES FROM CLASSIFICATIONS
// ============================================

function buildRewrites(slideData, classifications) {
    const classMap = new Map();
    for (const c of classifications) {
        classMap.set(makeKey(c.slideIndex, c.shapeIndex, c.row, c.col), c);
    }

    const rewrites = [];

    for (const slide of slideData) {
        for (const t of slide.texts) {
            const key = makeKey(slide.slideIndex, t.shapeIndex, t.row, t.col);
            const classification = classMap.get(key);

            if (!classification) continue;

            const { category, label } = classification;

            // Categories that should be kept as-is
            if (category === 'keep' || category === 'table_header' || category === 'section_label') {
                continue;
            }

            const text = t.text;
            let rewrittenText;

            if (category === 'body') {
                // Replace each paragraph with [Beskrivning], preserve empty lines
                const paragraphs = text.split('\r');
                rewrittenText = paragraphs.map(p =>
                    p.trim() ? PLACEHOLDER_MAP.body : ''
                ).join('\r');
            } else if (category === 'label_value') {
                // Preserve the label part, replace the value
                if (label) {
                    const labelPattern = new RegExp(`^(${escapeRegex(label)}\\s*[:;]\\s*)`, 'i');
                    const match = text.match(labelPattern);
                    if (match) {
                        rewrittenText = match[1] + '[Namn]';
                    } else {
                        rewrittenText = label + ': [Namn]';
                    }
                } else {
                    rewrittenText = PLACEHOLDER_MAP.body;
                }
            } else if (PLACEHOLDER_MAP[category]) {
                rewrittenText = PLACEHOLDER_MAP[category];
            } else {
                continue;
            }

            if (rewrittenText !== text) {
                rewrites.push({
                    slideIndex: slide.slideIndex,
                    shapeIndex: t.shapeIndex,
                    row: t.row,
                    col: t.col,
                    rewrittenText
                });
            }
        }
    }

    return rewrites;
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ============================================
// TEXT REPLACEMENT WITH FORMAT PRESERVATION
// ============================================

async function replaceAllShapes(slideData, rewrites) {
    let count = 0;

    const rewriteMap = new Map();
    for (const r of rewrites) {
        rewriteMap.set(makeKey(r.slideIndex, r.shapeIndex, r.row, r.col), r.rewrittenText);
    }

    const fontMap = new Map();
    for (const slide of slideData) {
        for (const t of slide.texts) {
            fontMap.set(makeKey(slide.slideIndex, t.shapeIndex, t.row, t.col), t.fontData);
        }
    }

    await PowerPoint.run(async (context) => {
        const slideCollection = context.presentation.slides;
        slideCollection.load('items');
        await context.sync();

        for (const slide of slideData) {
            const pptSlide = slideCollection.items[slide.slideIndex];
            const shapes = pptSlide.shapes;
            shapes.load('items');
            await context.sync();

            for (const t of slide.texts) {
                const key = makeKey(slide.slideIndex, t.shapeIndex, t.row, t.col);
                const newText = rewriteMap.get(key);
                if (!newText || newText === t.text) continue;

                try {
                    const shape = shapes.items[t.shapeIndex];
                    const fontData = fontMap.get(key);

                    let textFrame;
                    if (t.row !== undefined) {
                        // Table cell
                        const cell = shape.table.getCell(t.row, t.col);
                        textFrame = cell.body;
                    } else {
                        textFrame = shape.textFrame;
                    }

                    await replaceWithFormatting(textFrame, context, newText, fontData);
                    count++;
                } catch (e) {
                    console.warn('Could not replace shape:', e);
                }
            }
        }
    });

    return count;
}

async function replaceWithFormatting(textFrame, context, newText, fontData) {
    const textRange = textFrame.textRange;
    textRange.text = newText;
    await context.sync();

    if (!fontData) return;

    if (fontData.type === 'uniform') {
        for (const [prop, val] of Object.entries(fontData.font)) {
            textRange.font[prop] = val;
        }
        await context.sync();
        return;
    }

    if (fontData.type === 'perParagraph') {
        const paragraphs = newText.split('\r');
        let offset = 0;

        for (let i = 0; i < paragraphs.length && i < fontData.fonts.length; i++) {
            if (paragraphs[i].length > 0) {
                try {
                    const subRange = textRange.getSubstring(offset, paragraphs[i].length);
                    for (const [prop, val] of Object.entries(fontData.fonts[i])) {
                        subRange.font[prop] = val;
                    }
                } catch (e) {
                    console.warn('Could not apply paragraph formatting:', e);
                }
            }
            offset += paragraphs[i].length + 1;
        }
        await context.sync();
    }
}

// ============================================
// UI HELPERS
// ============================================

function showDoneSection(count) {
    document.getElementById('scan-section').classList.add('hidden');
    document.getElementById('done-section').classList.remove('hidden');
    document.getElementById('replaced-count').textContent = count;
}

function resetAndAnonymize() {
    document.getElementById('scan-section').classList.remove('hidden');
    document.getElementById('done-section').classList.add('hidden');
    hideStatus();
    anonymizePresentation();
}

function showLoading(text) {
    document.getElementById('loading-text').textContent = text;
    document.getElementById('loading').classList.remove('hidden');
}

function hideLoading() {
    document.getElementById('loading').classList.add('hidden');
}

function showStatus(message, type) {
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = `status ${type}`;
    status.classList.remove('hidden');
}

function hideStatus() {
    document.getElementById('status').classList.add('hidden');
}
