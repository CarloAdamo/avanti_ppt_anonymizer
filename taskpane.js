// Avanti PPT Anonymizer
// Anonymiserar presentationer genom att skriva om text med AI

const SUPABASE_URL = 'https://vnjcwffdhywckwnjothu.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZuamN3ZmZkaHl3Y2t3bmpvdGh1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU4NjA4MTAsImV4cCI6MjA4MTQzNjgxMH0.ETCptr-BYt7wunTOXVAsBCsv9L9kICR30GGHoC5X3ZQ';

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

        showLoading('AI skriver om text...');
        const rewrites = await anonymizeWithAI(slideData);

        if (rewrites.length === 0) {
            hideLoading();
            showStatus('Ingen text behövde anonymiseras.', 'success');
            return;
        }

        showLoading('Ersätter text...');
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

        for (let i = 0; i < slideCollection.items.length; i++) {
            const slide = slideCollection.items[i];
            const shapes = slide.shapes;
            shapes.load('items');
            await context.sync();

            const slideTexts = [];

            for (let j = 0; j < shapes.items.length; j++) {
                try {
                    const shape = shapes.items[j];
                    const textRange = shape.textFrame.textRange;
                    textRange.load('text');
                    await context.sync();

                    const text = textRange.text;
                    if (text && text.trim()) {
                        let fontData = null;
                        if (Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
                            fontData = await captureFormatting(shape, context);
                        }
                        slideTexts.push({ shapeIndex: j, text, fontData });
                    }
                } catch (e) {
                    continue;
                }
            }

            if (slideTexts.length > 0) {
                slides.push({ slideIndex: i, texts: slideTexts });
            }
        }
    });

    return slides;
}

// ============================================
// FORMAT CAPTURE
// ============================================

async function captureFormatting(shape, context) {
    const fontProps = ['bold', 'italic', 'color', 'size', 'name', 'underline'];

    try {
        const textRange = shape.textFrame.textRange;
        textRange.font.load(fontProps);
        await context.sync();

        // Check if formatting is uniform across the whole shape
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
// AI ANONYMIZATION
// ============================================

async function anonymizeWithAI(slideData) {
    const texts = slideData.flatMap(slide =>
        slide.texts.map(t => ({
            slideIndex: slide.slideIndex,
            shapeIndex: t.shapeIndex,
            text: t.text
        }))
    );

    try {
        const response = await fetch(`${SUPABASE_URL}/functions/v1/analyze-presentation`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify({ texts })
        });

        if (!response.ok) {
            throw new Error('AI analysis failed');
        }

        const result = await response.json();
        return result.rewrites || [];

    } catch (error) {
        console.warn('AI analysis unavailable, using local patterns:', error);
        return rewriteWithPatterns(texts);
    }
}

// Local pattern-based rewrite (fallback)
function rewriteWithPatterns(texts) {
    const patterns = [
        { regex: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, replacement: '[email]' },
        { regex: /(?:\+46|0)[\s-]?(?:\d[\s-]?){8,11}/g, replacement: '[telefon]' },
        { regex: /\d{6,8}[-\s]?\d{4}/g, replacement: '[personnummer]' },
        { regex: /\d+(?:[,.\s]\d{3})*(?:[,.\s]\d+)?\s*(?:kr|SEK|MSEK|TSEK|EUR|USD|miljoner|miljarder)/gi, replacement: '[belopp]' },
        { regex: /\d+(?:[,.]\d+)?\s*%/g, replacement: '[X]%' },
        { regex: /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi, replacement: '[URL]' }
    ];

    const rewrites = [];

    for (const item of texts) {
        let rewritten = item.text;
        for (const p of patterns) {
            rewritten = rewritten.replace(p.regex, p.replacement);
        }
        if (rewritten !== item.text) {
            rewrites.push({
                slideIndex: item.slideIndex,
                shapeIndex: item.shapeIndex,
                rewrittenText: rewritten
            });
        }
    }

    return rewrites;
}

// ============================================
// TEXT REPLACEMENT WITH FORMAT PRESERVATION
// ============================================

async function replaceAllShapes(slideData, rewrites) {
    let count = 0;

    const rewriteMap = new Map();
    for (const r of rewrites) {
        rewriteMap.set(`${r.slideIndex}-${r.shapeIndex}`, r.rewrittenText);
    }

    const fontMap = new Map();
    for (const slide of slideData) {
        for (const t of slide.texts) {
            fontMap.set(`${slide.slideIndex}-${t.shapeIndex}`, t.fontData);
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
                const key = `${slide.slideIndex}-${t.shapeIndex}`;
                const newText = rewriteMap.get(key);
                if (!newText || newText === t.text) continue;

                try {
                    const shape = shapes.items[t.shapeIndex];
                    const fontData = fontMap.get(key);
                    await replaceWithFormatting(shape, context, newText, fontData);
                    count++;
                } catch (e) {
                    console.warn('Could not replace shape:', e);
                }
            }
        }
    });

    return count;
}

async function replaceWithFormatting(shape, context, newText, fontData) {
    const textRange = shape.textFrame.textRange;
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
