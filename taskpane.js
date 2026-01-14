// Avanti PPT Anonymizer
// Skannar och anonymiserar känslig information i PowerPoint-presentationer

const SUPABASE_URL = 'https://vnjcwffdhywckwnjothu.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZuamN3ZmZkaHl3Y2t3bmpvdGh1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU4NjA4MTAsImV4cCI6MjA4MTQzNjgxMH0.ETCptr-BYt7wunTOXVAsBCsv9L9kICR30GGHoC5X3ZQ';

// State
let findings = [];
let slideData = [];

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log('Avanti Anonymizer loaded');
        initializeUI();
    }
});

function initializeUI() {
    // Scan button
    document.getElementById('scan-btn').addEventListener('click', scanPresentation);

    // Select all button
    document.getElementById('select-all-btn').addEventListener('click', toggleSelectAll);

    // Anonymize button
    document.getElementById('anonymize-btn').addEventListener('click', anonymizeSelected);

    // Rescan buttons
    document.getElementById('rescan-btn').addEventListener('click', resetAndScan);
    document.getElementById('new-scan-btn').addEventListener('click', resetAndScan);
}

// ============================================
// SCANNING
// ============================================

async function scanPresentation() {
    showLoading('Extraherar text från presentationen...');
    hideStatus();

    try {
        // Step 1: Extract all text from slides
        slideData = await extractAllText();

        if (slideData.length === 0) {
            hideLoading();
            showStatus('Ingen text hittades i presentationen.', 'info');
            return;
        }

        showLoading('Analyserar innehåll med AI...');

        // Step 2: Analyze with AI
        findings = await analyzeWithAI(slideData);

        hideLoading();

        if (findings.length === 0) {
            showStatus('Ingen känslig information hittades!', 'success');
            return;
        }

        // Step 3: Show findings
        displayFindings(findings);

    } catch (error) {
        hideLoading();
        console.error('Scan error:', error);
        showStatus('Ett fel uppstod: ' + error.message, 'error');
    }
}

async function extractAllText() {
    const slides = [];

    await PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const slideCollection = presentation.slides;
        slideCollection.load('items');
        await context.sync();

        // Load all shapes for all slides in one batch
        const allShapes = [];
        for (const slide of slideCollection.items) {
            slide.shapes.load('items');
            allShapes.push(slide.shapes);
        }
        await context.sync();

        // Load all text from all shapes in one batch
        const shapeTextMap = [];
        for (let i = 0; i < slideCollection.items.length; i++) {
            const shapes = allShapes[i];
            for (let j = 0; j < shapes.items.length; j++) {
                const shape = shapes.items[j];
                try {
                    shape.textFrame.load('hasText');
                    shapeTextMap.push({ slideIndex: i, shapeIndex: j, shape });
                } catch (e) {
                    // Shape doesn't support textFrame
                }
            }
        }
        await context.sync();

        // Now load text only for shapes that have text
        const textShapes = [];
        for (const item of shapeTextMap) {
            try {
                if (item.shape.textFrame.hasText) {
                    item.shape.textFrame.textRange.load('text');
                    textShapes.push(item);
                }
            } catch (e) {
                // Skip shapes without text
            }
        }
        await context.sync();

        // Collect results
        const slideTextsMap = new Map();
        for (const item of textShapes) {
            try {
                const text = item.shape.textFrame.textRange.text;
                if (text && text.trim()) {
                    if (!slideTextsMap.has(item.slideIndex)) {
                        slideTextsMap.set(item.slideIndex, []);
                    }
                    slideTextsMap.get(item.slideIndex).push({
                        shapeIndex: item.shapeIndex,
                        text: text
                    });
                }
            } catch (e) {
                // Skip
            }
        }

        // Convert to array
        for (const [slideIndex, texts] of slideTextsMap) {
            slides.push({ slideIndex, texts });
        }
        slides.sort((a, b) => a.slideIndex - b.slideIndex);
    });

    return slides;
}

// ============================================
// AI ANALYSIS
// ============================================

async function analyzeWithAI(slideData) {
    // Prepare text for analysis
    const allTexts = slideData.flatMap(slide =>
        slide.texts.map(t => ({
            slideIndex: slide.slideIndex,
            shapeIndex: t.shapeIndex,
            text: t.text
        }))
    );

    try {
        // Call Edge Function for AI analysis
        const response = await fetch(`${SUPABASE_URL}/functions/v1/analyze-presentation`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify({ texts: allTexts })
        });

        if (!response.ok) {
            throw new Error('AI analysis failed');
        }

        const result = await response.json();
        return result.findings || [];

    } catch (error) {
        console.warn('AI analysis unavailable, using local patterns:', error);
        // Fallback to local pattern matching
        return analyzeWithPatterns(allTexts);
    }
}

// Local pattern-based analysis (fallback)
function analyzeWithPatterns(texts) {
    const findings = [];
    const patterns = [
        {
            type: 'email',
            regex: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g,
            suggestion: '[email borttagen]'
        },
        {
            type: 'phone',
            regex: /(?:\+46|0)[\s-]?(?:\d[\s-]?){8,11}/g,
            suggestion: '[telefon borttagen]'
        },
        {
            type: 'personnummer',
            regex: /\d{6,8}[-\s]?\d{4}/g,
            suggestion: '[personnummer borttaget]'
        },
        {
            type: 'financial',
            regex: /\d+(?:[,.\s]\d{3})*(?:[,.\s]\d+)?\s*(?:kr|SEK|MSEK|TSEK|EUR|USD|miljoner|miljarder)/gi,
            suggestion: '[belopp borttaget]'
        },
        {
            type: 'percentage',
            regex: /\d+(?:[,.]\d+)?\s*%/g,
            suggestion: '[X] %'
        },
        {
            type: 'url',
            regex: /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi,
            suggestion: '[URL borttagen]'
        }
    ];

    // Track unique findings
    const foundItems = new Map();

    for (const textItem of texts) {
        for (const pattern of patterns) {
            const matches = textItem.text.match(pattern.regex);
            if (matches) {
                for (const match of matches) {
                    const key = `${pattern.type}:${match}`;
                    if (!foundItems.has(key)) {
                        foundItems.set(key, {
                            type: pattern.type,
                            original: match,
                            suggestion: pattern.suggestion,
                            occurrences: []
                        });
                    }
                    foundItems.get(key).occurrences.push({
                        slideIndex: textItem.slideIndex,
                        shapeIndex: textItem.shapeIndex
                    });
                }
            }
        }
    }

    return Array.from(foundItems.values());
}

// ============================================
// UI DISPLAY
// ============================================

function displayFindings(findings) {
    const findingsList = document.getElementById('findings-list');
    findingsList.innerHTML = '';

    findings.forEach((finding, index) => {
        const item = document.createElement('div');
        item.className = 'finding-item';
        item.innerHTML = `
            <input type="checkbox" class="finding-checkbox" data-index="${index}" checked>
            <div class="finding-content">
                <span class="finding-type ${finding.type}">${getTypeLabel(finding.type)}</span>
                <div>
                    <span class="finding-original">${escapeHtml(finding.original)}</span>
                </div>
                <div class="finding-replacement">
                    <span class="finding-arrow">→</span>
                    <input type="text" value="${escapeHtml(finding.suggestion)}" data-index="${index}">
                </div>
                <div class="finding-occurrences">
                    ${finding.occurrences.length} förekomst${finding.occurrences.length > 1 ? 'er' : ''}
                    på slide ${[...new Set(finding.occurrences.map(o => o.slideIndex + 1))].join(', ')}
                </div>
            </div>
        `;
        findingsList.appendChild(item);
    });

    // Update count
    document.getElementById('findings-count').textContent = findings.length;

    // Show findings section
    document.getElementById('scan-section').classList.add('hidden');
    document.getElementById('findings-section').classList.remove('hidden');
    document.getElementById('done-section').classList.add('hidden');

    // Enable/disable anonymize button based on selections
    updateAnonymizeButton();

    // Add event listeners for checkboxes and inputs
    document.querySelectorAll('.finding-checkbox').forEach(cb => {
        cb.addEventListener('change', updateAnonymizeButton);
    });

    document.querySelectorAll('.finding-replacement input').forEach(input => {
        input.addEventListener('input', (e) => {
            const index = parseInt(e.target.dataset.index);
            findings[index].suggestion = e.target.value;
        });
    });
}

function getTypeLabel(type) {
    const labels = {
        'company': 'Företag',
        'person': 'Person',
        'email': 'E-post',
        'phone': 'Telefon',
        'personnummer': 'Personnr',
        'financial': 'Belopp',
        'percentage': 'Procent',
        'date': 'Datum',
        'url': 'URL',
        'other': 'Övrigt'
    };
    return labels[type] || type;
}

function updateAnonymizeButton() {
    const checked = document.querySelectorAll('.finding-checkbox:checked');
    const btn = document.getElementById('anonymize-btn');
    btn.disabled = checked.length === 0;
    btn.textContent = checked.length > 0
        ? `🔒 Anonymisera ${checked.length} valda`
        : '🔒 Anonymisera valda';
}

function toggleSelectAll() {
    const checkboxes = document.querySelectorAll('.finding-checkbox');
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    checkboxes.forEach(cb => cb.checked = !allChecked);
    updateAnonymizeButton();

    const btn = document.getElementById('select-all-btn');
    btn.textContent = allChecked ? 'Välj alla' : 'Avmarkera alla';
}

// ============================================
// ANONYMIZATION
// ============================================

async function anonymizeSelected() {
    const selectedIndices = Array.from(document.querySelectorAll('.finding-checkbox:checked'))
        .map(cb => parseInt(cb.dataset.index));

    if (selectedIndices.length === 0) return;

    showLoading('Anonymiserar...');

    try {
        let replacedCount = 0;

        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            const slideCollection = presentation.slides;
            slideCollection.load('items');
            await context.sync();

            for (const index of selectedIndices) {
                const finding = findings[index];

                // Group occurrences by slide for efficiency
                const occurrencesBySlide = new Map();
                for (const occ of finding.occurrences) {
                    if (!occurrencesBySlide.has(occ.slideIndex)) {
                        occurrencesBySlide.set(occ.slideIndex, []);
                    }
                    occurrencesBySlide.get(occ.slideIndex).push(occ.shapeIndex);
                }

                // Replace in each slide
                for (const [slideIndex, shapeIndices] of occurrencesBySlide) {
                    const slide = slideCollection.items[slideIndex];
                    const shapes = slide.shapes;
                    shapes.load('items');
                    await context.sync();

                    // Get unique shape indices
                    const uniqueShapes = [...new Set(shapeIndices)];

                    for (const shapeIndex of uniqueShapes) {
                        const shape = shapes.items[shapeIndex];
                        try {
                            shape.textFrame.textRange.load('text');
                            await context.sync();

                            let text = shape.textFrame.textRange.text;
                            const originalText = text;

                            // Replace all occurrences in this shape
                            text = text.split(finding.original).join(finding.suggestion);

                            if (text !== originalText) {
                                shape.textFrame.textRange.text = text;
                                await context.sync();
                                replacedCount++;
                            }
                        } catch (e) {
                            console.warn('Could not replace in shape:', e);
                        }
                    }
                }
            }
        });

        hideLoading();
        showDoneSection(replacedCount);

    } catch (error) {
        hideLoading();
        console.error('Anonymization error:', error);
        showStatus('Ett fel uppstod vid anonymisering: ' + error.message, 'error');
    }
}

function showDoneSection(count) {
    document.getElementById('scan-section').classList.add('hidden');
    document.getElementById('findings-section').classList.add('hidden');
    document.getElementById('done-section').classList.remove('hidden');
    document.getElementById('replaced-count').textContent = count;
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function resetAndScan() {
    findings = [];
    slideData = [];

    document.getElementById('scan-section').classList.remove('hidden');
    document.getElementById('findings-section').classList.add('hidden');
    document.getElementById('done-section').classList.add('hidden');

    hideStatus();
    scanPresentation();
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

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
