// Avanti PPT Anonymizer — Local + AI classification, buildRewrites

import {
    SUPABASE_URL,
    SUPABASE_ANON_KEY,
    LOCAL_PATTERNS,
    PLACEHOLDER_MAP,
    makeKey,
    escapeRegex
} from './config.js';

export function classifyLocally(slideData) {
    const classified = [];
    const unclassified = [];

    for (const slide of slideData) {
        for (const t of slide.texts) {
            const text = t.text;

            // Table cells: classify entirely locally (no AI needed)
            if (t.row !== undefined) {
                if (t.row === 0) {
                    classified.push({
                        slideIndex: slide.slideIndex,
                        shapeIndex: t.shapeIndex,
                        groupChildIndex: t.groupChildIndex,
                        row: t.row,
                        col: t.col,
                        category: 'table_header'
                    });
                } else {
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
                        groupChildIndex: t.groupChildIndex,
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
                        groupChildIndex: t.groupChildIndex,
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
                    groupChildIndex: t.groupChildIndex,
                    text,
                    paragraphs: paragraphs.filter(p => p.trim())
                });
            }
        }
    }

    return { classified, unclassified };
}

export async function classifyWithAI(items) {
    // Send with sequential IDs — AI doesn't need to track slide/shape refs
    const payload = items.map((s, idx) => {
        const item = { id: idx };
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

    // Map AI results back to full references
    return (result.classifications || []).map(c => {
        const original = items[c.id];
        if (!original) return null;
        const classification = {
            slideIndex: original.slideIndex,
            shapeIndex: original.shapeIndex,
            category: c.category,
        };
        if (original.groupChildIndex !== undefined) {
            classification.groupChildIndex = original.groupChildIndex;
        }
        if (c.label) classification.label = c.label;
        return classification;
    }).filter(Boolean);
}

// Fallback when AI is unavailable: classify all unclassified as 'body'
export function fallbackClassify(unclassified) {
    return unclassified.map(s => ({
        slideIndex: s.slideIndex,
        shapeIndex: s.shapeIndex,
        groupChildIndex: s.groupChildIndex,
        category: 'body'
    }));
}

export function buildRewrites(slideData, classifications) {
    const classMap = new Map();
    for (const c of classifications) {
        classMap.set(makeKey(c.slideIndex, c.shapeIndex, c.row, c.col, c.groupChildIndex), c);
    }

    const rewrites = [];

    for (const slide of slideData) {
        for (const t of slide.texts) {
            const key = makeKey(slide.slideIndex, t.shapeIndex, t.row, t.col, t.groupChildIndex);
            const classification = classMap.get(key);

            if (!classification) continue;

            const { category, label } = classification;

            // Categories that should be kept as-is
            if (category === 'keep' || category === 'table_header' || category === 'section_label') {
                continue;
            }

            const text = t.text;
            const isTableCell = t.row !== undefined;
            let rewrittenText;

            if (isTableCell) {
                // Short placeholder for ALL table cell replacements to avoid overflow
                rewrittenText = '[...]';
            } else if (category === 'body') {
                const paragraphs = text.split('\r');
                rewrittenText = paragraphs.map(p =>
                    p.trim() ? PLACEHOLDER_MAP.body : ''
                ).join('\r');
            } else if (category === 'label_value') {
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
                    groupChildIndex: t.groupChildIndex,
                    rewrittenText
                });
            }
        }
    }

    return rewrites;
}
