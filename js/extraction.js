// Avanti PPT Anonymizer — Text extraction from shapes/tables/groups

async function captureFormatting(textFrame, context) {
    const fontProps = ['bold', 'italic', 'color', 'size', 'name', 'underline'];

    try {
        const textRange = textFrame.textRange;
        textRange.font.load(fontProps);
        await context.sync();

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
            offset += para.length + 1;
        }

        return { type: 'perParagraph', fonts: paragraphFonts };

    } catch (e) {
        console.warn('Could not capture formatting:', e);
        return null;
    }
}

async function extractTextShape(shape, shapeIndex, context, slideTexts, canCaptureFormatting) {
    const textRange = shape.textFrame.textRange;
    textRange.load('text');
    await context.sync();

    const text = textRange.text;
    if (text && text.trim()) {
        let fontData = null;
        if (canCaptureFormatting) {
            fontData = await captureFormatting(shape.textFrame, context);
        }
        slideTexts.push({ shapeIndex, text, fontData });
    }
}

async function extractTableCells(shape, shapeIndex, context, slideTexts) {
    const table = shape.table;

    // Get dimensions
    table.rows.load('count');
    const firstRow = table.rows.getItemAt(0);
    firstRow.load('cellCount');
    await context.sync();

    const rowCount = table.rows.count;
    const colCount = firstRow.cellCount;

    // Batch load all cell texts using getCell
    const cellRefs = [];
    for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
            try {
                const cell = table.getCell(r, c);
                cell.body.textRange.load('text');
                cellRefs.push({ row: r, col: c, cell });
            } catch (e) {
                // Merged or inaccessible cell
            }
        }
    }
    await context.sync();

    for (const { row, col, cell } of cellRefs) {
        try {
            const text = cell.body.textRange.text;
            if (text && text.trim()) {
                slideTexts.push({ shapeIndex, text, fontData: null, row, col });
            }
        } catch (e) {
            // Cell text couldn't be read
        }
    }
}

async function extractGroupShapes(shape, shapeIndex, context, slideTexts, canCaptureFormatting) {
    const group = shape.group;
    const childShapes = group.shapes;
    childShapes.load('items');
    await context.sync();

    // Load types for children
    for (const child of childShapes.items) {
        try { child.load('type'); } catch (e) { /* ignore */ }
    }
    try { await context.sync(); } catch (e) { /* ignore */ }

    for (let k = 0; k < childShapes.items.length; k++) {
        const child = childShapes.items[k];
        let childType = null;
        try { childType = child.type; } catch (e) { /* ignore */ }

        try {
            if (childType === 'Table') {
                // Table within a group — extract cells but tag with groupChildIndex
                const childTable = child.table;
                childTable.rows.load('count');
                const childFirstRow = childTable.rows.getItemAt(0);
                childFirstRow.load('cellCount');
                await context.sync();

                const rowCount = childTable.rows.count;
                const colCount = childFirstRow.cellCount;
                const cellRefs = [];
                for (let r = 0; r < rowCount; r++) {
                    for (let c = 0; c < colCount; c++) {
                        try {
                            const cell = childTable.getCell(r, c);
                            cell.body.textRange.load('text');
                            cellRefs.push({ row: r, col: c, cell });
                        } catch (e) { /* merged cell */ }
                    }
                }
                await context.sync();
                for (const { row, col, cell } of cellRefs) {
                    try {
                        const text = cell.body.textRange.text;
                        if (text && text.trim()) {
                            slideTexts.push({ shapeIndex, groupChildIndex: k, text, fontData: null, row, col });
                        }
                    } catch (e) { /* skip */ }
                }
            } else {
                // Text shape within a group
                const textRange = child.textFrame.textRange;
                textRange.load('text');
                await context.sync();

                const text = textRange.text;
                if (text && text.trim()) {
                    let fontData = null;
                    if (canCaptureFormatting) {
                        fontData = await captureFormatting(child.textFrame, context);
                    }
                    slideTexts.push({ shapeIndex, groupChildIndex: k, text, fontData });
                }
            }
        } catch (e) {
            continue;
        }
    }
}

export async function extractAllText() {
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

            // Batch load shape types for dispatch
            let typesLoaded = false;
            try {
                for (const s of shapes.items) {
                    s.load('type');
                }
                await context.sync();
                typesLoaded = true;
            } catch (e) {
                console.warn('Could not load shape types, using fallback detection');
            }

            const slideTexts = [];

            for (let j = 0; j < shapes.items.length; j++) {
                const shape = shapes.items[j];
                let shapeType = null;
                if (typesLoaded) {
                    try { shapeType = shape.type; } catch (e) { /* ignore */ }
                }

                try {
                    if (shapeType === 'Table') {
                        await extractTableCells(shape, j, context, slideTexts);
                    } else if (shapeType === 'Group') {
                        await extractGroupShapes(shape, j, context, slideTexts, canCaptureFormatting);
                    } else {
                        // Text shapes, unknown types, images (will throw) — try textFrame
                        await extractTextShape(shape, j, context, slideTexts, canCaptureFormatting);
                    }
                } catch (e) {
                    // If type-based dispatch failed, try fallback cascade
                    if (!typesLoaded) {
                        try {
                            await extractTableCells(shape, j, context, slideTexts);
                        } catch (e2) {
                            try {
                                await extractGroupShapes(shape, j, context, slideTexts, canCaptureFormatting);
                            } catch (e3) {
                                continue;
                            }
                        }
                    } else {
                        console.warn(`Skipping shape ${j} (type: ${shapeType}):`, e.message);
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
