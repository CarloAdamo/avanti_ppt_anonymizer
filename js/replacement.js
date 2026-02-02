// Avanti PPT Anonymizer — Text replacement with format preservation

import { makeKey } from './config.js';

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

export async function replaceAllShapes(slideData, rewrites) {
    let count = 0;

    const rewriteMap = new Map();
    for (const r of rewrites) {
        rewriteMap.set(makeKey(r.slideIndex, r.shapeIndex, r.row, r.col, r.groupChildIndex), r.rewrittenText);
    }

    const fontMap = new Map();
    for (const slide of slideData) {
        for (const t of slide.texts) {
            fontMap.set(makeKey(slide.slideIndex, t.shapeIndex, t.row, t.col, t.groupChildIndex), t.fontData);
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
                const key = makeKey(slide.slideIndex, t.shapeIndex, t.row, t.col, t.groupChildIndex);
                const newText = rewriteMap.get(key);
                if (!newText || newText === t.text) continue;

                try {
                    const shape = shapes.items[t.shapeIndex];
                    const fontData = fontMap.get(key);

                    let textFrame;
                    if (t.row !== undefined && t.groupChildIndex !== undefined) {
                        // Table cell within a group
                        const group = shape.group;
                        const childShapes = group.shapes;
                        childShapes.load('items');
                        await context.sync();
                        const child = childShapes.items[t.groupChildIndex];
                        const cell = child.table.getCell(t.row, t.col);
                        textFrame = cell.body;
                    } else if (t.row !== undefined) {
                        // Table cell
                        const cell = shape.table.getCell(t.row, t.col);
                        textFrame = cell.body;
                    } else if (t.groupChildIndex !== undefined) {
                        // Text shape within a group
                        const group = shape.group;
                        const childShapes = group.shapes;
                        childShapes.load('items');
                        await context.sync();
                        textFrame = childShapes.items[t.groupChildIndex].textFrame;
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
