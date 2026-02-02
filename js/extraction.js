// Avanti PPT Anonymizer — Text extraction from shapes/tables/groups

// ---- PPTX binary retrieval & table parsing (fallback for Table API) ----

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function getPptxFile() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: 65536 },
            async (result) => {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error(result.error.message));
                    return;
                }
                try {
                    const file = result.value;
                    const slices = [];
                    for (let i = 0; i < file.sliceCount; i++) {
                        slices.push(await new Promise((res, rej) => {
                            file.getSliceAsync(i, (r) =>
                                r.status === Office.AsyncResultStatus.Succeeded
                                    ? res(r.value.data)
                                    : rej(new Error(r.error.message))
                            );
                        }));
                    }
                    file.closeAsync();
                    let size = 0;
                    for (const s of slices) size += s.byteLength;
                    const buf = new Uint8Array(size);
                    let off = 0;
                    for (const s of slices) {
                        buf.set(new Uint8Array(s), off);
                        off += s.byteLength;
                    }
                    resolve(buf.buffer);
                } catch (e) {
                    reject(e);
                }
            }
        );
    });
}

function extractCellsFromTable(tbl) {
    const cells = [];
    let row = 0;
    for (const tr of tbl.childNodes) {
        if (tr.nodeType !== 1 || tr.namespaceURI !== A_NS || tr.localName !== 'tr') continue;
        let col = 0;
        for (const tc of tr.childNodes) {
            if (tc.nodeType !== 1 || tc.namespaceURI !== A_NS || tc.localName !== 'tc') continue;
            const parts = [];
            const tEls = tc.getElementsByTagNameNS(A_NS, 't');
            for (let i = 0; i < tEls.length; i++) {
                if (tEls[i].textContent) parts.push(tEls[i].textContent);
            }
            const text = parts.join('');
            if (text.trim()) cells.push({ row, col, text });
            col++;
        }
        row++;
    }
    return cells;
}

async function parseTablesFromPptx() {
    const data = await getPptxFile();
    const zip = await JSZip.loadAsync(data);
    const parser = new DOMParser();

    // Resolve slide order from presentation.xml
    const presXml = await zip.file('ppt/presentation.xml').async('text');
    const presDoc = parser.parseFromString(presXml, 'application/xml');
    const relsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');

    const rIdMap = {};
    const rels = relsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < rels.length; i++) {
        rIdMap[rels[i].getAttribute('Id')] = rels[i].getAttribute('Target');
    }

    const sldIds = presDoc.getElementsByTagNameNS(P_NS, 'sldId');
    const slideFiles = [];
    for (let i = 0; i < sldIds.length; i++) {
        const rId = sldIds[i].getAttributeNS(R_NS, 'id');
        if (rIdMap[rId]) slideFiles.push('ppt/' + rIdMap[rId]);
    }

    // Parse each slide — only top-level tables (direct children of spTree)
    const result = {};
    for (let si = 0; si < slideFiles.length; si++) {
        const f = zip.file(slideFiles[si]);
        if (!f) continue;
        const xml = await f.async('text');
        const doc = parser.parseFromString(xml, 'application/xml');

        const spTree = doc.getElementsByTagNameNS(P_NS, 'spTree')[0];
        if (!spTree) continue;

        const tables = [];
        for (const child of spTree.childNodes) {
            if (child.nodeType !== 1) continue;
            if (child.localName === 'graphicFrame' && child.namespaceURI === P_NS) {
                const tbl = child.getElementsByTagNameNS(A_NS, 'tbl')[0];
                if (!tbl) continue;
                const cells = extractCellsFromTable(tbl);
                if (cells.length > 0) tables.push(cells);
            }
        }
        if (tables.length > 0) result[si] = tables;
    }

    return result;
}

// ---- Shape extraction functions ----

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

async function extractTableCells(shape, shapeIndex, context, slideTexts, pptxCells) {
    // Try getTable() API (documented method)
    try {
        const table = shape.getTable();
        table.load('rowCount, columnCount');
        await context.sync();

        const allCells = [];
        for (let r = 0; r < table.rowCount; r++) {
            for (let c = 0; c < table.columnCount; c++) {
                const cell = table.getCellOrNullObject(r, c);
                cell.load('text');
                allCells.push({ row: r, col: c, cell });
            }
        }
        await context.sync();

        for (const { row, col, cell } of allCells) {
            if (!cell.isNullObject && cell.text && cell.text.trim()) {
                slideTexts.push({ shapeIndex, text: cell.text, fontData: null, row, col });
            }
        }
        return; // API succeeded
    } catch (e) {
        console.warn(`getTable() failed for shape ${shapeIndex}:`, e.message);
    }

    // Fallback: use pre-parsed PPTX table data
    if (pptxCells && pptxCells.length > 0) {
        console.log(`Shape ${shapeIndex}: recovered ${pptxCells.length} table cells from PPTX`);
        for (const cell of pptxCells) {
            slideTexts.push({
                shapeIndex,
                text: cell.text,
                fontData: null,
                row: cell.row,
                col: cell.col
            });
        }
    } else {
        console.warn(`Shape ${shapeIndex} (Table): no PPTX data available — skipping`);
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
                // Table within a group — use getTable() API
                const childTable = child.getTable();
                childTable.load('rowCount, columnCount');
                await context.sync();

                const allCells = [];
                for (let r = 0; r < childTable.rowCount; r++) {
                    for (let c = 0; c < childTable.columnCount; c++) {
                        const cell = childTable.getCellOrNullObject(r, c);
                        cell.load('text');
                        allCells.push({ row: r, col: c, cell });
                    }
                }
                await context.sync();

                for (const { row, col, cell } of allCells) {
                    if (!cell.isNullObject && cell.text && cell.text.trim()) {
                        slideTexts.push({ shapeIndex, groupChildIndex: k, text: cell.text, fontData: null, row, col });
                    }
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

    // Pre-parse PPTX for table text (fallback when Table API is unavailable)
    let pptxTables = null;
    try {
        pptxTables = await parseTablesFromPptx();
        console.log('PPTX parsed:', Object.keys(pptxTables).length, 'slide(s) with tables');
    } catch (e) {
        console.warn('PPTX table parsing unavailable:', e.message);
    }

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
            let tableOrdinal = 0;

            for (let j = 0; j < shapes.items.length; j++) {
                const shape = shapes.items[j];
                let shapeType = null;
                if (typesLoaded) {
                    try { shapeType = shape.type; } catch (e) { /* ignore */ }
                }

                try {
                    if (shapeType === 'Table') {
                        await extractTableCells(shape, j, context, slideTexts, pptxTables?.[i]?.[tableOrdinal]);
                        tableOrdinal++;
                    } else if (shapeType === 'Group') {
                        await extractGroupShapes(shape, j, context, slideTexts, canCaptureFormatting);
                    } else {
                        // Text shapes, unknown types, images (will throw) — try textFrame
                        await extractTextShape(shape, j, context, slideTexts, canCaptureFormatting);
                    }
                } catch (e) {
                    // Primary extraction failed — try other methods as fallback
                    let recovered = false;
                    if (shapeType !== 'Table') {
                        try {
                            await extractTableCells(shape, j, context, slideTexts, null);
                            recovered = true;
                        } catch (e2) { /* not a table */ }
                    }
                    if (!recovered && shapeType !== 'Group') {
                        try {
                            await extractGroupShapes(shape, j, context, slideTexts, canCaptureFormatting);
                            recovered = true;
                        } catch (e3) { /* not a group */ }
                    }
                    if (!recovered) {
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
