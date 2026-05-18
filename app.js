// ========================================================
// 1. OFFICE ENGINE INITIALIZATION
// ========================================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeDragEngine();
        initializeFinanceButtons();
    }
});

// ========================================================
// 2. THE DRAG & DROP ENGINE (The 6 Dots)
// ========================================================
function initializeDragEngine() {
    const dragHandle = document.getElementById("drag-handle");
    const toolbar = document.getElementById("toolbar");

    let isDragging = false;
    let startX, startY, initialLeft, initialTop;

    dragHandle.addEventListener("mousedown", (e) => {
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        
        const rect = toolbar.getBoundingClientRect();
        initialLeft = rect.left;
        initialTop = rect.top;
        
        document.addEventListener("mousemove", dragMove);
        document.addEventListener("mouseup", dragStop);
        e.preventDefault(); // Stop accidental text highlights
    });

    function dragMove(e) {
        if (!isDragging) return;
        const dx = e.clientX - startX;
        const dy = e.clientY - startY;
        
        toolbar.style.left = (initialLeft + dx) + "px";
        toolbar.style.top = (initialTop + dy) + "px";
    }

    function dragStop() {
        isDragging = false;
        document.removeEventListener("mousemove", dragMove);
        document.removeEventListener("mouseup", dragStop);
    }
}

// ========================================================
// 3. CORE FINANCE LOGIC
// ========================================================
function initializeFinanceButtons() {
    // --- Audit X-Ray (The "Bowtie") ---
    document.getElementById("btn-audit").onclick = () => runExcelAction(async (range) => {
        const formulas = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
        const constants = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.constants);
        await range.context.sync();

        if (!formulas.isNullObject) formulas.format.fill.color = "#cce5ff"; 
        if (!constants.isNullObject) constants.format.fill.color = "#ffe0b2"; 
    });

    // --- Professional Formatting ---
    document.getElementById("btn-format").onclick = () => runExcelAction(r => {
        r.numberFormat = [["£#,##0.00"]];
        r.format.font.bold = true;
    });

    // --- Smart AutoSum (Horizontal & Vertical) ---
    document.getElementById("btn-sum").onclick = () => runExcelAction(async (range) => {
        range.load(["address", "rowCount", "columnCount"]);
        await range.context.sync();
        
        let target = range.rowCount >= range.columnCount ? 
                     range.getLastRow().getOffsetRange(1, 0) : 
                     range.getLastColumn().getOffsetRange(0, 1);
        
        target.formulas = [["=SUM(" + range.address + ")"]];
        target.format.font.bold = true;
        target.format.fill.color = "#e2efda"; 
    });

    // --- Add 20% VAT ---
    document.getElementById("btn-vat").onclick = () => runExcelAction(async (r) => {
        r.load("address"); await r.context.sync();
        const target = r.getLastColumn().getOffsetRange(0, 1);
        target.formulas = [["=" + r.address + "*1.2"]];
        target.numberFormat = [["£#,##0.00"]];
        target.format.fill.color = "#fff2cc"; 
    });

    // --- Variance Δ% (Smart Direction) ---
    document.getElementById("btn-delta").onclick = () => runExcelAction(async (range) => {
        range.load(["address", "rowCount", "columnCount"]);
        await range.context.sync();
        const cells = range.address.split(":");
        const f = cells[0], s = cells.length > 1 ? cells[1] : cells[0];
        
        let target = range.rowCount >= range.columnCount ? 
                     range.getLastRow().getOffsetRange(1, 0) : 
                     range.getLastColumn().getOffsetRange(0, 1);
                     
        target.formulas = [["=(" + s + "-" + f + ")/" + f]];
        target.numberFormat = [["0.00%"]];
        target.format.fill.color = "#e0f2f1"; 
    });

    // --- Navigation & Clean ---
    document.getElementById("btn-freeze").onclick = () => runExcelAction(async (r) => { r.worksheet.freezePanes.freezeRows(1); });
    document.getElementById("btn-unfreeze").onclick = () => runExcelAction(async (r) => { r.worksheet.freezePanes.unfreeze(); });
    document.getElementById("btn-clean").onclick = () => runExcelAction(r => { 
        r.clear(Excel.ClearApplyTo.formats); 
        r.format.fill.clear(); 
    });
}

// ========================================================
// 4. CORE RUNTIME PIPELINE
// ========================================================
async function runExcelAction(callback) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            await callback(range);
            await context.sync();
        });
    } catch (e) { console.log("Logic Error: " + e.message); }
}