let sheetData = []; // Data storage for the spreadsheet
const numRows = 10; // Number of rows in the sheet
const numCols = 10; // Number of columns in the sheet
let dependencyGraph = {}; // Track cell dependencies

// Initialize the sheet with empty data
function initSheet() {
    const sheet = document.getElementById("sheet").getElementsByTagName("tbody")[0];
    sheet.innerHTML = ""; // Clear any existing rows

    // Create columns headers (A, B, C, ..., Z)
    const headerRow = sheet.insertRow();
    for (let col = 0; col < numCols; col++) {
        const th = document.createElement("th");
        th.innerText = String.fromCharCode(65 + col); // A, B, C, ...
        headerRow.appendChild(th);
    }

    // Create rows and cells
    for (let row = 0; row < numRows; row++) {
        const rowElement = sheet.insertRow();
        sheetData[row] = [];
        for (let col = 0; col < numCols; col++) {
            const cell = rowElement.insertCell();
            cell.contentEditable = true;
            cell.setAttribute("data-row", row);
            cell.setAttribute("data-col", col);
            cell.addEventListener("input", updateCell);
            cell.addEventListener("mousedown", startDrag);
            cell.addEventListener("mouseup", endDrag);
            cell.addEventListener("mousemove", dragFill);
        }
    }
}

// Drag-and-drop related variables
let isDragging = false;
let dragStartCell = null;

// Start drag operation
function startDrag(event) {
    isDragging = true;
    dragStartCell = event.target;
}

// End drag operation
function endDrag(event) {
    isDragging = false;
    dragStartCell = null;
}

// Drag functionality
function dragFill(event) {
    if (!isDragging || !dragStartCell) return;

    const startRow = parseInt(dragStartCell.getAttribute("data-row"));
    const startCol = parseInt(dragStartCell.getAttribute("data-col"));
    const endRow = parseInt(event.target.getAttribute("data-row"));
    const endCol = parseInt(event.target.getAttribute("data-col"));

    for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row++) {
        for (let col = Math.min(startCol, endCol); col <= Math.max(startCol, endCol); col++) {
            const draggedCell = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
            const startValue = sheetData[startRow][startCol];

            // Copy value or formula during drag
            if (startValue.startsWith("=")) {
                const relativeFormula = adjustFormulaForDrag(startValue, startRow, startCol, row, col);
                draggedCell.innerText = relativeFormula;
                sheetData[row][col] = relativeFormula;
            } else {
                draggedCell.innerText = startValue;
                sheetData[row][col] = startValue;
            }

            updateDependencies(row, col); // Update dependent cells
        }
    }
}

// Adjust formula for drag operation (relative references)
function adjustFormulaForDrag(formula, startRow, startCol, newRow, newCol) {
    return formula.replace(/[A-Z]+\d+/g, (cellRef) => {
        const [colLetter, rowNumber] = cellRef.match(/([A-Z]+)(\d+)/).slice(1);
        const colIndex = columnToIndex(colLetter);
        const rowIndex = parseInt(rowNumber) - 1;

        const newColIndex = colIndex + (newCol - startCol);
        const newRowIndex = rowIndex + (newRow - startRow);

        return `${indexToColumn(newColIndex)}${newRowIndex + 1}`;
    });
}

// Column letter to index (e.g., A -> 0, B -> 1)
function columnToIndex(col) {
    return col.charCodeAt(0) - 65;
}

// Index to column letter (e.g., 0 -> A, 1 -> B)
function indexToColumn(index) {
    return String.fromCharCode(65 + index);
}

// Update cell data and dependencies
function updateCell(event) {
    const row = parseInt(event.target.getAttribute("data-row"));
    const col = parseInt(event.target.getAttribute("data-col"));
    const value = event.target.innerText;

    sheetData[row][col] = value;
    updateDependencies(row, col); // Update dependent cells
}

// Update dependent cells dynamically
function updateDependencies(row, col) {
    const cellRef = `${indexToColumn(col)}${row + 1}`;
    if (dependencyGraph[cellRef]) {
        dependencyGraph[cellRef].forEach((dependentCell) => {
            const [depRow, depCol] = parseCellReference(dependentCell);
            const formula = sheetData[depRow][depCol];

            if (formula.startsWith("=")) {
                const result = evaluateFormula(formula.substring(1));
                sheetData[depRow][depCol] = result;
                const depCell = document.querySelector(`[data-row="${depRow}"][data-col="${depCol}"]`);
                depCell.innerText = result;
            }
        });
    }
}

// Parse cell reference (e.g., "A1" -> [0, 0])
function parseCellReference(cellRef) {
    const [colLetter, rowNumber] = cellRef.match(/([A-Z]+)(\d+)/).slice(1);
    return [parseInt(rowNumber) - 1, columnToIndex(colLetter)];
}

// Evaluate formula (basic implementation for demo purposes)
function evaluateFormula(formula) {
    return formula.replace(/[A-Z]+\d+/g, (cellRef) => {
        const [row, col] = parseCellReference(cellRef);
        return parseFloat(sheetData[row][col]) || 0;
    });
}
// Apply bold formatting
function applyBold() {
    document.execCommand('bold');
}

// Apply italic formatting
function applyItalic() {
    document.execCommand('italic');
}

// Increase font size
function increaseFontSize() {
    document.execCommand('fontSize', false, '5');
}

// Decrease font size
function decreaseFontSize() {
    document.execCommand('fontSize', false, '2');
}

// Clear all data in the sheet
function clearData() {
    const sheet = document.getElementById("sheet").getElementsByTagName("tbody")[0];
    const rows = sheet.rows;
    for (let row = 0; row < rows.length; row++) {
        for (let col = 0; col < rows[row].cells.length; col++) {
            rows[row].cells[col].innerText = "";
            sheetData[row][col] = "";
        }
    }
}

// Helper: Convert column letter to index (e.g., A -> 0, B -> 1)
function columnToIndex(col) {
    return col.charCodeAt(0) - 65;
}

// Mathematical Functions
function sum(range) {
    return calculateRange(range, (values) => values.reduce((a, b) => a + b, 0));
}

function average(range) {
    return calculateRange(range, (values) => values.reduce((a, b) => a + b, 0) / values.length);
}

function max(range) {
    return calculateRange(range, (values) => Math.max(...values));
}

function min(range) {
    return calculateRange(range, (values) => Math.min(...values));
}

function count(range) {
    const [start, end] = range.split(":");
    const startCell = start.match(/([A-Z]+)(\d+)/);
    const endCell = end.match(/([A-Z]+)(\d+)/);

    const startRow = parseInt(startCell[2]) - 1;
    const endRow = parseInt(endCell[2]) - 1;
    const startCol = columnToIndex(startCell[1]);
    const endCol = columnToIndex(endCell[1]);

    let numericCount = 0;

    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            const value = sheetData[i][j];
            if (!isNaN(value) && value.trim() !== "") {
                numericCount++;
            }
        }
    }

    return numericCount;
}


// Helper: Calculate range values
function calculateRange(range, callback) {
    const [start, end] = range.split(":");
    const startCell = start.match(/([A-Z]+)(\d+)/);
    const endCell = end.match(/([A-Z]+)(\d+)/);

    const startRow = parseInt(startCell[2]) - 1;
    const endRow = parseInt(endCell[2]) - 1;
    const startCol = columnToIndex(startCell[1]);
    const endCol = columnToIndex(endCell[1]);

    const values = [];
    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            const value = parseFloat(sheetData[i][j]) || 0;
            values.push(value);
        }
    }

    return callback(values);
}

// Data Quality Functions
function trim(cell) {
    const trimmedValue = cell.trim();
    return trimmedValue;
}

function upper(range) {
    applyToRange(range, (value) => value.toUpperCase());
}

function lower(range) {
    applyToRange(range, (value) => value.toLowerCase());
}

// Apply transformations to range
function applyToRange(range, callback) {
    const [start, end] = range.split(":");
    const startCell = start.match(/([A-Z]+)(\d+)/);
    const endCell = end.match(/([A-Z]+)(\d+)/);

    const startRow = parseInt(startCell[2]) - 1;
    const endRow = parseInt(endCell[2]) - 1;
    const startCol = columnToIndex(startCell[1]);
    const endCol = columnToIndex(endCell[1]);

    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            const currentCell = document.querySelector(`[data-row="${i}"][data-col="${j}"]`);
            const currentValue = sheetData[i][j] || "";
            const updatedValue = callback(currentValue);
            sheetData[i][j] = updatedValue;
            currentCell.innerText = updatedValue;
        }
    }
}
function removeDuplicates(range) {
    const [start, end] = range.split(":");
    const startCell = start.match(/([A-Z]+)(\d+)/);
    const endCell = end.match(/([A-Z]+)(\d+)/);

    // Convert cell references (e.g., A1, C5) to row and column indices
    const startRow = parseInt(startCell[2]) - 1; // Rows are 0-indexed in the array
    const endRow = parseInt(endCell[2]) - 1;
    const startCol = columnToIndex(startCell[1]);
    const endCol = columnToIndex(endCell[1]);

    const seenRows = new Set();

    for (let i = startRow; i <= endRow; i++) {
        const rowData = [];
        for (let j = startCol; j <= endCol; j++) {
            rowData.push(sheetData[i][j] || ""); // Collect cell data
        }
        const rowKey = rowData.join(","); // Create a unique key for the row

        if (seenRows.has(rowKey)) {
            // Clear duplicate row data in sheetData and the DOM
            for (let j = startCol; j <= endCol; j++) {
                sheetData[i][j] = ""; // Clear in the data array
                const cell = document.querySelector(`[data-row="${i}"][data-col="${j}"]`);
                if (cell) cell.innerText = ""; // Clear the visible table cell
            }
        } else {
            seenRows.add(rowKey); // Add unique row to the set
        }
    }
    alert("Duplicates removed successfully!");
}

// Helper function to convert column letters to indices (e.g., A -> 0, B -> 1, etc.)
function columnToIndex(column) {
    let index = 0;
    for (let i = 0; i < column.length; i++) {
        index = index * 26 + (column.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }
    return index - 1;
}

function findAndReplace(findText, replaceText, range) {
    // Parse the range into start and end cell references
    const [start, end] = range.split(":");
    const startCell = start.match(/([A-Z]+)(\d+)/);
    const endCell = end.match(/([A-Z]+)(\d+)/);

    const startRow = parseInt(startCell[2]) - 1; // Convert row to zero-index
    const endRow = parseInt(endCell[2]) - 1;
    const startCol = columnToIndex(startCell[1]); // Convert column to zero-index
    const endCol = columnToIndex(endCell[1]);

    // Loop through the range
    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            const currentCell = document.querySelector(`[data-row="${i}"][data-col="${j}"]`);
            const currentValue = sheetData[i][j] || ""; // Get current cell value or empty string

            // Replace text and update if the cell contains the findText
            if (currentValue.includes(findText)) {
                const updatedValue = currentValue.replace(new RegExp(findText, "g"), replaceText);
                sheetData[i][j] = updatedValue; // Update internal data
                currentCell.innerText = updatedValue; // Update cell in UI
            }
        }
    }
}
// Formula Execution
function executeFormula() {
    const formula = document.getElementById("formulaInput").value.trim();
    if (formula.startsWith("=")) {
        const parts = formula.substring(1).split("(");
        const functionName = parts[0];
        const args = parts[1].replace(")", "");

        let result = null;



        switch (functionName.toUpperCase()) {
            case "SUM":
                alert("Result: " + sum(args));
                break;
            case "AVERAGE":
                alert("Result: " + average(args));
                break;
            case "MAX":
                alert("Result: " + max(args));
                break;
            case "MIN":
                alert("Result: " + min(args));
                break;
            case "COUNT":
                alert("Result: " + count(args));
                break;
            case "TRIM":
                alert("Result: " + trim(args));
                break;
            case "UPPER":
                upper(args);
                alert("Applied UPPERCASE transformation.");
                break;
            case "LOWER":
                lower(args);
                alert("Applied lowercase transformation.");
                break;
            case "REMOVE_DUPLICATES":
                removeDuplicates(args);
                alert("Removed duplicates in the selected range.");
                break;
            case "FIND_AND_REPLACE":
                const params = args.split(",").map((param) => param.trim());
                if (params.length === 3) {
                    const findText = params[0].replace(/['"]+/g, ""); // Remove quotes
                    const replaceText = params[1].replace(/['"]+/g, ""); // Remove quotes
                    const range = params[2];
                    findAndReplace(findText, replaceText, range);
                    alert(`Replaced '${findText}' with '${replaceText}' in range ${range}.`);
                } else {
                    alert("Invalid FIND_AND_REPLACE syntax! Use: FIND_AND_REPLACE('findText', 'replaceText', 'range')");
                }
                break;
            default:
                alert("Invalid function!");
        }
    }
}
// Save the spreadsheet data to localStorage
function saveSheet() {
    localStorage.setItem("sheetData", JSON.stringify(sheetData));
    alert("Spreadsheet saved!");
}

// Load the spreadsheet data from localStorage
function loadSheet() {
    const storedData = localStorage.getItem("sheetData");
    if (storedData) {
        sheetData = JSON.parse(storedData);
        renderSheet();
        alert("Spreadsheet loaded!");
    } else {
        alert("No saved data found.");
    }
}

// Render the sheet with the saved data
function renderSheet() {
    const sheet = document.getElementById("sheet").getElementsByTagName("tbody")[0];
    const rows = sheet.rows;
    
    for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
            const cell = rows[i].cells[j];
            cell.innerText = sheetData[i][j] || "";
        }
    }
}

// Generate a chart based on the data
function generateChart() {
    const chartData = sheetData.map(row => row.map(cell => parseFloat(cell) || 0));
    const labels = Array.from({ length: numRows }, (_, i) => `Row ${i + 1}`);

    const ctx = document.getElementById('dataChart').getContext('2d');
    const data = {
        labels: labels,
        datasets: chartData.map((row, i) => ({
            label: `Column ${i + 1}`,
            data: row,
            borderColor: `hsl(${i * 360 / numCols}, 100%, 50%)`,
            fill: false
        }))
    };

    new Chart(ctx, {
        type: 'line',
        data: data
    });
}

// Initialize the sheet on page load
initSheet();
