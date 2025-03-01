let sheets = {};
let currentSheet = null;
let isEditing = false;
let editingCells = [];
let originalValues = {};
let sortColumn = null;
let sortDirection = null;
let currentPage = 1;
let rowsPerPage = 10;

const excelThemeColors = [
    '#FFFFFF', '#000000', '#EEECE1', '#1F497D', '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'
];

const STATIC_PASSWORD = "secret123";

function getDynamicTimeSuffix() {
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    return hours + minutes;
}

const CORRECT_PASSWORD = STATIC_PASSWORD + getDynamicTimeSuffix();

$(document).ready(function () {
    let storedData = localStorage.getItem("sheets");
    if (storedData) {
        sheets = JSON.parse(storedData);
        currentSheet = Object.keys(sheets)[0];
        generateTable();
        updateSheetList(); // Replace updateSheetSelector
    }

    if (window.location.pathname === "/export-code") {
        const expectedPassword = STATIC_PASSWORD + getDynamicTimeSuffix();
        const password = prompt(`Enter the password to export the project code (hint: secret123 + current time HHMM, e.g., secret123${getDynamicTimeSuffix()})`);
        if (password === expectedPassword) {
            exportProjectFolder();
        } else {
            alert("Incorrect password! Export denied.");
            window.history.pushState({}, document.title, '/');
        }
    }

    updateUIState();

    $("#excelFile").on("click", function (e) {
        if (isEditing && editingCells.length > 0) {
            let hasChanges = editingCells.some(cell => {
                let row = $(cell).data("row") !== undefined ? $(cell).data("row") : 0;
                let col = $(cell).data("col") - ($(cell).closest("th").hasClass("col-number") ? 1 : 0);
                let cellKey = `${row},${col}`;
                let text = $(cell).hasClass("editable-col-name") ? $(cell).text() : $(cell).text().replace(/^\d+: /, '').replace(/[▲▼]/g, '').trim();
                return text !== originalValues[cellKey];
            });
            if (hasChanges) {
                let proceed = confirm("You have unsaved changes. Do you want to save them before selecting a new file? Click OK to save and proceed, Cancel to discard and proceed.");
                if (proceed) {
                    saveToLocalStorage();
                } else {
                    let sheet = sheets[currentSheet];
                    editingCells.forEach(cell => {
                        let row = $(cell).data("row") !== undefined ? $(cell).data("row") : 0;
                        let col = $(cell).data("col") - ($(cell).closest("th").hasClass("col-number") ? 1 : 0);
                        let cellKey = `${row},${col}`;
                        sheet.tableData[row][col] = originalValues[cellKey];
                        if (row === 0) {
                            $(cell).text(originalValues[cellKey]);
                        } else {
                            $(cell).text(originalValues[cellKey]);
                        }
                    });
                    sheet.undoStack.pop();
                    sheet.redoStack = [];
                    editingCells = [];
                    originalValues = {};
                    isEditing = false;
                    $("#editBtn").text("Edit");
                    updateEditStatus();
                    generateTable();
                }
            }
        }
    });

    $("#excelFile").on("change", function (e) {
        let file = e.target.files[0];
        if (!file) return;
        processFile(file);
    });

    $("#dataTable").on("dragover", function (e) {
        e.preventDefault();
        e.stopPropagation();
        $(this).addClass("dragover");
    });

    $("#dataTable").on("dragleave", function (e) {
        e.preventDefault();
        e.stopPropagation();
        $(this).removeClass("dragover");
    });

    $("#dataTable").on("drop", function (e) {
        e.preventDefault();
        e.stopPropagation();
        $(this).removeClass("dragover");
        let file = e.originalEvent.dataTransfer.files[0];
        if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
            checkUnsavedChanges(() => {
                processFile(file);
            });
        } else {
            alert("Please drop a valid Excel file (.xlsx or .xls)");
        }
    });

    $("#dataTable").on("dblclick", "td:not(.row-number)", function () {
        if (isEditing) {
            let sheet = sheets[currentSheet];
            let row = $(this).data("row");
            let col = $(this).data("col");
            let cellKey = `${row},${col}`;
            if (!editingCells.includes(this)) {
                editingCells.push(this);
                originalValues[cellKey] = $(this).text();
            }
            $(this).attr("contenteditable", "true").focus();
        }
    });

    $("#dataTable").on("input", "td[contenteditable='true']", function () {
        let sheet = sheets[currentSheet];
        let row = $(this).data("row");
        let col = $(this).data("col");
        sheet.tableData[row][col] = $(this).text();
        sheet.mergedCells.forEach(merge => {
            if (row >= merge.s.r && row <= merge.e.r && col >= merge.s.c && col <= merge.e.c) {
                for (let r = merge.s.r; r <= merge.e.r; r++) {
                    for (let c = merge.s.c; c <= merge.e.c; c++) {
                        sheet.tableData[r][c] = $(this).text();
                    }
                }
            }
        });
    });

    $("#dataTable").on("dblclick", "th.col-number span.editable-col-name", function (e) {
        if (isEditing) {
            let colHeader = $(this).closest("th.col-number");
            let col = colHeader.data("col") - 1;
            let cellKey = `0,${col}`;
            if (!editingCells.includes(this)) {
                editingCells.push(this);
                originalValues[cellKey] = sheets[currentSheet].tableData[0][col] || '';
            }
            $(this).attr("contenteditable", "true").focus();
            e.stopPropagation();
        }
    });

    $("#dataTable").on("input", "th.col-number span.editable-col-name[contenteditable='true']", function () {
        let sheet = sheets[currentSheet];
        let col = $(this).closest("th.col-number").data("col") - 1;
        sheet.tableData[0][col] = $(this).text();
    });

    $("#dataTable").on("blur", "th.col-number span.editable-col-name[contenteditable='true']", function () {
        $(this).attr("contenteditable", "false");
        generateTable();
    });

    updateEditStatus();
    updateButtonStates();

    let popupCount = parseInt(sessionStorage.getItem("popupCount")) || 0;
    if (popupCount < 2) {
        $("#welcomeModal").css("display", "flex");
        popupCount++;
        sessionStorage.setItem("popupCount", popupCount);
    }
    $("#closeModal").on("click", function () {
        $("#welcomeModal").hide();
    });

    // Context menu for sheet renaming
    $(document).on("contextmenu", ".sheet-tab", function (e) {
        e.preventDefault();
        let sheetName = $(this).data("sheet");
        $("#contextMenu").remove();

        // Create context menu
        let contextMenu = $(`
            <div id="contextMenu" style="position: absolute; background: white; border: 1px solid #ccc; padding: 5px; z-index: 1000;">
                <div class="context-item" onclick="renameSheet('${sheetName}')">Edit Sheet Name</div>
            </div>
        `);
        $("body").append(contextMenu);

        // Calculate position to stay within viewport
        let menuHeight = contextMenu.outerHeight();
        let menuWidth = contextMenu.outerWidth();
        let viewportHeight = $(window).height();
        let viewportWidth = $(window).width();
        let topPos = e.pageY;
        let leftPos = e.pageMagickX;

        // Adjust top position (above if near bottom, below otherwise)
        if (topPos + menuHeight > viewportHeight) {
            topPos = e.pageY - menuHeight; // Position above the click
        }
        if (topPos < 0) topPos = 0; // Ensure it doesn’t go above the top

        // Adjust left position (left if near right edge)
        if (leftPos + menuWidth > viewportWidth) {
            leftPos = e.pageX - menuWidth; // Position to the left
        }
        if (leftPos < 0) leftPos = 0; // Ensure it doesn’t go beyond the left edge

        contextMenu.css({ top: topPos + "px", left: leftPos + "px" });
    });

    $(document).on("click", function (e) {
        if (!$(e.target).closest("#contextMenu, .sheet-tab").length) {
            $("#contextMenu").remove();
        }
    });
});

function processFile(file) {
    let reader = new FileReader();
    reader.readAsArrayBuffer(file);

    reader.onload = function (event) {
        let data = new Uint8Array(event.target.result);
        let workbook = XLSX.read(data, { type: "array" });
        sheets = {};

        workbook.SheetNames.forEach(sheetName => {
            let sheet = workbook.Sheets[sheetName];
            let range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
            let tableData = [];
            let mergedCells = sheet['!merges'] || [];
            let cellColors = {};

            for (let r = 0; r <= range.e.r; r++) {
                tableData[r] = [];
                for (let c = 0; c <= range.e.c; c++) {
                    tableData[r][c] = '';
                }
            }

            for (let cell in sheet) {
                if (cell[0] === '!') continue;
                let {r, c} = XLSX.utils.decode_cell(cell);
                tableData[r][c] = sheet[cell].w !== undefined ? sheet[cell].w : sheet[cell].v || '';
                if (sheet[cell].s && sheet[cell].s.bgColor) {
                    let bgColor = sheet[cell].s.bgColor;
                    let color;
                    if (bgColor.rgb) {
                        color = `#${bgColor.rgb}`;
                    } else if (bgColor.theme !== undefined && excelThemeColors[bgColor.theme]) {
                        color = excelThemeColors[bgColor.theme];
                    }
                    if (color) cellColors[`${r},${c}`] = color;
                }
            }

            mergedCells.forEach(merge => {
                let value = tableData[merge.s.r][merge.s.c] || '';
                let color = cellColors[`${merge.s.r},${merge.s.c}`] || '';
                for (let r = merge.s.r; r <= merge.e.r; r++) {
                    for (let c = merge.s.c; c <= merge.e.c; c++) {
                        tableData[r][c] = value;
                        if (color) cellColors[`${r},${c}`] = color;
                    }
                }
            });

            sheets[sheetName] = {
                tableData,
                mergedCells: JSON.parse(JSON.stringify(mergedCells)),
                cellColors: Object.assign({}, cellColors),
                undoStack: [{ tableData: JSON.parse(JSON.stringify(tableData)), mergedCells: JSON.parse(JSON.stringify(mergedCells)), cellColors: Object.assign({}, cellColors) }],
                redoStack: []
            };
        });

        currentSheet = workbook.SheetNames[0];
        currentPage = 1;
        isEditing = false;
        $("#editBtn").text("Edit");
        editingCells = [];
        originalValues = {};
        saveToLocalStorage();
        generateTable();
        updateSheetList(); // Replace updateSheetSelector
        updateEditStatus();
        updateButtonStates();
    };
}

function updateUIState() {
    const tableExists = Object.keys(sheets).length > 0;
    
    $("#editBtn").prop("disabled", !tableExists);
    $("#newSheetBtn").prop("disabled", !tableExists);
    $("#undoBtn").prop("disabled", !tableExists || sheets[currentSheet]?.undoStack.length <= 1);
    $("#redoBtn").prop("disabled", !tableExists || sheets[currentSheet]?.redoStack.length === 0);
    $("#clearBtn").prop("disabled", !tableExists);
    $("#exportExcelBtn").prop("disabled", !tableExists);
    $("#exportPDFBtn").prop("disabled", !tableExists);

    if (!tableExists) {
        $("#dataTable thead, #dataTable tbody").empty();
        $("#sheetSearchContainer, #paginationControls").remove();
    }
}

function checkUnsavedChanges(callback) {
    if (isEditing && editingCells.length > 0) {
        let newValues = {};
        editingCells.forEach(cell => {
            let row = $(cell).data("row") !== undefined ? $(cell).data("row") : 0;
            let col = $(cell).data("col") - ($(cell).closest("th").hasClass("col-number") ? 1 : 0);
            let cellKey = `${row},${col}`;
            let text = $(cell).hasClass("editable-col-name") ? $(cell).text() : $(cell).text().replace(/^\d+: /, '').replace(/[▲▼]/g, '').trim();
            newValues[cellKey] = text;
        });
        let hasChanges = Object.keys(newValues).some(key => newValues[key] !== originalValues[key]);
        if (hasChanges) {
            promptSaveOrDiscard(callback);
            return true;
        }
    }
    callback();
    return false;
}

function promptSaveOrDiscard(callback = () => {}) {
    if (editingCells.length === 0) {
        isEditing = false;
        $("#editBtn").text("Edit");
        updateEditStatus();
        callback();
        return;
    }

    let newValues = {};
    editingCells.forEach(cell => {
        let row = $(cell).data("row") !== undefined ? $(cell).data("row") : 0;
        let col = $(cell).data("col") - ($(cell).closest("th").hasClass("col-number") ? 1 : 0);
        let cellKey = `${row},${col}`;
        let text = $(cell).hasClass("editable-col-name") ? $(cell).text() : $(cell).text().replace(/^\d+: /, '').replace(/[▲▼]/g, '').trim();
        newValues[cellKey] = text;
        $(cell).attr("contenteditable", "false");
    });

    let hasChanges = Object.keys(newValues).some(key => newValues[key] !== originalValues[key]);
    if (hasChanges) {
        let response = confirm("You have unsaved changes. Do you want to save them before proceeding? Click OK to save, Cancel to discard.");
        if (response) {
            saveToLocalStorage();
        } else {
            let sheet = sheets[currentSheet];
            for (let cell of editingCells) {
                let row = $(cell).data("row") !== undefined ? $(cell).data("row") : 0;
                let col = $(cell).data("col") - ($(cell).closest("th").hasClass("col-number") ? 1 : 0);
                let cellKey = `${row},${col}`;
                sheet.tableData[row][col] = originalValues[cellKey];
                if (row === 0) {
                    $(cell).text(originalValues[cellKey]);
                } else {
                    $(cell).text(originalValues[cellKey]);
                }
            }
            sheet.undoStack.pop();
            sheet.redoStack = [];
        }
    } else {
        sheets[currentSheet].undoStack.pop();
        sheets[currentSheet].redoStack = [];
    }

    editingCells = [];
    originalValues = {};
    isEditing = false;
    $("#editBtn").text("Edit");
    updateEditStatus();
    generateTable();
    callback();
}

function updateEditStatus() {
    $("#editStatus").text(isEditing ? "Edit Mode: ON" : "Edit Mode: OFF")
        .removeClass("on off")
        .addClass(isEditing ? "on" : "off");
    if (isEditing) {
        $("#editInstruction").text("Double click on cells or headers to edit, then click Save")
            .show()
            .removeClass("bounce") 
            .addClass("bounce");   
    } else {
        $("#editInstruction").hide();
    }
}

function updateButtonStates() {
    if (!currentSheet || !sheets[currentSheet]) return;
    let sheet = sheets[currentSheet];
    $("#undoBtn").prop("disabled", sheet.undoStack.length <= 1 || !sheet.tableData.length);
    $("#redoBtn").prop("disabled", sheet.redoStack.length === 0 || !sheet.tableData.length);
}

// New function to update sheet list at bottom
function updateSheetList() {
    let sheetListContainer = $("#sheetListContainer");
    if (!sheetListContainer.length) {
        sheetListContainer = $('<div id="sheetListContainer" class="sheet-list-container"><div id="sheetList" class="sheet-list"></div></div>').appendTo("body");
    }
    let sheetList = $("#sheetList");
    sheetList.empty();
    Object.keys(sheets).forEach(sheetName => {
        let sheetTab = $(`<span class="sheet-tab" data-sheet="${sheetName}">${sheetName}</span>`);
        if (sheetName === currentSheet) {
            sheetTab.addClass("active-sheet");
        }
        sheetTab.on("click", function () {
            checkUnsavedChanges(() => {
                currentSheet = sheetName;
                currentPage = 1;
                sortColumn = null;
                sortDirection = null;
                generateTable();
                updateSheetList();
                updateButtonStates();
            });
        });
        sheetList.append(sheetTab);
    });
}

// Function to rename sheet
function renameSheet(oldName) {
    let newName = prompt("Enter new sheet name:", oldName);
    if (newName === null || newName === oldName) return;
    if (!newName || sheets[newName]) {
        alert("Sheet name cannot be empty or already exists!");
        return;
    }
    sheets[newName] = sheets[oldName];
    delete sheets[oldName];
    if (currentSheet === oldName) {
        currentSheet = newName;
    }
    saveToLocalStorage();
    generateTable();
    updateSheetList();
    $("#contextMenu").remove();
}

function generateTable() {
    if (!currentSheet || !sheets[currentSheet]) {
        $("#dataTable thead, #dataTable tbody").empty();
        $("#sheetSearchContainer, #paginationControls").remove();
        return;
    }

    let sheet = sheets[currentSheet];
    $("#dataTable thead, #dataTable tbody").empty();

    let tableHead = "<tr><th class='row-number' data-col='0'><span class='sort-arrows'><span class='asc' data-dir='asc'>▲</span><span class='desc' data-dir='desc'>▼</span></span></th>";
    let tableBody = "";
    let columnCount = Math.max(...sheet.tableData.map(row => row.length));
    let hiddenCells = new Set();
    let rows = sheet.tableData.slice(1).map((row, index) => ({ row, originalIndex: index + 1 }));
    let totalRows = rows.length;
    let start = rowsPerPage === Infinity ? 0 : (currentPage - 1) * rowsPerPage;
    let end = rowsPerPage === Infinity ? totalRows : Math.min(currentPage * rowsPerPage, totalRows);
    let filteredRows = rows;

    let container = $("#sheetSearchContainer");
    if (!container.length) {
        container = $('<div id="sheetSearchContainer" class="sheet-search-container"></div>').insertBefore("#dataTable");
        $('<input type="text" id="searchFilter" placeholder="Search...">').appendTo(container).on("input", function () {
            checkUnsavedChanges(() => {
                currentPage = 1;
                generateTable();
            });
        });
    }

    let searchTerm = $("#searchFilter").val().toLowerCase();
    if (searchTerm) {
        filteredRows = rows.filter(item => 
            item.row.some(cell => String(cell).toLowerCase().includes(searchTerm))
        );
        totalRows = filteredRows.length;
        start = rowsPerPage === Infinity ? 0 : (currentPage - 1) * rowsPerPage;
        end = rowsPerPage === Infinity ? totalRows : Math.min(currentPage * rowsPerPage, totalRows);
    }

    if (sortColumn !== null && sortDirection) {
        filteredRows.sort((a, b) => {
            let valA, valB;
            if (sortColumn === 0) {
                valA = a.originalIndex;
                valB = b.originalIndex;
            } else {
                valA = String(a.row[sortColumn - 1] || '').trim();
                valB = String(b.row[sortColumn - 1] || '').trim();
            }

            let dateA = Date.parse(valA);
            let dateB = Date.parse(valB);
            if (!isNaN(dateA) && !isNaN(dateB)) {
                return sortDirection === 'asc' ? dateA - dateB : dateB - dateA;
            }

            let numA = parseFloat(valA);
            let numB = parseFloat(valB);
            if (!isNaN(numA) && !isNaN(numB)) {
                return sortDirection === 'asc' ? numA - numB : numB - numA;
            }

            valA = valA.toLowerCase();
            valB = valB.toLowerCase();
            return sortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
        });
    }

    filteredRows = filteredRows.slice(start, end);

    for (let i = 0; i < columnCount; i++) {
        let cell = sheet.tableData[0][i] || '';
        let ascClass = sortColumn === (i + 1) && sortDirection === 'asc' ? 'active' : '';
        let descClass = sortColumn === (i + 1) && sortDirection === 'desc' ? 'active' : '';
        tableHead += `<th class='col-number' data-col="${i + 1}"><span class="col-num-style">${i + 1}:</span><span class="editable-col-name">${cell}</span><span class="sort-arrows"><span class="asc ${ascClass}" data-dir="asc">▲</span><span class="desc ${descClass}" data-dir="desc">▼</span></span><div class="addAtBtnColBox"><span class="add-btn" onclick="addColumnAt(${i})">+</span><span class="delete-btn" onclick="deleteColumnAt(${i})">−</span></div></th>`;
    }
    tableHead += "</tr>";

    filteredRows.forEach((item, index) => {
        let rowIndex = item.originalIndex;
        let displayIndex = start + index + 1;
        let rowHtml = `<tr><td class='row-number' data-row="${rowIndex}">${rowIndex}<div class="addAtBtnRowBox"><span class="add-btn" onclick="addRowAt(${rowIndex})">+</span><span class="delete-btn" onclick="deleteRowAt(${rowIndex})">−</span></div></td>`;
        for (let colIndex = 0; colIndex < columnCount; colIndex++) {
            let cellKey = `${rowIndex},${colIndex}`;
            if (hiddenCells.has(cellKey)) continue;

            let cell = item.row[colIndex] || '';
            let color = sheet.cellColors[cellKey] || '';
            let style = color ? `style="background-color: ${color};"` : '';
            let mergeInfo = sheet.mergedCells.find(m => m.s.r === rowIndex && m.s.c === colIndex);
            if (mergeInfo) {
                let rowspan = mergeInfo.e.r - mergeInfo.s.r + 1;
                let colspan = mergeInfo.e.c - mergeInfo.s.c + 1;
                rowHtml += `<td contenteditable="false" data-row="${rowIndex}" data-col="${colIndex}" rowspan="${rowspan}" colspan="${colspan}" ${style}>${cell}</td>`;
                for (let r = rowIndex; r <= mergeInfo.e.r; r++) {
                    for (let c = colIndex; c <= mergeInfo.e.c; c++) {
                        if (r !== rowIndex || c !== colIndex) {
                            hiddenCells.add(`${r},${c}`);
                        }
                    }
                }
            } else {
                rowHtml += `<td contenteditable="false" data-row="${rowIndex}" data-col="${colIndex}" ${style}>${cell}</td>`;
            }
        }
        rowHtml += "</tr>";
        tableBody += rowHtml;
    });

    $("#dataTable thead").html(tableHead);
    $("#dataTable tbody").html(tableBody);

    let controls = $("#paginationControls");
    if (!controls.length) {
        controls = $('<div id="paginationControls" class="pagination-controls"></div>').insertAfter("#dataTable");
    }
    controls.empty();
    let totalPages = rowsPerPage === Infinity ? 1 : Math.ceil(totalRows / rowsPerPage);
    controls.append(`<select id="rowsPerPage">
        <option value="10" ${rowsPerPage === 10 ? 'selected' : ''}>10</option>
        <option value="25" ${rowsPerPage === 25 ? 'selected' : ''}>25</option>
        <option value="50" ${rowsPerPage === 50 ? 'selected' : ''}>50</option>
        <option value="100" ${rowsPerPage === 100 ? 'selected' : ''}>100</option>
        <option value="Infinity" ${rowsPerPage === Infinity ? 'selected' : ''}>All</option>
    </select>`);
    controls.append(`<button id="prevPage" ${currentPage === 1 ? 'disabled' : ''}>Previous</button>`);
    controls.append(`<span>Page ${currentPage} of ${totalPages} (Showing ${filteredRows.length} of ${totalRows} entries)</span>`);
    controls.append(`<button id="nextPage" ${currentPage === totalPages ? 'disabled' : ''}>Next</button>`);

    $("#rowsPerPage").off("change").on("change", function () {
        checkUnsavedChanges(() => {
            rowsPerPage = $(this).val() === "Infinity" ? Infinity : parseInt($(this).val());
            currentPage = 1;
            generateTable();
        });
    });
    $("#prevPage").off("click").on("click", function () {
        checkUnsavedChanges(() => {
            if (currentPage > 1) {
                currentPage--;
                generateTable();
            }
        });
    });
    $("#nextPage").off("click").on("click", function () {
        checkUnsavedChanges(() => {
            if (currentPage < totalPages) {
                currentPage++;
                generateTable();
            }
        });
    });

    $(".col-number, .row-number").off("click").on("click", function (e) {
        if (e.target.tagName === "SPAN" && $(e.target).parent().hasClass("sort-arrows")) {
            checkUnsavedChanges(() => {
                let colIndex = parseInt($(this).data("col"));
                if (sortColumn === colIndex) {
                    if (sortDirection === 'asc') {
                        sortDirection = 'desc';
                    } else if (sortDirection === 'desc') {
                        sortColumn = null;
                        sortDirection = null;
                    }
                } else {
                    sortColumn = colIndex;
                    sortDirection = 'asc';
                }
                currentPage = 1;
                generateTable();
            });
        }
    });
}

function toggleEdit() {
    if (!currentSheet || !sheets[currentSheet]) {
        alert("Please load or create a table first to enable edit mode!");
        return;
    }

    if (isEditing && editingCells.length > 0) {
        promptSaveOrDiscard();
    } else {
        isEditing = !isEditing;
        $("#editBtn").text(isEditing ? "Save" : "Edit");
        if (isEditing) {
            let sheet = sheets[currentSheet];
            sheet.undoStack.push({
                tableData: JSON.parse(JSON.stringify(sheet.tableData)),
                mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
                cellColors: Object.assign({}, sheet.cellColors)
            });
            sheet.redoStack = [];
            editingCells = [];
            originalValues = {};
            generateTable();
        }
        updateEditStatus();
    }
    updateButtonStates();
}

function clearTable() {
    checkUnsavedChanges(() => {
        if (Object.keys(sheets).length > 0) {
            let savePrompt = confirm("Do you want to save the current table as an Excel file before clearing? Click OK to save, Cancel to clear without saving.");
            if (savePrompt) {
                exportToExcel();
            }
        }

        $("#dataTable thead, #dataTable tbody").empty();
        $("#sheetSearchContainer, #paginationControls").remove();
        localStorage.removeItem("sheets");
        sheets = {};
        currentSheet = null;
        $("#excelFile").val("");
        isEditing = false;
        $("#editBtn").text("Edit");
        editingCells = [];
        originalValues = {};
        sortColumn = null;
        sortDirection = null;
        currentPage = 1;
        rowsPerPage = 10;
        updateEditStatus();
        updateUIState();
        updateSheetList(); // Update sheet list
    });
}

function saveToLocalStorage() {
    localStorage.setItem("sheets", JSON.stringify(sheets));
}

function exportToExcel() {
    if (Object.keys(sheets).length === 0) {
        alert("No data to export!");
        return;
    }

    let defaultName = "UpdatedData";
    let fileName = prompt("Enter the file name for the Excel export:", defaultName);
    if (fileName === null) return;
    if (!fileName.trim()) fileName = defaultName;
    fileName = fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;

    let wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(sheetName => {
        let sheet = sheets[sheetName];
        let cleanData = sheet.tableData.map((row, index) => {
            if (index === 0) {
                return row.map(cell => cell.replace(/^\d+: /, ''));
            }
            return row.slice(0);
        });

        let ws = XLSX.utils.aoa_to_sheet(cleanData);
        ws['!merges'] = sheet.mergedCells;
        for (let key in sheet.cellColors) {
            let [r, c] = key.split(',').map(Number);
            let cellRef = XLSX.utils.encode_cell({r, c});
            if (!ws[cellRef]) ws[cellRef] = {v: cleanData[r][c]};
            ws[cellRef].s = ws[cellRef].s || {};
            ws[cellRef].s.bgColor = {rgb: sheet.cellColors[key].replace('#', '')};
        }
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, fileName);
}

function exportToPDF() {
    checkUnsavedChanges(() => {
        if (Object.keys(sheets).length === 0) {
            alert("No data to export!");
            return;
        }

        let defaultName = "ExportedData";
        let fileName = prompt("Enter the file name for the PDF export (will print with this title):", defaultName);
        if (fileName === null) return;
        if (!fileName.trim()) fileName = defaultName;

        let printWindow = window.open("", "", "width=800,height=600");
        printWindow.document.write("<html><head><title>" + fileName + "</title></head><body>");
        printWindow.document.write("<h2>" + fileName + "</h2>");

        Object.keys(sheets).forEach(sheetName => {
            let sheet = sheets[sheetName];
            printWindow.document.write(`<h3>${sheetName}</h3>`);
            let cleanTable = '<table style="width:100%; border-collapse: collapse; border: 1px solid #ddd; margin-bottom: 20px;">';
            cleanTable += '<thead><tr>';
            sheet.tableData[0].forEach(cell => {
                cleanTable += `<th style="padding: 12px; text-align: left; border: 1px solid #ddd;">${cell.replace(/^\d+: /, '')}</th>`;
            });
            cleanTable += '</tr></thead><tbody>';
            let hiddenCells = new Set();

            for (let i = 1; i < sheet.tableData.length; i++) {
                cleanTable += '<tr>';
                for (let j = 0; j < sheet.tableData[i].length; j++) {
                    let cellKey = `${i},${j}`;
                    if (hiddenCells.has(cellKey)) continue;
                    let color = sheet.cellColors[cellKey] || '';
                    let style = color ? `style="padding: 12px; text-align: left; border: 1px solid #ddd; background-color: ${color};"` : `style="padding: 12px; text-align: left; border: 1px solid #ddd;"`;
                    let mergeInfo = sheet.mergedCells.find(m => m.s.r === i && m.s.c === j);
                    if (mergeInfo) {
                        let rowspan = mergeInfo.e.r - mergeInfo.s.r + 1;
                        let colspan = mergeInfo.e.c - mergeInfo.s.c + 1;
                        cleanTable += `<td ${style} rowspan="${rowspan}" colspan="${colspan}">${sheet.tableData[i][j]}</td>`;
                        for (let r = i; r <= mergeInfo.e.r; r++) {
                            for (let c = j; c <= mergeInfo.e.c; c++) {
                                if (r !== i || c !== j) hiddenCells.add(`${r},${c}`);
                            }
                        }
                    } else {
                        cleanTable += `<td ${style}>${sheet.tableData[i][j]}</td>`;
                    }
                }
                cleanTable += '</tr>';
            }
            cleanTable += '</tbody></table>';
            printWindow.document.write(cleanTable);
        });

        printWindow.document.write("</body></html>");
        printWindow.document.close();
        printWindow.print();
    });
}

function undo() {
    checkUnsavedChanges(() => {
        if (!currentSheet || !sheets[currentSheet] || sheets[currentSheet].undoStack.length <= 1) {
            alert("No more actions to undo!");
            return;
        }
        let sheet = sheets[currentSheet];
        sheet.redoStack.push({
            tableData: JSON.parse(JSON.stringify(sheet.tableData)),
            mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
            cellColors: Object.assign({}, sheet.cellColors)
        });
        let previousState = sheet.undoStack.pop();
        sheet.tableData = previousState.tableData;
        sheet.mergedCells = previousState.mergedCells;
        sheet.cellColors = previousState.cellColors;
        saveToLocalStorage();
        generateTable();
        updateButtonStates();
    });
}

function redo() {
    checkUnsavedChanges(() => {
        if (!currentSheet || !sheets[currentSheet] || sheets[currentSheet].redoStack.length === 0) {
            alert("No more actions to redo!");
            return;
        }
        let sheet = sheets[currentSheet];
        sheet.undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheet.tableData)),
            mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
            cellColors: Object.assign({}, sheet.cellColors)
        });
        let nextState = sheet.redoStack.pop();
        sheet.tableData = nextState.tableData;
        sheet.mergedCells = nextState.mergedCells;
        sheet.cellColors = nextState.cellColors;
        saveToLocalStorage();
        generateTable();
        updateButtonStates();
    });
}

function createNewTable() {
    checkUnsavedChanges(() => {
        if (Object.keys(sheets).length > 0) {
            let confirmClear = confirm("A table already exists. Do you want to clear all existing data and create a new table? Click OK to clear and create, Cancel to keep existing data and use 'New Sheet' instead.");
            if (!confirmClear) {
                alert("Use 'New Sheet' to add another sheet without clearing existing data.");
                return;
            }
            sheets = {};
            localStorage.removeItem("sheets");
            $("#dataTable thead, #dataTable tbody").empty();
            $("#sheetSearchContainer, #paginationControls").remove();
            currentSheet = null;
        }

        let sheetName = "Sheet1";
        sheets[sheetName] = {
            tableData: [
                ["Column 1", "Column 2", "Column 3"],
                ["", "", ""],
                ["", "", ""]
            ],
            mergedCells: [],
            cellColors: {},
            undoStack: [],
            redoStack: []
        };
        sheets[sheetName].undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheets[sheetName].tableData)),
            mergedCells: [],
            cellColors: {}
        });
        currentSheet = sheetName;

        saveToLocalStorage();
        currentPage = 1;
        sortColumn = null;
        sortDirection = null;
        generateTable();
        updateSheetList();
        updateUIState();
        updateButtonStates();
    });
}

function createNewSheet() {
    if (!Object.keys(sheets).length) {
        alert("Please create or import a table first!");
        return;
    }
    checkUnsavedChanges(() => {
        let sheetName = prompt("Enter new sheet name:", "Sheet" + (Object.keys(sheets).length + 1));
        if (sheetName === null) return;
        if (!sheetName || sheets[sheetName]) {
            alert("Sheet name cannot be empty or already exists!");
            return;
        }
        sheets[sheetName] = {
            tableData: [
                ["Column 1", "Column 2", "Column 3"],
                ["", "", ""],
                ["", "", ""]
            ],
            mergedCells: [],
            cellColors: {},
            undoStack: [],
            redoStack: []
        };
        sheets[sheetName].undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheets[sheetName].tableData)),
            mergedCells: [],
            cellColors: {}
        });
        currentSheet = sheetName;
        saveToLocalStorage();
        currentPage = 1;
        sortColumn = null;
        sortDirection = null;
        generateTable();
        updateSheetList();
        updateButtonStates();
    });
}

function addColumnAt(colIndex) {
    if (!Object.keys(sheets).length) {
        alert("Please create or import a table first!");
        return;
    }
    checkUnsavedChanges(() => {
        let sheet = sheets[currentSheet];
        let choice = prompt(`Add a new column:\n1. Left of this column (${colIndex + 1}: ${sheet.tableData[0][colIndex]})\n2. Right of this column\nEnter 1 or 2:`);
        if (choice === null) return;

        let colNum;
        if (choice === "1") {
            colNum = colIndex;
        } else if (choice === "2") {
            colNum = colIndex + 1;
        } else {
            alert("Invalid choice! Please enter 1 or 2.");
            return;
        }

        sheet.undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheet.tableData)),
            mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
            cellColors: Object.assign({}, sheet.cellColors)
        });
        sheet.redoStack = [];

        sheet.tableData.forEach((row, index) => {
            if (index === 0) {
                row.splice(colNum, 0, "New Column");
            } else {
                row.splice(colNum, 0, '');
            }
        });

        sheet.mergedCells = sheet.mergedCells.map(merge => {
            if (merge.s.c >= colNum) merge.s.c++;
            if (merge.e.c >= colNum) merge.e.c++;
            return merge;
        });

        let newCellColors = {};
        for (let key in sheet.cellColors) {
            let [r, c] = key.split(',').map(Number);
            if (c >= colNum) {
                newCellColors[`${r},${c + 1}`] = sheet.cellColors[key];
            } else {
                newCellColors[key] = sheet.cellColors[key];
            }
        }
        sheet.cellColors = newCellColors;

        saveToLocalStorage();
        generateTable();
        updateButtonStates();
    });
}

function addRowAt(rowIndex) {
    if (!Object.keys(sheets).length) {
        alert("Please create or import a table first!");
        return;
    }
    checkUnsavedChanges(() => {
        let sheet = sheets[currentSheet];
        let choice = prompt(`Add a new row:\n1. Before this row (${rowIndex})\n2. After this row\nEnter 1 or 2:`);
        if (choice === null) return;

        let rowNum;
        if (choice === "1") {
            rowNum = rowIndex;
        } else if (choice === "2") {
            rowNum = rowIndex + 1;
        } else {
            alert("Invalid choice! Please enter 1 or 2.");
            return;
        }

        sheet.undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheet.tableData)),
            mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
            cellColors: Object.assign({}, sheet.cellColors)
        });
        sheet.redoStack = [];

        let columnCount = sheet.tableData[0].length;
        let newRow = Array(columnCount).fill('');
        sheet.tableData.splice(rowNum, 0, newRow);

        sheet.mergedCells = sheet.mergedCells.map(merge => {
            if (merge.s.r >= rowNum) merge.s.r++;
            if (merge.e.r >= rowNum) merge.e.r++;
            return merge;
        });

        let newCellColors = {};
        for (let key in sheet.cellColors) {
            let [r, c] = key.split(',').map(Number);
            if (r >= rowNum) {
                newCellColors[`${r + 1},${c}`] = sheet.cellColors[key];
            } else {
                newCellColors[key] = sheet.cellColors[key];
            }
        }
        sheet.cellColors = newCellColors;

        saveToLocalStorage();
        generateTable();
        updateButtonStates();
    });
}

function deleteColumnAt(colIndex) {
    if (!Object.keys(sheets).length) {
        alert("Please create or import a table first!");
        return;
    }
    checkUnsavedChanges(() => {
        let sheet = sheets[currentSheet];
        if (sheet.tableData[0].length <= 1) {
            alert("Cannot delete - table must have at least one column!");
            return;
        }
        let confirmDelete = confirm(`Are you sure you want to delete column ${colIndex + 1}: ${sheet.tableData[0][colIndex]}?`);
        if (!confirmDelete) return;

        sheet.undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheet.tableData)),
            mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
            cellColors: Object.assign({}, sheet.cellColors)
        });
        sheet.redoStack = [];

        sheet.tableData.forEach(row => row.splice(colIndex, 1));

        sheet.mergedCells = sheet.mergedCells.filter(merge => {
            if (merge.s.c === colIndex || merge.e.c === colIndex) return false;
            if (merge.s.c > colIndex) merge.s.c--;
            if (merge.e.c >= colIndex) merge.e.c--;
            return true;
        });

        let newCellColors = {};
        for (let key in sheet.cellColors) {
            let [r, c] = key.split(',').map(Number);
            if (c > colIndex) {
                newCellColors[`${r},${c - 1}`] = sheet.cellColors[key];
            } else if (c < colIndex) {
                newCellColors[key] = sheet.cellColors[key];
            }
        }
        sheet.cellColors = newCellColors;

        saveToLocalStorage();
        generateTable();
        updateButtonStates();
    });
}

function deleteRowAt(rowIndex) {
    if (!Object.keys(sheets).length) {
        alert("Please create or import a table first!");
        return;
    }
    checkUnsavedChanges(() => {
        let sheet = sheets[currentSheet];
        if (sheet.tableData.length <= 1) {
            alert("Cannot delete - table must have at least headers!");
            return;
        }
        let confirmDelete = confirm(`Are you sure you want to delete row ${rowIndex}?`);
        if (!confirmDelete) return;

        sheet.undoStack.push({
            tableData: JSON.parse(JSON.stringify(sheet.tableData)),
            mergedCells: JSON.parse(JSON.stringify(sheet.mergedCells)),
            cellColors: Object.assign({}, sheet.cellColors)
        });
        sheet.redoStack = [];

        sheet.tableData.splice(rowIndex, 1);

        sheet.mergedCells = sheet.mergedCells.filter(merge => {
            if (merge.s.r === rowIndex || merge.e.r === rowIndex) return false;
            if (merge.s.r > rowIndex) merge.s.r--;
            if (merge.e.r >= rowIndex) merge.e.r--;
            return true;
        });

        let newCellColors = {};
        for (let key in sheet.cellColors) {
            let [r, c] = key.split(',').map(Number);
            if (r > rowIndex) {
                newCellColors[`${r - 1},${c}`] = sheet.cellColors[key];
            } else if (r < rowIndex) {
                newCellColors[key] = sheet.cellColors[key];
            }
        }
        sheet.cellColors = newCellColors;

        saveToLocalStorage();
        generateTable();
        updateButtonStates();
    });
}

window.toggleEdit = toggleEdit;
window.clearTable = clearTable;
window.exportToExcel = exportToExcel;
window.exportToPDF = exportToPDF;
window.undo = undo;
window.redo = redo;
window.createNewTable = createNewTable;
window.createNewSheet = createNewSheet;
window.addColumnAt = addColumnAt;
window.addRowAt = addRowAt;
window.deleteColumnAt = deleteColumnAt;
window.deleteRowAt = deleteRowAt;
window.renameSheet = renameSheet;