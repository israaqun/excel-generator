let hot; //Handsontable
let gridGenerated = false;

function createGrid(numRows, numColumns) {
    const data = [];
    for (let i = 0; i < numRows; i++) {
        const row = Array(numColumns).fill("");
        data.push(row);
    }

    const container = document.getElementById("grid-container");
    hot = new Handsontable(container, {
        licenseKey: 'non-commercial-and-evaluation',
        data: data,
        colHeaders: true,
        rowHeaders: true,
        contextMenu: true,
        minSpareCols: 1,
        minSpareRows: 1,
        columns: Array(numColumns).fill({ editor: "text" }),
    });
}

function clearGrid() {
    if (hot) {
        hot.destroy();
    }
    gridGenerated = false;
    document.getElementById("num-columns").value = "";
    document.getElementById("num-rows").value = "";
}

function enableGenerateButton() {
    document.getElementById("generate-button").removeAttribute("disabled");
}

function enableExportButton() {
    document.getElementById("export-button").removeAttribute("disabled");
}

function generateGrid() {
    const numColumnsInput = document.getElementById("num-columns").value;
    const numRowsInput = document.getElementById("num-rows").value;

    if (numColumnsInput.trim() === '' || numRowsInput.trim() === '') {
        alert('Please enter valid values for rows and columns.');
        return;
    }

    const numColumns = parseInt(numColumnsInput);
    const numRows = parseInt(numRowsInput);

    if (isNaN(numColumns) || isNaN(numRows) || numColumns <= 0 || numRows <= 0) {
        alert('Please enter valid values for rows and columns.');
        return;
    }

    if (hot) {
        clearGrid();
    }

    createGrid(numRows, numColumns);
    gridGenerated = true;
    enableExportButton();
}

function exportToExcel() {
    if (!gridGenerated) {
        alert('Please generate the grid first.');
        return;
    }

    if (hot) {
        if (typeof XLSX !== 'undefined') {
            const data = hot.getData();
            const ws = XLSX.utils.aoa_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            XLSX.writeFile(wb, 'exported_data.xlsx');
        } else {
            alert('XLSX library is not loaded. Please check the library source.');
        }
    }
}

const generateButton = document.getElementById("generate-button");
generateButton.addEventListener("click", generateGrid);

const exportButton = document.getElementById("export-button");
exportButton.addEventListener("click", exportToExcel);

window.onload = function () {
    enableGenerateButton();
};