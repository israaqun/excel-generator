let hot,gridGenerated=!1;function createGrid(t,n){var o=[];for(let e=0;e<t;e++){var r=Array(n).fill("");o.push(r)}var e=document.getElementById("grid-container");hot=new Handsontable(e,{licenseKey:"non-commercial-and-evaluation",data:o,colHeaders:!0,rowHeaders:!0,contextMenu:!0,minSpareCols:1,minSpareRows:1,columns:Array(n).fill({editor:"text"})})}function clearGrid(){hot&&hot.destroy(),gridGenerated=!1,document.getElementById("num-columns").value="",document.getElementById("num-rows").value=""}function enableGenerateButton(){document.getElementById("generate-button").removeAttribute("disabled")}function enableExportButton(){document.getElementById("export-button").removeAttribute("disabled")}function generateGrid(){var e=document.getElementById("num-columns").value,t=document.getElementById("num-rows").value;""===e.trim()||""===t.trim()||(e=parseInt(e),t=parseInt(t),isNaN(e))||isNaN(t)||e<=0||t<=0?alert("Please enter valid values for rows and columns."):(hot&&clearGrid(),createGrid(t,e),gridGenerated=!0,enableExportButton())}function exportToExcel(){var e,t;gridGenerated?hot&&("undefined"!=typeof XLSX?(e=hot.getData(),e=XLSX.utils.aoa_to_sheet(e),t=XLSX.utils.book_new(),XLSX.utils.book_append_sheet(t,e,"Sheet1"),XLSX.writeFile(t,"exported_data.xlsx")):alert("XLSX library is not loaded. Please check the library source.")):alert("Please generate grid first.")}const generateButton=document.getElementById("generate-button"),exportButton=(generateButton.addEventListener("click",generateGrid),document.getElementById("export-button"));exportButton.addEventListener("click",exportToExcel),window.onload=function(){enableGenerateButton()};