
async function exportProjectFolder() {
    const zip = new JSZip();
    const projectFolder = zip.folder("SwipewireProject");

    const htmlContent = document.documentElement.outerHTML;
    projectFolder.file("index.html", htmlContent);

    try {
        const cssResponse = await fetch('/css/style.css');
        if (cssResponse.ok) {
            const cssContent = await cssResponse.text();
            projectFolder.folder("css").file("style.css", cssContent);
        } else {
            console.warn("Failed to fetch /css/style.css:", cssResponse.status);
        }
    } catch (error) {
        console.error("Error fetching CSS:", error);
    }

    try {
        const jsResponse = await fetch('/script/script.js');
        if (jsResponse.ok) {
            const jsContent = await jsResponse.text();
            projectFolder.folder("script").file("script.js", jsContent);
        } else {
            console.warn("Failed to fetch /script/script.js:", jsResponse.status);
        }
    } catch (error) {
        console.error("Error fetching JS:", error);
    }

    const readmeContent = `
Swipewire Project
================

This project includes:
- index.html: Main HTML file
- css/style.css: Stylesheet
- script/script.js: JavaScript logic

External dependencies (loaded via CDN in index.html):
- jQuery: https://code.jquery.com/jquery-3.6.0.min.js
- DataTables: https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js
- XLSX: https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
- JSZip: https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
- DataTables CSS: https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css

To run locally:
1. Extract this ZIP
2. Open index.html in a browser with internet access
    `;
    projectFolder.file("README.txt", readmeContent);

    zip.generateAsync({ type: "blob" }).then(function (content) {
        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        a.download = "SwipewireProject.zip";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });

    window.history.pushState({}, document.title, '/');
}

window.exportProjectFolder = exportProjectFolder;