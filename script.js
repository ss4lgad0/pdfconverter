document.getElementById("fileInput").addEventListener("change", function(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById("convertBtn").disabled = false;
    }
});

document.getElementById("convertBtn").addEventListener("click", async function() {
    const fileInput = document.getElementById("fileInput").files[0];

    if (!fileInput) {
        alert("Por favor, selecciona un archivo PDF.");
        return;
    }

    // Usar pdf.js cargado globalmente en index.html
    const fileReader = new FileReader();

    fileReader.onload = async function() {
        const pdfData = new Uint8Array(this.result);
        const pdf = await window.pdfjsLib.getDocument({ data: pdfData }).promise;

        let extractedData = [];

        for (let i = 0; i < pdf.numPages; i++) {
            const page = await pdf.getPage(i + 1);
            const textContent = await page.getTextContent();
            const textItems = textContent.items.map(item => item.str).join("\n");

            extractedData.push(textItems);
        }

        // Convertir a CSV para descargar como Excel
        const csvContent = "data:text/csv;charset=utf-8," + extractedData.join("\n");
        const encodedUri = encodeURI(csvContent);
        const link = document.getElementById("downloadLink");

        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "output.csv");
        link.style.display = "block";
        link.innerText = "Descargar Excel";
    };

    fileReader.readAsArrayBuffer(fileInput);
});
