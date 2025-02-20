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

    const fileReader = new FileReader();
    fileReader.onload = async function() {
        const pdfData = new Uint8Array(this.result);
        const pdf = await window.pdfjsLib.getDocument({ data: pdfData }).promise;
        const datos = await extraerDatos(pdf);

        crearExcel(datos);

        const link = document.getElementById("downloadLink");
        link.setAttribute("href", 'output.xlsx');
        link.setAttribute("download", 'output.xlsx');
        link.style.display = 'block';
        link.innerText = 'Descargar Excel';
    };
    fileReader.readAsArrayBuffer(fileInput);
});

async function extraerDatos(pdf) {
    const datos = [];
    for (let i = 0; i < pdf.numPages; i++) {
        const page = await pdf.getPage(i + 1);
        const textContent = await page.getTextContent();
        const lineas = textContent.items.map(item => item.str).join('\n').split('\n');
        let numeroAlbaran = "";
        for (let i = 0; i < lineas.length; i++) {
            if (lineas[i].includes("Decl. goods it Nr.")) {
                const partes = lineas[i + 1]?.trim().split(/\s+/) || [];
                if (partes.length >= 6) {
                    const numeroPartida = partes[0];
                    const numeroBultos = partes[4];
                    const descripcion = (partes[5].length === 10 && /^\d+$/.test(partes[5]))
                        ? partes.slice(6).join(' ')
                        : partes.slice(5).join(' ');

                    for (let j = 0; j < lineas.length; j++) {
                        if (lineas[j].includes("UCR [12 08] Gross mass [18 04]")) {
                            const partesPeso = lineas[j + 1]?.trim().split(/\s+/) || [];
                            if (partesPeso.length >= 2) {
                                numeroAlbaran = partesPeso[0];
                                const peso = parseFloat(partesPeso[1]);
                                if (peso > 0) {
                                    datos.push([numeroPartida, parseInt(numeroBultos), descripcion, numeroAlbaran, peso]);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return datos;
}

function crearExcel(datos) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
        ["Número de partida", "Número de bultos", "Descripción de la mercancía", "Número de albarán", "Peso"],
        ...datos
    ]);

    const estiloBultos = { fill: { fgColor: { rgb: "FFD700" } } };
    const estiloAlternate1 = { fill: { fgColor: { rgb: "ADD8E6" } } };
    const estiloAlternate2 = { fill: { fgColor: { rgb: "D3D3D3" } } };

    datos.forEach((fila, index) => {
        const fillStyle = fila[1] > 1 ? estiloBultos : (index % 2 === 0 ? estiloAlternate1 : estiloAlternate2);
        for (let i = 0; i < 5; i++) {
            const cellAddress = XLSX.utils.encode_cell({ r: index + 1, c: i });
            if (!ws[cellAddress].s) ws[cellAddress].s = {};
            ws[cellAddress].s.fill = fillStyle;
        }
    });

    const maxLengths = ws['!ref'].split(':').map(cell => cell.match(/[A-Z]+/)[0]);
    maxLengths.forEach((col, i) => {
        let maxLength = 0;
        for (let row = 1; row <= datos.length + 1; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: i });
            maxLength = Math.max(maxLength, (ws[cellAddress]?.v || '').toString().length);
        }
        ws['!cols'][i] = { wch: maxLength + 2 };
    });

    for (let row = 2; row <= datos.length + 1; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 });
        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        ws[cellAddress].s.alignment = { horizontal: "center" };
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Datos');
    XLSX.writeFile
