const green = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
const red = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
const orange = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
const labelRgb = "FFD9D9D9";

function isLabeled(cell) {
    const fill = cell.fill;
    return fill &&
         fill.type === 'pattern' &&
         fill.pattern === 'solid' &&
         fill.fgColor &&
         fill.fgColor.argb === labelRgb;
}

async function compare() {
    const masterFile = fileSelector.masterFile;
    const testFiles = fileSelector.testFiles;
    const zipOutput = document.getElementById('zipOutput').checked;

    if (!masterFile || testFiles.length === 0) {
        alert("Seleziona il file master e almeno un file di test!");
        return;
    }

    const wbMaster = new ExcelJS.Workbook();
    const wbTest = new ExcelJS.Workbook();

    const zip = new JSZip();

    for (const testFile of testFiles) {
        await wbMaster.xlsx.load(await masterFile.arrayBuffer());
        await wbTest.xlsx.load(await testFile.arrayBuffer());

        const wsMaster = wbMaster.worksheets[0];
        const wsTest = wbTest.worksheets[0];

        const wbOut = new ExcelJS.Workbook();
        const wsOut = wbOut.addWorksheet("Confronto");

        wsMaster.eachRow((row, rowIndex) => {
            const outRow = wsOut.getRow(rowIndex);
            const testRow = wsTest.getRow(rowIndex);
            row.eachCell((cell, colIndex) => {
                const cMaster = cell;
                const cTest = testRow.getCell(colIndex);
                const cOut = outRow.getCell(colIndex);
                cOut.value = cTest.value;

                if (!isLabeled(cMaster)) return;

                const v1 = cMaster.value.result;
                const v2 = cTest.value.result;
                const f1 = cMaster.formula;
                const f2 = cTest.formula;

                if (v1 === v2) {
                    cOut.fill = (f1 === f2) ? green : orange;
                } else {
                    cOut.fill = red;
                }
            });
            outRow.commit();
        });

        const buffer = await wbOut.xlsx.writeBuffer();
        const outputName = `out_${testFile.name}`;
        if (zipOutput) {
            zip.file(outputName, buffer);
        } else {
            saveAs(new Blob([buffer]), outputName);
        }
    }

    if (zipOutput) {
        const zipBlob = await zip.generateAsync({ type: "blob" });
        saveAs(zipBlob, "results.zip");
    }

}
