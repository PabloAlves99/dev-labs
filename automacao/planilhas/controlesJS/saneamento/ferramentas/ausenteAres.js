function ausentesAres() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testeSheet = ss.getSheetByName('xx');
    const abaSheet = ss.getSheetByName('yy');

    const testeData = testeSheet.getRange(2, 1, testeSheet.getLastRow() - 1).getValues();
    const abaData = abaSheet.getRange(2, 1, abaSheet.getLastRow() - 1, 12).getValues();

    const testeList = testeData.map(row => row[0].trim());
    const valoresIgnorados = [
        "LIBERADO (OK)",
        "DESISTÊNCIA (COM DECLARAÇÃO)",
        "DESISTÊNCIA (SEM DECLARAÇÃO)",
        "LIBERADO (DOCS. MÍNIMOS)",
        "LIBERADO (DOCS. PARCIAIS)",
        "DEVOLVIDO POR OMISSÃO",
        "DEVOLVIDO SEM CONTATO",
        "TRANSFERIDO",
        "CONCLUÍDO",
    ];


    let initRow = 2;

    abaData.forEach((col, index) => {
        const colunaA = col[0]?.trim();
        const colunaH = col[7];
        const colunaL = col[11];

        if (!testeList.includes(colunaA) && !valoresIgnorados.includes(colunaL) && colunaL !== "AUSENTE NO ARES") {
            testeSheet.getRange(initRow, 10).setValue(colunaA);
            testeSheet.getRange(initRow, 11).setValue(colunaH);
            initRow++;
        };
    });
};
