function marcarAusenteNoAres() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ausenteSheet = ss.getSheetByName('PABLO2');
    const abasParaVerificar = funcionariosSol();

    const ausenteData = ausenteSheet.getRange(2, 1, ausenteSheet.getLastRow() - 1).getValues();
    const ausenteProcessos = ausenteData.map(row => row[0]);

    abasParaVerificar.forEach(abaNome => {
        const abaSheet = ss.getSheetByName(abaNome);

        if (!abaSheet || abaSheet.getLastRow() < 2) return;

        const abaData = abaSheet.getRange(2, 1, abaSheet.getLastRow() - 1, 13).getValues();

        abaData.forEach((row, index) => {
            const processo = row[0];
            if (ausenteProcessos.includes(processo)) {
                abaSheet.getRange(index + 2, 13).setValue("AUSENTE NO ARES");
            }
        });
    });
};
