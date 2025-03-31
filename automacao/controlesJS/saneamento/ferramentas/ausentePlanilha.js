function ausentePlanilha() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testeSheet = ss.getSheetByName('xx');
    const abaSheet = ss.getSheetByName('yy');

    const testeData = testeSheet.getRange(2, 1, testeSheet.getLastRow() - 1).getValues();
    const abaData = abaSheet.getRange(2, 1, abaSheet.getLastRow() - 1).getValues();

    const abaList = abaData.map(row => row[0]);

    testeData.forEach((row, index) => {
        if (!abaList.includes(row[0])) {
            testeSheet.getRange(index + 2, 10).setValue('AUSENTE');
        }
    });
}
