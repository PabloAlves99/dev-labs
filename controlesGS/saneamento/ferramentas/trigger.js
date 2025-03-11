function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const allowedSheets = funcionariosSol();
    const range = e.range;
    const sheetName = sheet.getName();


    if (allowedSheets.includes(sheet.getName()) && range.getColumn() === 13 && range.getRow() > 1) {

        const currentDate = new Date();
        const newValue = range.getValue();

        if (newValue == "DEVOLVIDO PELO ADVOGADO") {
            return;

        } else if (newValue) {
            sheet.getRange(range.getRow(), 18)
                .setValue(currentDate);

        } else {
            sheet.getRange(range.getRow(), 18)
                .setValue("");

        }
    }
}