function conversao() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var funcionarios = listaDeFuncionarios();
    var sheetRelatoPrescricao = ss.getSheetByName("RELATO - PRESCRIÇÃO");
    var sheetRelatoProcon = ss.getSheetByName("RELATO - PROCON");
    let compl = "CNV - ";

    funcionarios.forEach(funcionario => {
        var sheet_agente = ss.getSheetByName(compl + funcionario);
        if (!sheet_agente) return;

        var range = sheet_agente.getRange("A2:T" + sheet_agente.getLastRow());
        var data = range.getValues();

        data.forEach(function (row, index) {
            if (row[3] !== "" && row[16] === "VERDADEIRO" && row[19] === "") {
                var rowToCopy = row.slice(0, 19);
                var today = new Date();
                var formattedDate = Utilities.formatDate(today, "GMT-0300", "yyyy-MM-dd");

                var sheetRelatoTarget;
                if (row[6] === "PRESCRIÇÃO") {
                    sheetRelatoTarget = sheetRelatoPrescricao;
                } else {
                    sheetRelatoTarget = sheetRelatoProcon;
                }

                var nextRow = sheetRelatoTarget.getLastRow() + 1;

                sheetRelatoTarget.getRange("C" + nextRow).setValue(funcionario);

                sheetRelatoTarget.getRange("D" + nextRow + ":V" + nextRow).setValues([rowToCopy]);
                sheetRelatoTarget.getRange("U" + nextRow).setValue(formattedDate);

                sheet_agente.getRange("T" + (index + 2)).setValue("VERDADEIRO");
            }
        });
    });
}
