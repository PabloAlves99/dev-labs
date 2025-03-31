// Script antigo e não atualizado, deve ser refatorado para melhorar performance e legibilidade.
// Entretanto supre a necessidade de atualizar pequenas quantidades de dados.

function atualizarInformacao() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ABA_TESTE = ss.getSheetByName("xx");
    var ABA_ADV = ss.getSheetByName("yy");

    var dadosTeste = ABA_TESTE.getRange("A:E").getValues();
    var dadosAdv = ABA_ADV.getRange("A:E").getValues();
    var mapAdv = {};

    for (var i = 0; i < dadosAdv.length; i++) {
        mapAdv[dadosAdv[i][0]] = i;
    }

    for (var j = 0; j < dadosTeste.length; j++) {
        var chaveTeste = dadosTeste[j][0];
        var valorTesteE = dadosTeste[j][4];

        if (mapAdv.hasOwnProperty(chaveTeste)) {
            var linhaAdv = mapAdv[chaveTeste];
            var valorAdvE = dadosAdv[linhaAdv][4];

            if (valorTesteE !== valorAdvE) {
                ABA_ADV.getRange(linhaAdv + 1, 5).setValue(valorTesteE);
            }
        }
    }
    atualizarDados(ABA_ADV, ABA_TESTE)
}

function atualizarDados(ABA_ADV, ABA_TESTE) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();


    const rangeTeste = ABA_TESTE.getRange("A2:D" + ABA_TESTE.getLastRow()).getValues();


    const mapTeste = {};
    rangeTeste.forEach(row => {
        const key = row[0];
        const date = new Date(row[3]);
        if (!mapTeste[key] || date < mapTeste[key]) {
            mapTeste[key] = date;
        }
    });

    const rangeAdv = ABA_ADV.getRange("A2:D" + ABA_ADV.getLastRow()).getValues();

    let newDates = [];
    rangeAdv.forEach((row, index) => {
        let keyAdv = row[0];
        let originalDate = row[3];
        if (mapTeste[keyAdv]) {
            newDates.push([mapTeste[keyAdv]]);
        } else {
            newDates.push([originalDate]);
        }
    });

    if (newDates.length > 0) {
        ABA_ADV.getRange(2, 4, newDates.length, 1).setValues(newDates);
    }
}
