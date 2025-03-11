function verificarErrosPreenchimento() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var funcionarios = listaDeFuncionarios();
    let compl = "BOAS-VINDAS ";

    funcionarios.forEach((funcionario) => {
        var tabBoasVindas = ss.getSheetByName(compl + funcionario);

        if (tabBoasVindas) {
            var lastRow = tabBoasVindas.getLastRow();
            var valores = tabBoasVindas.getRange(2, 1, lastRow - 1, 13).getValues();

            var cabecalho = tabBoasVindas.getRange(1, 1, 1, tabBoasVindas.getLastColumn()).getValues()[0];
            var indiceNViavelRelato = obterIndiceCabecalho(cabecalho, "MOTIVO NÃO VIÁVEL RELATO");
            var indiceStatus = obterIndiceCabecalho(cabecalho, "STATUS");
            var indiceViavel = obterIndiceCabecalho(cabecalho, "VIÁVEL?");


            for (var i = 0; i < valores.length; i++) {
                var motivoNaoViavelRelato = valores[i][indiceNViavelRelato];
                var status = valores[i][indiceStatus];
                var viavel = valores[i][indiceViavel];

                if (motivoNaoViavelRelato && status) {
                    tabBoasVindas.getRange(i + 2, indiceStatus + 1).setValue("");
                };

                if (status == "DOCUMENTAÇÃO" && !Boolean(viavel)) {
                    tabBoasVindas.getRange(i + 2, indiceViavel + 1).setValue(Boolean(true));
                }
            };

        } else {
            Logger.log("A aba " + compl + funcionario + " não foi encontrada.");
        }
    });
};

function obterIndiceCabecalho(allCabecalho, nomeColuna) {
    return allCabecalho.indexOf(nomeColuna);

};
