function devolverInicialRealizada() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaInicial = ss.getSheetByName("INICIAL REALIZADA");
    const abaDevolucoesInicial = ss.getSheetByName("DEVOLUÇÕES - INICIAL");

    if (!abaInicial || !abaDevolucoesInicial) {
        Logger.log("Erro: Uma ou ambas as abas não foram encontradas.");
        return;
    }

    const cabecalhoInicial = abaInicial.getRange(1, 1, 1, abaInicial.getLastColumn()).getValues()[0];
    const cabecalhoDevolucao = abaDevolucoesInicial.getRange(1, 1, 1, abaDevolucoesInicial.getLastColumn()).getValues()[0];

    const indiceStatus = cabecalhoInicial.indexOf("STATUS");
    const indiceJaFoiDevolvido = cabecalhoInicial.indexOf("JA FOI DEVOLVIDO?");
    const indiceDataDevolucao = cabecalhoDevolucao.indexOf("DEVOLUÇÃO DA INICIAL");

    if (indiceStatus === -1 || indiceJaFoiDevolvido === -1 || indiceDataDevolucao === -1) {
        Logger.log("Erro: Algumas colunas necessárias não foram encontradas.");
        return;
    }

    const dadosInicial = abaInicial.getRange(2, 1, abaInicial.getLastRow() - 1, abaInicial.getLastColumn()).getValues();
    const hojeFormatado = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    dadosInicial.forEach((row, index) => {
        const status = row[indiceStatus];
        const jaFoiDevolvido = row[indiceJaFoiDevolvido];

        if (jaFoiDevolvido !== "DEVOLVIDO" && (status === "DEVOLVIDO AO RELATO")) {
            row[indiceDataDevolucao] = hojeFormatado;
            abaDevolucoesInicial.appendRow(row);

            abaInicial.getRange(index + 2, indiceJaFoiDevolvido + 1).setValue("DEVOLVIDO");
        }
    });
};