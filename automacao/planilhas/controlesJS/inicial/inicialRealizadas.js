function inicialRealizada() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var origem = ss.getSheetByName("INICIAL");
    var destino = ss.getSheetByName("INICIAL REALIZADA");

    if (!origem || !destino) {
        Logger.log("Uma das abas não foi encontrada.");
        return;
    };

    var dados = origem.getDataRange().getValues();
    if (dados.length < 2) {
        Logger.log("Nenhum dado encontrado além do cabeçalho.");
        return;
    };

    var cabecalho = dados[0]; // Captura o cabeçalho para manter a estrutura
    var statusValidos = [
        "INICIAL REALIZADA",
        "INICIAL REALIZADA COM PROBLEMA SANADO",
        "INICIAL REDISTRIBUIDA"
    ];

    var colunaStatus = cabecalho.indexOf("STATUS");
    if (colunaStatus == -1) {
        Logger.log("Coluna 'STATUS' não encontrada.");
        return;
    };

    var linhasParaCopiar = [];
    var linhasParaRemover = [];
    var hojeFormatado = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    for (var i = 1; i < dados.length; i++) {
        if (statusValidos.includes(dados[i][colunaStatus])) {
            dados[i].push(hojeFormatado);
            linhasParaCopiar.push(dados[i]);
            linhasParaRemover.push(i + 1); // Guardar índice real da planilha (começa em 1)
        }
    };

    // Se houver linhas para copiar, adiciona na aba "INICIAL REALIZADA"
    if (linhasParaCopiar.length > 0) {
        var ultimaLinha = destino.getLastRow();
        destino.getRange(ultimaLinha + 1, 1, linhasParaCopiar.length, linhasParaCopiar[0].length).setValues(linhasParaCopiar);
    } else {
        Logger.log("Nenhum cliente com status válido encontrado.");
    };

    // Removendo as linhas de baixo para cima para evitar problemas com índices
    if (linhasParaRemover.length > 0) {
        linhasParaRemover.reverse().forEach(linha => {
            origem.deleteRow(linha);
        });
    };
};
