function compararEstoqueCopia() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaEstoque = ss.getSheetByName("estoque"); // Aba de Estoque com os clientes válidos
    let abaCopiaInicial = ss.getSheetByName("INICIAL"); // Aba que precisa ser comparada

    if (!abaEstoque || !abaCopiaInicial) {
        Logger.log("❌ Verifique se as abas 'Estoque' e 'Cópia de INICIAL' existem.");
        return;
    };

    let dadosEstoque = abaEstoque.getDataRange().getValues();
    let dadosCopia = abaCopiaInicial.getDataRange().getValues();

    let clientesCopia = new Set();
    let naoEncontrados = [];


    for (let i = 1; i < dadosCopia.length; i++) {
        let clienteCopia = limparTexto(dadosCopia[i][10]);
        if (clienteCopia) {
            clientesCopia.add(clienteCopia);
        };
    };

    for (let i = 1; i < dadosEstoque.length; i++) {
        let clienteEstoque = limparTexto(dadosEstoque[i][0]);
        if (clienteEstoque && !clientesCopia.has(clienteEstoque)) {
            naoEncontrados.push(clienteEstoque);
        };
    };

    if (naoEncontrados.length > 0) {
        Logger.log(`❌ Clientes encontrados em Estoque, mas não em Cópia de INICIAL:`);
        naoEncontrados.forEach(cliente => Logger.log(cliente));
    } else {
        Logger.log("✅ Todos os clientes de Estoque estão presentes em Cópia de INICIAL.");
    };
};

function limparTexto(texto) {
    return texto ? texto.toString().trim().replace(/\s+/g, ' ').toLowerCase() : "";
};
