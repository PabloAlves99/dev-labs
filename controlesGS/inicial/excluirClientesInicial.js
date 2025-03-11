function excluirClientesNaoEncontrados() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaEstoque = ss.getSheetByName("estoque");
    let abaCopiaInicial = ss.getSheetByName("INICIAL");

    if (!abaEstoque || !abaCopiaInicial) {
        Logger.log("❌ Verifique se as abas 'Estoque' e 'Cópia de INICIAL' existem.");
        return;
    }

    let dadosEstoque = abaEstoque.getDataRange().getValues();
    let dadosCopia = abaCopiaInicial.getDataRange().getValues();

    let clientesEstoque = new Set();


    for (let i = 1; i < dadosEstoque.length; i++) {
        let clienteEstoque = limparTexto(dadosEstoque[i][0]);
        if (clienteEstoque) {
            clientesEstoque.add(clienteEstoque);
        }
    }

    let linhasParaExcluir = [];
    let mantidas = 0;
    let removidas = 0;

    for (let i = 1; i < dadosCopia.length; i++) {
        let clienteCopia = limparTexto(dadosCopia[i][10]);
        if (!clientesEstoque.has(clienteCopia)) {
            linhasParaExcluir.push(i + 1);
            removidas++;
        } else {
            mantidas++;
        }
    }

    for (let i = linhasParaExcluir.length - 1; i >= 0; i--) {
        abaCopiaInicial.deleteRow(linhasParaExcluir[i]);
    }

    Logger.log(`✅ Filtragem concluída! ${mantidas} linhas mantidas.`);
    Logger.log(`❌ ${removidas} linhas removidas (clientes não encontrados no Estoque).`);
}

function limparTexto(texto) {
    return texto ? texto.toString().trim().replace(/\s+/g, ' ').toLowerCase() : "";
}
