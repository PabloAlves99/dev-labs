function finalizarAdvogados() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaEstoque = ss.getSheetByName("ESTOQUE - CLIENTES FINALIZADOS");
    const abaFinalizadosCorreto = ss.getSheetByName("CLIENTES FINALIZADOS CORRETO");
    const abaDevolvidos = ss.getSheetByName("CLIENTES DEVOLVIDOS PELO ADV")

    const dadosEstoque = abaEstoque.getDataRange().getValues();
    const statusManifestacao = [
        "MANIFESTAÇÃO REALIZADA COM DOCS. PARCIAIS",
        "MANIFESTAÇÃO REALIZADA SEM DOCUMENTOS",
        "MANIFESTAÇÃO REALIZADA COM DOCS. OK",
        "DESISTÊNCIA SOLICITADA"
    ];
    const statusDilacaoDevolvido = [
        "1ª DILAÇÃO SOLICITADA",
        "2ª DILAÇÃO SOLICITADA",
        "3ª DILAÇÃO SOLICITADA",
        "DEVOLVIDO PARA O AGENTE",
        "MANIFESTAÇÃO REALIZADA COM DOCS. PARCIAIS + DILAÇÃO"
    ];

    const agenteResponsavel = {
        'xxx': 'AGENTE1',
        'yyy': 'AGENTE2',
    };


    let linhasParaExcluir = [];

    const indiceProcedimentoAdvogado = obterIndiceCabecalho(dadosEstoque[0], "PROCEDIMENTO DO AVOGADO");
    const indiceAgenteResponsavel = obterIndiceCabecalho(dadosEstoque[0], "AGENTE RESPONSÁVEL");
    const indiceStatus = obterIndiceCabecalho(dadosEstoque[0], "STATUS");

    dadosEstoque.forEach((linha, i) => {
        let procedimentoAdv = linha[indiceProcedimentoAdvogado].trim().toUpperCase();
        let agente = linha[indiceAgenteResponsavel]?.trim().toUpperCase();

        if (!procedimentoAdv) return;

        if (statusManifestacao.includes(procedimentoAdv)) {
            abaFinalizadosCorreto.appendRow(linha);
            linhasParaExcluir.push(i + 1);

        } else if (statusDilacaoDevolvido.includes(procedimentoAdv)) {
            abaDevolvidos.appendRow(linha);
            linhasParaExcluir.push(i + 1);

            if (agente && agenteResponsavel[agente]) {
                let linhaAgente = linha.slice();
                linhaAgente[indiceStatus] = "DEVOLVIDO PELO ADVOGADO";

                const abaAgente = ss.getSheetByName(agenteResponsavel[agente]);
                if (abaAgente) {
                    abaAgente.appendRow(linhaAgente);
                }
            }
        };
    });

    for (let i = linhasParaExcluir.length - 1; i >= 0; i--) {
        abaEstoque.deleteRow(linhasParaExcluir[i]);
    }

    function obterIndiceCabecalho(allCabecalho, nomeColuna) {
        return allCabecalho.indexOf(nomeColuna);
    };
};
