function controleRelato() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaControle = ss.getSheetByName("ACOMPANHAMENTO - RELATO (QUADRO)");
    const abaRProcon = ss.getSheetByName("RELATO - PROCON");
    const abaRPrescricao = ss.getSheetByName("RELATO - PRESCRIÇÃO");
    const abaValidacao = ss.getSheetByName("VALIDAÇÃO");
    const abaArqRelato = ss.getSheetByName("CLIENTES ARQUIVADOS - RELATO");
    const abaDevRelato = ss.getSheetByName("DEVOLUÇÕES - RELATO");
    const funcionarios = ["aa", "bb", "cc", "ss", "ee"];
    let infoRelato = {};
    let todosDados = {};

    obterTodosDados();
    processarRelato();
    processarValidacao();
    processarArquivados();
    processarDevolucao();
    escreverDadosRelato();

    function obterTodosDados() {
        if (abaRProcon) todosDados["RELATO - PROCON"] = abaRProcon.getDataRange().getValues();
        if (abaValidacao) todosDados["VALIDAÇÃO"] = abaValidacao.getDataRange().getValues();
        if (abaRPrescricao) todosDados["RELATO - PRESCRIÇÃO"] = abaRPrescricao.getDataRange().getValues();
        if (abaArqRelato) todosDados["CLIENTES ARQUIVADOS - RELATO"] = abaArqRelato.getDataRange().getValues();
        if (abaDevRelato) todosDados["DEVOLUÇÕES - RELATO"] = abaDevRelato.getDataRange().getValues();
    }

    function processarRelato() {
        let abasRelato = [todosDados["RELATO - PROCON"], todosDados["RELATO - PRESCRIÇÃO"]];

        abasRelato.forEach(aba => {
            let dados = aba;
            let cabecalho = dados[0];
            let indResponsavelRelato = cabecalho.indexOf("RESPONSAVEL RELATO");
            let indStatus = cabecalho.indexOf("STATUS");
            let indDataAtribuicao = cabecalho.indexOf("DATA DA ATRIBUIÇÃO");
            let colunas = ["AC", "AE", "AG", "AI", "AK", "AM", "AO", "AQ", "AS", "AU"];
            let indicesColunas = colunas.map(col => (col.charCodeAt(0) - 65) * 26 + col.charCodeAt(1) - 65);

            if (!dados || dados.length === 0) {
                Logger.log(`Nenhum dado encontrado`);
                return;
            };

            for (let i = 1; i < dados.length; i++) {
                if (!dados[i][indResponsavelRelato]) continue;
                let responsavelRelato = dados[i][indResponsavelRelato].trim().toUpperCase();
                let status = dados[i][indStatus].trim();
                let dataAtribuicao = new Date(dados[i][indDataAtribuicao]);
                let dataFormatada = formatarData(dataAtribuicao);

                if (!responsavelRelato) continue;
                if (!infoRelato[responsavelRelato]) infoRelato[responsavelRelato] = {};
                if (!infoRelato[responsavelRelato][dataFormatada]) infoRelato[responsavelRelato][dataFormatada] = {};



                // LÓGICAS DE CONTROLE

                indicesColunas.forEach((indice, j) => {
                    let valor = dados[i][indice];
                    if (valor) infoRelato[responsavelRelato][dataFormatada][valor] = (infoRelato[responsavelRelato][dataFormatada][valor] || 0) + 1;
                });


                if (status) infoRelato[responsavelRelato][dataFormatada][status] = (infoRelato[responsavelRelato][dataFormatada][status] || 0) + 1;
            };
        });

    };

    function processarValidacao() {
        let abaValidacao = todosDados["VALIDAÇÃO"];
        let cabecalhoValidacao = abaValidacao[0];
        let indResponsavelRelatoVali = cabecalhoValidacao.indexOf("RESPONSAVEL RELATO");
        let indDataLiberacao = cabecalhoValidacao.indexOf("DATA DE LIBERAÇÃO RELATO");

        for (let i = 1; i < abaValidacao.length; i++) {
            let linhaValidacao = abaValidacao[i];
            let responsavelRelatoVali = linhaValidacao[indResponsavelRelatoVali].trim().toUpperCase();
            let dataLiberacao = new Date(linhaValidacao[indDataLiberacao]);
            let dataFormatada = formatarData(dataLiberacao);

            if (!responsavelRelatoVali) continue;
            if (!infoRelato[responsavelRelatoVali]) infoRelato[responsavelRelatoVali] = {};
            if (!infoRelato[responsavelRelatoVali][dataFormatada]) infoRelato[responsavelRelatoVali][dataFormatada] = {};

            infoRelato[responsavelRelatoVali][dataFormatada]["RELATO COLHIDO"] = (infoRelato[responsavelRelatoVali][dataFormatada]["RELATO COLHIDO"] || 0) + 1;
        };
    };

    function processarArquivados() {
        let abaArqRelato = todosDados["CLIENTES ARQUIVADOS - RELATO"]
        let cabecalhoArq = abaArqRelato[0];
        let indResponsavelRelatoArq = cabecalhoArq.indexOf("RESPONSAVEL RELATO");
        let indStatusArq = cabecalhoArq.indexOf("STATUS");
        let indDataArquivamentoArq = cabecalhoArq.indexOf("DATA ARQV");

        for (let i = 1; i < abaArqRelato.length; i++) {
            let linha = abaArqRelato[i];
            let responsavelRelatoArq = linha[indResponsavelRelatoArq].trim().toUpperCase();
            let status = linha[indStatusArq].trim();
            let dataArquivamento = new Date(linha[indDataArquivamentoArq]);
            let dataFormatada = formatarData(dataArquivamento);

            if (!infoRelato[responsavelRelatoArq]) infoRelato[responsavelRelatoArq] = {};
            if (!infoRelato[responsavelRelatoArq][dataFormatada]) infoRelato[responsavelRelatoArq][dataFormatada] = {};
            infoRelato[responsavelRelatoArq][dataFormatada][status] = (infoRelato[responsavelRelatoArq][dataFormatada][status] || 0) + 1;
        };
    };

    function processarDevolucao() {
        let abaDevRelato = todosDados["DEVOLUÇÕES - RELATO"];
        let cabecalhoArq = abaDevRelato[0];
        let indResponsavelRelatoDev = cabecalhoArq.indexOf("RESPONSAVEL RELATO");
        let indDataDevolucao = cabecalhoArq.indexOf("DATA DA DEVOLUÇÃO");

        for (let i = 1; i < abaDevRelato.length; i++) {
            let linhaDev = abaDevRelato[i];
            let responsavelRelatoDev = linhaDev[indResponsavelRelatoDev].trim().toUpperCase();
            let dataDevolucao = new Date(linhaDev[indDataDevolucao]);
            let dataFormatada = formatarData(dataDevolucao);

            if (!infoRelato[responsavelRelatoDev]) infoRelato[responsavelRelatoDev] = {};
            if (!infoRelato[responsavelRelatoDev][dataFormatada]) infoRelato[responsavelRelatoDev][dataFormatada] = {};
            infoRelato[responsavelRelatoDev][dataFormatada]["DEVOLVIDO"] = (infoRelato[responsavelRelatoDev][dataFormatada]["DEVOLVIDO"] || 0) + 1;
        }
    }


    function todosOsDiasNoMes() {
        var hoje = new Date();
        var primeiroDiaDoMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
        var ultimoDiaDoMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0); // Último dia do mês

        var todosOsDias = [];

        // Itera por todos os dias do mês
        for (var d = primeiroDiaDoMes; d <= ultimoDiaDoMes; d.setDate(d.getDate() + 1)) {
            todosOsDias.push(new Date(d)); // Adiciona a data à lista de todos os dias
        }

        return todosOsDias;
    };

    function todosOsDiasNoMesPassado() {
        var hoje = new Date();
        var mesAnterior = hoje.getMonth() - 1;
        var ano = hoje.getFullYear();

        // Se o mês anterior for dezembro (mês 11), ajusta o ano
        if (mesAnterior < 0) {
            mesAnterior = 11;
            ano -= 1;
        }

        var primeiroDiaDoMesPassado = new Date(ano, mesAnterior, 1);
        var ultimoDiaDoMesPassado = new Date(ano, mesAnterior + 1, 0); // Último dia do mês passado

        var todosOsDias = [];

        // Itera por todos os dias do mês passado
        for (var d = primeiroDiaDoMesPassado; d <= ultimoDiaDoMesPassado; d.setDate(d.getDate() + 1)) {
            todosOsDias.push(new Date(d)); // Adiciona a data à lista de todos os dias
        }

        return todosOsDias;
    };

    function formatarData(data) {
        var dia = String(data.getDate()).padStart(2, '0');
        var mes = String(data.getMonth() + 1).padStart(2, '0');
        var ano = data.getFullYear();
        return `${dia}/${mes}/${ano}`;
    };

    function escreverDadosRelato() {
        let todosOsDias = todosOsDiasNoMes();
        let linhaInicial = 3;

        if (!abaControle) {
            Logger.log("A aba 'teste relato' não foi encontrada.");
            return;
        }


        funcionarios.forEach(funcionario => {

            abaControle.getRange(`A${linhaInicial - 2}`).setValue(funcionario);
            abaControle.getRange(`A${linhaInicial - 1}`).setValue("TIPO DE EVENTO");
            abaControle.getRange(`A${linhaInicial}`).setValue("AGUARDANDO ÁUDIO DE CLIENTES (AUTORIZAÇÃO)");
            abaControle.getRange(`A${linhaInicial + 1}`).setValue("AGUARDANDO LIGAÇÃO (AUTORIZAÇÃO)");
            abaControle.getRange(`A${linhaInicial + 2}`).setValue("CLIENTES EM TRATATIVA");
            abaControle.getRange(`A${linhaInicial + 3}`).setValue("CLIENTES COM RELATO COLHIDO");
            abaControle.getRange(`A${linhaInicial + 4}`).setValue("CLIENTES ARQUIVADOS");
            abaControle.getRange(`A${linhaInicial + 5}`).setValue("TOTAL DE CLIENTES TRATADOS");
            abaControle.getRange(`A${linhaInicial + 6}`).setValue("TOTAL DE RELATOS COLHIDOS (VÁLIDOS)");
            abaControle.getRange(`A${linhaInicial + 7}`).setValue("TOTAL DE RELATOS COLHIDOS (VÁLIDOS E DEVOLVIDOS)");
            abaControle.getRange(`A${linhaInicial + 8}`).setValue("COBRANÇA INDEVIDA");
            abaControle.getRange(`A${linhaInicial + 9}`).setValue("CONSIGNADO");
            abaControle.getRange(`A${linhaInicial + 10}`).setValue("FALHA NA PRESTAÇÃO DOS SERVIÇOS");
            abaControle.getRange(`A${linhaInicial + 11}`).setValue("HABILITAÇÃO FRAUDULENTA");
            abaControle.getRange(`A${linhaInicial + 12}`).setValue("MANUTENÇÃO INDEVIDA");
            abaControle.getRange(`A${linhaInicial + 13}`).setValue("ÔNIBUS");
            abaControle.getRange(`A${linhaInicial + 14}`).setValue("OUTROS");
            abaControle.getRange(`A${linhaInicial + 15}`).setValue("PRESCRIÇÃO");
            abaControle.getRange(`A${linhaInicial + 16}`).setValue("PROCON");
            abaControle.getRange(`A${linhaInicial + 17}`).setValue("CLIENTES PERDIDOS");
            abaControle.getRange(`A${linhaInicial + 18}`).setValue("CLIENTES NÃO RESPONDERAM");
            abaControle.getRange(`A${linhaInicial + 19}`).setValue("CLIENTES QUE DESISTIRAM");
            abaControle.getRange(`A${linhaInicial + 20}`).setValue("CLIENTES COMARCA PROIBIDA");
            abaControle.getRange(`A${linhaInicial + 21}`).setValue("CLIENTES TODAS DÍVIDAS DEVIDAS");
            abaControle.getRange(`A${linhaInicial + 22}`).setValue("CLIENTE COM DÍVIDA DO PROCON NÃO VIÁVEL");
            abaControle.getRange(`A${linhaInicial + 23}`).setValue("DEVOLVIDOS À CONVERSÃO");
            abaControle.getRange(`A${linhaInicial + 24}`).setValue("DEVOLVIDOS À CONVERSÃO");

            todosOsDias.forEach((data, index) => {
                let coluna = index + 2;
                let dataFormatada = formatarData(data);
                abaControle.getRange(linhaInicial - 1, coluna).setValue(dataFormatada); // Define a data na linha 2
            });


            todosOsDias.forEach((data, index) => {
                let coluna = index + 2;
                let linha = linhaInicial;
                let dataFormatada = formatarData(data);

                let valorAudio = infoRelato[funcionario]?.[dataFormatada]?.["AGUARDANDO ÁUDIO"] || 0;
                let valorLigacao = infoRelato[funcionario]?.[dataFormatada]?.["AGUARDANDO LIGAÇÃO"] || 0;
                let valorTratativa = infoRelato[funcionario]?.[dataFormatada]?.["EM TRATATIVA"] || 0;
                let valorRC = infoRelato[funcionario]?.[dataFormatada]?.["RELATO COLHIDO"] || 0;
                let valorCI = infoRelato[funcionario]?.[dataFormatada]?.["COBRANÇA INDEVIDA"] || 0;
                let valorCons = infoRelato[funcionario]?.[dataFormatada]?.["CONSIGNADO"] || 0;
                let valorFPS = infoRelato[funcionario]?.[dataFormatada]?.["FALHA NA PRESTAÇÃO DOS SERVIÇOS"] || 0;
                let valorHF = infoRelato[funcionario]?.[dataFormatada]?.["HABILITAÇÃO FRAUDULENTA"] || 0;
                let valorMI = infoRelato[funcionario]?.[dataFormatada]?.["MANUTENÇÃO INDEVIDA"] || 0;
                let valorCO = infoRelato[funcionario]?.[dataFormatada]?.["COMPANHIA DE ÔNIBUS"] || 0;
                let valorO = infoRelato[funcionario]?.[dataFormatada]?.["OUTROS"] || 0;
                let valorPresc = infoRelato[funcionario]?.[dataFormatada]?.["PRESCRIÇÃO"] || 0;
                let valorProcon = infoRelato[funcionario]?.[dataFormatada]?.["PROCON"] || 0;
                let valorNaoR = infoRelato[funcionario]?.[dataFormatada]?.["ARQUIVADO - CLIENTE NÃO RESPONDE"] || 0;
                let valorCliD = infoRelato[funcionario]?.[dataFormatada]?.["ARQUIVADO - CLIENTE DESISTIU"] || 0;
                let valorComPro = infoRelato[funcionario]?.[dataFormatada]?.["ARQUIVADO - CLIENTE COMARCA PROIBIDA"] || 0;
                let valorCTDD = infoRelato[funcionario]?.[dataFormatada]?.["ARQUIVADO - CLIENTE TODAS DÍVIDAS DEVIDAS"] || 0;
                let valorDPNV = infoRelato[funcionario]?.[dataFormatada]?.["ARQUIVADO - DÍVIDA DO PROCON NÃO VIÁVEL"] || 0;
                let devolucaoValor = infoRelato[funcionario]?.[dataFormatada]?.["DEVOLVIDO"] || 0;


                abaControle.getRange(linha, coluna).setValue(valorAudio);
                abaControle.getRange(linha + 1, coluna).setValue(valorLigacao);
                abaControle.getRange(linha + 2, coluna).setValue(valorTratativa);
                abaControle.getRange(linha + 3, coluna).setValue(valorRC);
                abaControle.getRange(linha + 8, coluna).setValue(valorCI);
                abaControle.getRange(linha + 9, coluna).setValue(valorCons);
                abaControle.getRange(linha + 10, coluna).setValue(valorFPS);
                abaControle.getRange(linha + 11, coluna).setValue(valorHF);
                abaControle.getRange(linha + 12, coluna).setValue(valorMI);
                abaControle.getRange(linha + 13, coluna).setValue(valorCO);
                abaControle.getRange(linha + 14, coluna).setValue(valorO);
                abaControle.getRange(linha + 15, coluna).setValue(valorPresc);
                abaControle.getRange(linha + 16, coluna).setValue(valorProcon);
                abaControle.getRange(linha + 18, coluna).setValue(valorNaoR);
                abaControle.getRange(linha + 19, coluna).setValue(valorCliD);
                abaControle.getRange(linha + 20, coluna).setValue(valorComPro);
                abaControle.getRange(linha + 21, coluna).setValue(valorCTDD);
                abaControle.getRange(linha + 22, coluna).setValue(valorDPNV);
                abaControle.getRange(linha + 24, coluna).setValue(devolucaoValor);
            });

            linhaInicial = linhaInicial + 28;
        });
    };

    function limparAbaControle() {
        if (abaControle) {
            let ultimaLinha = abaControle.getLastRow();
            let ultimaColuna = abaControle.getLastColumn();
            if (ultimaLinha > 0 && ultimaColuna > 0) {
                abaControle.getRange(1, 1, ultimaLinha, ultimaColuna).clearContent(); // Limpa apenas os valores, mantendo formatação
            };
        };
    };

};
