function controleInicial() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaInicial = ss.getSheetByName('INICIAL');
    const abaInicialRealizada = ss.getSheetByName('INICIAL REALIZADA');
    const abaDevolucoesInicial = ss.getSheetByName('DEVOLUÇÕES - INICIAL');
    const abaAcompanhamentoInicial = ss.getSheetByName('ACOMPANHAMENTO INICIAL');
    const funcionariosInicial = [
        'xx',
        'yy',
        'zz',
        'ww',
        'll',
    ].sort()

    if (!abaInicial || !abaInicialRealizada || !abaDevolucoesInicial) {
        console.log('Por favor, certifique-se de que as abas INICIAL, INICIAL REALIZADA e DEVOLUÇÕES - INICIAL existam na planilha.');
        return;
    }

    let todosOsDados = {};
    var dadosAgentesInicial = {};

    obterTodosOsDados();
    processarInicial();
    processarFeitos();
    processarDevolucoesInicial();

    escreverQuadroFeitos();
    escreverDevolucoes();
    escreverAcompanhamentoGeral();
    escreverQuadroAcompanhamentoTipo();


    function processarInicial() {
        let dados = todosOsDados["INICIAL"];
        const cabecalhoInicial = dados[0];

        let indexAgente = cabecalhoInicial.indexOf('agente inicial');
        let indexCliente = cabecalhoInicial.indexOf('CLIENTE');
        let indexStatus = cabecalhoInicial.indexOf('STATUS');
        let indexAcao = cabecalhoInicial.indexOf('AÇÃO');

        for (let i = 1; i < dados.length; i++) {
            let linha = dados[i];
            let agente = linha[indexAgente] || 'SEM AGENTE';
            let cliente = linha[indexCliente];
            let status = linha[indexStatus];
            let acao = linha[indexAcao];

            if (!cliente || !acao) continue;

            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            if (!dadosAgentesInicial[agente][acao]) dadosAgentesInicial[agente][acao] = { 'ESTOQUE': 0 };


            dadosAgentesInicial[agente]['ESTOQUE'] = (dadosAgentesInicial[agente]['ESTOQUE'] || 0) + 1;
            dadosAgentesInicial[agente][acao]['ESTOQUE']++;

            if (status && status !== 'DEVOLVIDO AO RELATO') {
                dadosAgentesInicial[agente][acao][status] = (dadosAgentesInicial[agente][acao][status] || 0) + 1;
            };
        };
    };

    function processarFeitos() {
        let listaAbas = [todosOsDados["INICIAL REALIZADA"]];

        listaAbas.forEach((aba) => {
            const cabecalho = aba[0];

            let indexAgente = cabecalho.indexOf('agente inicial');
            let indexCliente = cabecalho.indexOf('CLIENTE');
            let indexStatus = cabecalho.indexOf('STATUS');
            let indexRealizacao = cabecalho.indexOf('REALIZAÇÃO');
            let indexAcao = cabecalho.indexOf('AÇÃO');

            for (let i = 1; i < aba.length; i++) {
                let linha = aba[i];
                let agente = linha[indexAgente] || 'SEM AGENTE';
                let cliente = linha[indexCliente];
                let status = linha[indexStatus];
                let realizacao = formatarData(linha[indexRealizacao]);
                let acao = linha[indexAcao];

                if (!cliente || !acao) continue;

                if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
                if (!dadosAgentesInicial[agente][acao]) dadosAgentesInicial[agente][acao] = {};
                if (!dadosAgentesInicial[agente]["PRESCRIÇÃO"]) dadosAgentesInicial[agente]["PRESCRIÇÃO"] = {};
                if (!dadosAgentesInicial[agente]["NÃO PRESCRIÇÃO"]) dadosAgentesInicial[agente]["NÃO PRESCRIÇÃO"] = {};


                if (status) {
                    dadosAgentesInicial[agente][status] = (dadosAgentesInicial[agente][status] || 0) + 1;
                    dadosAgentesInicial[agente][acao][status] = (dadosAgentesInicial[agente][acao][status] || 0) + 1;

                    if (realizacao) {

                        if (acao === 'PRESCRIÇÃO') {
                            dadosAgentesInicial[agente]["PRESCRIÇÃO"][realizacao] = (dadosAgentesInicial[agente]["PRESCRIÇÃO"][realizacao] || 0) + 1;
                        } else if (acao != 'PRESCRIÇÃO') {
                            dadosAgentesInicial[agente]["NÃO PRESCRIÇÃO"][realizacao] = (dadosAgentesInicial[agente]["NÃO PRESCRIÇÃO"][realizacao] || 0) + 1;
                        };
                    };
                };
            };
        });
    };

    function processarDevolucoesInicial() {
        let dados = todosOsDados["DEVOLUÇÕES - INICIAL"];
        const cabecalhoDevolucoes = dados[0];

        let indexAgente = cabecalhoDevolucoes.indexOf('agente inicial');
        let indexCliente = cabecalhoDevolucoes.indexOf('CLIENTE');
        let indexStatus = cabecalhoDevolucoes.indexOf('STATUS');
        let indexAcao = cabecalhoDevolucoes.indexOf('AÇÃO');
        let indexDevolucaoInicial = cabecalhoDevolucoes.indexOf('DEVOLUÇÃO DA INICIAL');

        for (let i = 1; i < dados.length; i++) {
            let linha = dados[i];
            let agente = linha[indexAgente] || 'SEM AGENTE';
            let cliente = linha[indexCliente];
            let status = linha[indexStatus];
            let acao = linha[indexAcao];
            let realizacao = formatarData(linha[indexDevolucaoInicial]);

            if (!cliente) break;

            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            if (!dadosAgentesInicial[agente]['DEVOLUÇÕES']) dadosAgentesInicial[agente]['DEVOLUÇÕES'] = {};

            dadosAgentesInicial[agente]['DEVOLUÇÕES'][realizacao] = (dadosAgentesInicial[agente]['DEVOLUÇÕES'][realizacao] || 0) + 1;
        };
    }

    function escreverQuadroFeitos() {

        todosOsDias = todosOsDiasNoMes();
        linhaInicial = 5;

        let metaPresc = 0
        let metaNPresc = 0

        todosOsDias.forEach((data, index) => {
            let coluna = index + 2;
            let dataFormatada = formatarData(data);

            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaInicial, coluna)
                .setValue(dataFormatada)
                .setBackground('#8989EB'));

            let diaDaSemana = data.getDay(); // 0 = domingo, 6 = sábado

            if (diaDaSemana > 0 && diaDaSemana < 6) {
                metaPresc += 30; // Adiciona 30 apenas para dias úteis)
                metaNPresc += 26;
            };

            // Escreve a meta na linha 2
            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(2, coluna)
                .setValue(metaPresc)
                .setBackground('#FFF'));
            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(3, coluna)
                .setValue(metaNPresc)
                .setBackground('#FFF'));
        });

        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('A1')
            .setValue('FEITAS')
            .setBackground('#8989EB'));
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial
            .getRange(1, 2, 1, abaAcompanhamentoInicial.getLastColumn())
            .merge()
            .setValue('.')
            .setBackground('#8989EB'));
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('A2')
            .setValue('TOTAL PRESCRIÇÃO')
            .setBackground('#8989EB'));
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('A3')
            .setValue('TOTAL NÃO PRESCRIÇÃO')
            .setBackground('#8989EB'));
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('A4')
            .setValue('TOTAL')
            .setBackground('#8989EB'));
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('A5')
            .setValue('DIAS')
            .setBackground('#8989EB'));


        let linhas = [];

        funcionariosInicial.forEach(agente => {
            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            linhas.push([`${agente} NÃO PRESCRIÇÃO`]);
            linhas.push([`${agente} PRESCRIÇÃO`]);
        });

        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaInicial + 1, 1, linhas.length, 1).setValues(linhas).setBackground('#E8E7FC'));

        funcionariosInicial.forEach(agente => {
            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            let linhaAgente = linhaInicial + 1;

            // Itera sobre cada dia do mês
            todosOsDias.forEach((data, index) => {
                let coluna = index + 2;
                let dataFormatada = formatarData(data);

                // Preenche PRESCRIÇÃO (ou 0 se não houver dados)
                let prescricao = dadosAgentesInicial[agente]['PRESCRIÇÃO']?.[dataFormatada] || 0;
                aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAgente, coluna).setValue(prescricao).setBackground('white'));

                // Preenche NÃO PRESCRIÇÃO (ou 0 se não houver dados)
                let naoPrescricao = dadosAgentesInicial[agente]['NÃO PRESCRIÇÃO']?.[dataFormatada] || 0;
                aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAgente + 1, coluna).setValue(naoPrescricao).setBackground('white'));
            });

            // Avança para o próximo agente
            linhaInicial += 2;


        });

        let linhaTotal = linhaInicial + 1; // Linha após os dados dos agentes
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaTotal, 1).setValue('TOTAL').setBackground('#8989EB'));

        // Adiciona a fórmula de soma para cada coluna
        todosOsDias.forEach((data, index) => {
            let coluna = index + 2;

            // Define o intervalo da soma (da linha 6 até a linha anterior ao total)
            let intervaloInicial = 6;
            let intervaloFinal = linhaTotal - 1;

            // Insere a fórmula de soma
            let colunaLetra = abaAcompanhamentoInicial.getRange(1, coluna).getA1Notation().replace(/\d+/g, ''); // Converte número de coluna para letra correta

            abaAcompanhamentoInicial.getRange(linhaTotal, coluna).setFormula(`=SUM(${colunaLetra}${intervaloInicial}:${colunaLetra}${intervaloFinal})`);

            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaTotal, coluna).setBackground('#8989EB'));

        });

    };

    function escreverDevolucoes() {
        let linhaAtual = 28;
        let todosOsDias = todosOsDiasNoMes();
        let linhas = [];

        funcionariosInicial.forEach(agente => {
            linhas.push([agente]);
        });
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAtual, 1, linhas.length, 1).setValues(linhas).setBackground('#E8E7FC'));

        funcionariosInicial.forEach(agente => {
            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            let linhaAgente = linhaAtual;

            // Itera sobre cada dia do mês
            todosOsDias.forEach((data, index) => {
                let coluna = index + 2;
                let dataFormatada = formatarData(data);

                let devolucao = dadosAgentesInicial[agente]['DEVOLUÇÕES']?.[dataFormatada] || 0;
                aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAgente, coluna).setValue(devolucao).setBackground('white'));
            });

            linhaAtual++;

        });

        let linhaTotal = linhaInicial + 1; // Linha após os dados dos agentes
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaTotal, 1).setValue('TOTAL').setBackground('#8989EB'));
    };

    function escreverAcompanhamentoGeral() {
        if (!abaAcompanhamentoInicial) {
            console.log('A aba ACOMPANHAMENTO INICIAL não foi encontrada.');
            return;
        };

        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('B38:F38').merge().setValue('ACOMPANHAMENTO GERAL')
            .setFontSize(24)
            .setBackground('#304673')
        );

        let cabecalhos = ['AGENTE', 'ESTOQUE', 'INICIAL REALIZADA', 'INICIAL REALIZADA COM PROBLEMA SANADO', 'DEVOLVIDO AO RELATO'];
        aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange('B39:F39')
            .setValues([cabecalhos])
            .setBackground('#5176A6')
        );

        let linhas = [];
        funcionariosInicial.forEach(agente => {
            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            let estoque = dadosAgentesInicial[agente]['ESTOQUE'] || 0;
            let inicialRealizada = dadosAgentesInicial[agente]['INICIAL REALIZADA'] || 0;
            let problemaSanado = dadosAgentesInicial[agente]['INICIAL REALIZADA COM PROBLEMA SANADO'] || 0;
            let devolvido = dadosAgentesInicial[agente]['DEVOLVIDO AO RELATO'] || 0;

            linhas.push([agente, estoque, inicialRealizada, problemaSanado, devolvido]);
        });

        if (linhas.length > 0) {
            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(40, 2, linhas.length, 5)
                .setValues(linhas)
                .setBackground('#A7BDD9')
            );
        };
    };

    function escreverQuadroAcompanhamentoTipo() {

        let cabecalhos = ['AÇÃO', 'ESTOQUE', 'INICIAL REALIZADA', 'INICIAL REALIZADA COM PROBLEMA SANADO', 'DEVOLVIDO AO RELATO'];
        let linhaInicial = 46;
        let ultimaColuna = 2; // Primeira posição para o primeiro quadro

        funcionariosInicial.forEach(agente => {
            if (!dadosAgentesInicial[agente]) dadosAgentesInicial[agente] = {};
            let linhaAtual = linhaInicial;

            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAtual, ultimaColuna, 1, 5).merge()
                .setValue(agente)
                .setFontSize(24)
                .setBackground('#304673')
            );

            linhaAtual++;

            aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAtual, ultimaColuna, 1, cabecalhos.length)
                .setValues([cabecalhos])
                .setBackground('#5176A6')
            );

            linhaAtual++;

            let linhas = [];

            for (let acao in dadosAgentesInicial[agente]) {
                if (acao === 'ESTOQUE') continue;

                let estoque = dadosAgentesInicial[agente][acao]['ESTOQUE'] || 0;
                let inicialRealizada = dadosAgentesInicial[agente][acao]['INICIAL REALIZADA'] || 0;
                let problemaSanado = dadosAgentesInicial[agente][acao]['INICIAL REALIZADA COM PROBLEMA SANADO'] || 0;
                let devolvido = dadosAgentesInicial[agente][acao]['DEVOLVIDO AO RELATO'] || 0;

                linhas.push([acao, estoque, inicialRealizada, problemaSanado, devolvido]);

            };

            // Escreve os dados na planilha
            if (linhas.length > 0) {
                aplicarFormatacaoPadrao(abaAcompanhamentoInicial.getRange(linhaAtual, ultimaColuna, linhas.length, 5)
                    .setValues(linhas)
                    .setBackground('#A7BDD9')
                );
            };

            ultimaColuna += 6;
        });

    };




    function aplicarFormatacaoPadrao(range) {
        // Centraliza o conteúdo tanto vertical quanto horizontal
        range.setHorizontalAlignment('center')
            .setVerticalAlignment('middle');

        range.setBorder(true, true, true, true, true, true);
        range.setFontWeight('bold');
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    };

    function obterTodosOsDados() {
        if (abaInicial) todosOsDados["INICIAL"] = abaInicial.getDataRange().getValues();
        if (abaInicialRealizada) todosOsDados["INICIAL REALIZADA"] = abaInicialRealizada.getDataRange().getValues();
        if (abaDevolucoesInicial) todosOsDados["DEVOLUÇÕES - INICIAL"] = abaDevolucoesInicial.getDataRange().getValues();
    };

    function formatarData(data) {
        if (!data) return "DATA INVÁLIDA"; // Retorna uma string padrão para evitar objetos vazios
        let novaData = new Date(data);
        if (isNaN(novaData)) return "DATA INVÁLIDA"; // Garante que não armazene objetos inválidos
        return Utilities.formatDate(novaData, 'GMT-3', 'dd/MM/yyyy');
    };

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

};
