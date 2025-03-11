function controleConversao() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let funcionarios = listaDeFuncionarios();
    let abaAcompanhamento = ss.getSheetByName("ACOMPANHAMENTO - CONVERSÃO");
    let feriadosNoMes = 2;
    let compl = "CNV - ";

    var hoje = new Date();

    let mapAgentes = new Map();
    let logs = [];

    processarAbaAgentesCV();
    processarAbaCanceladosCV();
    processarAbaRelatoCV();
    processarValidacaoCV();
    processarAbaDevolucoesRelatoCV();
    processarAbaArqRelatoCV();
    escreverQuadros();
    diasUteisNoMes();
    escreverLogs();

    function processarAbaAgentesCV() {

        funcionarios.forEach(agente => {
            let abaAgente = ss.getSheetByName(compl + agente);
            let dadosAgente = abaAgente.getDataRange().getValues();
            let cabecalhoAgente = dadosAgente[0];

            let complBV = "BOAS-VINDAS ";
            let dadosBoasVindasAgente = ss.getSheetByName(complBV + agente).getDataRange().getValues();

            let indiceData = cabecalhoAgente.indexOf('DATA')
            let indiceNome = cabecalhoAgente.indexOf('NOME');
            let indiceLead = cabecalhoAgente.indexOf('LEAD');
            let indiceAcao = cabecalhoAgente.indexOf('TIPO DE AÇÃO');
            let indiceProcuracao = cabecalhoAgente.indexOf('PROCURAÇÃO');
            let indiceLiberado = cabecalhoAgente.indexOf('PASTA LIBERADA?');

            if (indiceData == -1) {
                logs.push(`Erro na aba ${agente}: CABEÇALHO INVÁLIDO - A COLUNA 'DATA' NÃO FOI ENCONTRADA.`);
                return;
            };

            if (indiceNome == -1) {
                logs.push(`Erro na aba ${agente}: CABEÇALHO INVÁLIDO - A COLUNA 'NOME' NÃO FOI ENCONTRADA.`);
                return;
            };

            if (indiceAcao == -1) {
                logs.push(`Erro na aba ${agente}: CABEÇALHO INVÁLIDO - A COLUNA 'TIPO DE AÇÃO' NÃO FOI ENCONTRADA.`);
                return;
            };

            if (indiceProcuracao == -1) {
                logs.push(`Erro na aba ${agente}: CABEÇALHO INVÁLIDO - A COLUNA 'PROCURAÇÃO' NÃO FOI ENCONTRADA.`);
                return;
            };

            if (indiceLiberado == -1) {
                logs.push(`Erro na aba ${agente}: CABEÇALHO INVÁLIDO - A COLUNA 'PASTA LIBERADA?' NÃO FOI ENCONTRADA.`);
                return;
            };

            if (indiceLead == -1) {
                logs.push(`Erro na aba ${agente}: CABEÇALHO INVÁLIDO - A COLUNA 'LEAD' NÃO FOI ENCONTRADA.`);
                return;
            };



            for (i = 1; i < dadosAgente.length; i++) {
                let linha = dadosAgente[i];
                let data = linha[indiceData];
                let nome = linha[indiceNome];
                let lead = linha[indiceLead];
                let acao = linha[indiceAcao];
                let procuracao = linha[indiceProcuracao];
                let liberado = linha[indiceLiberado];


                if (!nome) return;
                if (!data) {
                    logs.push(`DATA está vazia na linha ${i + 1} da aba ${agente}`);
                    continue;
                } else if (!(data instanceof Date)) {
                    logs.push(`DATA está em formato inválido na linha ${i + 1} da aba ${agente}`);
                    continue;
                };

                if (isThisMonth(data)) {

                    // let boolVABV = verificarAusentesBoasVindas(dadosBoasVindasAgente, nome);

                    // if (!boolVABV){
                    //   logs.push(`Cliente ${nome} não escontrado em boas-vindas de ${agente}.`);
                    // };

                    if (!mapAgentes[agente]) mapAgentes[agente] = {};
                    if (!mapAgentes[agente]['GERAL']) mapAgentes[agente]['GERAL'] = {};

                    mapAgentes[agente]['GERAL']['PROCURAÇÃO GERADA'] = (mapAgentes[agente]['GERAL']['PROCURAÇÃO GERADA'] || 0) + 1;

                    if (procuracao == 'CANCELAR - ASSINATURA') {
                        mapAgentes[agente]['GERAL']['PERDA DE ASSINATURA'] = (mapAgentes[agente]['GERAL']['PERDA DE ASSINATURA'] || 0) + 1;
                    } else if (procuracao == 'CANCELAR - PASTA') {
                        mapAgentes[agente]['GERAL']['PERDA PÓS ASSINATURA FC'] = (mapAgentes[agente]['GERAL']['PERDA PÓS ASSINATURA FC'] || 0) + 1;
                    } else if (procuracao == 'OK' && liberado == 'VERDADEIRO') {
                        mapAgentes[agente]['GERAL']['PASTA LIBERADA'] = (mapAgentes[agente]['GERAL']['PASTA LIBERADA'] || 0) + 1;
                    } else if (procuracao == '') {
                        mapAgentes[agente]['GERAL']['PENDENTE DE ASSINATURA'] = (mapAgentes[agente]['GERAL']['PENDENTE DE ASSINATURA'] || 0) + 1;
                    };

                    if (procuracao == 'OK' && liberado == 'FALSO') {
                        mapAgentes[agente]['GERAL']['PASTAS PENDENTES'] = (mapAgentes[agente]['GERAL']['PASTAS PENDENTES'] || 0) + 1;
                    };

                    if (lead) {
                        if (!mapAgentes[agente]) mapAgentes[agente] = {};
                        if (!mapAgentes[agente][lead]) mapAgentes[agente][lead] = {};

                        mapAgentes[agente][lead]['PROCURAÇÃO GERADA'] = (mapAgentes[agente][lead]['PROCURAÇÃO GERADA'] || 0) + 1;

                        if (procuracao == 'CANCELAR - ASSINATURA') {
                            mapAgentes[agente][lead]['PERDA DE ASSINATURA'] = (mapAgentes[agente][lead]['PERDA DE ASSINATURA'] || 0) + 1;
                        } else if (procuracao == 'CANCELAR - PASTA') {
                            mapAgentes[agente][lead]['PERDA PÓS ASSINATURA FC'] = (mapAgentes[agente][lead]['PERDA PÓS ASSINATURA FC'] || 0) + 1;
                        } else if (procuracao == 'OK' && liberado == 'VERDADEIRO') {
                            mapAgentes[agente][lead]['PASTA LIBERADA'] = (mapAgentes[agente][lead]['PASTA LIBERADA'] || 0) + 1;
                        } else if (procuracao == '') {
                            mapAgentes[agente][lead]['PENDENTE DE ASSINATURA'] = (mapAgentes[agente][lead]['PENDENTE DE ASSINATURA'] || 0) + 1;
                        };

                        if (procuracao == 'OK' && liberado == 'FALSO') {
                            mapAgentes[agente][lead]['PASTAS PENDENTES'] = (mapAgentes[agente][lead]['PASTAS PENDENTES'] || 0) + 1;
                        };
                    };

                };


            };

        });

    };

    function processarAbaCanceladosCV() {
        let abaCancelados = ss.getSheetByName("CANCELADAS");
        let dadosCancelados = abaCancelados.getDataRange().getValues();
        let cabecalhoCancelados = dadosCancelados[0];

        let indiceData = cabecalhoCancelados.indexOf('DATA DA CONVERSÃO');
        let indiceAgente = cabecalhoCancelados.indexOf('AGENTE RESPONSÁVEL');
        let indiceLead = cabecalhoCancelados.indexOf('LEAD');
        let indiceTipoPerda = cabecalhoCancelados.indexOf('TIPO DE PERDA');

        for (i = 1; i < dadosCancelados.length; i++) {
            let linha = dadosCancelados[i];
            let data = linha[indiceData];
            let agente = linha[indiceAgente];
            let tipoPerda = linha[indiceTipoPerda];
            let lead = linha[indiceLead];

            if (!agente) return;
            if (!data) {
                logs.push(`Erro: DATA DA CONVERSÃO está vazia na linha ${i + 1} da aba CANCELADAS`);
                continue;
            };

            if (isThisMonth(data)) {

                if (!mapAgentes[agente]) mapAgentes[agente] = {};
                if (!mapAgentes[agente]['GERAL']) mapAgentes[agente]['GERAL'] = {};

                mapAgentes[agente]['GERAL']['PROCURAÇÃO GERADA'] = (mapAgentes[agente]['GERAL']['PROCURAÇÃO GERADA'] || 0) + 1;

                if (tipoPerda == 'Assinatura Perdida') {
                    mapAgentes[agente]['GERAL']['PERDA DE ASSINATURA'] = (mapAgentes[agente]['GERAL']['PERDA DE ASSINATURA'] || 0) + 1;
                } else if (tipoPerda == 'Pasta Perdida') {
                    mapAgentes[agente]['GERAL']['PERDA PÓS ASSINATURA FC'] = (mapAgentes[agente]['GERAL']['PERDA PÓS ASSINATURA FC'] || 0) + 1;
                };

                if (lead) {

                    if (!mapAgentes[agente]) mapAgentes[agente] = {};
                    if (!mapAgentes[agente][lead]) mapAgentes[agente][lead] = {};

                    mapAgentes[agente][lead]['PROCURAÇÃO GERADA'] = (mapAgentes[agente][lead]['PROCURAÇÃO GERADA'] || 0) + 1;

                    if (tipoPerda == 'Assinatura Perdida') {
                        mapAgentes[agente][lead]['PERDA DE ASSINATURA'] = (mapAgentes[agente][lead]['PERDA DE ASSINATURA'] || 0) + 1;
                    } else if (tipoPerda == 'Pasta Perdida') {
                        mapAgentes[agente][lead]['PERDA PÓS ASSINATURA FC'] = (mapAgentes[agente][lead]['PERDA PÓS ASSINATURA FC'] || 0) + 1;
                    };

                };

            };

        };

    };

    function processarAbaRelatoCV() {
        let listaAbas = ['RELATO - PRESCRIÇÃO', 'RELATO - PROCON'];
        listaAbas.forEach(aba => {
            let abaRelato = ss.getSheetByName(aba);
            let dadosRelato = abaRelato.getDataRange().getValues();
            let cabecalhoRelato = dadosRelato[0];

            let indiceData = cabecalhoRelato.indexOf('DATA DA CONVERSÃO');
            let indiceAgenteConversao = cabecalhoRelato.indexOf('RESPONSÁVEL CONVERSÃO');
            let indiceCliente = cabecalhoRelato.indexOf('CLIENTE');
            let indiceStatus = cabecalhoRelato.indexOf('STATUS');
            let indiceLead = cabecalhoRelato.indexOf('TIPO DE LEAD');

            for (i = 1; i < dadosRelato.length; i++) {
                let linha = dadosRelato[i];
                let data = linha[indiceData];
                let agenteConversao = linha[indiceAgenteConversao];
                let cliente = linha[indiceCliente];
                let status = linha[indiceStatus];
                let lead = linha[indiceLead];

                if (!cliente) return;
                if (!data) {
                    logs.push(`Erro: DATA DA CONVERSÃO está vazia na linha ${i + 1} da aba ${aba}`);
                    continue;
                };
                if (!agenteConversao) {
                    logs.push(`Erro: AGENTE DA CONVERSÃO está vazio na linha ${i + 1} da aba ${aba}`);
                    continue;
                };



                if (isThisMonth(data)) {

                    if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                    if (!mapAgentes[agenteConversao]['GERAL']) mapAgentes[agenteConversao]['GERAL'] = {};

                    if (status != 'FINALIZADO - RELATO COLHIDO') {
                        mapAgentes[agenteConversao]['GERAL']['PENDENTE DE RELATO'] = (mapAgentes[agenteConversao]['GERAL']['PENDENTE DE RELATO'] || 0) + 1;
                    };

                    if (lead) {

                        if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                        if (!mapAgentes[agenteConversao][lead]) mapAgentes[agenteConversao][lead] = {};

                        if (status != 'FINALIZADO - RELATO COLHIDO') {
                            mapAgentes[agenteConversao][lead]['PENDENTE DE RELATO'] = (mapAgentes[agenteConversao][lead]['PENDENTE DE RELATO'] || 0) + 1;
                        };

                    };
                };
            };
        });
    };

    function processarValidacaoCV() {

        let abaValidacao = ss.getSheetByName("VALIDAÇÃO");
        let dadosValidacao = abaValidacao.getDataRange().getValues();
        let cabecalhoValidacao = dadosValidacao[0];

        let indiceData = cabecalhoValidacao.indexOf('DATA DA CONVERSÃO');
        let indiceAgenteConversao = cabecalhoValidacao.indexOf('RESPONSÁVEL CONVERSÃO');
        let indiceCliente = cabecalhoValidacao.indexOf('CLIENTE');
        let indiceStatus = cabecalhoValidacao.indexOf('STATUS');
        let indiceLead = cabecalhoValidacao.indexOf('TIPO DE LEAD');

        for (i = 1; i < dadosValidacao.length; i++) {
            let linha = dadosValidacao[i];
            let data = linha[indiceData];
            let agenteConversao = linha[indiceAgenteConversao];
            let cliente = linha[indiceCliente];
            let status = linha[indiceStatus];
            let lead = linha[indiceLead];

            if (!cliente) return;
            if (!data) {
                logs.push(`Erro: DATA DA CONVERSÃO está vazia na linha ${i + 1} da aba VALIDAÇÃO`);
                continue;
            };
            if (!agenteConversao) {
                logs.push(`Erro: AGENTE DA CONVERSÃO está vazio na linha ${i + 1} da aba VALIDAÇÃO`);
                continue;
            };

            if (isThisMonth(data)) {

                if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                if (!mapAgentes[agenteConversao]['GERAL']) mapAgentes[agenteConversao]['GERAL'] = {};

                if (status != 'LIBERADO') {
                    mapAgentes[agenteConversao]['GERAL']['PENDENTE DE VALIDAÇÃO'] = (mapAgentes[agenteConversao]['GERAL']['PENDENTE DE VALIDAÇÃO'] || 0) + 1;
                } else if (status == 'LIBERADO') {
                    mapAgentes[agenteConversao]['GERAL']['VALIDAÇÃO LIBERADO'] = (mapAgentes[agenteConversao]['GERAL']['VALIDAÇÃO LIBERADO'] || 0) + 1;
                };

                if (status == 'PENDÊNCIA - CONVERSÃO') {
                    mapAgentes[agenteConversao]['GERAL']['PENDÊNCIA - CONVERSÃO'] = (mapAgentes[agenteConversao]['GERAL']['PENDÊNCIA - CONVERSÃO'] || 0) + 1;
                } else if (status == 'CLIENTE ARQUIVADO') {
                    mapAgentes[agenteConversao]['GERAL']['PERDA PÓS ASSINATURA PC'] = (mapAgentes[agenteConversao]['GERAL']['PERDA PÓS ASSINATURA PC'] || 0) + 1;
                };

                if (lead) {
                    if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                    if (!mapAgentes[agenteConversao][lead]) mapAgentes[agenteConversao][lead] = {};

                    if (status != 'LIBERADO') {
                        mapAgentes[agenteConversao][lead]['PENDENTE DE VALIDAÇÃO'] = (mapAgentes[agenteConversao][lead]['PENDENTE DE VALIDAÇÃO'] || 0) + 1;
                    } else if (status == 'LIBERADO') {
                        mapAgentes[agenteConversao][lead]['VALIDAÇÃO LIBERADO'] = (mapAgentes[agenteConversao][lead]['VALIDAÇÃO LIBERADO'] || 0) + 1;
                    };

                    if (status == 'PENDÊNCIA - CONVERSÃO') {
                        mapAgentes[agenteConversao][lead]['PENDÊNCIA - CONVERSÃO'] = (mapAgentes[agenteConversao][lead]['PENDÊNCIA - CONVERSÃO'] || 0) + 1;
                    } else if (status == 'CLIENTE ARQUIVADO') {
                        mapAgentes[agenteConversao][lead]['PERDA PÓS ASSINATURA PC'] = (mapAgentes[agenteConversao][lead]['PERDA PÓS ASSINATURA PC'] || 0) + 1;
                    };
                };
            };
        };

    };

    function processarAbaDevolucoesRelatoCV() {
        let abaDevolucoesRelato = ss.getSheetByName("DEVOLUÇÕES - RELATO");
        let dadosDevolucoesRelato = abaDevolucoesRelato.getDataRange().getValues();
        let cabecalhoDevolucoesRelato = dadosDevolucoesRelato[0];

        let indiceData = cabecalhoDevolucoesRelato.indexOf('DATA DA CONVERSÃO');
        let indiceAgenteConversao = cabecalhoDevolucoesRelato.indexOf('RESPONSÁVEL CONVERSÃO');
        let indiceCliente = cabecalhoDevolucoesRelato.indexOf('CLIENTE');
        let indiceLead = cabecalhoDevolucoesRelato.indexOf('TIPO DE LEAD');

        for (i = 1; i < dadosDevolucoesRelato.length; i++) {
            let linha = dadosDevolucoesRelato[i];
            let data = linha[indiceData];
            let agenteConversao = linha[indiceAgenteConversao];
            let cliente = linha[indiceCliente];
            let lead = linha[indiceLead];

            if (!cliente) return;
            if (!data) {
                logs.push(`Erro: DATA DA CONVERSÃO está vazia na linha ${i + 1} da aba DEVOLUÇÕES - RELATO`);
                continue;
            };
            if (!agenteConversao) {
                logs.push(`Erro: AGENTE DA CONVERSÃO está vazio na linha ${i + 1} da aba DEVOLUÇÕES - RELATO`);
                continue;
            };

            if (isThisMonth(data)) {

                if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                if (!mapAgentes[agenteConversao]['GERAL']) mapAgentes[agenteConversao]['GERAL'] = {};

                mapAgentes[agenteConversao]['GERAL']['PENDÊNCIA - CONVERSÃO'] = (mapAgentes[agenteConversao]['GERAL']['PENDÊNCIA - CONVERSÃO'] || 0) + 1;

                if (lead) {

                    if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                    if (!mapAgentes[agenteConversao][lead]) mapAgentes[agenteConversao][lead] = {};

                    mapAgentes[agenteConversao][lead]['PENDÊNCIA - CONVERSÃO'] = (mapAgentes[agenteConversao][lead]['PENDÊNCIA - CONVERSÃO'] || 0) + 1;

                };
            };
        };
    };

    function processarAbaArqRelatoCV() {
        let abaArqRelato = ss.getSheetByName("CLIENTES ARQUIVADOS - RELATO");
        let dadosArqRelato = abaArqRelato.getDataRange().getValues();
        let cabecalhoArqRelato = dadosArqRelato[0];

        let indiceData = cabecalhoArqRelato.indexOf('DATA DA CONVERSÃO');
        let indiceAgenteConversao = cabecalhoArqRelato.indexOf('RESPONSÁVEL CONVERSÃO');
        let indiceCliente = cabecalhoArqRelato.indexOf('CLIENTE');
        let indiceLead = cabecalhoArqRelato.indexOf('TIPO DE LEAD');

        for (i = 1; i < dadosArqRelato.length; i++) {
            let linha = dadosArqRelato[i];
            let data = linha[indiceData];
            let agenteConversao = linha[indiceAgenteConversao];
            let cliente = linha[indiceCliente];
            let lead = linha[indiceLead];

            if (!cliente) return;
            if (!data) {
                logs.push(`Erro: DATA DA CONVERSÃO está vazia na linha ${i + 1} da aba CLIENTES ARQUIVADOS - RELATO`);
                continue;
            };
            if (!agenteConversao) {
                logs.push(`Erro: AGENTE DA CONVERSÃO está vazio na linha ${i + 1} da aba CLIENTES ARQUIVADOS - RELATO`);
                continue;
            };

            if (isThisMonth(data)) {

                if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                if (!mapAgentes[agenteConversao]['GERAL']) mapAgentes[agenteConversao]['GERAL'] = {};

                mapAgentes[agenteConversao]['GERAL']['PERDA PÓS ASSINATURA PC'] = (mapAgentes[agenteConversao]['GERAL']['PERDA PÓS ASSINATURA PC'] || 0) + 1;

                if (lead) {
                    if (!mapAgentes[agenteConversao]) mapAgentes[agenteConversao] = {};
                    if (!mapAgentes[agenteConversao][lead]) mapAgentes[agenteConversao][lead] = {};

                    mapAgentes[agenteConversao][lead]['PERDA PÓS ASSINATURA PC'] = (mapAgentes[agenteConversao][lead]['PERDA PÓS ASSINATURA PC'] || 0) + 1;
                };
            };
        };

    };

    function escreverQuadros() {
        let listaLeads = [
            'GERAL',
            'PROCON',
            'ROBÔ/CONSUMIDOR',
            'INDICAÇÃO',
            'RECONVERSÃO',
            'COMPANHIA DE ÔNIBUS',
            'CONSIGNADO',
            'BANCOS',
        ];
        let linhaInicial = 9;

        if (!abaAcompanhamento) {
            Logger.log("A aba 'teste relato' não foi encontrada.");
            return;
        };
        listaLeads.forEach(lead => {

            funcionarios.forEach(funcionario => {
                let procuracaoGerada = mapAgentes[funcionario]?.[lead]?.['PROCURAÇÃO GERADA'] || 0;
                let perdaAssinatura = mapAgentes[funcionario]?.[lead]?.['PERDA DE ASSINATURA'] || 0;
                let perdaPosAssinaturaFC = mapAgentes[funcionario]?.[lead]?.['PERDA PÓS ASSINATURA FC'] || 0;
                let pendenteAssinatura = mapAgentes[funcionario]?.[lead]?.['PENDENTE DE ASSINATURA'] || 0;
                let pastaLiberada = mapAgentes[funcionario]?.[lead]?.['PASTA LIBERADA'] || 0;
                let aguardandoRelato = mapAgentes[funcionario]?.[lead]?.['PENDENTE DE RELATO'] || 0;
                let pendenteValidacao = mapAgentes[funcionario]?.[lead]?.['PENDENTE DE VALIDAÇÃO'] || 0;
                let validacaoLiberado = mapAgentes[funcionario]?.[lead]?.['VALIDAÇÃO LIBERADO'] || 0;
                let pendenciaConversao = mapAgentes[funcionario]?.[lead]?.['PENDÊNCIA - CONVERSÃO'] || 0;
                let perdaPosAssinaturaPC = mapAgentes[funcionario]?.[lead]?.['PERDA PÓS ASSINATURA PC'] || 0;
                let pastasPendentes = mapAgentes[funcionario]?.[lead]?.['PASTAS PENDENTES'] || 0;

                abaAcompanhamento.getRange(linhaInicial, 2).setValue(funcionario);
                abaAcompanhamento.getRange(linhaInicial, 5).setValue(procuracaoGerada);
                abaAcompanhamento.getRange(linhaInicial, 6).setValue(perdaAssinatura);
                abaAcompanhamento.getRange(linhaInicial, 7).setValue(perdaPosAssinaturaFC);
                abaAcompanhamento.getRange(linhaInicial, 8).setValue(pendenteAssinatura);
                abaAcompanhamento.getRange(linhaInicial, 10).setValue(pastaLiberada);
                abaAcompanhamento.getRange(linhaInicial, 11).setValue(aguardandoRelato);
                abaAcompanhamento.getRange(linhaInicial, 12).setValue(pendenteValidacao);
                abaAcompanhamento.getRange(linhaInicial, 13).setValue(validacaoLiberado);
                abaAcompanhamento.getRange(linhaInicial, 14).setValue(pendenciaConversao);
                abaAcompanhamento.getRange(linhaInicial, 15).setValue(perdaPosAssinaturaPC);
                abaAcompanhamento.getRange(linhaInicial, 16).setValue(pastasPendentes);

                linhaInicial++;
            });

            linhaInicial += 4;
        });
    };

    function isThisMonth(date) {
        date = new Date(date);
        return (date.getMonth() == hoje.getMonth() && date.getFullYear() == hoje.getFullYear());
    };

    function isLastMonth(date) {
        date = new Date(date);
        let lastMonth = hoje.getMonth() - 1;
        let lastYear = hoje.getFullYear();
        if (lastMonth == -1) {
            lastMonth = 11;
            lastYear = hoje.getFullYear() - 1;
        };
        return (date.getMonth() == lastMonth && date.getFullYear() == lastYear);
    };

    function diasUteisNoMes() {
        var primeiroDiaDoMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
        var diasUteis = 0;

        // Itera pelos dias do mês até hoje
        for (var d = primeiroDiaDoMes; d <= hoje; d.setDate(d.getDate() + 1)) {

            var diaDaSemana = d.getDay();

            if (diaDaSemana !== 0 && diaDaSemana !== 6) {
                diasUteis++;
            };

        };

        return diasUteis - feriadosNoMes;
    };

    function escreverLogs() {
        let abaLogs = ss.getSheetByName("LOGS");
        let linha = 3;
        let ultimaLinha = abaLogs.getLastRow() || 3;
        abaLogs.getRange(3, 2, ultimaLinha, 3).clearContent();
        const horaAtual = Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "HH:mm"
        );

        // for (i = linha; i <= 18; i++) {
        //     abaLogs.getRange(i, 2, 1, 3).merge();
        // };

        logs.forEach(log => {
            abaLogs.getRange(linha, 2, 1, 3).merge();
            abaLogs.getRange(linha, 2).setValue(log);
            linha++;
        });

        abaLogs.getRange(1, 1).setValue('ULTIMA EXECUÇÃO').setBackground("#F06060");
        abaLogs.getRange(2, 1).setValue(horaAtual).setBackground("#FFF");

    };

    function verificarAusentesBoasVindas(dadosBoasVindasAgente, nome) {
        let indiceNome = dadosBoasVindasAgente[0].indexOf('NOME');
        let indiceData = dadosBoasVindasAgente[0].indexOf('DATA');

        let nomesReceberamBV = new Set();

        for (let i = 1; i < dadosBoasVindasAgente.length; i++) {
            let linha = dadosBoasVindasAgente[i];
            if (isThisMonth(linha[indiceData])) {
                nomesReceberamBV.add(linha[indiceNome]);
            };
        };

        return nomesReceberamBV.has(nome);
    };


    abaAcompanhamento.getRange(1, 1).setValue('ULTIMA EXECUÇÃO');
    abaAcompanhamento.getRange(2, 1).setValue(hoje);
    abaAcompanhamento.getRange(3, 8).setValue(diasUteisNoMes());

};

function listaDeFuncionarios() {
    return [
        "ANNA",
        "ISABELLA",
        "JADE",
        "JOÃO",
        "LAURA",
        "LIVIA",
        "MARIANA",
        "MARIA EDUARDA",
        "RAYSSA",
        "THAYNA",
        "VIVIAN"
    ].sort();
};
