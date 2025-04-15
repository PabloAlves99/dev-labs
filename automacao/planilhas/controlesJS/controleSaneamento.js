function controle() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaControle = ss.getSheetByName("CONTROLE");
    const abasAgentes = funcionariosSol();
    const abaFinalizados = [
        "ESTOQUE - CLIENTES FINALIZADOS",
        "CLIENTES FINALIZADOS CORRETO",
    ];
    const mapAgentes = {
        "xx": "xx",
        "yy": "yy",
        "ll": "ll",
        "kk": "kk",
        "mm": "mm",
        "oo": "oo"
    }

    const hoje = formatarData(new Date());

    const ontem = new Date(hoje);
    formatarData(new Date(ontem.setDate(ontem.getDate() - 1)));

    const amanha = new Date(hoje);
    formatarData(new Date(amanha.setDate(amanha.getDate() + 1)));

    const tresDias = new Date(hoje);
    formatarData(new Date(tresDias.setDate(tresDias.getDate() + 3)));

    const seteDias = new Date(hoje);
    formatarData(new Date(seteDias.setDate(seteDias.getDate() + 7)));

    let omissosGeral = [];
    let sContatoGeral = [];


    // STATUS
    const statusTratativas = [
        "BOAS-VINDAS",
        "1ª TRATATIVA",
        "2ª TRATATIVA",
        "3ª TRATATIVA",
        "OMISSO",
        "SEM CONTATO",
    ];
    const statusLiberados = [
        "LIBERADO (OK)",
        "LIBERADO (DOCS. PARCIAIS)",
        "LIBERADO (DOCS. MÍNIMOS)",
    ];
    const statusDesistencias = [
        "DESISTÊNCIA (COM DECLARAÇÃO)",
        "DESISTÊNCIA (SEM DECLARAÇÃO)",
    ];
    const statusDevolvidos = [
        "DEVOLVIDO SEM CONTATO",
        "DEVOLVIDO POR OMISSÃO",
    ];
    const statusSemAtv = [
        "AUSENTE NO ARES",
        "CONCLUÍDO",
        "TRANSFERIDO",
    ];

    let indiceProcesso, indiceFatal, indiceTipo, indiceCliente, indiceReu,
        indiceAgenteResponsavel, indiceDataInicio, indiceDocSolicitado,
        indiceDespacho, indiceStatus, indiceProcedimentoAdvogado;

    let infoAgentes = {};

    tratarAbasAgentes();
    tratarAbasFinalizados();
    tratarAbaDevolvidos();
    tratarFinalizadosCorretos();
    exibirControle();
    atividadesDoMes();
    limparOmisso();
    limpasSContato();

    function tratarAbasAgentes() {
        abasAgentes.forEach(aba => {

            let dados = ss.getSheetByName(aba).getDataRange().getValues();


            let abaEstoque = ss.getSheetByName("ESTOQUE - CLIENTES FINALIZADOS")
            let abaOmisso = ss.getSheetByName("OMISSO")
            let abaSContato = ss.getSheetByName("SEM CONTATO")

            let linhaEstoque = abaEstoque.getLastRow() + 1;
            let linhaOmisso = abaOmisso.getLastRow() + 1;
            let linhaSContato = abaSContato.getLastRow() + 1;


            let linhasParaExcluir = [];

            if (dados && dados.length > 0) {
                const cabecalho = dados[0];

                indiceProcesso = obterIndiceCabecalho(cabecalho, "PROCESSO");
                indiceFatal = obterIndiceCabecalho(cabecalho, "FATAL");
                indiceTipo = obterIndiceCabecalho(cabecalho, "TIPO");
                indiceCliente = obterIndiceCabecalho(cabecalho, "CLIENTE");
                indiceReu = obterIndiceCabecalho(cabecalho, "RÉU");
                indiceAgenteResponsavel = obterIndiceCabecalho(cabecalho, "AGENTE RESPONSÁVEL");
                indiceDataInicio = obterIndiceCabecalho(cabecalho, "DATA DO INÍCIO DA TRATATIVA");
                indiceDocSolicitado = obterIndiceCabecalho(cabecalho, "DOC. SOLICITADO");
                indiceDespacho = obterIndiceCabecalho(cabecalho, "DESPACHO");
                indiceStatus = obterIndiceCabecalho(cabecalho, "STATUS");
                indiceProcedimentoAdvogado = obterIndiceCabecalho(cabecalho, "PROCEDIMENTO DO AVOGADO");
                indiceTrigger = obterIndiceCabecalho(cabecalho, "TRIGGER");
            };

            for (i = 1; i < dados.length; i++) {
                let fatal = formatarData(dados[i][indiceFatal]);
                let trigger = formatarData(dados[i][indiceTrigger]);
                let tipo = dados[i][indiceTipo].trim().toUpperCase();
                let agente = dados[i][indiceAgenteResponsavel].trim().toUpperCase();
                let docs = dados[i][indiceDocSolicitado];
                let status = dados[i][indiceStatus].trim().toUpperCase();
                let cliente = dados[i][indiceCliente].trim().toUpperCase();

                if (!agente || !cliente) continue;
                if (!infoAgentes[agente]) infoAgentes[agente] = {};

                if (cliente) infoAgentes[agente]["EM ABERTO"] = (infoAgentes[agente]["EM ABERTO"] || 0) + 1;

                if (trigger) {

                    if (trigger.toISOString() == hoje.toISOString()) {
                        infoAgentes[agente]["TRATATIVAS HOJE"] = (infoAgentes[agente]["TRATATIVAS HOJE"] || 0) + 1;

                        if (statusLiberados.includes(status)) {
                            infoAgentes[agente]["LIBERADOS HOJE"] = (infoAgentes[agente]["LIBERADOS HOJE"] || 0) + 1;
                        };
                    };

                };

                if (status) {

                    if (statusSemAtv.includes(status)) linhasParaExcluir.push(i + 1);

                    if (statusLiberados.includes(status) || statusDesistencias.includes(status) || statusDevolvidos.includes(status)) {
                        linhasParaExcluir.push(i + 1);
                        dados[i][indiceProcedimentoAdvogado] = "";
                        dados[i][indiceTrigger] = hoje;
                        abaEstoque.getRange(linhaEstoque, 1, 1, dados[i].length).setValues([dados[i]]);
                        linhaEstoque++;

                    } else if (statusTratativas.includes(status)) {
                        infoAgentes[agente]["TRATATIVA"] = (infoAgentes[agente]["TRATATIVA"] || 0) + 1;

                    } else if (status == 'ENROLADO') {
                        infoAgentes[agente]["ENROLADO"] = (infoAgentes[agente]["ENROLADO"] || 0) + 1;
                    };


                    if (status == "OMISSO") {
                        infoAgentes[agente]["OMISSO"] = (infoAgentes[agente]["OMISSO"] || 0) + 1;

                        let processosOmisso = [];
                        if (abaOmisso.getLastRow() > 1) {
                            processosOmisso = abaOmisso.getRange(2, indiceProcesso + 1, abaOmisso.getLastRow() - 1).getValues().flat();
                        };

                        if (!omissosGeral.includes(dados[i][indiceProcesso])) {
                            omissosGeral.push(dados[i][indiceProcesso]);
                        };

                        if (!processosOmisso.includes(dados[i][indiceProcesso])) {
                            abaOmisso.getRange(linhaOmisso, 1, 1, dados[i].length).setValues([dados[i]]);
                            linhaOmisso++;
                        };

                    } else if (status == "SEM CONTATO") {
                        infoAgentes[agente]["SEM CONTATO"] = (infoAgentes[agente]["SEM CONTATO"] || 0) + 1;

                        let processosSemContato = [];
                        if (abaSContato.getLastRow() > 1) {
                            processosSemContato = abaSContato.getRange(2, indiceProcesso + 1, abaSContato.getLastRow() - 1).getValues().flat();
                        };
                        if (!processosSemContato.includes(dados[i][indiceProcesso])) {
                            abaSContato.getRange(linhaSContato, 1, 1, dados[i].length).setValues([dados[i]]);
                            linhaSContato++;
                        }

                        if (!sContatoGeral.includes(dados[i][indiceProcesso])) {
                            sContatoGeral.push(dados[i][indiceProcesso]);
                        };

                    };
                };

                if (tipo) {
                    if (!infoAgentes[agente][tipo]) infoAgentes[agente][tipo] = {};
                    infoAgentes[agente][tipo]["TOTAL"] = (infoAgentes[agente][tipo]["TOTAL"] || 0) + 1;

                    if (statusLiberados.includes(status)) {
                        infoAgentes[agente][tipo]["FINALIZADOS"] = (infoAgentes[agente][tipo]["FINALIZADOS"] || 0) + 1;

                    } else if (statusDesistencias.includes(status)) {
                        infoAgentes[agente][tipo]["DESISTÊNCIAS"] = (infoAgentes[agente][tipo]["DESISTÊNCIAS"] || 0) + 1;

                    };
                };

                if (fatal && tipo != "DILAÇÃO") {

                    if (statusTratativas.includes(status) || status == '') {

                        if (fatal.toISOString() == ontem.toISOString()) {
                            infoAgentes[agente]["VENCIDOS ONTEM"] = (infoAgentes[agente]["VENCIDOS ONTEM"] || 0) + 1;

                        } else if (fatal.toISOString() == amanha.toISOString()) {
                            infoAgentes[agente]["VENCE AMANHA"] = (infoAgentes[agente]["VENCE AMANHA"] || 0) + 1;

                        } else if (fatal.toISOString() == hoje.toISOString()) {
                            infoAgentes[agente]["VENCE HOJE"] = (infoAgentes[agente]["VENCE HOJE"] || 0) + 1;
                        };

                        if (fatal >= hoje) {
                            if (fatal <= tresDias) infoAgentes[agente]["VENCE EM TRES"] = (infoAgentes[agente]["VENCE EM TRES"] || 0) + 1;
                            if (fatal <= seteDias) infoAgentes[agente]["VENCE EM SETE"] = (infoAgentes[agente]["VENCE EM SETE"] || 0) + 1;
                        };

                        if (fatal < hoje) infoAgentes[agente]["VENCIDOS"] = (infoAgentes[agente]["VENCIDOS"] || 0) + 1;

                    };
                };

                if (docs) {
                    listaDocs = docs.split(",")
                    listaDocs.forEach(doc => {
                        if (!infoAgentes[agente][doc]) infoAgentes[agente][doc] = {};
                        infoAgentes[agente][doc]["TOTAL"] = (infoAgentes[agente][doc]["TOTAL"] || 0) + 1;
                    });
                };

            };

            for (let j = linhasParaExcluir.length - 1; j >= 0; j--) {
                ss.getSheetByName(aba).deleteRow(linhasParaExcluir[j]);
            };
        });

    };

    function tratarAbasFinalizados() {

        abaFinalizados.forEach(aba => {
            let dados = ss.getSheetByName(aba).getDataRange().getValues();

            if (dados && dados.length > 0) {
                const cabecalho = dados[0];

                indiceProcesso = obterIndiceCabecalho(cabecalho, "PROCESSO");
                indiceFatal = obterIndiceCabecalho(cabecalho, "FATAL");
                indiceTipo = obterIndiceCabecalho(cabecalho, "TIPO");
                indiceCliente = obterIndiceCabecalho(cabecalho, "CLIENTE");
                indiceReu = obterIndiceCabecalho(cabecalho, "RÉU");
                indiceAgenteResponsavel = obterIndiceCabecalho(cabecalho, "AGENTE RESPONSÁVEL");
                indiceDataInicio = obterIndiceCabecalho(cabecalho, "DATA DO INÍCIO DA TRATATIVA");
                indiceDocSolicitado = obterIndiceCabecalho(cabecalho, "DOC. SOLICITADO");
                indiceDespacho = obterIndiceCabecalho(cabecalho, "DESPACHO");
                indiceStatus = obterIndiceCabecalho(cabecalho, "STATUS");
                indiceProcedimentoAdvogado = obterIndiceCabecalho(cabecalho, "PROCEDIMENTO DO AVOGADO");
                indiceTrigger = cabecalho.indexOf('TRIGGER')
            };

            for (i = 1; i < dados.length; i++) {
                let fatal = formatarData(dados[i][indiceFatal]);
                let tipo = dados[i][indiceTipo].trim().toUpperCase();
                let agente = dados[i][indiceAgenteResponsavel].trim().toUpperCase();
                let docs = dados[i][indiceDocSolicitado];
                let status = dados[i][indiceStatus].trim().toUpperCase();
                let cliente = dados[i][indiceCliente].trim().toUpperCase();
                let trigger = formatarData(dados[i][indiceTrigger]);

                if (!agente || !cliente) continue;
                if (!infoAgentes[agente]) infoAgentes[agente] = {};

                // fazer uma condição caso tenha trigger e a data do trigger for igual a data de hoje
                if (trigger) {
                    if (isThisMonth(trigger)) {

                        if (status) {

                            if (statusLiberados.includes(status)) {
                                infoAgentes[agente]["FINALIZADOS"] = (infoAgentes[agente]["FINALIZADOS"] || 0) + 1;

                                if (trigger) {

                                    if (trigger.toISOString() == hoje.toISOString()) {
                                        infoAgentes[agente]["LIBERADOS HOJE"] = (infoAgentes[agente]["LIBERADOS HOJE"] || 0) + 1;
                                    };

                                };

                            } else if (statusDesistencias.includes(status)) {
                                infoAgentes[agente]["DESISTÊNCIA"] = (infoAgentes[agente]["DESISTÊNCIA"] || 0) + 1;

                            } else if (statusDevolvidos.includes(status)) {
                                infoAgentes[agente]["DEVOLVIDOSOSC"] = (infoAgentes[agente]["DEVOLVIDOSOSC"] || 0) + 1;

                            }

                        };

                        if (tipo) {
                            if (!infoAgentes[agente][tipo]) infoAgentes[agente][tipo] = {};
                            infoAgentes[agente][tipo]["TOTAL"] = (infoAgentes[agente][tipo]["TOTAL"] || 0) + 1;

                            if (statusLiberados.includes(status)) {
                                infoAgentes[agente][tipo]["FINALIZADOS"] = (infoAgentes[agente][tipo]["FINALIZADOS"] || 0) + 1;

                            } else if (statusDesistencias.includes(status)) {
                                infoAgentes[agente][tipo]["DESISTÊNCIAS"] = (infoAgentes[agente][tipo]["DESISTÊNCIAS"] || 0) + 1;

                            };
                        };

                        if (docs) {
                            listaDocs = docs.split(",")
                            listaDocs.forEach(doc => {
                                if (!infoAgentes[agente][doc]) infoAgentes[agente][doc] = {};
                                infoAgentes[agente][doc]["TOTAL"] = (infoAgentes[agente][doc]["TOTAL"] || 0) + 1;

                                if (statusLiberados.includes(status)) {
                                    infoAgentes[agente][doc]["FINALIZADOS"] = (infoAgentes[agente][doc]["FINALIZADOS"] || 0) + 1;
                                };
                            });
                        };
                    };
                };

            };
        });

    };

    function tratarAbaDevolvidos() {
        let dados = ss.getSheetByName("CLIENTES DEVOLVIDOS PELO ADV").getDataRange().getValues();
        if (dados && dados.length > 0) {
            const cabecalho = dados[0];
            indiceCliente = obterIndiceCabecalho(cabecalho, "CLIENTE");
            indiceAgenteResponsavel = obterIndiceCabecalho(cabecalho, "AGENTE RESPONSÁVEL");
        };

        for (i = 1; i < dados.length; i++) {
            let agente = dados[i][indiceAgenteResponsavel].trim().toUpperCase();
            let cliente = dados[i][indiceCliente].trim().toUpperCase();

            if (!agente || !cliente) continue;
            if (!infoAgentes[agente]) infoAgentes[agente] = {};

            if (cliente) {
                infoAgentes[agente]["DEVOLVIDOS ADV"] = (infoAgentes[agente]["DEVOLVIDOS ADV"] || 0) + 1;
            };

        };
    };

    function tratarFinalizadosCorretos() {
        let dados = ss.getSheetByName("CLIENTES FINALIZADOS CORRETO").getDataRange().getValues();

        if (dados && dados.length > 0) {
            const cabecalho = dados[0];
            let indiceAgenteResponsavel = obterIndiceCabecalho(cabecalho, "AGENTE RESPONSÁVEL");
            let cliente = obterIndiceCabecalho(cabecalho, "CLIENTE");
            let indiceTrigger = obterIndiceCabecalho(cabecalho, "TRIGGER");
        };

        for (i = 1; i < dados.length; i++) {
            let agente = dados[i][indiceAgenteResponsavel].trim().toUpperCase();
            let trigger = formatarData(dados[i][indiceTrigger]);
            let cliente = dados[i][indiceCliente].trim().toUpperCase();

            if (!cliente) break;
            if (!agente) continue;
            if (!infoAgentes[agente]) infoAgentes[agente] = {}

            if (trigger) {
                if (trigger.toISOString() == hoje.toISOString()) {
                    infoAgentes[agente]["LIBERADOS HOJE"] = (infoAgentes[agente]["LIBERADOS HOJE"] || 0) + 1;
                };
            };
        };
    };

    function obterIndiceCabecalho(allCabecalho, nomeColuna) {
        return allCabecalho.indexOf(nomeColuna);
    };

    function formatarData(data) {
        if (data) return new Date(data.setHours(0, 0, 0, 0));
    };

    function exibirControle() {

        abaControle.getRange("B1:K55").clearContent();

        inserirTitulos();
        inserirCabecalho();

        inserirQuadroPrincipal();
        inserirQuadroDatas();
        inserirQuadroTipoPrazos();
        inserirQuadroTipoDocs();

    };

    function inserirQuadroPrincipal() {

        // QUADRO PRINCIAL
        let linhaInicial = 5;
        for (let iAgente in abasAgentes) {
            let agente = mapAgentes[abasAgentes[iAgente]];
            let dadosLinha = [
                agente,
                infoAgentes[agente]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["DESISTÊNCIA"] || 0,
                infoAgentes[agente]?.["OMISSO"] || 0,
                infoAgentes[agente]?.["SEM CONTATO"] || 0,
                infoAgentes[agente]?.["ENROLADO"] || 0,
                infoAgentes[agente]?.["EM ABERTO"] || 0,
                infoAgentes[agente]?.["TRATATIVA"] || 0,
                infoAgentes[agente]?.["DEVOLVIDOS ADV"] || 0,
                infoAgentes[agente]?.["VENCIDOS"] || 0,
            ];

            abaControle
                .getRange(linhaInicial, 2, 1, dadosLinha.length)
                .setBackground("#CEDAF2")
                .setValues([dadosLinha]);

            linhaInicial++;
        };

        // TOTAL QUADRO PRINCIPAL
        abaControle.getRange(linhaInicial, 2).setValue("TOTAL").setBackground("#0CF25D");
        for (let col = 3; col <= 11; col++) {
            const nomeColuna = obterNomeColuna(col);
            abaControle
                .getRange(linhaInicial, col)
                .setBackground("#0CF25D")
                .setFormula(`=SUM(${nomeColuna}5:${nomeColuna}${linhaInicial - 1})`);
        };
    };

    function inserirQuadroDatas() {
        // QUADRO DATAS
        let linhaAtrasados = 17;
        for (let iAgente in abasAgentes) {
            let agente = mapAgentes[abasAgentes[iAgente]];
            let dadosAtrasadosLinha = [
                agente,
                infoAgentes[agente]?.["TRATATIVAS HOJE"] || 0,
                infoAgentes[agente]?.["LIBERADOS HOJE"] || 0,
                infoAgentes[agente]?.["VENCIDOS ONTEM"] || 0,
                infoAgentes[agente]?.["VENCE HOJE"] || 0,
                infoAgentes[agente]?.["VENCE AMANHA"] || 0,
                infoAgentes[agente]?.["VENCE EM TRES"] || 0,
                infoAgentes[agente]?.["VENCE EM SETE"] || 0,
            ];

            abaControle
                .getRange(linhaAtrasados, 2, 1, dadosAtrasadosLinha.length)
                .setBackground("#CEDAF2")
                .setValues([dadosAtrasadosLinha]);

            linhaAtrasados++;
        };

        // TOTAL CONTROLE POR DATA
        abaControle.getRange(linhaAtrasados, 2).setValue("TOTAL").setBackground("#0CF25D");
        for (let col = 3; col <= 9; col++) {
            const nomeColuna = obterNomeColuna(col);
            abaControle
                .getRange(linhaAtrasados, col)
                .setBackground("#0CF25D")
                .setFormula(`=SUM(${nomeColuna}17:${nomeColuna}${linhaAtrasados - 1})`);
        };
    };

    function inserirQuadroTipoPrazos() {
        // QUADRO TIPO DE PRAZOS
        let linhaTiposPrazos = 30
        for (let iAgente in abasAgentes) {
            let agente = mapAgentes[abasAgentes[iAgente]];
            let dadosTiposPrazos = [
                agente,
                infoAgentes[agente]?.["SOLICITAÇÃO DE DOCUMENTOS"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["SOLICITAÇÃO DE DOCUMENTOS"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["SOLICITAÇÃO DE DOCUMENTOS"]?.["DESISTÊNCIAS"] || 0,
                infoAgentes[agente]?.["DILAÇÃO"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["DILAÇÃO"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["DILAÇÃO"]?.["DESISTÊNCIAS"] || 0,
                infoAgentes[agente]?.["DILAÇÃO DEFERIDA"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["DILAÇÃO DEFERIDA"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["DILAÇÃO DEFERIDA"]?.["DESISTÊNCIAS"] || 0,
            ];

            abaControle
                .getRange(linhaTiposPrazos, 2, 1, 10)
                .setBackground("#CEDAF2")
                .setValues([dadosTiposPrazos]);

            linhaTiposPrazos++;
        };

        // TOTAL CONTROLE TIPO DE PRAZOS
        abaControle.getRange(linhaTiposPrazos, 2).setValue("TOTAL").setBackground("#0CF25D");
        for (let col = 3; col <= 11; col++) {
            const nomeColuna = obterNomeColuna(col);
            abaControle
                .getRange(linhaTiposPrazos, col)
                .setBackground("#0CF25D")
                .setFormula(`=SUM(${nomeColuna}30:${nomeColuna}${linhaTiposPrazos - 1})`);
        };
    };

    function inserirQuadroTipoDocs() {

        let linhaTipoDocs = 43
        for (let iAgente in abasAgentes) {
            let agente = mapAgentes[abasAgentes[iAgente]];
            let dadosTiposDocs = [
                agente,
                infoAgentes[agente]?.["JUSTIÇA GRATUITA"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["JUSTIÇA GRATUITA"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["PROCURAÇÃO"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["PROCURAÇÃO"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["PROCURAÇÃO COM FIRMA"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["PROCURAÇÃO COM FIRMA"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["COMP. RESIDÊNCIA"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["COMP. RESIDÊNCIA"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["PEDIDO ADM"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["PEDIDO ADM"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["COMPARECIMENTO"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["COMPARECIMENTO"]?.["FINALIZADOS"] || 0,
                infoAgentes[agente]?.["DOC. IDENTIFICAÇÃO"]?.["TOTAL"] || 0,
                infoAgentes[agente]?.["DOC. IDENTIFICAÇÃO"]?.["FINALIZADOS"] || 0,
            ];

            abaControle
                .getRange(linhaTipoDocs, 2, 1, 15)
                .setBackground("#CEDAF2")
                .setValues([dadosTiposDocs]);

            linhaTipoDocs++;
        };

        abaControle.getRange(linhaTipoDocs, 2).setValue("TOTAL").setBackground("#0CF25D");
        for (let col = 3; col <= 16; col++) {
            const nomeColuna = obterNomeColuna(col);
            abaControle
                .getRange(linhaTipoDocs, col)
                .setBackground("#0CF25D")
                .setFormula(`=SUM(${nomeColuna}43:${nomeColuna}${linhaTipoDocs - 1})`);
        };

    };

    function inserirTitulos() {
        abaControle
            .getRange("B2:K3")
            .merge()
            .setValue("QUADRO MENSAL DOS PRAZOS DA SOLICITAÇÃO")
            .setFontWeight("bold")
            .setBackground("#668CD9")
            .setHorizontalAlignment("center");

        abaControle
            .getRange("B14:I15")
            .merge()
            .setValue("CONTROLE DOS PRAZOS POR DATA")
            .setFontWeight("bold")
            .setBackground("#668CD9")
            .setHorizontalAlignment("center");

        abaControle
            .getRange("B26:K27")
            .merge()
            .setValue("CONTROLE POR TIPO DE PRAZO")
            .setFontWeight("bold")
            .setBackground("#668CD9")
            .setHorizontalAlignment("center");

        abaControle
            .getRange("B39:P40")
            .merge()
            .setValue("CONTROLE POR TIPO DE DOCUMENTO")
            .setFontWeight("bold")
            .setBackground("#668CD9")
            .setHorizontalAlignment("center");
    };

    function inserirCabecalho() {
        const cabecalhoQuadroPrincipal = [
            "AGENTES",
            "PRAZOS LIBERADOS",
            "DESISTÊNCIAS",
            "OMISSOS",
            "SEM CONTATO",
            "ENROLADOS",
            "PRAZOS EM ABERTO",
            "PRAZOS EM TRATATIVA",
            "DEVOLVIDOS PELO ADV",
            "PRAZOS VENCIDOS",
        ];

        const cabecalhoQuadroAtrasados = [
            "AGENTES",
            "TRATATIVAS HOJE",
            "LIBERADOS HOJE",
            "VENCIDOS ONTEM",
            "FATAL HOJE",
            "FATAL AMANHA",
            "VENCE EM 3 DIAS",
            "VENCE EM 7 DIAS",
        ];

        const cabecalhoTipoPrazo = [
            "QUANTIDADE",
            "LIBERADOS",
            "DESISTÊNCIAS"
        ];

        const cabecalhoTipoDoc = [
            "QUANTIDADE",
            "LIBERADOS",
        ];

        abaControle
            .getRange("B4:K4")
            .setValues([cabecalhoQuadroPrincipal])
            .setBackground("#FFF")
            .setFontWeight("bold");

        abaControle
            .getRange("B16:I16")
            .setValues([cabecalhoQuadroAtrasados])
            .setBackground("#FFF")
            .setFontWeight("bold");

        abaControle
            .getRange("B28:B28")
            .setValue(".")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("C28:E28")
            .setValue("SOLICITAÇÃO DE DOCUMENTOS")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("F28:H28")
            .setValue("DILAÇÃO")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("I28:K28")
            .setValue("DILAÇÃO DEFERIDA")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");

        abaControle
            .getRange("B29:B29")
            .setValue("AGENTE")
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("C29:E29")
            .setValues([cabecalhoTipoPrazo])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("F29:H29")
            .setValues([cabecalhoTipoPrazo])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("I29:K29")
            .setValues([cabecalhoTipoPrazo])
            .setBackground("#FFF")
            .setFontWeight("bold");

        abaControle
            .getRange("B41:B41")
            .setValue(".")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("C41:D41")
            .setValue("JUSTIÇA GRATUITA")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("E41:F41")
            .setValue("PROCURAÇÃO")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("G41:H41")
            .setValue("PROCURAÇÃO COM FIRMA")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("I41:J41")
            .setValue("COMP. RESIDÊNCIA")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("K41:L41")
            .setValue("PEDIDO ADM")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("M41:N41")
            .setValue("COMPARECIMENTO")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");
        abaControle
            .getRange("O41:P41")
            .setValue("DOC. IDENTIFICAÇÃO")
            .setBackground("#CEDAF2")
            .setFontWeight("bold");

        abaControle
            .getRange("B42:B42")
            .setValue("AGENTE")
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("C42:D42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("E42:F42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("G42:H42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("I42:J42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("K42:L42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("M42:N42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");
        abaControle
            .getRange("O42:P42")
            .setValues([cabecalhoTipoDoc])
            .setBackground("#FFF")
            .setFontWeight("bold");


    };

    function obterNomeColuna(indice) {
        let coluna = "";
        while (indice > 0) {
            let resto = (indice - 1) % 26;
            coluna = String.fromCharCode(65 + resto) + coluna;
            indice = Math.floor((indice - 1) / 26);
        }
        return coluna;
    };

    function atividadesDoMes() {
        let startingColumn = 19; // Coluna L
        let index = 0;

        abaControle.getRange(1, 18, 2, 8)
            .setValue("ATIVIDADE MENSAL DA SOLICITAÇÃO")
            .setBackground("#668cd9");

        // Escreve a data atual na planilha
        abaControle.getRange(3, 18).setValue("DIAS ÚTEIS");
        abaControle.getRange(4, 18).setValue(hoje).setBackground('#cedaf2');


        for (let iAgente in abasAgentes) {
            let agente = mapAgentes[abasAgentes[iAgente]];
            abaControle.getRange(3, startingColumn + index).setValue(agente).setBackground('#fff');
            abaControle.getRange(4, startingColumn + index).setValue(infoAgentes[agente]?.["TRATATIVAS HOJE"] || 0).setBackground('#cedaf2');
            index++;
        };

    };

    function limparOmisso() {
        let abaOmisso = ss.getSheetByName("OMISSO");
        let dadosOmisso = abaOmisso.getDataRange().getValues();

        let linhasParaExcluir = [];
        let indiceProcesso = dadosOmisso[0].indexOf("PROCESSO");


        for (let i = 1; i < dadosOmisso.length; i++) {
            let processo = dadosOmisso[i][indiceProcesso];

            if (!omissosGeral.includes(processo)) {
                linhasParaExcluir.push(i + 1);
            };
        };

        for (let j = linhasParaExcluir.length - 1; j >= 0; j--) {
            abaOmisso.deleteRow(linhasParaExcluir[j]);
        };

    };

    function limpasSContato() {
        let abaSContato = ss.getSheetByName("SEM CONTATO");
        let dadosSContato = abaSContato.getDataRange().getValues();

        let linhasParaExcluir = [];
        let indiceProcesso = dadosSContato[0].indexOf("PROCESSO");


        for (let i = 1; i < dadosSContato.length; i++) {
            let processo = dadosSContato[i][indiceProcesso];

            if (!sContatoGeral.includes(processo)) {
                linhasParaExcluir.push(i + 1);
            };
        };

        for (let j = linhasParaExcluir.length - 1; j >= 0; j--) {
            abaSContato.deleteRow(linhasParaExcluir[j]);
        };

    };

    function isThisMonth(date) {
        date = new Date(date);
        return (date.getMonth() == hoje.getMonth() && date.getFullYear() == hoje.getFullYear());
    };



    abaControle.getRange(1, 1).setValue("Ultima execução").setBackground('#0CF25S');
    abaControle.getRange(2, 1).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm")).setBackground('#0CF25S');


};

function funcionariosSol() {
    return ["x", "y", "z", "j", "k", "l"].sort();
};

function clearAndCopyAllData() {
    const current_spread_sheet = SpreadsheetApp.getActiveSpreadsheet();
    const activitiesSheet = current_spread_sheet.getSheetByName("Controle");

    // Define a linha e coluna inicial
    const startRow = 4; // Linha 4
    const startColumn = 18; // Coluna K

    const row4Values = activitiesSheet.getRange(startRow, startColumn, 1, activitiesSheet.getLastColumn() - startColumn + 1).getValues()[0];
    const hasDataInRow4 = row4Values.some(cell => cell !== "");
    if (!hasDataInRow4) {
        return;
    }

    // Obtém a última linha com dados
    const lastRow = activitiesSheet.getLastRow();

    // Verifica se há dados a partir da linha 4
    if (lastRow >= startRow) {
        // Obtém os dados a partir da linha 4
        const dataRange = activitiesSheet.getRange(
            startRow,
            startColumn,
            lastRow - startRow + 1,
            activitiesSheet.getLastColumn() - startColumn + 1
        );
        const dataValues = dataRange.getValues();

        // Copia os dados para a linha seguinte
        const destinationRange = activitiesSheet.getRange(
            startRow + 1,
            startColumn,
            dataValues.length,
            dataValues[0].length
        );
        destinationRange.setValues(dataValues);

        // Limpa os dados na linha 4
        activitiesSheet
            .getRange(startRow, startColumn, 1, dataValues[0].length)
            .clearContent();
    };
};


