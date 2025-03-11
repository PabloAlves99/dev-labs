function ControleBoasVindas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaControle = ss.getSheetByName("CONTROLE - BOAS VINDAS");
    const abaAcompanhamentoConv = ss.getSheetByName("ACOMPANHAMENTO - CONVERSÃO")
    const abaLogs = ss.getSheetByName("LOGS");

    let funcionarios = listaDeFuncionarios();
    let funcionariosPrescricao = ["JADE", "MARIA EDUARDA"];
    let compl = "BOAS-VINDAS ";

    let funcionariosInfo = {};
    let motivosPorFase = {};
    let logs = [];

    const clientesPerdidos = [
        "PAROU DE RESPONDER",
        "NÃO POSSUI INTERESSE",
        "NÃO VIÁVEL",
        "RESOLVEU",
        "COMARCA PROIBIDA",
        "SEM PROVAS",
    ];

    const cabecalhoProcon = [
        "FUNCIONÁRIO",
        "TOTAL DE CLIENTES",
        "DESCARTADO RELATO",
        "BOAS-VINDAS",
        "EM TRATATIVA",
        "VIÁVEIS",
        "NÃO VIÁVEIS",
        "DÚVIDA DA VIABILIDADE",
        "VIÁVEIS PERDIDOS",
        "VIÁVEIS EM TRATATIVA",
    ];
    const cabecalhoPrescricao = [
        "FUNCIONÁRIO",
        "TOTAL DE CLIENTES",
        "BOAS-VINDAS",
        "EM TRATATIVA",
        "VIÁVEIS",
    ];
    const cabecalhoLeads = [
        "FUNCIONÁRIOS",
        "ROBÔ/CONSUMIDOR",
        "PROCON",
        "INDICAÇÃO",
        "RECONVERSÃO",
        "CONSIGNADO",
        "COMPANHIA DE ÔNIBUS",
        "BANCOS",
    ];
    const cabecalhoNaoViavel = [
        "FUNCIONARIOS",
        "EMPRESA PEQUENA",
        "EMPRESA NÃO VIÁVEL",
        "RELATO NÃO VIÁVEL",
        "NÃO ATUAMOS",
    ];
    const cabecalhoFaseQueParou = [
        "FUNCIORARIOS",
        "NOVOS",
        "QUER PROSSEGUIR",
        "FALAR COM ATENDENTE",
        "PROVAS",
    ];
    const cabecalhoTotalFaseQueParou = [
        "MOTIVOS",
        "NOVOS",
        "QUER PROSSEGUIR",
        "FALAR COM ATENDENTE",
        "PROVAS",
    ];
    const cabecalhoStatus = [
        "FUNCIONÁRIO",
        "PAROU DE RESPONDER",
        "NÃO POSSUI INTERESSE",
        "NÃO VIÁVEL",
        "RESOLVEU",
        "COMARCA PROIBIDA"
    ];

    let cabecalhoClientesPerdidos = [
        "FUNCIONÁRIOS",
        "PAROU DE RESPONDER",
        "NÃO POSSUI INTERESSE",
        "NÃO VIÁVEL",
        "RESOLVEU",
        "COMARCA PROIBIDA",
        "SEM PROVAS",
    ];

    let linhaPrimeirosQuadros = 5;
    let linhaPrescricao = 22;

    processarAba();
    escreverLogsBoasVindas();
    escreverQuadrosAcompanhamento();
    escreverQuadroBoasVindas();
    escreverTotalPorLeads();
    escreverQuadroMotivoNVRelatao();
    escreverQuadroClientesPerdidos();
    escreverQuadroMotivoPorFase();
    escreverQuadroBoasVindasLeads();
    escreverHoraExec();



    function processarAba() {
        funcionarios.forEach((funcionario) => {
            let tabBoasVindas = ss
                .getSheetByName(compl + funcionario)
                .getDataRange()
                .getValues();

            if (!tabBoasVindas) {
                logs.push(`Aba não encontrada para ${funcionario}`);
                return;
            };

            if (!funcionariosInfo[funcionario]) funcionariosInfo[funcionario] = {};

            let cabecalho = tabBoasVindas[0];
            let indiceData = cabecalho.indexOf("DATA");
            let indiceNome = cabecalho.indexOf("NOME");
            let indicefaseQueParou = cabecalho.indexOf("FASE QUE PAROU");
            let indiceLeads = cabecalho.indexOf("LEADS");
            let indiceNViavelRelato = cabecalho.indexOf("MOTIVO NÃO VIÁVEL RELATO");
            let indiceDuvidaViavel = cabecalho.indexOf("DÚVIDA SOBRE VIABILIDADE");
            let indiceStatus = cabecalho.indexOf("STATUS");
            let indiceViavel = cabecalho.indexOf("VIÁVEL?");
            let indiceTrigger = cabecalho.indexOf("TRIGGER");


            for (let index = 1; index < tabBoasVindas.length; index++) {
                let data = tabBoasVindas[index][indiceData];
                let nome = tabBoasVindas[index][indiceNome];
                let leadType = tabBoasVindas[index][indiceLeads];
                let viavel = tabBoasVindas[index][indiceViavel];
                let faseQueParou = tabBoasVindas[index][indicefaseQueParou];
                let motivoNaoViavelRelato = tabBoasVindas[index][indiceNViavelRelato];
                let duvidaSobreViabilidade = tabBoasVindas[index][indiceDuvidaViavel];
                let status = tabBoasVindas[index][indiceStatus];
                let dataTrigger = tabBoasVindas[index][indiceTrigger];

                if (!nome) return;
                if (!data && (motivoNaoViavelRelato || status)) logs.push(`aba de ${funcionario} não possui DATA na linha ${index + 1}`);

                // QUADRO PRINCIPAL + QUADROS POR CAMPANHA

                if (leadType) {
                    if (!funcionariosInfo[funcionario][leadType])
                        funcionariosInfo[funcionario][leadType] = {};
                    if (!funcionariosInfo[funcionario]["GERAL"])
                        funcionariosInfo[funcionario]["GERAL"] = {};

                    funcionariosInfo[funcionario]["GERAL"]["CLIENTES"] =
                        (funcionariosInfo[funcionario]["GERAL"]["CLIENTES"] || 0) + 1;

                    funcionariosInfo[funcionario][leadType]["CLIENTES"] =
                        (funcionariosInfo[funcionario][leadType]["CLIENTES"] || 0) + 1;

                    if (motivoNaoViavelRelato !== "") {
                        funcionariosInfo[funcionario]["GERAL"]["DESCARTADO"] =
                            (funcionariosInfo[funcionario]["GERAL"]["DESCARTADO"] || 0) + 1;

                        funcionariosInfo[funcionario][leadType]["DESCARTADO"] =
                            (funcionariosInfo[funcionario][leadType]["DESCARTADO"] || 0) + 1;
                    };

                    if (data && !motivoNaoViavelRelato) {
                        funcionariosInfo[funcionario]["GERAL"]["BOAS-VINDAS"] =
                            (funcionariosInfo[funcionario]["GERAL"]["BOAS-VINDAS"] || 0) + 1;

                        funcionariosInfo[funcionario][leadType]["BOAS-VINDAS"] =
                            (funcionariosInfo[funcionario][leadType]["BOAS-VINDAS"] || 0) + 1;
                    };

                    if (Boolean(viavel)) {
                        funcionariosInfo[funcionario]["GERAL"]["VIAVEL"] =
                            (funcionariosInfo[funcionario]["GERAL"]["VIAVEL"] || 0) + 1;

                        funcionariosInfo[funcionario][leadType]["VIAVEL"] =
                            (funcionariosInfo[funcionario][leadType]["VIAVEL"] || 0) + 1;

                        if (clientesPerdidos.includes(status)) {
                            funcionariosInfo[funcionario]["GERAL"]["VIAVEIS PERDIDOS"] =
                                (funcionariosInfo[funcionario]["GERAL"]["VIAVEIS PERDIDOS"] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["VIAVEIS PERDIDOS"] =
                                (funcionariosInfo[funcionario][leadType]["VIAVEIS PERDIDOS"] || 0) + 1;
                        };
                    };

                    if (clientesPerdidos.includes(status)) {
                        funcionariosInfo[funcionario]["GERAL"]["CLIENTES PERDIDOS"] =
                            (funcionariosInfo[funcionario]["GERAL"]["CLIENTES PERDIDOS"] ||
                                0) + 1;

                        funcionariosInfo[funcionario][leadType]["CLIENTES PERDIDOS"] =
                            (funcionariosInfo[funcionario][leadType]["CLIENTES PERDIDOS"] ||
                                0) + 1;

                        funcionariosInfo[funcionario]["GERAL"][status] =
                            (funcionariosInfo[funcionario]["GERAL"][status] || 0) + 1;

                        funcionariosInfo[funcionario][leadType][status] =
                            (funcionariosInfo[funcionario][leadType][status] || 0) + 1;
                    };

                    if (status == "EM TRATATIVA") {
                        funcionariosInfo[funcionario]["GERAL"]["TRATATIVAS"] =
                            (funcionariosInfo[funcionario]["GERAL"]["TRATATIVAS"] || 0) + 1;

                        funcionariosInfo[funcionario][leadType]["TRATATIVAS"] =
                            (funcionariosInfo[funcionario][leadType]["TRATATIVAS"] || 0) + 1;

                        if (Boolean(viavel)) {
                            funcionariosInfo[funcionario]["GERAL"]["VIAVEIS EM TRATATIVA"] =
                                (funcionariosInfo[funcionario]["GERAL"][
                                    "VIAVEIS EM TRATATIVA"
                                ] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["VIAVEIS EM TRATATIVA"] =
                                (funcionariosInfo[funcionario][leadType][
                                    "VIAVEIS EM TRATATIVA"
                                ] || 0) + 1;
                        };
                    } else if (status == "DOCUMENTAÇÃO") {
                        funcionariosInfo[funcionario]["GERAL"]["CONVERTIDO"] =
                            (funcionariosInfo[funcionario]["GERAL"]["CONVERTIDO"] || 0) + 1;

                        funcionariosInfo[funcionario][leadType]["CONVERTIDO"] =
                            (funcionariosInfo[funcionario][leadType]["CONVERTIDO"] || 0) + 1;
                    };

                    if (Boolean(duvidaSobreViabilidade)) {
                        funcionariosInfo[funcionario]["GERAL"]["DUVIDA SOBRE VIABILIDADE"] =
                            (funcionariosInfo[funcionario]["GERAL"][
                                "DUVIDA SOBRE VIABILIDADE"
                            ] || 0) + 1;

                        funcionariosInfo[funcionario][leadType][
                            "DUVIDA SOBRE VIABILIDADE"
                        ] =
                            (funcionariosInfo[funcionario][leadType][
                                "DUVIDA SOBRE VIABILIDADE"
                            ] || 0) + 1;
                    };
                } else {
                    if (data instanceof Date) {
                        logs.push(`aba de ${funcionario} não possui LEAD na linha ${index + 1}`);
                    };
                };

                // FIM DO QUADRO PRINCIPAL + QUADROS POR CAMPANHA

                if (motivoNaoViavelRelato) {
                    if (!funcionariosInfo[funcionario]["GERAL"])
                        funcionariosInfo[funcionario]["GERAL"] = {};
                    if (!funcionariosInfo[funcionario]["GERAL"]["MOTIVO NAO VIAVEL RELATO"])
                        funcionariosInfo[funcionario]["GERAL"]["MOTIVO NAO VIAVEL RELATO"] = {};

                    funcionariosInfo[funcionario]["GERAL"]["MOTIVO NAO VIAVEL RELATO"][motivoNaoViavelRelato] =
                        (funcionariosInfo[funcionario]["GERAL"]["MOTIVO NAO VIAVEL RELATO"][motivoNaoViavelRelato] || 0) + 1;

                    if (faseQueParou) {
                        if (!motivosPorFase[faseQueParou])
                            motivosPorFase[faseQueParou] = {};

                        motivosPorFase[faseQueParou][motivoNaoViavelRelato] =
                            (motivosPorFase[faseQueParou][motivoNaoViavelRelato] || 0) + 1;
                    };
                };

                if (faseQueParou) {
                    if (!funcionariosInfo[funcionario]["GERAL"])
                        funcionariosInfo[funcionario]["GERAL"] = {};
                    if (!funcionariosInfo[funcionario]["GERAL"]["FASE QUE PAROU"])
                        funcionariosInfo[funcionario]["GERAL"]["FASE QUE PAROU"] = {};

                    funcionariosInfo[funcionario]["GERAL"]["FASE QUE PAROU"][faseQueParou] =
                        (funcionariosInfo[funcionario]["GERAL"]["FASE QUE PAROU"][faseQueParou] || 0) + 1;

                };

                if (Boolean(viavel) && motivoNaoViavelRelato) {
                    logs.push(
                        `Descartado relato e viavel: ${funcionario}: linha ${index + 1} `
                    );
                };


            };

        });
    };

    function escreverQuadrosAcompanhamento() {
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

        if (!abaAcompanhamentoConv) {
            Logger.log("ACOMPANHAMENTO - CONVERSÃO' não foi encontrada.");
            return;
        };
        listaLeads.forEach(lead => {

            funcionarios.forEach(funcionario => {
                let boasVindas = funcionariosInfo[funcionario]?.[lead]?.['BOAS-VINDAS'] || 0;
                let viaveis = funcionariosInfo[funcionario]?.[lead]?.['VIAVEL'] || 0;


                abaAcompanhamentoConv.getRange(linhaInicial, 3).setValue(boasVindas);
                abaAcompanhamentoConv.getRange(linhaInicial, 4).setValue(viaveis);

                linhaInicial++;
            });

            linhaInicial += 4;
        });
    };

    function escreverQuadroBoasVindas() {
        let linha = linhaPrimeirosQuadros;
        abaControle
            .getRange("B2:K3")
            .merge()
            .setValue("CONTROLE NÃO PRESCRIÇÃO - BOAS-VINDAS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 2, 1, cabecalhoProcon.length)
            .setValues([cabecalhoProcon])
            .setBackground("#fff");

        abaControle
            .getRange("B20:F20")
            .merge()
            .setValue("CONTROLE PRESCRIÇÃO - BOAS-VINDAS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(21, 2, 1, cabecalhoPrescricao.length)
            .setValues([cabecalhoPrescricao])
            .setBackground("#FFF");

        for (let funcionario in funcionariosInfo) {
            if (!funcionariosPrescricao.includes(funcionario)) {
                const dados = [
                    funcionario,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["CLIENTES"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["DESCARTADO"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["BOAS-VINDAS"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["TRATATIVAS"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["VIAVEL"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["CLIENTES PERDIDOS"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["DUVIDA SOBRE VIABILIDADE"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["VIAVEIS PERDIDOS"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["VIAVEIS EM TRATATIVA"] || 0,
                ];

                abaControle.getRange(linha, 2, 1, cabecalhoProcon.length).setValues([dados]).setBackground("#D9E5FF");;
                linha++;

            } else {

                const dados = [
                    funcionario,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["CLIENTES"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["BOAS-VINDAS"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["TRATATIVAS"] || 0,
                    funcionariosInfo[funcionario]?.["GERAL"]?.["VIAVEL"] || 0,
                ];

                abaControle.getRange(linhaPrescricao, 2, 1, cabecalhoPrescricao.length).setValues([dados]).setBackground("#D9E5FF");;
                linhaPrescricao++;
            };
        };
    };

    function escreverTotalPorLeads() {
        let linha = linhaPrimeirosQuadros;
        abaControle
            .getRange("O2:V3")
            .merge()
            .setValue("TOTAL DE LEADS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 15, 1, cabecalhoLeads.length)
            .setValues([cabecalhoLeads])
            .setBackground("#fff");

        for (funcionarios in funcionariosInfo) {
            const dados = [
                funcionarios,
                funcionariosInfo[funcionarios]?.["ROBÔ/CONSUMIDOR"]?.["CLIENTES"] || 0,
                funcionariosInfo[funcionarios]?.["PROCON"]?.["CLIENTES"] || 0,
                funcionariosInfo[funcionarios]?.["INDICAÇÃO"]?.["CLIENTES"] || 0,
                funcionariosInfo[funcionarios]?.["RECONVERSÃO"]?.["CLIENTES"] || 0,
                funcionariosInfo[funcionarios]?.["CONSIGNADO"]?.["CLIENTES"] || 0,
                funcionariosInfo[funcionarios]?.["COMPANHIA DE ÔNIBUS"]?.["CLIENTES"] || 0,
                funcionariosInfo[funcionarios]?.["BANCOS"]?.["CLIENTES"] || 0,
            ];

            abaControle.getRange(linha, 15, 1, cabecalhoLeads.length).setValues([dados]).setBackground("#D9E5FF");;
            linha++;
        };
    };

    function escreverQuadroMotivoNVRelatao() {
        let linha = linhaPrimeirosQuadros;
        abaControle
            .getRange("X2:AB3")
            .merge()
            .setValue("MOTIVO NÃO VIAVEIS RELATO")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 24, 1, cabecalhoNaoViavel.length)
            .setValues([cabecalhoNaoViavel])
            .setBackground("#fff");

        for (funcionarios in funcionariosInfo) {
            const dados = [
                funcionarios,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["MOTIVO NAO VIAVEL RELATO"]?.["EMPRESA PEQUENA"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["MOTIVO NAO VIAVEL RELATO"]?.["EMPRESA NÃO VIÁVEL"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["MOTIVO NAO VIAVEL RELATO"]?.["RELATO NÃO VIÁVEL"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["MOTIVO NAO VIAVEL RELATO"]?.["NÃO ATUAMOS"] || 0,
            ];

            abaControle.getRange(linha, 24, 1, cabecalhoNaoViavel.length).setValues([dados]).setBackground("#D9E5FF");
            linha++;
        };

    };

    function escreverQuadroClientesPerdidos() {
        let linha = linhaPrimeirosQuadros;
        abaControle
            .getRange("AD2:AJ3")
            .merge()
            .setValue("MOTIVOS DE CLIENTES PERDIDOS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 30, 1, cabecalhoClientesPerdidos.length)
            .setValues([cabecalhoClientesPerdidos])
            .setBackground("#fff");

        for (funcionarios in funcionariosInfo) {
            const dados = [
                funcionarios,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["PAROU DE RESPONDER"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["NÃO POSSUI INTERESSE"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["NÃO VIÁVEL"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["RESOLVEU"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["COMARCA PROIBIDA"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["SEM PROVAS"] || 0,
            ];

            abaControle.getRange(linha, 30, 1, cabecalhoClientesPerdidos.length).setValues([dados]).setBackground("#D9E5FF");
            linha++;
        };
    };

    function escreverQuadroMotivoPorFase() {
        let linha = linhaPrimeirosQuadros;
        abaControle
            .getRange("AL2:AP3")
            .merge()
            .setValue("FASE QUE PAROU")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 38, 1, cabecalhoFaseQueParou.length)
            .setValues([cabecalhoFaseQueParou])
            .setBackground("#fff");

        for (funcionarios in funcionariosInfo) {
            const dados = [
                funcionarios,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["FASE QUE PAROU"]?.["NOVOS"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["FASE QUE PAROU"]?.["QUER PROSSEGUIR"] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["FASE QUE PAROU"]?.["FALAR COM ATENDENTE "] || 0,
                funcionariosInfo[funcionarios]?.["GERAL"]?.["FASE QUE PAROU"]?.["PROVAS"] || 0,
            ];

            abaControle.getRange(linha, 38, 1, cabecalhoFaseQueParou.length).setValues([dados]).setBackground("#D9E5FF");
            linha++;
        };

    };

    function escreverQuadroBoasVindasLeads() {
        let linhaInicial = 29;
        let listaLeads = [
            'PROCON',
            'ROBÔ/CONSUMIDOR',
            'INDICAÇÃO',
            'RECONVERSÃO',
            'COMPANHIA DE ÔNIBUS',
            'CONSIGNADO',
            'BANCOS',
        ];

        listaLeads.forEach((lead, index) => {
            let t = linhaInicial + (index * (9 + funcionarios.length));
            let c = t + 1;
            let linha = c + 1;

            abaControle
                .getRange(`B${t}:K${t}`)
                .merge()
                .setValue(`CONTROLE BOAS-VINDAS: ${lead}`)
                .setHorizontalAlignment("center")
                .setBackground("#668CD9");
            abaControle
                .getRange(c, 2, 1, cabecalhoProcon.length)
                .setValues([cabecalhoProcon])
                .setBackground("#fff");

            for (let funcionario in funcionariosInfo) {
                const dados = [
                    funcionario,
                    funcionariosInfo[funcionario]?.[lead]?.["CLIENTES"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["DESCARTADO"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["BOAS-VINDAS"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["TRATATIVAS"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["VIAVEL"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["CLIENTES PERDIDOS"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["DUVIDA SOBRE VIABILIDADE"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["VIAVEIS PERDIDOS"] || 0,
                    funcionariosInfo[funcionario]?.[lead]?.["VIAVEIS EM TRATATIVA"] || 0,
                ];

                abaControle.getRange(linha, 2, 1, cabecalhoProcon.length).setValues([dados]).setBackground("#D9E5FF");;
                linha++;
            };

        });
    };

    function escreverLogsBoasVindas() {
        let linha = 3;
        let ultimaLinha = abaLogs.getLastRow();
        abaLogs.getRange(linha, 6, ultimaLinha, 1).clearContent();


        // for (i = linha; i <= 18; i++) {
        //     abaLogs.getRange(i, 2, 1, 3).merge();
        // };

        logs.forEach(log => {
            abaLogs.getRange(linha, 6).setValue(log).setBackground('#D9E5FF');
            linha++;
        });

    };

    function escreverHoraExec() {
        const horaAtual = Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "HH:mm"
        );


        abaLogs
            .getRange("E1")
            .setValue("ULTIMA EXECUÇÃO")
            .setBackground("#F06060");
        abaLogs.getRange("E2").setValue(horaAtual).setBackground("#FFF");
        abaControle
            .getRange("A1")
            .setValue("ULTIMA EXECUÇÃO")
            .setBackground("#668CD9");
        abaControle.getRange("A2").setValue(horaAtual).setBackground("#FFF");
    };
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

