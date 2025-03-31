function processos_repetidos() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const _abas = funcionariosSol();
    const abas = _abas.filter(item => item != 'xx')
    const ignorar = [
        "LIBERADO (OK)",
        "LIBERADO (DOCS. PARCIAIS)",
        "LIBERADO (DOCS. MÍNIMOS)",
        "CONCLUÍDO",
        "TRANSFERIDO",
        "DESISTÊNCIA (COM DECLARAÇÃO)",
        "DESISTÊNCIA (SEM DECLARAÇÃO)",
        "DEVOLVIDO SEM CONTATO",
        "DEVOLVIDO POR OMISSÃO",
        "OMISSO / SEM CONTATO",
    ];

    let processos_map = {};

    abas.forEach(sheet_aba => {
        let aba = ss.getSheetByName(sheet_aba);
        if (!aba || aba.getLastRow() < 2) return;
        let ultima_linha = aba.getLastRow();

        const processos = aba
            .getRange(11, 1, ultima_linha - 10, 1)
            .getValues()
            .flat();
        const status = aba
            .getRange(11, 12, ultima_linha - 10, 1)
            .getValues()
            .flat();

        for (let i = 0; i < processos.length; i++) {
            let processo = processos[i];
            let status_atual = status[i];

            if (!processo || ignorar.includes(status_atual)) {
                continue;
            }

            if (processo in processos_map) {
                let linha_original = processos_map[processo].linha;

                Logger.log(
                    `Processo repetido encontrado: ${processo}. Constante nas abas: ${processos_map[processo].aba.getName()} e ${sheet_aba}`
                );
            } else {
                processos_map[processo] = {
                    aba: aba,
                    linha: 11 + i,
                };
            };
        };
    });

    Logger.log("Verificação concluída!");
}