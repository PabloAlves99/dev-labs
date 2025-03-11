function nomes_repetidos() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abas = funcionariosSol();


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
        "OMISSO",
        "SEM CONTATO",
    ];

    let clientes_map = {};
    let duplicados = {};

    abas.forEach(sheet_aba => {
        let aba = ss.getSheetByName(sheet_aba);
        if (!aba || aba.getLastRow() < 2) return;

        let ultima_linha = aba.getLastRow();

        const clientes = aba
            .getRange(11, 6, ultima_linha - 10, 1)
            .getValues()
            .flat();
        const status = aba
            .getRange(11, 12, ultima_linha - 10, 1)
            .getValues()
            .flat();

        clientes.forEach((cliente, index) => {
            if (!cliente || ignorar.includes(status[index])) return;

            if (clientes_map[cliente]) {
                if (clientes_map[cliente] !== sheet_aba) {
                    duplicados[cliente] = duplicados[cliente] || new Set();
                    duplicados[cliente].add(clientes_map[cliente]);
                    duplicados[cliente].add(sheet_aba);
                }
            } else {
                // Registrar cliente e aba onde foi encontrado
                clientes_map[cliente] = sheet_aba;
            }
        });
    });

    if (Object.keys(duplicados).length > 0) {

        let i = 1;
        let resultado = "Clientes duplicados encontrados:\n";
        for (let cliente in duplicados) {
            resultado += `${i} - ${cliente}: ${[...duplicados[cliente]].join(", ")}\n`;
            i++
        }
        Logger.log(resultado);
    } else {
        Logger.log("Nenhum cliente em mais de uma aba encontrado.");
    }
};


