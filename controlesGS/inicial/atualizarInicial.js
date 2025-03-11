function atualizarInicial() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let abaInicial = ss.getSheetByName("INICIAL");
  let dadosInicial = abaInicial.getDataRange().getValues();
  let abaOdilene = ss.getSheetByName("Odilene").getDataRange().getValues();
  let abaJulia = ss.getSheetByName("Julia").getDataRange().getValues();

  if (!dadosInicial || !abaOdilene || !abaJulia) {
    Logger.log("❌ Verifique se todas as abas estão corretas.");
    return;
  }

  let abas = [
    { nome: "Cópia de Odilene", dados: abaOdilene, nomePessoa: "ODILENE" },
    { nome: "Cópia de Julia", dados: abaJulia, nomePessoa: "JULIA" }
  ];

  let atualizacoes = [];
  let naoEncontrados = [];
  let totalAtualizados = 0;
  let totalNaoEncontrados = 0;

  abas.forEach(abaObj => {
    let dados = abaObj.dados;

    Logger.log(`\n🔎 --- Verificando ${abaObj.nome} ---`);

    for (let i = 1; i < dados.length; i++) {
      let cliente = limparTexto(dados[i][1]); // Coluna B (Cliente)
      let reu = limparTexto(dados[i][2]); // Coluna C (Réu)
      let status = dados[i][5]; // Coluna F (Status)

      if (cliente && reu && status) {
        let encontrado = false;

        for (let j = 1; j < dadosInicial.length; j++) {
          let clienteInicial = limparTexto(dadosInicial[j][10]); // Coluna K (Cliente)
          let reuInicial = limparTexto(dadosInicial[j][11]); // Coluna L (Réu)

          if (clienteInicial.includes(cliente) && reuInicial.includes(reu)) {
            // Adiciona a atualização com o nome da pessoa que atualizou e o status
            atualizacoes.push({ linha: j + 1, status: status, nome: abaObj.nomePessoa });
            Logger.log(`✅ Atualizando linha ${j + 1} da aba INICIAL com status '${status}' e nome '${abaObj.nomePessoa}'.`);
            totalAtualizados++;
            encontrado = true;
            break; // Para evitar múltiplas atualizações na mesma linha
          }
        }

        if (!encontrado) {
          naoEncontrados.push({ linha: i + 1, cliente, reu, aba: abaObj.nome });
          totalNaoEncontrados++;
        }
      }
    }
  });

  // Aplicando as atualizações na aba INICIAL
  if (atualizacoes.length > 0) {
    let rangeO = abaInicial.getRange(2, 15, dadosInicial.length - 1, 1); // Coluna O (índice 15 = O)
    let valoresO = rangeO.getValues();

    let rangeA = abaInicial.getRange(2, 1, dadosInicial.length - 1, 1); // Coluna A (índice 1 = A)
    let valoresA = rangeA.getValues();

    atualizacoes.forEach(atualizacao => {
      valoresO[atualizacao.linha - 2][0] = atualizacao.status; // Ajuste para índice correto
      valoresA[atualizacao.linha - 2][0] = atualizacao.nome; // Ajuste para o nome da pessoa que atualizou
    });

    rangeO.setValues(valoresO);
    rangeA.setValues(valoresA); // Atualiza os nomes na coluna A

    Logger.log(`✅ ${totalAtualizados} linhas foram atualizadas com sucesso!\n`);
  }

  // Exibindo as linhas que não foram encontradas
  if (naoEncontrados.length > 0) {
    Logger.log(`❌ ${totalNaoEncontrados} linhas NÃO foram encontradas na aba INICIAL:`);
    naoEncontrados.forEach(item => {
      Logger.log(`🚨 Linha ${item.linha} (${item.aba}): Cliente '${item.cliente}', Réu '${item.reu}'`);
    });
  }

  // Resumo final
  Logger.log(`\n📊 --- Resumo Final ---`);
  Logger.log(`✅ Total de linhas atualizadas: ${totalAtualizados}`);
  Logger.log(`❌ Total de linhas não encontradas: ${totalNaoEncontrados}`);
}

// 🔹 Função para limpar texto (remover espaços extras e padronizar letras)
function limparTexto(texto) {
  return texto ? texto.toString().trim().replace(/\s+/g, ' ').toLowerCase() : "";
}
