function montarEstruturaFinalComFormulasCorretas() {
  const TEMPO_LIMITE = 290000;
  const INICIO = Date.now();
  //const pastaRaizId = "Adicionar url da cota";
  const pastaRaiz = DriveApp.getFolderById(pastaRaizId);
  const properties = PropertiesService.getScriptProperties();

  const modeloFile = pastaRaiz.getFilesByName("MODELO_00");
  if (!modeloFile.hasNext()) {
    Logger.log("❌ MODELO_00 não encontrada.");
    return;
  }
  const modeloId = modeloFile.next().getId();
  const planilhaModelo = SpreadsheetApp.openById(modeloId);
  const abaOriginal = planilhaModelo.getSheetByName("ANALISTA_00");

  const totalColunas = abaOriginal.getLastColumn();
  const totalLinhas = abaOriginal.getLastRow();

  const dados = abaOriginal.getRange(3, 1, totalLinhas - 2, totalColunas).getValues();
  const formulas = abaOriginal.getRange(3, 1, totalLinhas - 2, totalColunas).getFormulasR1C1();

  const estrutura = {};

  dados.forEach((linha, idx) => {
    const valorA = linha[0];
    const linguagem = linha[2];
    const gt = linha[6];
    if (!linguagem || !gt || valorA === "") return;

    const chave = `${linguagem}|${gt}`;
    if (!estrutura[chave]) estrutura[chave] = [];

    estrutura[chave].push({
      valores: linha,
      formulas: formulas[idx],
      linhaOriginal: idx + 3 // linha real da aba original
    });
  });

  const chaves = Object.keys(estrutura);
  const totalGTs = chaves.length;
  let idxInicio = parseInt(properties.getProperty("IDX_INICIO") || "0");

  Logger.log(`🚀 Iniciando processamento dos GTs... (Total: ${totalGTs})`);
  Logger.log(`🔁 Retomando do índice ${idxInicio}`);

  for (let i = idxInicio; i < chaves.length; i++) {
    const chave = chaves[i];
    const [linguagem, gt] = chave.split("|");
    const linhasValidas = estrutura[chave];

    Logger.log(`🛠️ Processando (${i + 1}/${totalGTs}): ${linguagem} | ${gt} - ${linhasValidas.length} linha(s)`);

    let pastaLinguagem = pastaRaiz.getFoldersByName(linguagem);
    pastaLinguagem = pastaLinguagem.hasNext() ? pastaLinguagem.next() : pastaRaiz.createFolder(linguagem);

    let pastaGT = pastaLinguagem.getFoldersByName(gt);
    pastaGT = pastaGT.hasNext() ? pastaGT.next() : pastaLinguagem.createFolder(gt);

    const nomesAnalistas = ["ANALISTA_01", "ANALISTA_02", "ANALISTA_03"];
    const pastasAnalistas = nomesAnalistas.map(nome => {
      let p = pastaGT.getFoldersByName(nome);
      return p.hasNext() ? p.next() : pastaGT.createFolder(nome);
    });

    const copiaTemp = DriveApp.getFileById(modeloId).makeCopy(`TEMP_${linguagem}_${gt}`, pastaRaiz);
    const planilhaTemp = SpreadsheetApp.openById(copiaTemp.getId());
    const abaTemp = planilhaTemp.getSheetByName("ANALISTA_00");
    abaTemp.clearContents();

    const titulo = abaOriginal.getRange(2, 1, 1, totalColunas).getValues();
    abaTemp.getRange(2, 1, 1, totalColunas).setValues(titulo);

    let linhaDestino = 3;
    linhasValidas.forEach(({ valores, formulas, linhaOriginal }) => {
      for (let col = 0; col < totalColunas; col++) {
        const celula = abaTemp.getRange(linhaDestino, col + 1);

        if (col === 9) { // Coluna J
          const validacaoOrigem = abaOriginal.getRange(linhaOriginal, col + 1).getDataValidation();
          if (validacaoOrigem) {
            celula.setDataValidation(validacaoOrigem);
          }
        }

        if (formulas[col]) {
          celula.setFormulaR1C1(formulas[col]);
        } else {
          celula.setValue(valores[col]);
        }
      }
      linhaDestino++;
    });

    const ultimaLinhaAtual = abaTemp.getMaxRows();
    for (let r = ultimaLinhaAtual; r >= 3; r--) {
      const valA = abaTemp.getRange(r, 1).getDisplayValue().toString().trim();
      if (valA === "") abaTemp.deleteRow(r);
    }

    pastasAnalistas.forEach((pasta, idx) => {
      const nomeAnalista = nomesAnalistas[idx];
      const nomeArquivo = `${nomeAnalista}-${gt}`;
      const copiaFinal = DriveApp.getFileById(planilhaTemp.getId()).makeCopy(nomeArquivo, pasta);
      const planilhaFinal = SpreadsheetApp.openById(copiaFinal.getId());
      const abaFinal = planilhaFinal.getSheetByName("ANALISTA_00");
      if (abaFinal) abaFinal.setName(nomeAnalista);
    });

    DriveApp.getFileById(planilhaTemp.getId()).setTrashed(true);
    Logger.log(`✅ Finalizado: ${linguagem} | ${gt}`);

    properties.setProperty("IDX_INICIO", (i + 1).toString());

    // ⏱️ Verifica o tempo e agenda continuação se necessário
    if (Date.now() - INICIO >= TEMPO_LIMITE) {
      Logger.log("⏳ Tempo quase esgotado! Agendando continuação em 10 segundos...");

      // 🧹 Remove triggers antigos deste script
      ScriptApp.getProjectTriggers()
        .filter(t => t.getHandlerFunction() === "montarEstruturaFinalComFormulasCorretas")
        .forEach(trigger => ScriptApp.deleteTrigger(trigger));

      // ⏳ Cria novo trigger
      ScriptApp.newTrigger("montarEstruturaFinalComFormulasCorretas")
               .timeBased()
               .after(10 * 1000)
               .create();

      return;
    }
  }

  properties.deleteAllProperties();
  Logger.log("🎯 Todos os GTs foram processados com sucesso!");
}
