function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('index');

  // Carrega Dados da aba "Dados"
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDados = planilha.getSheetByName('Dados');
  const data = sheetDados.getDataRange().getValues();
  htmlTemplate.dadosBanco = JSON.stringify(data);

  // Gera todos os endereços com tratamento especial para ruas 115–129
  const enderecos = [];

  // Gera endereços das ruas 100 a 129 e 200 a 204
    const faixasDeRuas = [
      { inicio: 100, fim: 129, ajustar: (rua) => (rua >= 115 ? rua - 100 : rua) },
      { inicio: 200, fim: 204, ajustar: (rua) => rua } // Não altera nada
    ];

    faixasDeRuas.forEach(faixa => {
      for (let rua = faixa.inicio; rua <= faixa.fim; rua++) {
        const ruaFormatada = faixa.ajustar(rua);
        for (let estante = 1; estante <= 46; estante++) {
          for (let prateleira = 1; prateleira <= 30; prateleira++) {
            const endereco = `${ruaFormatada}-0-${estante}-${prateleira}`;
            enderecos.push(endereco);
          }
        }
      }
    });


  htmlTemplate.todosEnderecos = JSON.stringify(enderecos);

  return htmlTemplate.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Auditoria de Estoque")
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
}


function salvarCodigosInseridos(dados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = planilha.getSheetByName('Rascunho');
  if (!sheet) throw new Error("Aba 'Rascunho' não encontrada.");

  const bancoSheet = planilha.getSheetByName('Dados');
  if (!bancoSheet) throw new Error("Aba 'Dados' não encontrada.");

  const bancoDados = bancoSheet.getRange(2, 1, bancoSheet.getLastRow() - 1, 5).getValues();

  const enderecoAtual = (dados && dados.length > 0) ? dados[0].endereco : null;
  if (!enderecoAtual) {
    throw new Error("Endereço não informado. Bipagem vazia e nenhum endereço foi enviado.");
  }

  // Mapas para armazenar os finais do sistema e auditados por CF
  const sistemaPorCF = new Map();
  const auditadoPorCF = new Map();
  const cfEnderecoMap = new Map();

  // Construir dados do sistema
  bancoDados.forEach(row => {
    const [cf, codigo, final, descricao, endereco] = row;
    if (endereco !== enderecoAtual) return;

    const chave = `${cf}|${endereco}`;
    if (!cfEnderecoMap.has(chave)) {
      cfEnderecoMap.set(chave, { qtdSistema: 0, qtdAuditada: 0 });
    }
    cfEnderecoMap.get(chave).qtdSistema++;

    if (!sistemaPorCF.has(cf)) sistemaPorCF.set(cf, new Set());
    sistemaPorCF.get(cf).add(final);
  });

  // Construir dados auditados
  dados.forEach(item => {
    const codigo = String(item.codigo).trim();
    const registro = bancoDados.find(row => String(row[1]) === codigo || String(row[0]) === codigo);
    if (!registro) return;

    const [cf, , final, , endereco] = registro;
    const chave = `${cf}|${endereco}`;
    if (cfEnderecoMap.has(chave)) {
      cfEnderecoMap.get(chave).qtdAuditada++;
    }

    if (!auditadoPorCF.has(cf)) auditadoPorCF.set(cf, new Set());
    auditadoPorCF.get(cf).add(final);
  });

  // Garante que todos os CFs do endereço atual estejam no mapa
  bancoDados.forEach(row => {
    const [cf, , final, , endereco] = row;
    const chave = `${cf}|${endereco}`;

    if (endereco === enderecoAtual && !cfEnderecoMap.has(chave)) {
      cfEnderecoMap.set(chave, { qtdSistema: 1, qtdAuditada: 0 });
      if (!sistemaPorCF.has(cf)) sistemaPorCF.set(cf, new Set());
      sistemaPorCF.get(cf).add(final);
    }
  });

  // Monta as linhas para salvar
  const linhasParaSalvar = Array.from(cfEnderecoMap.entries()).map(([chave, info]) => {
    const [cf, endereco] = chave.split('|');

    const finaisSistema = sistemaPorCF.get(cf) || new Set();
    const finaisAuditados = auditadoPorCF.get(cf) || new Set();

    // Detectar divergências reais
    const finaisDivergentes = [...finaisSistema].filter(f => !finaisAuditados.has(f))
      .concat([...finaisAuditados].filter(f => !finaisSistema.has(f)));

    const finaisStr = info.qtdAuditada !== info.qtdSistema ? finaisDivergentes.join(", ") : "";

    return [cf, endereco, info.qtdSistema, info.qtdAuditada, finaisStr];
  });

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([["CF", "ENDEREÇO", "QTD SISTEMA", "QTD AUDITADA", "FINAL (Divergente)"]]);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, linhasParaSalvar.length, 5).setValues(linhasParaSalvar);
  return `CFs salvos com sucesso: ${linhasParaSalvar.length}`;
}


function adicionarNovoRegistro(novoRegistro) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dados');

    if (!sheet) {
      throw new Error("Aba 'Dados' não encontrada!");
    }

    sheet.appendRow([
      novoRegistro.cf,
      novoRegistro.codigo,
      novoRegistro.final,
      novoRegistro.descricao,
      novoRegistro.endereco
    ]);

    SpreadsheetApp.flush();

    const novosDados = sheet.getDataRange().getValues();
    return JSON.stringify(novosDados);

  } catch (erro) {
    Logger.log("Erro em adicionarNovoRegistro: " + erro.message);
    throw erro;
  }
}

function SalvarDivergencia(dados) {
  const outraPlanilhaID = '1WGm_PdDJUb9Jf7YABuotMdU7XWv9ckyH4BgWzoxi2ws';
  const outraPlanilha = SpreadsheetApp.openById(outraPlanilhaID);
  const abaDestino = outraPlanilha.getSheetByName('2025');

  if (!abaDestino) throw new Error('Aba de destino não encontrada!');

  const novaLinha = [
    dados.Data,
    dados.CF,
    dados.Final,
    dados.CdMercadoria,
    dados.Descricao,
    dados.Qtd,
    dados.EndeAud,
    dados.Lista
  ];

  abaDestino.appendRow(novaLinha);
  return 'Divergência registrada com sucesso na outra planilha!';
}

function SalvarObservacao(dadosObs) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ITENS OBSERVAÇÃO');

  const novaLinha = [
    dadosObs.CdMercadoria,
    dadosObs.CF,
    dadosObs.Final,
    dadosObs.Descricao,
    dadosObs.Endereco,
    dadosObs.MotivoObs,
    dadosObs.Data
  ];

  planilha.appendRow(novaLinha);
  return 'Observação Salva';
}
