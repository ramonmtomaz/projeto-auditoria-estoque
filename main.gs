/**
 * Verifica se o usuário atual tem acesso ao sistema
 * @returns {Object} Objeto com status de acesso e informações do usuário
 */
function verificarAcessoUsuario() {
  try {
    const email = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '';
    
    // Lista de emails bloqueados
    const emailsBloqueados = [
      'prev.viana@kabum.com.br'
    ];
    
    console.log(`🔐 Verificando acesso para: ${email}`);
    
    if (emailsBloqueados.includes(email.toLowerCase())) {
      console.log(`❌ Acesso negado para: ${email}`);
      return {
        permitido: false,
        email: email,
        motivo: 'Email bloqueado'
      };
    }
    
    console.log(`✅ Acesso permitido para: ${email}`);
    return {
      permitido: true,
      email: email,
      motivo: 'Acesso autorizado'
    };
    
  } catch (error) {
    console.error('Erro ao verificar acesso do usuário:', error);
    return {
      permitido: false,
      email: 'desconhecido',
      motivo: 'Erro na verificação'
    };
  }
}

/**
 * Obtém informações agregadas de um CF em um endereço/tipo, com cache para acelerar chamadas repetidas.
 * Retorna { descricao, custo:number|0, qtd:number, finais:string[] } ou null.
 */
function obterInfoCfPorEndereco(endereco, tipoEstoque, cfAlvo) {
  const cache = CacheService.getDocumentCache();
  const key = `resumo:${String(tipoEstoque||'principal')}:${String(endereco).trim()}`;
  let resumo = null;
  try { resumo = JSON.parse(cache.get(key) || 'null'); } catch (_) { resumo = null; }

  if (!resumo) {
    // Construir resumo uma vez
    const linhas = buscarProdutosPorEndereco(endereco, tipoEstoque);
    if (!linhas || linhas.length <= 1) return null;
    const parseMoedaLocal = (v) => {
      const s0 = String(v == null ? '' : v).trim();
      if (!s0) return NaN;
      const s = s0.replace(/[R$\s]/g, '');
      const hasComma = s.includes(',');
      const hasDot = s.includes('.');
      if (hasComma && hasDot) {
        const lastComma = s.lastIndexOf(',');
        const lastDot = s.lastIndexOf('.');
        if (lastComma > lastDot) {
          const noThousands = s.replace(/\./g, '');
          const normalized = noThousands.replace(/,/g, '.');
          const n = parseFloat(normalized);
          return isNaN(n) ? NaN : n;
        } else {
          const noThousands = s.replace(/,/g, '');
          const n = parseFloat(noThousands);
          return isNaN(n) ? NaN : n;
        }
      }
      if (hasComma && !hasDot) {
        const noThousands = s.replace(/\./g, '');
        const normalized = noThousands.replace(/,/g, '.');
        const n = parseFloat(normalized);
        return isNaN(n) ? NaN : n;
      }
      if (hasDot && !hasComma) {
        const noThousands = s.replace(/,/g, '');
        const n = parseFloat(noThousands);
        return isNaN(n) ? NaN : n;
      }
      const n = parseFloat(s);
      return isNaN(n) ? NaN : n;
    };

    const mapa = {};
    linhas.slice(1).forEach(r => {
      const cf = String(r[0] || '').trim();
      if (!cf) return;
      const desc = String(r[3] || '').trim();
      const fin = String(r[2] || '').trim();
      const custoRaw = r.length > 5 ? r[5] : '';
      const custoNum = parseMoedaLocal(custoRaw);
      if (!mapa[cf]) mapa[cf] = { descricao: '', custo: 0, qtd: 0, finais: {} };
      const item = mapa[cf];
      if (!item.descricao && desc) item.descricao = desc;
      if (!isNaN(custoNum) && custoNum > 0 && item.custo === 0) item.custo = custoNum;
      item.qtd += 1;
      if (fin) item.finais[fin] = true;
    });

    // Compacta para salvar no cache
    const compacto = Object.keys(mapa).map(cf => ({
      cf,
      d: mapa[cf].descricao,
      c: mapa[cf].custo,
      q: mapa[cf].qtd,
      f: Object.keys(mapa[cf].finais)
    }));
    try {
      const json = JSON.stringify(compacto);
      if (json.length < 95000) cache.put(key, json, 900); // 15 minutos
    } catch (_) {}
    resumo = compacto;
  }

  if (!resumo) return null;
  const item = Array.isArray(resumo) ? resumo.find(x => String(x.cf || x.CF || x.c || '').trim() === String(cfAlvo).trim()) : null;
  if (!item) return null;
  return {
    descricao: item.d || item.descricao || '',
    custo: typeof item.c === 'number' ? item.c : Number(item.c || 0) || 0,
    qtd: typeof item.q === 'number' ? item.q : Number(item.q || 0) || 0,
    finais: Array.isArray(item.f) ? item.f : []
  };
}

// ==================== FUNÇÕES DE AUDITORIA DE TRANSITÓRIOS ====================

/**
 * Obtém lista de endereços únicos de uma aba específica
 * @param {string} nomeAba - Nome da aba para buscar endereços
 * @returns {Array} Lista de endereços únicos
 */
function obterEnderecosPorAba(nomeAba) {
  try {
    console.log(` Obtendo endereços da aba: ${nomeAba}`);
    
    // NOVO: Se for Base CD Principal, usar função otimizada
    if (nomeAba === 'Base CD Principal') {
      console.log(' Carregando endereços otimizados do Drive...');
      
      try {
        // Usar função específica para carregar apenas endereços
        const enderecos = carregarApenasEnderecosCDPrincipal();
        console.log(' Endereços carregados:', enderecos ? enderecos.length : 'null');
        return enderecos;
        
      } catch (driveError) {
        console.error(' Erro específico do Drive:', driveError);
        console.error('Stack do Drive:', driveError.stack);
        return [];
      }
    }
    
    // Para outras abas, buscar da planilha local
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(nomeAba);
    
    if (!aba) {
      console.error(` Aba "${nomeAba}" não encontrada`);
      return [];
    }
    
    const dados = aba.getDataRange().getValues();
    
    if (dados.length <= 1) {
      console.log(` Aba "${nomeAba}" está vazia ou só tem cabeçalho`);
      return [];
    }
    
    // Assumindo que endereços estão na coluna E (índice 4)
    const enderecos = dados.slice(1) // Pular cabeçalho
      .map(row => row[4]) // Coluna E (endereço)
      .filter(endereco => endereco && endereco.toString().trim() !== '') // Filtrar vazios
      .map(endereco => endereco.toString().trim()); // Converter para string e trimmar
    
    // Remover duplicatas e ordenar
    const enderecosUnicos = [...new Set(enderecos)].sort();
    
    console.log(` ${enderecosUnicos.length} endereços únicos encontrados na aba "${nomeAba}"`);
    return enderecosUnicos;
    
  } catch (error) {
    console.error(` Erro ao obter endereços da aba "${nomeAba}":`, error);
    return [];
  }
}

/**
 * Salva um registro auditado (CONFORME) de smartphone na aba 'Smartphones Auditados',
 * alinhando o esquema e cálculos com 'Estoque Auditados'.
 * @param {Object} dados - { qtd, codigo, finais, endereco, tipoEstoque }
 *  - qtd: quantidade no sistema (contagem no endereço)
 *  - codigo: CF do produto
 *  - finais: string com finais auditados (ex: 'A,B,C')
 *  - endereco: endereço auditado
 *  - tipoEstoque: 'principal' | 'rma'
 */
function salvarSmartphoneAuditado(dados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = planilha.getSheetByName('Smartphones Auditados');
  if (!sheet) {
    sheet = planilha.insertSheet('Smartphones Auditados');
  }

  // Garante cabeçalho no padrão de 11 colunas
  const header11 = [
    'DATA', 'CF', 'Finais Auditados', 'Descrição', 'Qtd Sis', 'Qtd Aud',
    'Valor Auditado', 'Valor Divergencia', 'Endereço', 'Divergencia', 'Auditor'
  ];
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < 11) {
    sheet.getRange(1, 1, 1, 11).setValues([header11]);
  }

  // Utilitários locais
  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const dataStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yy');
  let operador = '';
  try { operador = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || ''; } catch {}
  const operadorNome = extrairNomeOperador(operador || '');

  const tipoEstoque = (dados && dados.tipoEstoque) ? String(dados.tipoEstoque) : 'principal';
  const enderecoAtual = dados && dados.endereco ? String(dados.endereco).trim() : '';
  const cfAlvo = dados && dados.codigo ? String(dados.codigo).trim() : '';
  const finaisAuditadosStr = dados && dados.finais ? String(dados.finais) : '';
  const qtdSisInformada = Number(dados && dados.qtd != null ? dados.qtd : 0);

  if (!enderecoAtual || !cfAlvo) {
    throw new Error('Endereço e CF são obrigatórios.');
  }

  // Buscar base do endereço para obter descrição e (se possível) custo
  const parseMoeda = (v) => {
    const s0 = String(v == null ? '' : v).trim();
    if (!s0) return NaN;
    const s = s0.replace(/[R$\s]/g, '');
    const hasComma = s.includes(',');
    const hasDot = s.includes('.');
    if (hasComma && hasDot) {
      const lastComma = s.lastIndexOf(',');
      const lastDot = s.lastIndexOf('.');
      if (lastComma > lastDot) {
        const noThousands = s.replace(/\./g, '');
        const normalized = noThousands.replace(/,/g, '.');
        const n = parseFloat(normalized);
        return isNaN(n) ? NaN : n;
      } else {
        const noThousands = s.replace(/,/g, '');
        const n = parseFloat(noThousands);
        return isNaN(n) ? NaN : n;
      }
    }
    if (hasComma && !hasDot) {
      const noThousands = s.replace(/\./g, '');
      const normalized = noThousands.replace(/,/g, '.');
      const n = parseFloat(normalized);
      return isNaN(n) ? NaN : n;
    }
    if (hasDot && !hasComma) {
      const noThousands = s.replace(/,/g, '');
      const n = parseFloat(noThousands);
      return isNaN(n) ? NaN : n;
    }
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
  };

  let descricaoCF = '';
  let valorUnitario = 0;
  try {
    const info = obterInfoCfPorEndereco(enderecoAtual, tipoEstoque, cfAlvo);
    if (info) {
      descricaoCF = info.descricao || '';
      if (typeof info.custo === 'number' && info.custo > 0) valorUnitario = info.custo;
    }
  } catch (e) { /* fallback silencioso */ }

  // Quantidades: para confirmação, Qtd Aud = Qtd Sis
  const qtdSis = qtdSisInformada > 0 ? qtdSisInformada : 0;
  const qtdAud = qtdSis;
  const delta = qtdAud - qtdSis;
  const valorAuditado = valorUnitario * qtdAud;
  const valorDivergencia = valorUnitario * Math.abs(delta); // 0 em conformidade

  const linha = [
    dataStr,             // DATA
    cfAlvo,              // CF
    finaisAuditadosStr,  // FINAIS AUDITADOS
    descricaoCF,         // DESCRIÇÃO
    qtdSis,              // QTD SIS
    qtdAud,              // QTD AUD
    valorAuditado,       // VALOR AUDITADO
    valorDivergencia,    // VALOR DIVERGENCIA
    enderecoAtual,       // ENDEREÇO
    'CONFORME',          // DIVERGENCIA (texto)
    operadorNome         // AUDITOR
  ];

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, 11).setValues([linha]);
  try { sheet.getRange(startRow, 7, 1, 2).setNumberFormat('"R$" #,##0.00'); } catch (e) {}

  return 'Registro de smartphone salvo (CONFORME) com sucesso!';
}

/**
 * Carrega toda a Base CD Principal em cache para otimização
 * @returns {Array} Array com todos os dados da Base CD Principal
 */
function carregarBaseCDPrincipalCompleta() {
  try {
    console.log("🚀 BACKEND: Iniciando carregamento completo da Base CD Principal...");
    
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // Lista de nomes possíveis para o arquivo
    const nomesPossiveis = [
      'Base CD Principal.csv',
      'Base CD Principal - Parte 1.csv',
      'Base CD Principal - Parte1.csv',
      'Base CD Principal - Parte 2.csv', 
      'Base CD Principal - Parte2.csv'
    ];
    
    let arquivoEncontrado = null;
    
    for (const nome of nomesPossiveis) {
      console.log(` Procurando arquivo: "${nome}"`);
      const arquivos = pasta.getFilesByName(nome);
      
      if (arquivos.hasNext()) {
        arquivoEncontrado = arquivos.next();
        console.log(` Arquivo encontrado: "${nome}" (${(arquivoEncontrado.getSize() / 1024 / 1024).toFixed(2)} MB)`);
        break;
      }
    }
    
    if (!arquivoEncontrado) {
      throw new Error('Nenhum arquivo da Base CD Principal encontrado');
    }
    
    // Verificar tamanho do arquivo
    const tamanhoMB = arquivoEncontrado.getSize() / 1024 / 1024;
    console.log(` Tamanho do arquivo: ${tamanhoMB.toFixed(2)} MB`);
    
    if (tamanhoMB > 50) {
      throw new Error(`Arquivo muito grande (${tamanhoMB.toFixed(2)} MB). Limite: 50MB`);
    }
    
    // Ler o arquivo CSV completo
    console.log(' Lendo conteúdo do arquivo CSV...');
    const conteudoCsv = arquivoEncontrado.getBlob().getDataAsString('UTF-8');
    
    // Dividir em linhas
    const linhas = conteudoCsv.split('\n');
    console.log(` Total de linhas no CSV: ${linhas.length}`);
    
    // Processar cada linha
    const dadosCompletos = [];
    let linhaProcesada = 0;
    
    for (let i = 0; i < linhas.length; i++) {
      const linha = linhas[i].trim();
      if (!linha) continue;
      
      // Dividir por vírgula (CSV)
      const colunas = linha.split(',').map(col => col.replace(/"/g, '').trim());
      
      if (colunas.length >= 5) {
        dadosCompletos.push(colunas);
        linhaProcesada++;
      }
      
      // Log de progresso a cada 10000 linhas
      if (linhaProcesada % 10000 === 0) {
        console.log(` Processadas ${linhaProcesada} linhas...`);
      }
    }
    
    console.log(` ✅ BACKEND: Base CD Principal completa carregada: ${dadosCompletos.length} registros`);
    console.log(` Primeiros 3 registros:`, dadosCompletos.slice(0, 3));
    
    return dadosCompletos;
    
  } catch (error) {
    console.error('❌ BACKEND: Erro ao carregar Base CD Principal completa:', error);
    throw new Error(`Erro ao carregar Base CD Principal: ${error.message}`);
  }
}

/**
 * Busca um produto específico na Base CD Principal (otimizado para arquivos grandes)
 * @param {string} codigo - Código do produto a buscar
 * @returns {Array|null} Dados do produto ou null se não encontrado
 */
function buscarProdutoBaseCDPrincipal(codigo) {
  try {
    console.log(` Buscando produto "${codigo}" na Base CD Principal...`);
    
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // Estratégia 1: Tentar CSV primeiro (mais eficiente para busca)
    const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
    if (csvArquivos.hasNext()) {
      const csvArquivo = csvArquivos.next();
      const csvTamanhoMB = csvArquivo.getSize() / 1024 / 1024;
      
      if (csvTamanhoMB < 50) { // Limite seguro
        console.log(' Buscando no arquivo CSV...');
        const csvConteudo = csvArquivo.getBlob().getDataAsString();
        
        // Buscar linha por linha sem carregar tudo na memória
        const linhas = csvConteudo.split('\n');
        
        for (let i = 1; i < linhas.length; i++) { // Pular cabeçalho
          // CORREÇÃO: usar ; como separador, não vírgula
          const colunas = linhas[i].split(';').map(col => col.replace(/"/g, '').trim());
          
          if (colunas.length >= 2) {
            const codigoProduto = colunas[0];  // Código Produto está na coluna 0
            const etiqueta = colunas[1];       // Etiqueta na coluna 1
            
            // Comparar código ou etiqueta
            if (String(codigoProduto).trim() === String(codigo).trim() || 
                String(etiqueta).trim() === String(codigo).trim()) {
              
              console.log(` Produto encontrado: ${codigoProduto} - ${etiqueta}`);
              return colunas; // Retorna array com todas as colunas
            }
          }
        }
        
        console.log(` Produto "${codigo}" não encontrado no CSV`);
        return null;
      }
    }
    
    // Estratégia 2: Fallback para aba local
    console.log(' Buscando na aba local como fallback...');
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Base CD Principal');
    
    if (aba) {
      const dados = aba.getDataRange().getValues();
      
      const produto = dados.find(row => 
        String(row[1]).trim() === String(codigo).trim() || // Código
        String(row[0]).trim() === String(codigo).trim()    // CF
      );
      
      if (produto) {
        console.log(` Produto encontrado na aba local: ${produto[0]} - ${produto[1]}`);
        return produto;
      }
    }
    
    console.log(` Produto "${codigo}" não encontrado`);
    return null;
    
  } catch (error) {
    console.error(` Erro ao buscar produto "${codigo}":`, error);
    return null;
  }
}

function removerTransitorioAuditado(codigo) {
  try {
    console.log(' Removendo produto auditado:', codigo);
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Transitorios Auditados');
    
    if (!aba) {
      throw new Error('Aba "Transitorios Auditados" não encontrada');
    }
    
    const dados = aba.getDataRange().getValues();
    
    for (let i = dados.length - 1; i >= 1; i--) { // Começar do final para não afetar índices
      if (dados[i][1] && dados[i][1].toString() === codigo.toString()) {
        aba.deleteRow(i + 1);
        console.log(' Produto removido da linha:', i + 1);
        return { sucesso: true, linha: i + 1 };
      }
    }
    
    console.log(' Produto não encontrado para remoção:', codigo);
    return { sucesso: false, erro: 'Produto não encontrado' };
    
  } catch (error) {
    console.error(' Erro ao remover produto auditado:', error);
    return { sucesso: false, erro: error.message };
  }
}

// ==================== FUNÇÃO PARA SALVAR PRODUTO LOCALIZADO ====================

function salvarProdutoLocalizado(dados) {
  try {
    console.log(' Salvando produto localizado na planilha Insucessos:', dados);
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let aba = planilha.getSheetByName('Insucessos');
    
    // Criar aba se não existir
    if (!aba) {
      console.log(' Criando aba "Insucessos"...');
      aba = planilha.insertSheet('Insucessos');
      
      // Adicionar cabeçalhos
      const cabecalhos = ['DATA', 'ETIQUETA', 'CODIGO/CF', 'FINAL', 'DESCRICAO', 'STATUS'];
      aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      
      // Formatar cabeçalho
      const cabecalhoRange = aba.getRange(1, 1, 1, cabecalhos.length);
      cabecalhoRange.setFontWeight('bold');
      cabecalhoRange.setBackground('#E8F4FD');
      
      console.log(' Aba "Insucessos" criada com cabeçalhos');
    }
    
    // Preparar dados para inserção
    // Ordem: DATA / ETIQUETA / CODIGO OU CF / FINAL / DESCRICAO / STATUS
    const linha = [
      dados.data,
      dados.etiqueta,
      dados.codigoCf,
      dados.final,
      dados.descricao,
      dados.status
    ];
    
    // Adicionar nova linha no final
    const proximaLinha = aba.getLastRow() + 1;
    aba.getRange(proximaLinha, 1, 1, linha.length).setValues([linha]);
    
    console.log(` Produto localizado salvo na linha ${proximaLinha}:`, linha);
    
    return {
      sucesso: true,
      linha: proximaLinha,
      dados: linha
    };
    
  } catch (error) {
    console.error(' Erro ao salvar produto localizado:', error);
    throw new Error('Erro ao salvar na planilha: ' + error.message);
  }
}

// ==================== FUNÇÃO PARA SALVAR AUDITORIA ====================

/**
 * Salva uma divergência de smartphone na aba 'Smartphones Auditados'.
 * @param {Object} dados - { qtd, qtdAuditada, codigo, finais, endereco, divergencia }
 */
function salvarSmartphoneDivergencia(dados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = planilha.getSheetByName('Smartphones Auditados');
  if (!sheet) {
    sheet = planilha.insertSheet('Smartphones Auditados');
  }

  // Garante cabeçalho no padrão de 11 colunas
  const header11 = [
    'DATA', 'CF', 'Finais Auditados', 'Descrição', 'Qtd Sis', 'Qtd Aud',
    'Valor Auditado', 'Valor Divergencia', 'Endereço', 'Divergencia', 'Auditor'
  ];
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < 11) {
    sheet.getRange(1, 1, 1, 11).setValues([header11]);
  }

  // Utilitários e parâmetros
  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const dataStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yy');
  let operador = '';
  try { operador = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || ''; } catch {}
  const operadorNome = extrairNomeOperador(operador || '');

  const tipoEstoque = (dados && dados.tipoEstoque) ? String(dados.tipoEstoque) : 'principal';
  const enderecoAtual = dados && dados.endereco ? String(dados.endereco).trim() : '';
  const cfAlvo = dados && dados.codigo ? String(dados.codigo).trim() : '';
  const finaisAuditadosStr = dados && dados.finais ? String(dados.finais) : '';
  const qtdSis = Number(dados && dados.qtd != null ? dados.qtd : 0);
  const qtdAud = Number(dados && dados.qtdAuditada != null ? dados.qtdAuditada : 0);
  if (!enderecoAtual || !cfAlvo) throw new Error('Endereço e CF são obrigatórios.');

  // Buscar descrição e custo do CF no endereço
  const parseMoeda = (v) => {
    const s0 = String(v == null ? '' : v).trim();
    if (!s0) return NaN;
    const s = s0.replace(/[R$\s]/g, '');
    const hasComma = s.includes(',');
    const hasDot = s.includes('.');
    if (hasComma && hasDot) {
      const lastComma = s.lastIndexOf(',');
      const lastDot = s.lastIndexOf('.');
      if (lastComma > lastDot) {
        const noThousands = s.replace(/\./g, '');
        const normalized = noThousands.replace(/,/g, '.');
        const n = parseFloat(normalized);
        return isNaN(n) ? NaN : n;
      } else {
        const noThousands = s.replace(/,/g, '');
        const n = parseFloat(noThousands);
        return isNaN(n) ? NaN : n;
      }
    }
    if (hasComma && !hasDot) {
      const noThousands = s.replace(/\./g, '');
      const normalized = noThousands.replace(/,/g, '.');
      const n = parseFloat(normalized);
      return isNaN(n) ? NaN : n;
    }
    if (hasDot && !hasComma) {
      const noThousands = s.replace(/,/g, '');
      const n = parseFloat(noThousands);
      return isNaN(n) ? NaN : n;
    }
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
  };

  let descricaoCF = '';
  let valorUnitario = 0;
  try {
    const info = obterInfoCfPorEndereco(enderecoAtual, tipoEstoque, cfAlvo);
    if (info) {
      descricaoCF = info.descricao || '';
      if (typeof info.custo === 'number' && info.custo > 0) valorUnitario = info.custo;
    }
  } catch (e) {}

  const delta = qtdAud - qtdSis;
  const valorAuditado = valorUnitario * qtdAud;
  const valorDivergencia = valorUnitario * Math.abs(delta);
  const status = delta > 0 ? 'SOBRA' : (delta < 0 ? 'FALTA' : 'SEM DIVERGÊNCIA');
  const divergenciaTexto = `${dados && dados.divergencia ? String(dados.divergencia) : ''}${delta !== 0 ? ` — ${status}` : ''}`.trim();

  const linha = [
    dataStr,             // DATA
    cfAlvo,              // CF
    finaisAuditadosStr,  // FINAIS AUDITADOS
    descricaoCF,         // DESCRIÇÃO
    qtdSis,              // QTD SIS
    qtdAud,              // QTD AUD
    valorAuditado,       // VALOR AUDITADO
    valorDivergencia,    // VALOR DIVERGENCIA
    enderecoAtual,       // ENDEREÇO
    divergenciaTexto,    // DIVERGENCIA (texto + SOBRA/FALTA)
    operadorNome         // AUDITOR
  ];

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, 11).setValues([linha]);
  try { sheet.getRange(startRow, 7, 1, 2).setNumberFormat('"R$" #,##0.00'); } catch (e) {}

  return 'Divergência de smartphone registrada com sucesso!';
}

/**
 * Remove um registro específico da aba 'Smartphones Auditados'.
 * @param {Object} dados - { codigo, endereco }
 */
function removerSmartphoneAuditado(dados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = planilha.getSheetByName('Smartphones Auditados');
  if (!sheet) throw new Error("Aba 'Smartphones Auditados' não encontrada.");

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'Nenhum registro para remover.';

  // Descobre índices das colunas pelo cabeçalho para suportar esquemas antigos/novos
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v || '').trim().toUpperCase());
  const idxCF = Math.max(0, header.indexOf('CF')) + 1; // 1-based
  const idxEnd = Math.max(0, header.indexOf('ENDEREÇO')) + 1; // 1-based
  const width = sheet.getLastColumn();

  for (let i = 2; i <= lastRow; i++) {
    const range = sheet.getRange(i, 1, 1, width);
    const values = range.getValues()[0];
    const cfPlanilha = String(values[idxCF - 1] || '').trim();
    const enderecoPlanilha = String(values[idxEnd - 1] || '').trim();
    if (cfPlanilha === String(dados.codigo || '').trim() && enderecoPlanilha === String(dados.endereco || '').trim()) {
      sheet.deleteRow(i);
      return 'Registro removido com sucesso!';
    }
  }
  return 'Registro não encontrado na planilha.';
}
/**
 * Carrega dados da Base CD Principal do Google Drive
 * @returns {Array} Dados da Base CD Principal em formato array
 */
// FUNÇÃO FORÇADA - Para extrair endereços ignorando erros
function extrairEnderecosForcado() {
  console.log(' === EXTRAÇÃO FORÇADA DE ENDEREÇOS ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // Buscar JSON
    const jsonArquivos = pasta.getFilesByName('Base CD Principal.json');
    
    if (!jsonArquivos.hasNext()) {
      console.error(' JSON não encontrado');
      return ['ERRO_JSON_NAO_ENCONTRADO'];
    }
    
    const arquivo = jsonArquivos.next();
    console.log(' Arquivo:', arquivo.getName());
    console.log(' Tamanho:', (arquivo.getSize() / 1024 / 1024).toFixed(2), 'MB');
    
    // Verificar se arquivo é muito grande
    const tamanhoMB = arquivo.getSize() / 1024 / 1024;
    if (tamanhoMB > 50) {
      console.warn(' Arquivo muito grande! Tentando CSV...');
      
      // Tentar CSV como alternativa
      const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
      if (csvArquivos.hasNext()) {
        const csvArquivo = csvArquivos.next();
        console.log(' Usando CSV:', csvArquivo.getName());
        
        const csvConteudo = csvArquivo.getBlob().getDataAsString();
        const linhas = csvConteudo.split('\n');
        
        const enderecos = new Set();
        
        for (let i = 1; i < Math.min(1000, linhas.length); i++) {
          const colunas = linhas[i].split(';');
          if (colunas.length >= 5) {
            const endereco = colunas[4]; // Assumindo que endereço está na coluna 4
            if (endereco && /^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(endereco.trim())) {
              enderecos.add(endereco.trim());
            }
          }
        }
        
        const enderecosArray = Array.from(enderecos).sort();
        console.log(' Endereços extraídos do CSV:', enderecosArray.length);
        return enderecosArray;
      }
      
      return ['ARQUIVO_MUITO_GRANDE'];
    }
    
    // Arquivo JSON pequeno o suficiente
    console.log(' Carregando JSON...');
    const conteudo = arquivo.getBlob().getDataAsString();
    const dados = JSON.parse(conteudo);
    
    console.log(' Registros no JSON:', dados.length);
    
    const enderecos = new Set();
    let processados = 0;
    
    // Processar apenas uma amostra para teste
    const limite = Math.min(10000, dados.length);
    
    for (let i = 0; i < limite; i++) {
      const item = dados[i];
      processados++;
      
      // Tentar todos os possíveis campos de endereço
      let endereco = item.endereco || item.ENDERECO || item.Endereco || 
                     item.end || item.END || item.End ||
                     item.address || item.ADDRESS || item.Address ||
                     item.enderec || item.ENDEREC ||
                     item.local || item.LOCAL || item.Local ||
                     item.posicao || item.POSICAO || item.Posicao;
      
      if (endereco) {
        const endLimpo = endereco.toString().trim();
        if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(endLimpo)) {
          enderecos.add(endLimpo);
          
          if (endLimpo.includes('29-0-2')) {
            console.log(` Encontrado: ${endLimpo}`);
          }
        }
      }
      
      if (processados % 1000 === 0) {
        console.log(` Processados: ${processados}, Endereços: ${enderecos.size}`);
      }
    }
    
    const enderecosArray = Array.from(enderecos).sort((a, b) => {
      const parseEndereco = (end) => {
        const partes = end.split('-').map(Number);
        return partes[0] * 1000000 + partes[1] * 10000 + partes[2] * 100 + partes[3];
      };
      return parseEndereco(a) - parseEndereco(b);
    });
    
    console.log(' Total de endereços únicos:', enderecosArray.length);
    console.log(' Primeiros 10:', enderecosArray.slice(0, 10));
    
    // Verificar endereço específico
    const endereco29021 = enderecosArray.find(end => end === '29-0-2-1');
    if (endereco29021) {
      console.log(' ENCONTRADO: 29-0-2-1');
    } else {
      console.log(' 29-0-2-1 não encontrado na amostra');
    }
    
    return enderecosArray;
    
  } catch (error) {
    console.error(' Erro na extração forçada:', error);
    return ['ERRO: ' + error.message];
  }
}

// FUNÇÃO DE TESTE SUPER SIMPLES - Para executar no console
function testeSimples() {
  console.log(' === TESTE SUPER SIMPLES ===');
  
  try {
    // 1. Testar acesso à pasta
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    console.log(' Pasta acessada:', pasta.getName());
    
    // 2. Listar TODOS os arquivos
    console.log(' LISTANDO TODOS OS ARQUIVOS:');
    const arquivos = pasta.getFiles();
    let contador = 0;
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      contador++;
      console.log(`${contador}. "${arquivo.getName()}" (${(arquivo.getSize() / 1024 / 1024).toFixed(2)} MB)`);
    }
    
    // 3. Buscar especificamente o JSON
    console.log('\n Procurando "Base CD Principal.json"...');
    const jsonArquivos = pasta.getFilesByName('Base CD Principal.json');
    
    if (jsonArquivos.hasNext()) {
      console.log(' Arquivo JSON encontrado!');
      const arquivo = jsonArquivos.next();
      console.log(' Tamanho:', (arquivo.getSize() / 1024 / 1024).toFixed(2), 'MB');
      
      // Tentar ler uma pequena parte
      const conteudo = arquivo.getBlob().getDataAsString();
      console.log(' Conteúdo carregado, tamanho:', conteudo.length, 'caracteres');
      
      // Tentar fazer parse JSON
      const dados = JSON.parse(conteudo);
      console.log(' JSON parseado! Total de registros:', dados.length);
      
      if (dados.length > 0) {
        console.log(' Primeiro registro:', dados[0]);
        console.log(' Chaves disponíveis:', Object.keys(dados[0]));
        
        // Verificar campo de endereço
        const item = dados[0];
        let endereco = item.endereco || item.ENDERECO || item.Endereco || 
                       item.end || item.END || item.End ||
                       item.address || item.ADDRESS || item.Address;
        
        if (endereco) {
          console.log(' Campo de endereço encontrado:', endereco);
        } else {
          console.warn(' Nenhum campo de endereço encontrado!');
        }
      }
      
      return 'SUCESSO';
    } else {
      console.error(' Arquivo JSON NÃO encontrado!');
      return 'JSON_NAO_ENCONTRADO';
    }
    
  } catch (error) {
    console.error(' Erro:', error);
    return 'ERRO: ' + error.message;
  }
}

// NOVA FUNÇÃO DE DEBUG - Para testar extração de endereços com logs detalhados
function testarExtracaoEnderecos() {
  console.log(' === TESTE DE EXTRAÇÃO DE ENDEREÇOS ===');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    console.log(' Pasta acessada:', pasta.getName());
    
    const nomeArquivo = 'Base CD Principal.json';
    const arquivos = pasta.getFilesByName(nomeArquivo);
    
    if (!arquivos.hasNext()) {
      console.error(' Arquivo não encontrado!');
      return null;
    }
    
    const arquivo = arquivos.next();
    console.log(' Arquivo encontrado:', arquivo.getName());
    
    const conteudo = arquivo.getBlob().getDataAsString();
    console.log(' Tamanho do conteúdo:', conteudo.length);
    
    const dados = JSON.parse(conteudo);
    console.log(' Total de registros:', dados.length);
    
    if (dados.length > 0) {
      console.log(' Primeiro registro:', dados[0]);
      console.log(' Chaves:', Object.keys(dados[0]));
      
      // Testar extração de endereços
      const enderecos = new Set();
      let contador = 0;
      
      for (let i = 0; i < Math.min(100, dados.length); i++) {
        const item = dados[i];
        
        // Tentar encontrar campo de endereço
        let endereco = item.endereco || item.ENDERECO || item.Endereco || 
                       item.end || item.END || item.End ||
                       item.address || item.ADDRESS || item.Address;
        
        if (endereco) {
          const endLimpo = endereco.toString().trim();
          if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(endLimpo)) {
            enderecos.add(endLimpo);
            contador++;
            
            if (contador <= 10) {
              console.log(` Endereço ${contador}: ${endLimpo}`);
            }
          }
        }
      }
      
      const enderecosArray = Array.from(enderecos).sort();
      console.log(` Total de endereços únicos (primeiro 100 registros): ${enderecosArray.length}`);
      console.log(' Primeiros 10 endereços:', enderecosArray.slice(0, 10));
      
      return enderecosArray;
    }
    
  } catch (error) {
    console.error(' Erro:', error);
    return null;
  }
}

// FUNÇÃO DE TESTE - Para simular doGet() e debug completo
function testarProcessoCompleto() {
  console.log(' === TESTE DO PROCESSO COMPLETO ===');
  
  try {
    console.log('1️⃣ Testando carregarApenasEnderecosCDPrincipal()...');
    const enderecos = carregarApenasEnderecosCDPrincipal();
    
    console.log(' Resultado de carregarApenasEnderecosCDPrincipal():');
    console.log('- Tipo:', typeof enderecos);
    console.log('- É array?', Array.isArray(enderecos));
    console.log('- Length:', enderecos ? enderecos.length : 'null/undefined');
    
    if (enderecos && Array.isArray(enderecos) && enderecos.length > 0) {
      console.log(' Endereços carregados com sucesso!');
      console.log(' Primeiros 10:', enderecos.slice(0, 10));
      
      // Testar JSON.stringify
      console.log('2️⃣ Testando JSON.stringify...');
      const enderecosJson = JSON.stringify(enderecos);
      console.log(' JSON.stringify funcionou!');
      console.log(' Tamanho do JSON:', enderecosJson.length, 'caracteres');
      console.log(' Início do JSON:', enderecosJson.substring(0, 200));
      
      return {
        sucesso: true,
        enderecos: enderecos,
        json: enderecosJson
      };
    } else {
      console.error(' Endereços não carregados corretamente!');
      return {
        sucesso: false,
        erro: 'Endereços vazios ou inválidos'
      };
    }
    
  } catch (error) {
    console.error(' Erro no teste completo:', error);
    console.error('Stack:', error.stack);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// FUNÇÃO DE TESTE SIMPLIFICADA - Para verificar acesso ao arquivo JSON específico
function testarArquivoJSON() {
  console.log(' === TESTE SIMPLIFICADO DO ARQUIVO JSON ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    console.log(' Acessando pasta ID:', pastaId);
    
    const pasta = DriveApp.getFolderById(pastaId);
    console.log(' Pasta acessada:', pasta.getName());
    
    // Buscar especificamente por "Base CD Principal.json"
    const nomeArquivo = 'Base CD Principal.json';
    console.log(' Procurando arquivo:', nomeArquivo);
    
    const arquivos = pasta.getFilesByName(nomeArquivo);
    
    if (!arquivos.hasNext()) {
      console.error(` Arquivo "${nomeArquivo}" NÃO ENCONTRADO!`);
      
      // Listar TODOS os arquivos para debug
      console.log(' TODOS os arquivos na pasta:');
      const todosArquivos = pasta.getFiles();
      let contador = 0;
      while (todosArquivos.hasNext()) {
        const arquivo = todosArquivos.next();
        console.log(`${contador + 1}. "${arquivo.getName()}" (${arquivo.getSize()} bytes)`);
        contador++;
      }
      
      return false;
    }
    
    const arquivo = arquivos.next();
    console.log(' Arquivo encontrado:', arquivo.getName());
    console.log(' Tamanho:', arquivo.getSize(), 'bytes');
    console.log(' Última modificação:', arquivo.getLastUpdated());
    
    // Tentar ler uma pequena parte
    console.log(' Tentando ler conteúdo...');
    const conteudo = arquivo.getBlob().getDataAsString();
    console.log(' Tamanho do conteúdo:', conteudo.length, 'caracteres');
    console.log(' Primeiros 500 caracteres:', conteudo.substring(0, 500));
    
    // Tentar fazer parse JSON
    console.log(' Tentando parse JSON...');
    const dados = JSON.parse(conteudo);
    console.log(' JSON parseado com sucesso!');
    console.log(' Total de registros:', dados.length);
    
    if (dados.length > 0) {
      console.log(' Primeiro registro:', dados[0]);
      console.log(' Chaves do primeiro registro:', Object.keys(dados[0]));
      
      // Verificar se tem campo de endereço
      const primeiroItem = dados[0];
      let campoEndereco = null;
      
      const possiveisCampos = ['endereco', 'ENDERECO', 'Endereco', 'end', 'END', 'End', 'address', 'ADDRESS', 'Address'];
      for (const campo of possiveisCampos) {
        if (primeiroItem[campo]) {
          campoEndereco = campo;
          console.log(` Campo de endereço encontrado: "${campo}" = "${primeiroItem[campo]}"`);
          break;
        }
      }
      
      if (!campoEndereco) {
        console.warn(' Nenhum campo de endereço encontrado no primeiro registro!');
      }
    }
    
    return true;
    
  } catch (error) {
    console.error(' Erro no teste:', error);
    console.error('Stack trace:', error.stack);
    return false;
  }
}

// FUNÇÃO DE TESTE - Para verificar estrutura do CSV
function analisarEstruturaCDPrincipal() {
  console.log(' === ANÁLISE DA ESTRUTURA DO CSV ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
    
    if (!csvArquivos.hasNext()) {
      console.error(' Arquivo CSV não encontrado');
      return;
    }
    
    const csvArquivo = csvArquivos.next();
    const csvConteudo = csvArquivo.getBlob().getDataAsString();
    const linhas = csvConteudo.split('\n');
    
    if (linhas.length < 2) {
      console.error(' CSV vazio ou sem dados');
      return;
    }
    
    // Analisar cabeçalho
    const cabecalho = linhas[0].split(',').map(col => col.replace(/"/g, '').trim());
    console.log(' CABEÇALHO:');
    cabecalho.forEach((col, index) => {
      console.log(`  [${index}] "${col}"`);
    });
    
    // Analisar primeira linha de dados
    console.log('\n PRIMEIRA LINHA DE DADOS:');
    const primeiraLinha = linhas[1].split(',').map(col => col.replace(/"/g, '').trim());
    primeiraLinha.forEach((col, index) => {
      console.log(`  [${index}] "${col}"`);
    });
    
    // Analisar segunda linha de dados
    if (linhas.length > 2) {
      console.log('\n SEGUNDA LINHA DE DADOS:');
      const segundaLinha = linhas[2].split(',').map(col => col.replace(/"/g, '').trim());
      segundaLinha.forEach((col, index) => {
        console.log(`  [${index}] "${col}"`);
      });
    }
    
    // Procurar qual coluna tem padrão de endereço
    console.log('\n ANÁLISE DE PADRÕES DE ENDEREÇO:');
    for (let col = 0; col < cabecalho.length; col++) {
      let enderecosEncontrados = 0;
      
      // Verificar 10 linhas
      for (let linha = 1; linha <= Math.min(10, linhas.length - 1); linha++) {
        const colunas = linhas[linha].split(',').map(col => col.replace(/"/g, '').trim());
        if (colunas[col]) {
          const valor = colunas[col];
          // Verificar se parece endereço
          if (/^[A-Z]{2,4}-[A-Z0-9]+(-[0-9]+)*$/i.test(valor)) {
            enderecosEncontrados++;
          }
        }
      }
      
      if (enderecosEncontrados > 0) {
        console.log(`  Coluna [${col}] "${cabecalho[col]}": ${enderecosEncontrados}/10 valores parecem endereços`);
      }
    }
    
    return true;
    
  } catch (error) {
    console.error(' Erro na análise:', error);
    return false;
  }
}

// FUNÇÃO DE TESTE - Para diagnosticar o problema
function testarDriveIntegracao() {
  console.log(' === TESTE DE INTEGRAÇÃO DRIVE ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    console.log(' Tentando acessar pasta:', pastaId);
    
    const pasta = DriveApp.getFolderById(pastaId);
    console.log(' Pasta acessada:', pasta.getName());
    
    // Listar todos os arquivos da pasta
    const arquivos = pasta.getFiles();
    console.log(' Arquivos na pasta:');
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      console.log(`- ${arquivo.getName()} (${arquivo.getSize()} bytes)`);
    }
    
    // Tentar carregar especificamente o JSON
    const nomeArquivo = 'Base CD Principal.json';
    const arquivosJSON = pasta.getFilesByName(nomeArquivo);
    
    if (!arquivosJSON.hasNext()) {
      console.error(` Arquivo "${nomeArquivo}" não encontrado`);
      
      // Tentar variações do nome
      const variacoes = [
        'Base CD Principal.json',
        'base cd principal.json',
        'Base_CD_Principal.json',
        'BaseCDPrincipal.json'
      ];
      
      console.log(' Testando variações do nome...');
      for (const variacao of variacoes) {
        const teste = pasta.getFilesByName(variacao);
        if (teste.hasNext()) {
          console.log(` Encontrado com nome: ${variacao}`);
        }
      }
      
      return false;
    }
    
    const arquivo = arquivosJSON.next();
    console.log(` Arquivo JSON encontrado: ${arquivo.getName()}`);
    console.log(` Última modificação: ${arquivo.getLastUpdated()}`);
    console.log(` Tamanho: ${arquivo.getSize()} bytes`);
    
    // Tentar ler o JSON
    const conteudo = arquivo.getBlob().getDataAsString();
    console.log(` Primeiros 200 caracteres:`, conteudo.substring(0, 200));
    
    const dados = JSON.parse(conteudo);
    console.log(` JSON parseado: ${dados.length} registros`);
    
    if (dados.length > 0) {
      console.log(' Estrutura do primeiro registro:', dados[0]);
      if (dados.length > 1) {
        console.log(' Estrutura do segundo registro:', dados[1]);
      }
    }
    
    return true;
    
  } catch (error) {
    console.error(' Erro no teste:', error);
    console.error('Stack:', error.stack);
    return false;
  }
}

function carregarBaseCDPrincipalDoDrive() {
  try {
    console.log(' Carregando Base CD Principal do Google Drive...');
    
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const nomeArquivo = 'Base CD Principal.json';
    
    // Buscar arquivo na pasta
    const pasta = DriveApp.getFolderById(pastaId);
    const arquivos = pasta.getFilesByName(nomeArquivo);
    
    if (!arquivos.hasNext()) {
      console.error(` Arquivo "${nomeArquivo}" não encontrado na pasta`);
      return [];
    }
    
    const arquivo = arquivos.next();
    const tamanhoMB = arquivo.getSize() / 1024 / 1024;
    console.log(` Arquivo encontrado: ${arquivo.getName()}`);
    console.log(` Última modificação: ${arquivo.getLastUpdated()}`);
    console.log(` Tamanho: ${tamanhoMB.toFixed(2)} MB`);
    
    // NOVO: Verificar se arquivo é muito grande para Apps Script
    const LIMITE_MB = 45; // Limite seguro para Apps Script
    
    if (tamanhoMB > LIMITE_MB) {
      console.log(` Arquivo muito grande (${tamanhoMB.toFixed(2)}MB > ${LIMITE_MB}MB)`);
      console.log(' Tentando carregar apenas os endereços únicos...');
      
      // Estratégia: usar CSV menor se disponível
      const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
      if (csvArquivos.hasNext()) {
        const csvArquivo = csvArquivos.next();
        const csvTamanhoMB = csvArquivo.getSize() / 1024 / 1024;
        console.log(` Tentando CSV: ${csvTamanhoMB.toFixed(2)} MB`);
        
        if (csvTamanhoMB < LIMITE_MB) {
          console.log(' Carregando dados do CSV...');
          const csvConteudo = csvArquivo.getBlob().getDataAsString();
          const linhas = csvConteudo.split('\n');
          
          const dados = linhas.map(linha => {
            // Parse básico de CSV (pode precisar ajustar dependendo do formato)
            return linha.split(',').map(campo => campo.replace(/"/g, '').trim());
          }).filter(linha => linha.length > 1 && linha[0]); // Filtrar linhas válidas
          
          console.log(` ${dados.length} registros carregados do CSV`);
          return dados;
        }
      }
      
      // Se nem CSV funcionar, retornar dados limitados do cache/fallback
      console.log(' Arquivo muito grande, usando fallback da aba local...');
      return carregarFallbackLocal();
    }
    
    // Arquivo pequeno o suficiente, carregar normalmente
    console.log(' Carregando arquivo JSON...');
    const conteudoJSON = arquivo.getBlob().getDataAsString();
    const dados = JSON.parse(conteudoJSON);
    
    console.log(` ${dados.length} registros carregados da Base CD Principal`);
    return dados;
    
  } catch (error) {
    console.error(' Erro ao carregar Base CD Principal do Drive:', error);
    console.error('Stack:', error.stack);
    
    // Fallback em caso de erro
    return carregarFallbackLocal();
  }
}

/**
 * Função auxiliar para carregar dados da aba local como fallback
 */
function carregarFallbackLocal() {
  try {
    console.log(' Carregando fallback da aba local...');
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Base CD Principal');
    
    if (aba) {
      const dadosFallback = aba.getDataRange().getValues();
      console.log(` Fallback: ${dadosFallback.length} registros da aba local`);
      return dadosFallback;
    } else {
      console.log(' Aba local também não encontrada, retornando dados vazios');
      return [];
    }
  } catch (fallbackError) {
    console.error(' Fallback também falhou:', fallbackError);
    return [];
  }
}

/**
 * Carrega uma amostra pequena da Base CD Principal para compatibilidade com frontend
 * @returns {Array} Amostra limitada dos dados
 */
/**
 * Extrai endereços únicos apenas dos dados da amostra
 * Garante que todos os endereços listados tenham produtos correspondentes
 */
function extrairEnderecosDeAmostra(dadosAmostra) {
  console.log(' Extraindo endereços únicos da amostra AMPLIADA...');
  
  if (!dadosAmostra || dadosAmostra.length <= 1) {
    console.warn(' Amostra vazia ou apenas cabeçalho');
    return [];
  }
  
  const enderecos = new Set();
  let processados = 0;
  
  // Processar dados da amostra (pular cabeçalho)
  for (let i = 1; i < dadosAmostra.length; i++) {
    const linha = dadosAmostra[i];
    processados++;
    
    if (linha && linha.length >= 5) {
      const endereco = linha[4]; // Endereço está na coluna 4
      
      if (endereco && endereco.toString().trim() !== '') {
        const enderecoLimpo = endereco.toString().trim();
        
        // Validar formato XX-X-XX-X
        if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(enderecoLimpo)) {
          enderecos.add(enderecoLimpo);
        }
      }
    }
  }
  
  // Converter para array e ordenar
  const enderecosArray = Array.from(enderecos).sort((a, b) => {
    const parseEndereco = (end) => {
      const partes = end.split('-').map(Number);
      return partes[0] * 1000000 + partes[1] * 10000 + partes[2] * 100 + partes[3];
    };
    return parseEndereco(a) - parseEndereco(b);
  });
  
  console.log(` EXTRAÇÃO DE ENDEREÇOS FINALIZADA:`);
  console.log(` Linhas processadas: ${processados}`);
  console.log(`  Endereços únicos extraídos: ${enderecosArray.length}`);
  console.log(` Primeiros 10 endereços:`, enderecosArray.slice(0, 10));
  console.log(` Últimos 10 endereços:`, enderecosArray.slice(-10));
  
  // Verificar cobertura por corredor
  const corredores = {};
  enderecosArray.forEach(end => {
    const corredor = end.split('-')[0];
    corredores[corredor] = (corredores[corredor] || 0) + 1;
  });
  
  console.log(` Cobertura por corredor:`, corredores);
  console.log(` Total de corredores cobertos: ${Object.keys(corredores).length}`);
  
  return enderecosArray;
}

/**
 * Extrai TODOS os endereços únicos do CSV completo
 * Para fornecer lista completa no frontend
 */
function extrairTodosEnderecosCsvCompleto() {
  console.log(' === EXTRAINDO TODOS OS ENDEREÇOS DO CSV COMPLETO ===');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    console.log(' Buscando arquivo CSV...');
    const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
    
    if (!csvArquivos.hasNext()) {
      console.warn(' Arquivo CSV não encontrado!');
      return [];
    }
    
    const arquivoCsv = csvArquivos.next();
    const tamanhoMB = arquivoCsv.getSize() / 1024 / 1024;
    
    console.log(` CSV encontrado: ${tamanhoMB.toFixed(2)} MB`);
    console.log(' ATENÇÃO: Processando arquivo COMPLETO para extrair TODOS os endereços');
    
    // Carregar CSV completo
    console.log(' Carregando CSV completo...');
    const conteudoCsv = arquivoCsv.getBlob().getDataAsString();
    const linhas = conteudoCsv.split('\n');
    
    console.log(` Total de linhas no arquivo: ${linhas.length}`);
    
    const enderecos = new Set();
    let processados = 0;
    let enderecosValidos = 0;
    let enderecosInvalidos = 0;
    
    // Processar TODAS as linhas (exceto cabeçalho)
    for (let i = 1; i < linhas.length; i++) {
      processados++;
      
      // Separar por ; (ponto e vírgula)
      const colunas = linhas[i].split(';').map(col => col.replace(/"/g, '').trim());
      
      if (colunas.length >= 5) {
        const endereco = colunas[4]; // Endereço está na coluna 4
        
        if (endereco && endereco.toString().trim() !== '') {
          const enderecoLimpo = endereco.toString().trim();
          
          // Validar formato XX-X-XX-X
          if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(enderecoLimpo)) {
            enderecos.add(enderecoLimpo);
            enderecosValidos++;
          } else {
            enderecosInvalidos++;
          }
        }
      }
      
      // Log a cada 50000 linhas para acompanhar progresso
      if (processados % 50000 === 0) {
        console.log(` Processadas ${processados} linhas - Endereços únicos: ${enderecos.size}`);
      }
    }
    
    // Converter para array e ordenar
    const enderecosArray = Array.from(enderecos).sort((a, b) => {
      const parseEndereco = (end) => {
        const partes = end.split('-').map(Number);
        return partes[0] * 1000000 + partes[1] * 10000 + partes[2] * 100 + partes[3];
      };
      return parseEndereco(a) - parseEndereco(b);
    });
    
    console.log(` PROCESSAMENTO COMPLETO FINALIZADO:`);
    console.log(` Total de linhas processadas: ${processados}`);
    console.log(` TODOS os endereços únicos extraídos: ${enderecosArray.length}`);
    console.log(` Endereços válidos: ${enderecosValidos}`);
    console.log(` Endereços inválidos: ${enderecosInvalidos}`);
    console.log(` Primeiros 10:`, enderecosArray.slice(0, 10));
    console.log(` Últimos 10:`, enderecosArray.slice(-10));
    
    // Análise estatística
    const corredores = {};
    enderecosArray.forEach(end => {
      const corredor = end.split('-')[0];
      corredores[corredor] = (corredores[corredor] || 0) + 1;
    });
    
    console.log(` ESTATÍSTICAS COMPLETAS:`);
    console.log(` Total de corredores: ${Object.keys(corredores).length}`);
    console.log(` Corredores com mais endereços:`, 
      Object.entries(corredores)
        .sort(([,a], [,b]) => b - a)
        .slice(0, 10)
        .map(([corredor, count]) => `${corredor}: ${count}`)
    );
    
    return enderecosArray;
    
  } catch (error) {
    console.error(' Erro ao extrair TODOS os endereços do CSV:', error);
    console.error('Stack:', error.stack);
    return [];
  }
}

function carregarAmostraBaseCDPrincipal() {
  try {
    console.log(' Carregando APENAS cabeçalho inicial (dados sob demanda)...');
    
    // NOVA ESTRATÉGIA: Apenas cabeçalho inicial
    // Os dados serão carregados sob demanda quando o endereço for selecionado
    const dadosAmostra = [['CF', 'CODIGO', 'FINAL', 'DESCRICAO', 'ENDERECO']];
    
    console.log(' Amostra mínima carregada - dados serão carregados sob demanda');
    return dadosAmostra;
    
  } catch (error) {
    console.error(' Erro ao carregar amostra:', error);
    return [['CF', 'CODIGO', 'FINAL', 'DESCRICAO', 'ENDERECO']];
  }
}

/**
 * VERSÃO ROBUSTA - Carrega apenas os endereços únicos da Base CD Principal
 * @returns {Array} Lista de endereços únicos
 */
function carregarApenasEnderecosCDPrincipalRobusta() {
  console.log(' === VERSÃO ROBUSTA DE EXTRAÇÃO DE ENDEREÇOS ===');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    console.log(' Pasta acessada:', pasta.getName());
    
    // Estratégia 1: Tentar JSON primeiro
    console.log(' Tentativa 1: Arquivo JSON...');
    const jsonArquivos = pasta.getFilesByName('Base CD Principal.json');
    
    if (jsonArquivos.hasNext()) {
      const arquivoJson = jsonArquivos.next();
      const tamanhoMB = arquivoJson.getSize() / 1024 / 1024;
      
      console.log(` JSON encontrado: ${tamanhoMB.toFixed(2)} MB`);
      
      // Se arquivo for muito grande, pular para CSV
      if (tamanhoMB > 50) {
        console.warn(' JSON muito grande, pulando para CSV...');
      } else {
        try {
          console.log(' Carregando JSON...');
          const conteudoJson = arquivoJson.getBlob().getDataAsString();
          const dadosJson = JSON.parse(conteudoJson);
          
          console.log(` ${dadosJson.length} registros no JSON`);
          
          if (dadosJson.length > 0) {
            return extrairEnderecosDoDados(dadosJson, 'JSON');
          }
        } catch (jsonError) {
          console.error(' Erro no JSON:', jsonError);
        }
      }
    }
    
    // Estratégia 2: Tentar CSV
    console.log(' Tentativa 2: Arquivo CSV...');
    const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
    
    if (csvArquivos.hasNext()) {
      const arquivoCsv = csvArquivos.next();
      const tamanhoMB = arquivoCsv.getSize() / 1024 / 1024;
      
      console.log(` CSV encontrado: ${tamanhoMB.toFixed(2)} MB`);
      
      if (tamanhoMB < 50) {
        try {
          console.log(' Carregando CSV...');
          const conteudoCsv = arquivoCsv.getBlob().getDataAsString();
          const linhas = conteudoCsv.split('\n');
          
          console.log(` ${linhas.length} linhas no CSV`);
          
          return extrairEnderecosDoCsv(linhas);
        } catch (csvError) {
          console.error(' Erro no CSV:', csvError);
        }
      }
    }
    
    // Estratégia 3: Aba local
    console.log(' Tentativa 3: Aba local...');
    return extrairEnderecosDaAbaLocal();
    
  } catch (error) {
    console.error(' Erro geral:', error);
    return gerarEnderecosFallback();
  }
}

// Função auxiliar para extrair endereços de dados JSON
function extrairEnderecosDoDados(dados, tipo) {
  console.log(` Extraindo endereços de ${tipo}...`);
  
  const enderecos = new Set();
  let processados = 0;
  const limite = Math.min(50000, dados.length); // Limitar para evitar timeout
  
  for (let i = 0; i < limite; i++) {
    const item = dados[i];
    processados++;
    
    // Tentar múltiplos campos possíveis
    let endereco = item.endereco || item.ENDERECO || item.Endereco || 
                   item.end || item.END || item.End ||
                   item.address || item.ADDRESS || item.Address ||
                   item.local || item.LOCAL || item.Local ||
                   item.posicao || item.POSICAO || item.Posicao;
    
    if (endereco) {
      const endLimpo = endereco.toString().trim();
      if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(endLimpo)) {
        enderecos.add(endLimpo);
      }
    }
    
    if (processados % 10000 === 0) {
      console.log(` ${tipo}: ${processados} processados, ${enderecos.size} endereços`);
    }
  }
  
  const enderecosArray = Array.from(enderecos).sort((a, b) => {
    const parseEndereco = (end) => {
      const partes = end.split('-').map(Number);
      return partes[0] * 1000000 + partes[1] * 10000 + partes[2] * 100 + partes[3];
    };
    return parseEndereco(a) - parseEndereco(b);
  });
  
  console.log(` ${enderecosArray.length} endereços únicos de ${tipo}`);
  return enderecosArray;
}

// Função auxiliar para extrair endereços de CSV
function extrairEnderecosDoCsv(linhas) {
  console.log(' Extraindo endereços do CSV...');
  
  const enderecos = new Set();
  const limite = Math.min(10000, linhas.length); // Limitar para evitar timeout
  
  for (let i = 1; i < limite; i++) { // Pular cabeçalho
    const colunas = linhas[i].split(';').map(col => col.replace(/"/g, '').trim());
    
    if (colunas.length >= 5) {
      const endereco = colunas[4]; // Assumindo coluna 4 para endereço
      if (endereco && /^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(endereco)) {
        enderecos.add(endereco);
      }
    }
  }
  
  const enderecosArray = Array.from(enderecos).sort();
  console.log(` ${enderecosArray.length} endereços únicos do CSV`);
  return enderecosArray;
}

// Função auxiliar para extrair da aba local
function extrairEnderecosDaAbaLocal() {
  console.log(' Extraindo endereços da aba local...');
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Base CD Principal');
    
    if (!aba) {
      console.warn(' Aba local não encontrada');
      return gerarEnderecosFallback();
    }
    
    const dados = aba.getDataRange().getValues();
    console.log(` ${dados.length} linhas na aba local`);
    
    if (dados.length > 1) {
      return extrairEnderecosCDPrincipal(dados);
    }
    
  } catch (error) {
    console.error(' Erro na aba local:', error);
  }
  
  return gerarEnderecosFallback();
}

// Função para gerar endereços de fallback mais realistas
function gerarEnderecosFallback() {
  console.log(' Gerando endereços de fallback...');
  
  const enderecosFallback = [];
  
  // Gerar alguns endereços de exemplo mais realistas
  for (let corredor = 1; corredor <= 5; corredor++) {
    for (let nivel = 1; nivel <= 3; nivel++) {
      for (let prateleira = 1; prateleira <= 10; prateleira++) {
        for (let posicao = 1; posicao <= 2; posicao++) {
          enderecosFallback.push(`${corredor}-${nivel}-${prateleira}-${posicao}`);
        }
      }
    }
  }
  
  console.log(` ${enderecosFallback.length} endereços de fallback gerados`);
  return enderecosFallback;
}

/**
 * Carrega apenas os endereços únicos da Base CD Principal (otimizado para arquivos grandes)
 * @returns {Array} Lista de endereços únicos
 */
/**
 * Extrair endereços únicos diretamente do CSV da Base CD Principal
 * Versão otimizada para contornar limitação de tamanho do JSON (104MB)
 */
function extrairEnderecosCsvDireto() {
  console.log(' === NOVA ESTRATÉGIA: Extraindo endereços direto do CSV ===');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    console.log(' Buscando arquivo CSV...');
    const csvArquivos = pasta.getFilesByName('Base CD Principal.csv');
    
    if (!csvArquivos.hasNext()) {
      console.warn(' Arquivo CSV não encontrado!');
      return [];
    }
    
    const arquivoCsv = csvArquivos.next();
    const tamanhoMB = arquivoCsv.getSize() / 1024 / 1024;
    
    console.log(` CSV encontrado: ${tamanhoMB.toFixed(2)} MB`);
    
    // Carregar CSV (46MB - dentro do limite)
    console.log(' Carregando CSV...');
    const conteudoCsv = arquivoCsv.getBlob().getDataAsString();
    const linhas = conteudoCsv.split('\n');
    
    console.log(` Total de linhas: ${linhas.length}`);
    
    const enderecos = new Set();
    let processados = 0;
    let enderecosValidos = 0;
    let enderecosInvalidos = 0;
    
    // Processar cada linha (exceto cabeçalho)
    for (let i = 1; i < linhas.length; i++) {
      processados++;
      
      // Separar por ; (ponto e vírgula)
      const colunas = linhas[i].split(';').map(col => col.replace(/"/g, '').trim());
      
      if (colunas.length >= 5) {
        const endereco = colunas[4]; // Endereço está na coluna 4
        
        if (endereco && endereco.toString().trim() !== '') {
          const enderecoLimpo = endereco.toString().trim();
          
          // Validar formato XX-X-XX-X
          if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(enderecoLimpo)) {
            enderecos.add(enderecoLimpo);
            enderecosValidos++;
            
            // Log específico para endereços 29-0-2-X
            if (enderecoLimpo.startsWith('29-0-2-')) {
              console.log(` ENCONTRADO 29-0-2-X: ${enderecoLimpo} (linha ${i})`);
            }
          } else {
            enderecosInvalidos++;
          }
        }
      }
      
      // Log a cada 10000 linhas
      if (processados % 10000 === 0) {
        console.log(` Processadas ${processados} linhas - Endereços únicos: ${enderecos.size}`);
      }
    }
    
    // Converter para array e ordenar
    const enderecosArray = Array.from(enderecos).sort((a, b) => {
      const parseEndereco = (end) => {
        const partes = end.split('-').map(Number);
        return partes[0] * 1000000 + partes[1] * 10000 + partes[2] * 100 + partes[3];
      };
      return parseEndereco(a) - parseEndereco(b);
    });
    
    console.log(` RESULTADO FINAL:`);
    console.log(` Linhas processadas: ${processados}`);
    console.log(` Endereços únicos extraídos: ${enderecosArray.length}`);
    console.log(` Endereços válidos: ${enderecosValidos}`);
    console.log(` Endereços inválidos: ${enderecosInvalidos}`);
    console.log(` Primeiros 10:`, enderecosArray.slice(0, 10));
    console.log(` Últimos 10:`, enderecosArray.slice(-10));
    
    // Verificar endereços que começam com "29"
    const enderecos29 = enderecosArray.filter(end => end.startsWith('29-'));
    console.log(` Endereços "29-": ${enderecos29.length}`);
    
    // Verificar especificamente "29-0-2-1"
    if (enderecosArray.includes('29-0-2-1')) {
      console.log(` CONFIRMADO: "29-0-2-1" EXISTE!`);
    } else {
      console.log(` "29-0-2-1" NÃO encontrado`);
    }
    
    return enderecosArray;
    
  } catch (error) {
    console.error(' Erro ao extrair endereços do CSV:', error);
    return [];
  }
}

function carregarApenasEnderecosCDPrincipal() {
  //  NOVA VERSÃO INSTANTÂNEA - Igual ao RMA
  return extrairEnderecosCDPrincipalInstantaneo();
}

/**
 * NOVA FUNÇÃO INSTANTÂNEA - Extrai endereços do CD Principal usando a mesma estratégia simples do RMA
 * @returns {Array} Array com endereços únicos ordenados
 */
function extrairEnderecosCDPrincipalInstantaneo() {
  console.log(' === CARREGAMENTO INSTANTÂNEO CD PRINCIPAL ===');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    // Lista de arquivos CSV (priorizando partes menores para rapidez)
    const nomesArquivos = [
      'Base CD Principal - Parte 1.csv',
      'Base CD Principal - Parte1.csv',
      'Base_CD_Principal_Parte1.csv',
      'Base CD Principal - Parte 2.csv',
      'Base CD Principal - Parte2.csv', 
      'Base_CD_Principal_Parte2.csv',
      'Base CD Principal.csv'
    ];
    
    const enderecos = new Set();
    let arquivosProcessados = 0;
    
    for (const nomeArquivo of nomesArquivos) {
      const arquivos = pasta.getFilesByName(nomeArquivo);
      if (!arquivos.hasNext()) continue;
      
      const arquivo = arquivos.next();
      const tamanhoMB = arquivo.getSize() / 1024 / 1024;
      
      // Só processar arquivos até 45MB para evitar timeout
      if (tamanhoMB > 45) {
        console.log(` Arquivo muito grande, pulando: ${nomeArquivo} (${tamanhoMB.toFixed(2)}MB)`);
        continue;
      }
      
      console.log(` Processando: ${nomeArquivo} (${tamanhoMB.toFixed(2)}MB)`);
      
      try {
        // MESMO MÉTODO DO RMA - SIMPLES E DIRETO
        const csv = arquivo.getBlob().getDataAsString();
        const linhas = csv.split('\n');
        
        // Loop simples - sem validações complexas (igual ao RMA)
        for (let i = 1; i < linhas.length; i++) {
          const linha = linhas[i];
          if (!linha || linha.trim() === '') continue;
          
          const cols = linha.split(';').map(c => c.replace(/"/g,'').trim());
          
          // Endereço na coluna 4 (índice 4) - igual ao RMA
          if (cols.length >= 5 && cols[4] && cols[4] !== 'ENDERECO' && cols[4] !== 'ENDEREÇO') {
            enderecos.add(cols[4]);
          }
        }
        
        arquivosProcessados++;
        console.log(` ${nomeArquivo}: Total ${enderecos.size} endereços únicos acumulados`);
        
      } catch (arquivoError) {
        console.warn(` Erro ao processar ${nomeArquivo}:`, arquivoError.message);
        continue;
      }
    }
    
    if (arquivosProcessados === 0) {
      console.error(' Nenhum arquivo CSV processável encontrado!');
      return [];
    }
    
    // Converter para array e ordenar (ordenação simples)
    const enderecosArray = Array.from(enderecos).sort();
    
    console.log(` CARREGAMENTO INSTANTÂNEO CONCLUÍDO:`);
    console.log(`- Arquivos processados: ${arquivosProcessados}`);
    console.log(`- Endereços únicos: ${enderecosArray.length}`);
    console.log(`- Primeiros 10: [${enderecosArray.slice(0, 10).join(', ')}]`);
    console.log(`- Últimos 10: [${enderecosArray.slice(-10).join(', ')}]`);
    
    return enderecosArray;
    
  } catch (error) {
    console.error(' Erro no carregamento instantâneo:', error);
    
    // Fallback para método mais simples ainda
    console.log(' Tentando fallback simples...');
    return extrairEnderecosCDPrincipalFallback();
  }
}

/**
 * Fallback ainda mais simples - apenas um arquivo
 */
function extrairEnderecosCDPrincipalFallback() {
  console.log(' Fallback simples para CD Principal...');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    // Tentar apenas o arquivo mais provável de existir
    const nomeArquivo = 'Base CD Principal - Parte 1.csv';
    const arquivos = pasta.getFilesByName(nomeArquivo);
    
    if (!arquivos.hasNext()) {
      console.log(' Arquivo de fallback não encontrado, retornando endereços de exemplo');
      return gerarEnderecosFallback();
    }
    
    const arquivo = arquivos.next();
    const csv = arquivo.getBlob().getDataAsString();
    const linhas = csv.split('\n');
    
    const enderecos = new Set();
    
    // Processamento mínimo - apenas primeiras 1000 linhas para garantir rapidez
    const maxLinhas = Math.min(1000, linhas.length);
    
    for (let i = 1; i < maxLinhas; i++) {
      const cols = linhas[i].split(';').map(c => c.replace(/"/g,'').trim());
      if (cols.length >= 5 && cols[4]) {
        enderecos.add(cols[4]);
      }
    }
    
    const enderecosArray = Array.from(enderecos).sort();
    console.log(` Fallback: ${enderecosArray.length} endereços extraídos`);
    
    return enderecosArray;
    
  } catch (error) {
    console.error(' Erro no fallback:', error);
    return gerarEnderecosFallback();
  }
}

/**
 * Extrai endereços únicos diretamente do CSV da Base RMA (ou aba local como fallback)
 */
function extrairEnderecosRmaDireto() {
  console.log(' Extraindo endereços da Base RMA (CSV/Planilha)...');
  const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
  const nomeArquivo = 'Base RMA.csv';
  const enderecos = new Set();

  try {
    const pasta = DriveApp.getFolderById(folderId);
    const arquivos = pasta.getFilesByName(nomeArquivo);
    if (arquivos.hasNext()) {
      const csv = arquivos.next().getBlob().getDataAsString();
      const linhas = csv.split('\n');
      for (let i = 1; i < linhas.length; i++) {
        const cols = linhas[i].split(';').map(c => c.replace(/"/g,'').trim());
        if (cols.length >= 5 && cols[4]) enderecos.add(cols[4]);
      }
      console.log(' Endereços RMA via CSV:', enderecos.size);
      return Array.from(enderecos).sort();
    }
  } catch (e) {
    console.warn(' Falha ao ler CSV RMA:', e);
  }

  // Fallback: aba local
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Base RMA');
    if (aba) {
      const dados = aba.getDataRange().getValues();
      dados.slice(1).forEach(row => { if (row[4]) enderecos.add(String(row[4]).trim()); });
      console.log(' Endereços RMA via Planilha:', enderecos.size);
    }
  } catch (e) {
    console.error(' Falha fallback planilha RMA:', e);
  }
  return Array.from(enderecos).sort();
}

/**
 * Carrega dados RMA de um endereço específico para cache do frontend
 * Otimizado para performance - carrega uma vez e mantém no frontend
 */
function carregarDadosRmaPorEndereco(endereco) {
  console.log(` carregarDadosRmaPorEndereco: ${endereco}`);
  
  if (!endereco) {
    return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO']];
  }
  
  try {
    const resultado = buscarProdutosPorEndereco(endereco, 'rma');
    console.log(` Carregados ${resultado.length - 1} produtos RMA para endereço ${endereco}`);
    return resultado;
  } catch (error) {
    console.error(' Erro ao carregar dados RMA por endereço:', error);
    return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO']];
  }
}

/**
 * Retorna os endereços do RMA contendo apenas itens da SEC "SMARTPHONES"
 * Busca no CSV mais recente "PosicaoEstoqueOpenbox*.csv" na pasta configurada
 */
function carregarEnderecosRmaSmartphones() {
  console.log(' carregarEnderecosRmaSmartphones()');
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);

    // Localiza o arquivo "PosicaoEstoqueOpenbox*.csv" mais recente
    const iter = pasta.getFiles();
    let arquivoMaisRecente = null;
    let dataMaisRecente = null;
    while (iter.hasNext()) {
      const arq = iter.next();
      const nome = arq.getName();
      if (nome.startsWith('PosicaoEstoqueOpenbox') && nome.endsWith('.csv')) {
        const dt = arq.getLastUpdated();
        if (!arquivoMaisRecente || dt > dataMaisRecente) {
          arquivoMaisRecente = arq;
          dataMaisRecente = dt;
        }
      }
    }

    if (!arquivoMaisRecente) {
      console.warn(' Nenhum arquivo PosicaoEstoqueOpenbox*.csv encontrado para RMA. Fallback: retorna todos endereços RMA');
      return extrairEnderecosRmaDireto();
    }

    const csv = arquivoMaisRecente.getBlob().getDataAsString();
    const linhas = csv.split('\n');
    if (linhas.length === 0) return [];

    const header = linhas[0].replace(/^\uFEFF/, '').split(';').map(c => c.replace(/"/g,'').trim().toUpperCase());
    const findIdx = (cands) => {
      const set = new Set(cands.map(s => s.toUpperCase()));
      for (let i = 0; i < header.length; i++) {
        if (set.has(header[i])) return i;
      }
      return -1;
    };
  const idxSEC = findIdx(['SEC','SECAO','SEÇÃO','SECAO/SEC','SEÇÃO/SEC']);
    const idxEND = findIdx(['ENDERECO','ENDEREÇO','END','ENDERECO DO PRODUTO','ENDEREÇO DO PRODUTO']);

    if (idxEND < 0) {
      console.warn(' Cabeçalho ENDERECO não encontrado no CSV RMA. Retornando fallback.');
      return extrairEnderecosRmaDireto();
    }

    const enderecos = new Set();
    for (let i = 1; i < linhas.length; i++) {
      const row = linhas[i];
      if (!row || row.trim()==='') continue;
      const cols = row.split(';').map(c => c.replace(/"/g,'').trim());
      const sec = idxSEC >= 0 ? String(cols[idxSEC] || '').toUpperCase() : '';
      const endereco = cols[idxEND];
      if (!endereco) continue;
      // Filtra somente SEC exatamente "SMARTPHONES" (ignora acentos/caixa)
      if (idxSEC >= 0) {
        const secNorm = sec.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        const secClean = secNorm.replace(/\s+/g, ' ').trim();
        if (secClean !== 'SMARTPHONES') continue;
      }
      enderecos.add(endereco);
    }

    const lista = Array.from(enderecos).sort((a,b) => {
      const pa = String(a).split('-').map(n=>parseInt(n,10)||0);
      const pb = String(b).split('-').map(n=>parseInt(n,10)||0);
      for (let i=0;i<4;i++){ if (pa[i]!==pb[i]) return pa[i]-pb[i]; }
      return 0;
    });
    console.log(` Enderecos RMA (Smartphones) encontrados: ${lista.length}`);
    return lista;
  } catch (e) {
    console.error(' Erro em carregarEnderecosRmaSmartphones:', e);
    return [];
  }
}

/**
 * Carrega dados de RMA para um endereço filtrando apenas SEC "SMARTPHONES"
 * Retorna matriz [ [CF,CODIGO,FINAL,DESCRICAO,ENDERECO], ... ]
 */
function carregarDadosRmaSmartphonesPorEndereco(enderecoBusca) {
  console.log(' carregarDadosRmaSmartphonesPorEndereco()', enderecoBusca);
  const cab = [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO']];
  if (!enderecoBusca) return cab;
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    // Localiza CSV mais recente da reversa (PosicaoEstoqueOpenbox*.csv) com cache do arquivo mais recente
    const cache = CacheService.getDocumentCache();
    const rmaKey = 'rma:latestFile';
    let arquivoMaisRecente = null;
    try {
      const meta = JSON.parse(cache.get(rmaKey) || 'null');
      if (meta && meta.id) {
        arquivoMaisRecente = DriveApp.getFileById(meta.id);
      }
    } catch (_) { arquivoMaisRecente = null; }
    if (!arquivoMaisRecente) {
      const iter = pasta.getFiles();
      let dataMaisRecente = null;
      while (iter.hasNext()) {
        const arq = iter.next();
        const nome = arq.getName();
        if (nome.startsWith('PosicaoEstoqueOpenbox') && nome.endsWith('.csv')) {
          const dt = arq.getLastUpdated();
          if (!arquivoMaisRecente || dt > dataMaisRecente) { arquivoMaisRecente = arq; dataMaisRecente = dt; }
        }
      }
      try { if (arquivoMaisRecente) cache.put(rmaKey, JSON.stringify({ id: arquivoMaisRecente.getId() }), 600); } catch (_) {}
    }

    if (!arquivoMaisRecente) {
      console.warn(' CSV reversa não encontrado. Fallback: buscarProdutosRmaPorEndereco (sem filtro SEC)');
      return buscarProdutosRmaPorEndereco(enderecoBusca);
    }

  // Cache de linhas por endereço
    
  // reusa "cache" já criado acima
    const cacheKey = `RMA_SMART:${String(enderecoBusca).trim()}`;
    const hit = cache.get(cacheKey);
    if (hit) {
      try { const parsed = JSON.parse(hit); if (Array.isArray(parsed) && parsed.length>0) { console.log(' Cache HIT', cacheKey); return parsed; } } catch(e){}
    }

    const csv = arquivoMaisRecente.getBlob().getDataAsString();
    const linhas = csv.split('\n');
    if (linhas.length === 0) return cab;

    const header = linhas[0].replace(/^\uFEFF/, '').split(';').map(c=>c.replace(/"/g,'').trim().toUpperCase());
    const findIdx = (cands) => { const set=new Set(cands.map(s=>s.toUpperCase())); for (let i=0;i<header.length;i++){ if (set.has(header[i])) return i; } return -1; };
    // Para RMA: CF = Código Produto | CODIGO = Etiqueta
    let idxCF = findIdx(['CODIGO PRODUTO','CÓDIGO PRODUTO','CD MERCADORIA','CDMERCADORIA','CODIGO','CÓDIGO','COD','CF']);
    let idxCOD = findIdx(['ETIQUETA DO PRODUTO','ETIQUETA PRODUTO','ETIQUETA','ETIQUETA DO PROD', 'CODIGO BARRAS', 'CÓDIGO BARRAS','CODIGO','CÓDIGO']);
    let idxFINAL = findIdx(['FINAL','FINAIS','FIN','SÉRIE','SERIE','SERIAL','IMEI']);
    let idxDESC = findIdx(['DESCRICAO','DESCRIÇÃO','DESCR','DESCRICAO DO PRODUTO','DESCRIÇÃO DO PRODUTO']);
    let idxEND = findIdx(['ENDERECO','ENDEREÇO','END','ENDERECO DO PRODUTO','ENDEREÇO DO PRODUTO']);
  let idxSEC = findIdx(['SEC','SECAO','SEÇÃO','SECAO/SEC','SEÇÃO/SEC']);
    if (idxCF<0) idxCF=0; if (idxCOD<0) idxCOD=1; if (idxFINAL<0) idxFINAL=2; if (idxDESC<0) idxDESC=3; if (idxEND<0) idxEND=4;

    const out = cab.slice();
    let encontrados = 0;
    for (let i=1;i<linhas.length;i++){
      const row = linhas[i];
      if (!row || row.trim()==='') continue;
      const cols = row.split(';').map(c=>c.replace(/"/g,'').trim());
      if (cols.length <= idxEND) continue;
      const endereco = cols[idxEND];
      if (!endereco || String(endereco).trim() !== String(enderecoBusca).trim()) continue;
      // Filtra SEC: somente exatamente "SMARTPHONES" (ignora acentos/caixa)
      if (idxSEC >= 0) {
        const sec = String(cols[idxSEC] || '').toUpperCase();
        const secNorm = sec.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        const secClean = secNorm.replace(/\s+/g, ' ').trim();
        if (secClean !== 'SMARTPHONES') continue;
      }
      out.push([cols[idxCF]||'', cols[idxCOD]||'', cols[idxFINAL]||'', cols[idxDESC]||'', cols[idxEND]||'']);
      encontrados++;
    }
    console.log(` RMA SMART: ${encontrados} linhas para endereço ${enderecoBusca}`);

    // Cache
    try { const json = JSON.stringify(out); if (json.length < 90000) cache.put(cacheKey, json, 300); } catch(e){}
    return out;
  } catch (e) {
    console.error(' Erro em carregarDadosRmaSmartphonesPorEndereco:', e);
    return cab;
  }
}
function buscarProdutosRmaPorEndereco(enderecoBusca) {
  console.log(' buscarProdutosRmaPorEndereco()', enderecoBusca);
  const cab = [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO']];
  if (!enderecoBusca) return cab;

  const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
  const nomeArquivo = 'Base RMA.csv';

  try {
    // Cache por endereço para acelerar chamadas repetidas
    const cache = CacheService.getDocumentCache();
    const cacheKey = `RMA:${String(enderecoBusca).trim()}`;
    const cacheHit = cache.get(cacheKey);
    if (cacheHit) {
      try {
        const parsed = JSON.parse(cacheHit);
        if (Array.isArray(parsed) && parsed.length > 0) {
          console.log(' Cache HIT RMA para', enderecoBusca, 'linhas:', parsed.length-1);
          return parsed;
        }
      } catch (e) { /* ignora parse */ }
    }

    const pasta = DriveApp.getFolderById(folderId);
    const arquivos = pasta.getFilesByName(nomeArquivo);
    if (arquivos.hasNext()) {
      const csv = arquivos.next().getBlob().getDataAsString();
      const linhas = csv.split('\n');
      const out = cab.slice();
      if (linhas.length > 0) {
        const header = linhas[0].split(';').map(c=>c.replace(/"/g,'').trim().toUpperCase());
    const findIdx = (cands) => { const set=new Set(cands.map(s=>s.toUpperCase())); for (let i=0;i<header.length;i++){ if (set.has(header[i])) return i; } return -1; };
    // Para RMA, CF := Código Produto e CODIGO := Etiqueta do Produto
    const cfCands = ['CODIGO PRODUTO','CÓDIGO PRODUTO','CD MERCADORIA','CDMERCADORIA','CODIGO','CÓDIGO','COD','CF'];
    const codCands = ['ETIQUETA DO PRODUTO','ETIQUETA PRODUTO','ETIQUETA','ETIQUETA DO PROD', 'CODIGO BARRAS', 'CÓDIGO BARRAS','CODIGO','CÓDIGO'];
    let idxCF = findIdx(cfCands);
    let idxCOD = findIdx(codCands);
    let idxFINAL = findIdx(['FINAL','FINAIS','FIN','SÉRIE','SERIE','SERIAL']);
    let idxDESC = findIdx(['DESCRICAO','DESCRIÇÃO','DESCR','DESCRICAO DO PRODUTO','DESCRIÇÃO DO PRODUTO']);
    let idxEND = findIdx(['ENDERECO','ENDEREÇO','END','ENDERECO DO PRODUTO','ENDEREÇO DO PRODUTO']);
        if (idxCF<0) idxCF=0; if (idxCOD<0) idxCOD=1; if (idxFINAL<0) idxFINAL=2; if (idxDESC<0) idxDESC=3; if (idxEND<0) idxEND=4;
        for (let i=1;i<linhas.length;i++){
          const cols = linhas[i].split(';').map(c=>c.replace(/"/g,'').trim());
          if (cols.length>=Math.max(idxCF,idxCOD,idxFINAL,idxDESC,idxEND)+1 && String(cols[idxEND]).trim()===String(enderecoBusca).trim()){
            out.push([cols[idxCF], cols[idxCOD], cols[idxFINAL], cols[idxDESC], cols[idxEND]]);
          }
        }
      }
      console.log(' RMA via CSV linhas:', out.length-1);
      // Grava em cache se couber (limite ~100KB)
      try {
        const json = JSON.stringify(out);
        if (json.length < 90000) cache.put(cacheKey, json, 300); // 5 min
      } catch(e) { /* ignora cache */ }
      return out;
    }
  } catch(e){
    console.warn(' Falha CSV RMA:', e);
  }

  // Fallback planilha
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Base RMA');
    if (aba){
      const dados = aba.getDataRange().getValues();
      const out = cab.slice();
      dados.slice(1).forEach(r=>{
        if (String(r[4]).trim()===String(enderecoBusca).trim()) {
          out.push([r[0], r[1], r[2], r[3], r[4]]);
        }
      });
      console.log(' RMA via Planilha linhas:', out.length-1);
      return out;
    }
  } catch(e){
    console.error(' Falha planilha RMA:', e);
  }
  return cab;
}

/**
 * Teste da nova estratégia CSV direto
 */
function testeExtrairCsvDireto() {
  console.log(' === TESTE: CSV DIRETO ===');
  
  const resultados = extrairEnderecosCsvDireto();
  
  console.log(' RESULTADOS DO TESTE:');
  console.log('- Tipo:', typeof resultados);
  console.log('- É array?', Array.isArray(resultados));
  console.log('- Total:', resultados.length);
  
  if (resultados.length > 0) {
    console.log('- Primeiros 10:', resultados.slice(0, 10));
    console.log('- Últimos 10:', resultados.slice(-10));
    
    // Verificar "29-0-2-1"
    if (resultados.includes('29-0-2-1')) {
      console.log(' SUCESSO: "29-0-2-1" encontrado!');
    } else {
      console.log(' "29-0-2-1" não encontrado');
    }
  }
  
  return resultados;
}

/**
 * Busca produtos (linhas CF,CODIGO,FINAL,DESCRICAO,ENDERECO) por endereço diretamente do CSV no Drive
 * Evita depender de amostras em memória no cliente.
 * @param {string} enderecoBusca Endereço no formato N-N-N-N
 * @param {string} tipoEstoque 'principal' | 'rma' (por padrão usa 'principal')
 * @returns {Array<Array<string>>} Matriz com cabeçalho + linhas filtradas
 */
function buscarProdutosPorEndereco(enderecoBusca, tipoEstoque) {
  console.log(' === INICIO buscarProdutosPorEndereco ===');
  console.log(' Parâmetros:', { enderecoBusca, tipoEstoque });
  
  if (!enderecoBusca) return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO','CUSTO']];

  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);

    // Seleciona arquivo conforme o tipo
    let nomeArquivo;
    let arquivos;
    
    if (tipoEstoque === 'rma') {
      // Para RMA: procura pelo arquivo PosicaoEstoqueOpenbox mais recente
      console.log(' Procurando arquivo PosicaoEstoqueOpenbox mais recente...');
      // Cache do arquivo RMA mais recente por 10 minutos para evitar varrer a pasta toda vez
      const cache = CacheService.getDocumentCache();
      const rmaKey = 'rma:latestFile';
      let rmaMeta = null;
      try { rmaMeta = JSON.parse(cache.get(rmaKey) || 'null'); } catch (_) { rmaMeta = null; }

      let arquivoRmaMaisRecente = null;
      if (rmaMeta && rmaMeta.id) {
        try {
          arquivoRmaMaisRecente = DriveApp.getFileById(rmaMeta.id);
          nomeArquivo = arquivoRmaMaisRecente.getName();
          console.log(` RMA (cache) selecionado: "${nomeArquivo}"`);
        } catch (e) { arquivoRmaMaisRecente = null; }
      }

      if (!arquivoRmaMaisRecente) {
        const todosArquivos = pasta.getFiles();
        let dataMaisRecente = null;
        while (todosArquivos.hasNext()) {
          const arquivo = todosArquivos.next();
          const nome = arquivo.getName();
          if (nome.startsWith('PosicaoEstoqueOpenbox') && nome.endsWith('.csv')) {
            const dataModificacao = arquivo.getLastUpdated();
            if (!arquivoRmaMaisRecente || dataModificacao > dataMaisRecente) {
              arquivoRmaMaisRecente = arquivo;
              dataMaisRecente = dataModificacao;
            }
          }
        }
        if (arquivoRmaMaisRecente) {
          try { cache.put(rmaKey, JSON.stringify({ id: arquivoRmaMaisRecente.getId() }), 600); } catch (_) {}
        }
      }

      if (arquivoRmaMaisRecente) {
        nomeArquivo = arquivoRmaMaisRecente.getName();
        console.log(` Arquivo RMA selecionado: "${nomeArquivo}"`);
        arquivos = { hasNext: () => true, next: () => arquivoRmaMaisRecente };
      } else {
        console.log(' Nenhum arquivo PosicaoEstoqueOpenbox encontrado. Usando fallback...');
        nomeArquivo = 'Base CD Principal.csv';
        arquivos = pasta.getFilesByName(nomeArquivo);
      }
    } else {
      // Para Principal: usar busca otimizada em múltiplos arquivos
      console.log(' Usando busca otimizada para Base CD Principal...');
  return buscarProdutosPorEnderecoOtimizado(enderecoBusca, pasta);
    }
    
    console.log(` Arquivo selecionado: "${nomeArquivo}"`);
    
    if (!arquivos.hasNext()) {
      console.error(` Arquivo não encontrado: ${nomeArquivo}`);
      console.log(' Listando arquivos na pasta:');
      const todosArquivos = pasta.getFiles();
      while (todosArquivos.hasNext()) {
        const arquivo = todosArquivos.next();
        console.log(`  - ${arquivo.getName()}`);
      }
  return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO','CUSTO']];
    }

    const arquivoCsv = arquivos.next();
    const tamanhoMB = (arquivoCsv.getSize()/1024/1024).toFixed(2);
    console.log(` Lendo ${nomeArquivo} (${tamanhoMB} MB)`);

    // Cache por endereço + tipo
    const cache = CacheService.getDocumentCache();
    const cacheKey = `${tipoEstoque||'principal'}:${String(enderecoBusca).trim()}`;
    const cacheHit = cache.get(cacheKey);
    if (cacheHit) {
      try {
        const parsed = JSON.parse(cacheHit);
        if (Array.isArray(parsed) && parsed.length > 0) {
          console.log(' Cache HIT para', cacheKey, 'linhas:', parsed.length-1);
          return parsed;
        }
      } catch (e) { /* ignora parse */ }
    }

    const conteudoCsv = arquivoCsv.getBlob().getDataAsString();
    const linhas = conteudoCsv.split('\n');
    console.log(` CSV lido: ${linhas.length} linhas`);

  if (linhas.length === 0) return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO','CUSTO']];

    // Cabeçalho
    const headerLine = (linhas[0] || '').replace(/^\uFEFF/, '');
    const header = headerLine.split(';').map(c=>c.replace(/"/g,'').trim().toUpperCase());
    console.log(` Cabeçalho original: "${linhas[0]}"`);
    console.log(` Cabeçalho parseado: [${header.join(', ')}]`);
    
    const findIdx = (cands) => { 
      const set=new Set(cands.map(s=>s.toUpperCase())); 
      for (let i=0;i<header.length;i++){ 
        if (set.has(header[i])) {
          console.log(` Encontrado cabeçalho "${header[i]}" na posição ${i} (candidatos: ${cands.join(', ')})`);
          return i; 
        }
      } 
      console.log(` Nenhum cabeçalho encontrado para candidatos: ${cands.join(', ')}`);
      return -1; 
    };
    
    const isRma = (tipoEstoque === 'rma');
    console.log(` Tipo RMA: ${isRma}`);
    
    let idxCF = findIdx(isRma ? ['CODIGO PRODUTO','CÓDIGO PRODUTO','CD MERCADORIA','CDMERCADORIA','CODIGO','CÓDIGO','COD','CF'] : ['CF','CD MERCADARIA','CDMERCADORIA','CODIGO','CÓDIGO','COD']);
    let idxCOD = findIdx(isRma ? ['ETIQUETA DO PRODUTO','ETIQUETA PRODUTO','ETIQUETA','ETIQUETA DO PROD', 'CODIGO BARRAS', 'CÓDIGO BARRAS','CODIGO','CÓDIGO'] : ['CODIGO','CÓDIGO','CD MERCADARIA','CDMERCADORIA','COD']);
    let idxFINAL = findIdx([
      'FINAL','FINAL DO PRODUTO','FINAIS','FINAIS DO PRODUTO','FIN',
      'SÉRIE','SERIE','Nº SÉRIE','Nº SERIE','NO SÉRIE','NO SERIE',
      'NUMERO DE SERIE','NÚMERO DE SÉRIE','NUMERO SERIE','NUMERO SÉRIE',
      'SERIAL','SERIAL NUMBER','IMEI'
    ]);
    let idxDESC = findIdx(['DESCRICAO','DESCRIÇÃO','DESCR','DESCRICAO DO PRODUTO','DESCRIÇÃO DO PRODUTO']);
    let idxEND = findIdx(['ENDERECO','ENDEREÇO','END','ENDERECO DO PRODUTO','ENDEREÇO DO PRODUTO']);
    let idxCUSTO = findIdx([
      'CUSTO DO PRODUTO','CUSTO','PRECO','PREÇO','VALOR','VALOR UNITARIO','VALOR UNITÁRIO','CUSTO PRODUTO','CUSTO_UNITARIO','CUSTO UNITARIO','CUSTO UND','CUSTO UNIT'
    ]);
    
    // Fallbacks
    if (idxCF<0) { console.log(' CF não encontrado, usando índice 0'); idxCF=0; }
    if (idxCOD<0) { console.log(' CODIGO não encontrado, usando índice 1'); idxCOD=1; }
    if (idxFINAL<0) { console.log(' FINAL não encontrado, usando índice 2'); idxFINAL=2; }
    if (idxDESC<0) { console.log(' DESCRICAO não encontrado, usando índice 3'); idxDESC=3; }
    if (idxEND<0) { console.log(' ENDERECO não encontrado, usando índice 4'); idxEND=4; }
    
    console.log('🧭 Mapeamento final:', { idxCF, idxCOD, idxFINAL, idxDESC, idxEND });

    // Cabeçalho padrão
  const resultado = [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO','CUSTO']];
    let encontrados = 0;
    let totalLinhasProcessadas = 0;
    
    console.log(` Procurando endereço: "${enderecoBusca}"`);

    for (let i = 1; i < linhas.length; i++) {
      const row = linhas[i];
      if (!row || row.trim() === '') continue;
      
      totalLinhasProcessadas++;
      const colunas = row.split(';').map(col => col.replace(/"/g, '').trim());
      
      // Log das primeiras 3 linhas para debug
      if (totalLinhasProcessadas <= 3) {
        console.log(` Linha ${i}: [${colunas.join(' | ')}]`);
        if (colunas.length > idxEND) {
          console.log(`   Endereço na posição ${idxEND}: "${colunas[idxEND]}"`);
          console.log(`   Match com "${enderecoBusca}": ${String(colunas[idxEND]).trim() === String(enderecoBusca).trim()}`);
        } else {
          console.log(`    Linha muito curta (${colunas.length} colunas, precisa de pelo menos ${idxEND + 1})`);
        }
      }
      
      if (colunas.length <= idxEND) continue;
      const endereco = colunas[idxEND];
      if (!endereco) continue;
      
      if (String(endereco).trim() === String(enderecoBusca).trim()) {
  const linha = [ colunas[idxCF], colunas[idxCOD], colunas[idxFINAL], colunas[idxDESC], colunas[idxEND], (idxCUSTO>=0? colunas[idxCUSTO] : '') ];
        resultado.push(linha);
        encontrados++;
        
        // Log da primeira linha encontrada
        if (encontrados === 1) {
          console.log(` Primeira linha encontrada:`, linha);
        }
      }
    }

    console.log(` Resultado busca:`);
    console.log(`- Total linhas processadas: ${totalLinhasProcessadas}`);
    console.log(`- Linhas encontradas: ${encontrados}`);
    console.log(`- Endereço buscado: "${enderecoBusca}"`);
    
    if (encontrados === 0) {
      console.log(' Amostra de endereços únicos no arquivo (primeiros 10):');
      const enderecosAmostra = new Set();
      for (let i = 1; i < Math.min(100, linhas.length); i++) {
        const row = linhas[i];
        if (!row) continue;
        const colunas = row.split(';').map(col => col.replace(/"/g, '').trim());
        if (colunas.length > idxEND && colunas[idxEND]) {
          enderecosAmostra.add(colunas[idxEND]);
          if (enderecosAmostra.size >= 10) break;
        }
      }
      console.log(`   [${Array.from(enderecosAmostra).join(', ')}]`);
    }

    // Grava em cache se couber
    try {
      const json = JSON.stringify(resultado);
      if (json.length < 90000) cache.put(cacheKey, json, 300);
    } catch(e) { /* ignora cache */ }
    
    console.log(' === FIM buscarProdutosPorEndereco ===');
    return resultado;
  } catch (err) {
    console.error(' Erro em buscarProdutosPorEndereco:', err);
    return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO']];
  }
}

/**
 * Busca produtos por endereço de forma otimizada para múltiplos arquivos CSV
 * @param {string} enderecoBusca - Endereço a ser buscado
 * @param {GoogleAppsScript.Drive.Folder} pasta - Pasta do Drive contendo os arquivos
 * @returns {Array} Array com os produtos encontrados
 */
function buscarProdutosPorEnderecoOtimizado(enderecoBusca, pasta) {
  console.log(' === BUSCA OTIMIZADA PARA MÚLTIPLOS CSVS ===');
  console.log(` Buscando endereço: "${enderecoBusca}"`);
  
  // Cache por endereço
  const cache = CacheService.getDocumentCache();
  const cacheKey = `principal_otimizado:${String(enderecoBusca).trim()}`;
  const cacheHit = cache.get(cacheKey);
  if (cacheHit) {
    try {
      const parsed = JSON.parse(cacheHit);
      if (Array.isArray(parsed) && parsed.length > 0) {
        console.log(' Cache HIT para busca otimizada, linhas:', parsed.length-1);
        return parsed;
      }
    } catch (e) { /* ignora parse */ }
  }

  // Lista de possíveis nomes de arquivo para Base CD Principal
  const nomesArquivos = [
    'Base CD Principal.csv',
    'Base CD Principal - Parte 1.csv',
    'Base CD Principal - Parte 2.csv',
    'Base CD Principal - Parte1.csv',
    'Base CD Principal - Parte2.csv',
    'Base_CD_Principal_Parte1.csv',
    'Base_CD_Principal_Parte2.csv'
  ];

  let arquivosEncontrados = [];
  
  // Buscar todos os arquivos disponíveis
  console.log(' Procurando arquivos CSV da Base CD Principal...');
  for (const nomeArquivo of nomesArquivos) {
    const arquivos = pasta.getFilesByName(nomeArquivo);
    if (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const tamanhoMB = (arquivo.getSize() / 1024 / 1024).toFixed(2);
      console.log(` Encontrado: ${nomeArquivo} (${tamanhoMB} MB)`);
      arquivosEncontrados.push({ nome: nomeArquivo, arquivo: arquivo, tamanho: parseFloat(tamanhoMB) });
    }
  }

  if (arquivosEncontrados.length === 0) {
    console.error(' Nenhum arquivo CSV da Base CD Principal encontrado!');
    return [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO']];
  }

  // Ordenar por tamanho (menores primeiro para eficiência)
  arquivosEncontrados.sort((a, b) => a.tamanho - b.tamanho);
  console.log(` Total de arquivos encontrados: ${arquivosEncontrados.length}`);

  const resultado = [['CF','CODIGO','FINAL','DESCRICAO','ENDERECO','CUSTO']];
  let totalEncontrados = 0;
  let headerMapeado = null;

  // Processar cada arquivo
  for (const {nome, arquivo, tamanho} of arquivosEncontrados) {
    console.log(` Processando ${nome} (${tamanho} MB)...`);
    
    try {
      // Verificar se arquivo não é muito grande
      if (tamanho > 45) {
        console.warn(` Arquivo ${nome} muito grande (${tamanho} MB), pulando...`);
        continue;
      }

      const conteudoCsv = arquivo.getBlob().getDataAsString();
      const linhas = conteudoCsv.split('\n');
      console.log(`   ${linhas.length} linhas carregadas`);

      if (linhas.length === 0) continue;

      // Mapear cabeçalho apenas uma vez
      if (!headerMapeado) {
        headerMapeado = mapearCabecalhoCSV(linhas[0]);
        console.log('🧭 Cabeçalho mapeado:', headerMapeado);
      }

      // Buscar endereço específico no arquivo
      let encontradosNoArquivo = 0;
      
      for (let i = 1; i < linhas.length; i++) {
        const row = linhas[i];
        if (!row || row.trim() === '') continue;
        
        const colunas = row.split(';').map(col => col.replace(/"/g, '').trim());
        
        if (colunas.length <= headerMapeado.idxEND) continue;
        const endereco = colunas[headerMapeado.idxEND];
        if (!endereco) continue;
        
        if (String(endereco).trim() === String(enderecoBusca).trim()) {
          const linha = [
            colunas[headerMapeado.idxCF] || '',
            colunas[headerMapeado.idxCOD] || '',
            colunas[headerMapeado.idxFINAL] || '',
            colunas[headerMapeado.idxDESC] || '',
            colunas[headerMapeado.idxEND] || '',
            headerMapeado.idxCUSTO != null && headerMapeado.idxCUSTO >= 0 ? (colunas[headerMapeado.idxCUSTO] || '') : ''
          ];
          resultado.push(linha);
          encontradosNoArquivo++;
          totalEncontrados++;
        }
      }
      
      console.log(`   Encontrados neste arquivo: ${encontradosNoArquivo}`);
      
    } catch (error) {
      console.error(` Erro ao processar ${nome}:`, error);
      continue;
    }
  }

  console.log(` RESULTADO FINAL:`);
  console.log(`- Total de arquivos processados: ${arquivosEncontrados.length}`);
  console.log(`- Total de produtos encontrados: ${totalEncontrados}`);
  console.log(`- Endereço buscado: "${enderecoBusca}"`);

  // Grava em cache se couber
  try {
    const json = JSON.stringify(resultado);
    if (json.length < 90000) {
      cache.put(cacheKey, json, 300);
      console.log(' Resultado salvo em cache');
    }
  } catch(e) { 
    console.log(' Não foi possível salvar em cache:', e.message);
  }
  
  console.log(' === FIM BUSCA OTIMIZADA ===');
  return resultado;
}

/**
 * Mapeia o cabeçalho do CSV para encontrar as colunas corretas
 * @param {string} headerLine - Primeira linha do CSV (cabeçalho)
 * @returns {Object} Objeto com os índices das colunas mapeadas
 */
function mapearCabecalhoCSV(headerLine) {
  const headerCleaned = headerLine.replace(/^\uFEFF/, '');
  const header = headerCleaned.split(';').map(c => c.replace(/"/g, '').trim().toUpperCase());
  
  console.log(` Cabeçalho detectado: [${header.join(', ')}]`);
  
  const findIdx = (candidatos) => {
    const set = new Set(candidatos.map(s => s.toUpperCase()));
    for (let i = 0; i < header.length; i++) {
      if (set.has(header[i])) {
        console.log(` Encontrado "${header[i]}" na posição ${i}`);
        return i;
      }
    }
    console.log(` Nenhum cabeçalho encontrado para: ${candidatos.join(', ')}`);
    return -1;
  };

  let idxCF = findIdx(['CF', 'CD MERCADARIA', 'CDMERCADORIA', 'CODIGO', 'CÓDIGO', 'COD']);
  let idxCOD = findIdx(['CODIGO', 'CÓDIGO', 'CD MERCADARIA', 'CDMERCADORIA', 'COD']);
  let idxFINAL = findIdx([
    'FINAL', 'FINAL DO PRODUTO', 'FINAIS', 'FINAIS DO PRODUTO', 'FIN',
    'SÉRIE', 'SERIE', 'Nº SÉRIE', 'Nº SERIE', 'NO SÉRIE', 'NO SERIE',
    'NUMERO DE SERIE', 'NÚMERO DE SÉRIE', 'NUMERO SERIE', 'NUMERO SÉRIE',
    'SERIAL', 'SERIAL NUMBER', 'IMEI'
  ]);
  let idxDESC = findIdx(['DESCRICAO', 'DESCRIÇÃO', 'DESCR', 'DESCRICAO DO PRODUTO', 'DESCRIÇÃO DO PRODUTO']);
  let idxEND = findIdx(['ENDERECO', 'ENDEREÇO', 'END', 'ENDERECO DO PRODUTO', 'ENDEREÇO DO PRODUTO']);
  let idxCUSTO = findIdx(['CUSTO DO PRODUTO','CUSTO','PRECO','PREÇO','VALOR','VALOR UNITARIO','VALOR UNITÁRIO','CUSTO PRODUTO','CUSTO_UNITARIO','CUSTO UNITARIO','CUSTO UND','CUSTO UNIT']);
  
  // Fallbacks com base na estrutura padrão
  if (idxCF < 0) { console.log(' CF não encontrado, usando índice 0'); idxCF = 0; }
  if (idxCOD < 0) { console.log(' CODIGO não encontrado, usando índice 1'); idxCOD = 1; }
  if (idxFINAL < 0) { console.log(' FINAL não encontrado, usando índice 2'); idxFINAL = 2; }
  if (idxDESC < 0) { console.log(' DESCRICAO não encontrado, usando índice 3'); idxDESC = 3; }
  if (idxEND < 0) { console.log(' ENDERECO não encontrado, usando índice 4'); idxEND = 4; }
  
  return { idxCF, idxCOD, idxFINAL, idxDESC, idxEND, idxCUSTO };
}

/**
 * Carrega endereços únicos de múltiplos arquivos CSV da Base CD Principal
 * ATUALIZADA: Agora usa a versão instantânea para melhor performance
 * @returns {Array} Array com endereços únicos ordenados
 */
function carregarEnderecosCDPrincipalOtimizado() {
  console.log(' === REDIRECIONANDO PARA VERSÃO INSTANTÂNEA ===');
  
  //  Usar a nova versão instantânea ao invés da complexa
  return extrairEnderecosCDPrincipalInstantaneo();
}

/**
 * Extrai TODOS os endereços dos CSVs de forma otimizada para múltiplos arquivos
 * ATUALIZADA: Agora usa a versão instantânea
 * @returns {Array} Array com todos os endereços únicos ordenados
 */
function extrairTodosEnderecosCsvCompleto() {
  console.log(' === EXTRAÇÃO COMPLETA DE ENDEREÇOS (INSTANTÂNEA) ===');
  
  try {
    //  Usar a função instantânea diretamente
    return extrairEnderecosCDPrincipalInstantaneo();
    
  } catch (error) {
    console.error(' Erro na extração completa de endereços:', error);
    
    // Fallback: tentar método ainda mais simples
    console.log(' Tentando fallback simples...');
    return extrairEnderecosCDPrincipalFallback();
  }
}

/**
 * Carrega uma amostra pequena da Base CD Principal para compatibilidade inicial
 * @returns {Array} Array com amostra dos dados (cabeçalho + algumas linhas)
 */
function carregarAmostraBaseCDPrincipal() {
  console.log(' Carregando amostra inicial...');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    // Lista de possíveis arquivos, priorizando os menores
    const nomesArquivos = [
      'Base CD Principal - Parte 1.csv',
      'Base CD Principal - Parte1.csv',
      'Base_CD_Principal_Parte1.csv',
      'Base CD Principal.csv'
    ];

    for (const nomeArquivo of nomesArquivos) {
      const arquivos = pasta.getFilesByName(nomeArquivo);
      if (arquivos.hasNext()) {
        const arquivo = arquivos.next();
        const tamanhoMB = (arquivo.getSize() / 1024 / 1024).toFixed(2);
        
        console.log(` Tentando arquivo: ${nomeArquivo} (${tamanhoMB} MB)`);
        
        // Só processar arquivos pequenos para amostra
        if (parseFloat(tamanhoMB) <= 30) {
          const conteudoCsv = arquivo.getBlob().getDataAsString();
          const linhas = conteudoCsv.split('\n');
          
          if (linhas.length > 0) {
            // Cabeçalho + primeiras 100 linhas
            const amostra = linhas.slice(0, Math.min(101, linhas.length))
              .map(linha => linha.split(';').map(col => col.replace(/"/g, '').trim()));
            
            console.log(` Amostra carregada: ${amostra.length - 1} registros de dados`);
            return amostra;
          }
        } else {
          console.log(` Arquivo muito grande para amostra: ${tamanhoMB} MB`);
        }
      }
    }
    
    // Se nenhum arquivo foi processado, retornar estrutura básica
    console.log(' Nenhum arquivo processável encontrado, retornando estrutura básica');
    return [['CF', 'CODIGO', 'FINAL', 'DESCRICAO', 'ENDERECO']];
    
  } catch (error) {
    console.error(' Erro ao carregar amostra:', error);
    return [['CF', 'CODIGO', 'FINAL', 'DESCRICAO', 'ENDERECO']];
  }
}

/**
 * Gera endereços de fallback realistas para casos de erro
 * @returns {Array} Array com endereços de exemplo
 */
function gerarEnderecosFallback() {
  console.log(' Gerando endereços de fallback...');
  
  const enderecosFallback = [];
  
  // Gerar alguns endereços de exemplo realistas
  for (let corredor = 1; corredor <= 50; corredor++) {
    for (let nivel = 0; nivel <= 2; nivel++) {
      for (let prateleira = 1; prateleira <= 20; prateleira++) {
        for (let posicao = 1; posicao <= 4; posicao++) {
          enderecosFallback.push(`${corredor}-${nivel}-${prateleira}-${posicao}`);
          
          // Limitar para não gerar milhares de endereços
          if (enderecosFallback.length >= 1000) {
            break;
          }
        }
        if (enderecosFallback.length >= 1000) break;
      }
      if (enderecosFallback.length >= 1000) break;
    }
    if (enderecosFallback.length >= 1000) break;
  }
  
  console.log(` Gerados ${enderecosFallback.length} endereços de fallback`);
  return enderecosFallback;
}

/**
 * Função de teste para comparar performance: RMA vs CD Principal
 */
function testarPerformanceRMAvsCD() {
  console.log(' === TESTE DE PERFORMANCE RMA vs CD PRINCIPAL ===');
  
  try {
    // Teste 1: RMA (já rápido)
    console.log('1️⃣ Testando carregamento RMA (referência)...');
    const inicioRMA = new Date().getTime();
    const enderecosRMA = extrairEnderecosRmaDireto();
    const tempoRMA = new Date().getTime() - inicioRMA;
    console.log(` RMA: ${enderecosRMA.length} endereços em ${tempoRMA}ms`);
    
    // Teste 2: CD Principal (nova versão instantânea)
    console.log('2️⃣ Testando CD Principal (nova versão instantânea)...');
    const inicioCD = new Date().getTime();
    const enderecosCD = extrairEnderecosCDPrincipalInstantaneo();
    const tempoCD = new Date().getTime() - inicioCD;
    console.log(` CD Principal: ${enderecosCD.length} endereços em ${tempoCD}ms`);
    
    // Comparação
    console.log(' === COMPARAÇÃO DE PERFORMANCE ===');
    console.log(`RMA:          ${enderecosRMA.length} endereços em ${tempoRMA}ms`);
    console.log(`CD Principal: ${enderecosCD.length} endereços em ${tempoCD}ms`);
    
    const diferencaTempo = tempoCD - tempoRMA;
    const percentualDiferenca = ((diferencaTempo / tempoRMA) * 100).toFixed(1);
    
    if (diferencaTempo > 0) {
      console.log(` CD Principal é ${diferencaTempo}ms mais lento (${percentualDiferenca}%)`);
    } else {
      console.log(` CD Principal é ${Math.abs(diferencaTempo)}ms mais rápido (${Math.abs(percentualDiferenca)}%)`);
    }
    
    // Verificar se ambos retornaram dados válidos
    const rmaValido = enderecosRMA && enderecosRMA.length > 0;
    const cdValido = enderecosCD && enderecosCD.length > 0;
    
    console.log(' === VALIDAÇÃO DOS RESULTADOS ===');
    console.log(`RMA válido: ${rmaValido ? '' : ''}`);
    console.log(`CD Principal válido: ${cdValido ? '' : ''}`);
    
    if (rmaValido && cdValido) {
      console.log(' AMBOS OS CARREGAMENTOS FUNCIONARAM!');
    }
    
    return {
      rma: { enderecos: enderecosRMA.length, tempo: tempoRMA },
      cd: { enderecos: enderecosCD.length, tempo: tempoCD },
      diferencaTempo: diferencaTempo,
      percentualDiferenca: percentualDiferenca
    };
    
  } catch (error) {
    console.error(' Erro no teste de performance:', error);
    return { erro: error.message };
  }
}

/**
 * Função de teste para verificar a otimização de múltiplos CSVs
 */
function testarOtimizacaoMultiplosCSVs() {
  console.log(' === TESTE OTIMIZAÇÃO MÚLTIPLOS CSVs ===');
  
  try {
    // Teste 1: Carregamento de endereços otimizado
    console.log('1️⃣ Testando carregamento otimizado de endereços...');
    const enderecos = carregarEnderecosCDPrincipalOtimizado();
    console.log(` Endereços carregados: ${enderecos ? enderecos.length : 'null'}`);
    
    if (enderecos && enderecos.length > 0) {
      console.log(' Primeiros 10:', enderecos.slice(0, 10));
      console.log(' Últimos 10:', enderecos.slice(-10));
    }
    
    // Teste 2: Busca otimizada por endereço específico
    console.log('2️⃣ Testando busca otimizada por endereço...');
    const enderecoDeTeste = enderecos && enderecos.length > 0 ? enderecos[0] : '1-0-1-1';
    console.log(` Buscando produtos no endereço: ${enderecoDeTeste}`);
    
    const resultadoBusca = buscarProdutosPorEndereco(enderecoDeTeste, 'principal');
    console.log(` Produtos encontrados: ${resultadoBusca ? resultadoBusca.length - 1 : 'null'}`);
    
    if (resultadoBusca && resultadoBusca.length > 1) {
      console.log(' Primeira linha encontrada:', resultadoBusca[1]);
    }
    
    // Teste 3: Carregamento de amostra
    console.log('3️⃣ Testando carregamento de amostra...');
    const amostra = carregarAmostraBaseCDPrincipal();
    console.log(` Linhas na amostra: ${amostra ? amostra.length : 'null'}`);
    
    // Teste 4: Endereços completos
    console.log('4️⃣ Testando extração completa de endereços...');
    const enderecosCompletos = extrairTodosEnderecosCsvCompleto();
    console.log(` Total de endereços únicos: ${enderecosCompletos ? enderecosCompletos.length : 'null'}`);
    
    console.log(' === FIM TESTE OTIMIZAÇÃO ===');
    return {
      enderecosCarregados: enderecos ? enderecos.length : 0,
      produtosEncontrados: resultadoBusca ? resultadoBusca.length - 1 : 0,
      amostraCarregada: amostra ? amostra.length : 0,
      enderecosCompletos: enderecosCompletos ? enderecosCompletos.length : 0
    };
    
  } catch (error) {
    console.error(' Erro no teste:', error);
    return { erro: error.message };
  }
}

/**
 * DOCUMENTAÇÃO: Como dividir CSV grande em múltiplas partes
 * 
 * Para resolver o problema de CSVs maiores que 50MB no Google Apps Script:
 * 
 * 1. DIVIDIR O ARQUIVO:
 *    - Abra o CSV original no Excel/Google Sheets
 *    - Copie o cabeçalho (primeira linha) 
 *    - Divida os dados em 2 ou mais partes
 *    - Salve como:
 *      • "Base CD Principal - Parte 1.csv" (com cabeçalho + primeira metade)
 *      • "Base CD Principal - Parte 2.csv" (com cabeçalho + segunda metade)
 * 
 * 2. UPLOAD NO DRIVE:
 *    - Faça upload dos arquivos na mesma pasta (ID: 1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE)
 *    - Mantenha o arquivo original se necessário
 * 
 * 3. NOMES SUPORTADOS:
 *    O sistema procura automaticamente por estes nomes:
 *    • Base CD Principal.csv (arquivo único)
 *    • Base CD Principal - Parte 1.csv
 *    • Base CD Principal - Parte 2.csv
 *    • Base CD Principal - Parte1.csv
 *    • Base CD Principal - Parte2.csv
 *    • Base_CD_Principal_Parte1.csv
 *    • Base_CD_Principal_Parte2.csv
 * 
 * 4. FUNCIONAMENTO:
 *    - O sistema detecta automaticamente quantas partes existem
 *    - Carrega apenas arquivos até 45MB para evitar timeout
 *    - Busca em todos os arquivos quando necessário
 *    - Combina resultados automaticamente
 *    - Usa cache para otimizar performance
 * 
 * 5. TESTE:
 *    Execute testarOtimizacaoMultiplosCSVs() para verificar funcionamento
 */

/**
 * Verifica se um produto existe em todas as bases (OTIMIZADO PARA VELOCIDADE)
 * Busca primeiro nas abas locais (instantâneo) depois no CSV (mais lento)
 * @param {string} codigo - Código/etiqueta do produto a ser verificado
 * @returns {Object} Resultado da verificação com dados do produto se encontrado
 */
function verificarProdutoEmTodasAsBases(codigo) {
  console.log(` === VERIFICAÇÃO INSTANTÂNEA PRODUTO SETADO: ${codigo} ===`);
  
  if (!codigo || codigo.trim() === '') {
    return { encontrado: false, motivo: 'Código vazio' };
  }
  
  const codigoLimpo = codigo.toString().trim();
  
  try {
    // =============================================
    // FASE 1: BUSCA INSTANTÂNEA NAS ABAS LOCAIS
    // =============================================
    
    console.log(' FASE 1: Buscando nas abas locais (instantâneo)...');
    
    // 1.1 - Base Smartphones CD Principal (aba local)
    console.log('1.1️⃣ Verificando aba "Base Smartphones CD Principal"...');
    const resultadoSmartCD = buscarNaAbaLocal('Base Smartphones CD Principal', codigoLimpo);
    if (resultadoSmartCD.encontrado) {
      console.log(' Produto encontrado na aba Base Smartphones CD Principal!');
      return {
        encontrado: true,
        base: 'Base Smartphones CD Principal (aba)',
        dados: resultadoSmartCD.dados
      };
    }
    
    // 1.2 - Base RMA (aba local)
    console.log('1.2️⃣ Verificando aba "Base RMA"...');
    const resultadoRMAAba = buscarNaAbaLocal('Base RMA', codigoLimpo);
    if (resultadoRMAAba.encontrado) {
      console.log(' Produto encontrado na aba Base RMA!');
      return {
        encontrado: true,
        base: 'Base RMA (aba)',
        dados: resultadoRMAAba.dados
      };
    }
    
    // 1.3 - Base Transitórios (abas locais - Parte 1 e 2)
    console.log('1.3️⃣ Verificando abas "Base Transitorios - Parte 1" e "Base Transitorios - Parte 2"...');
    
    // Verificar Base Transitorios - Parte 1
    console.log('1.3.1️⃣ Verificando "Base Transitorios - Parte 1"...');
    const resultadoTransitorios1 = buscarNaAbaLocal('Base Transitorios - Parte 1', codigoLimpo);
    if (resultadoTransitorios1.encontrado) {
      console.log(' Produto encontrado na aba Base Transitórios - Parte 1!');
      return {
        encontrado: true,
        base: 'Base Transitórios - Parte 1 (aba)',
        dados: resultadoTransitorios1.dados
      };
    }
    
    // Verificar Base Transitorios - Parte 2
    console.log('1.3.2️⃣ Verificando "Base Transitorios - Parte 2"...');
    const resultadoTransitorios2 = buscarNaAbaLocal('Base Transitorios - Parte 2', codigoLimpo);
    if (resultadoTransitorios2.encontrado) {
      console.log(' Produto encontrado na aba Base Transitórios - Parte 2!');
      return {
        encontrado: true,
        base: 'Base Transitórios - Parte 2 (aba)',
        dados: resultadoTransitorios2.dados
      };
    }
    
    // 1.4 - Base NL (aba local)
    console.log('1.4️⃣ Verificando aba "Base NL"...');
    const resultadoNL = buscarNaAbaLocal('Base NL', codigoLimpo);
    if (resultadoNL.encontrado) {
      console.log(' Produto encontrado na aba Base NL!');
      return {
        encontrado: true,
        base: 'Base NL (aba)',
        dados: resultadoNL.dados
      };
    }
    
    // =============================================
    // FASE 2: BUSCA NO CSV DO DRIVE (MAIS LENTA)
    // =============================================
    
    console.log(' FASE 2: Buscando no CSV do Drive...');
    
    // 2.1 - Base CD Principal (CSV do Drive) - BUSCA OTIMIZADA
    console.log('2.1️⃣ Verificando Base CD Principal (CSV)...');
    const resultadoCD = buscarNaBaseCDPrincipalCSV(codigoLimpo);
    if (resultadoCD.encontrado) {
      console.log(' Produto encontrado na Base CD Principal (CSV)!');
      return {
        encontrado: true,
        base: 'Base CD Principal (CSV)',
        dados: resultadoCD.dados
      };
    }
    
    // =============================================
    // RESULTADO: NÃO ENCONTRADO - POSSÍVEL SETADO
    // =============================================
    
    console.log(' Produto NÃO encontrado em nenhuma base - Possível setado!');
    return {
      encontrado: false,
      motivo: 'Produto não localizado em nenhuma base',
      possivel_setado: true
    };
    
  } catch (error) {
    console.error(' Erro ao verificar produto em todas as bases:', error);
    return {
      encontrado: false,
      erro: error.message,
      motivo: 'Erro na verificação'
    };
  }
}

/**
 * Busca um produto em uma aba local da planilha (INSTANTÂNEO)
 * @param {string} nomeAba - Nome da aba para buscar
 * @param {string} codigo - Código do produto
 * @returns {Object} {encontrado: boolean, dados: Object}
 */
function buscarNaAbaLocal(nomeAba, codigo) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(nomeAba);
    
    if (!aba) {
      console.log(` Aba "${nomeAba}" não encontrada`);
      return { encontrado: false };
    }
    
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) {
      console.log(` Aba "${nomeAba}" está vazia ou só tem cabeçalho`);
      return { encontrado: false };
    }
    
    console.log(` Aba "${nomeAba}": ${dados.length - 1} linhas para verificar`);
    
    // Busca otimizada - verifica múltiplas colunas onde o código pode estar
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      
      // Verificar nas primeiras 5 colunas (onde geralmente estão CF, CODIGO, FINAL, DESCRICAO, ENDERECO)
      for (let j = 0; j < Math.min(5, linha.length); j++) {
        const valorCelula = String(linha[j] || '').trim();
        
        if (valorCelula === codigo) {
          console.log(` Código "${codigo}" encontrado na aba "${nomeAba}", linha ${i + 1}, coluna ${j + 1}`);
          
          return {
            encontrado: true,
            dados: {
              cf: String(linha[0] || '').trim(),
              codigo: String(linha[1] || '').trim(),
              final: String(linha[2] || '').trim(),
              descricao: String(linha[3] || '').trim(),
              endereco: String(linha[4] || '').trim(),
              linhaEncontrada: i + 1,
              colunaEncontrada: j + 1
            }
          };
        }
      }
    }
    
    console.log(` Código "${codigo}" não encontrado na aba "${nomeAba}"`);
    return { encontrado: false };
    
  } catch (error) {
    console.error(` Erro ao buscar na aba "${nomeAba}":`, error);
    return { encontrado: false, erro: error.message };
  }
}

/**
 * Busca OTIMIZADA na Base CD Principal (CSV do Drive)
 * @param {string} codigo - Código do produto
 * @returns {Object} {encontrado: boolean, dados: Object}
 */
function buscarNaBaseCDPrincipalCSV(codigo) {
  try {
    console.log(` Buscando código "${codigo}" na Base CD Principal (CSV)...`);
    
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // Lista de nomes possíveis (priorizando verificação de AMBAS as partes)
    const nomesPossiveis = [
      'Base CD Principal - Parte 1.csv',
      'Base CD Principal - Parte1.csv',
      'Base CD Principal - Parte 2.csv', 
      'Base CD Principal - Parte2.csv',
      'Base CD Principal.csv'
    ];
    
    // CORREÇÃO: Buscar em TODOS os arquivos disponíveis (Parte 1 + Parte 2)
    let arquivosEncontrados = 0;
    let arquivosProcessados = 0;
    
    for (const nome of nomesPossiveis) {
      console.log(`\n 🔍 [${arquivosProcessados + 1}/${nomesPossiveis.length}] Verificando: "${nome}"`);
      const arquivos = pasta.getFilesByName(nome);
      
      if (!arquivos.hasNext()) {
        console.log(`   ❌ "${nome}" não encontrado no Drive`);
        continue;
      }
      
      const arquivo = arquivos.next();
      const tamanhoMB = arquivo.getSize() / 1024 / 1024;
      arquivosEncontrados++;
      console.log(`   ✅ "${nome}" encontrado (${tamanhoMB.toFixed(2)} MB)`);
      
      // Ajustar limite para arquivos da Parte 2 (podem ser maiores)
      if (tamanhoMB > 80) {
        console.log(`   ⚠️ Arquivo muito grande (${tamanhoMB.toFixed(2)} MB > 80MB), pulando por segurança...`);
        continue;
      }
      
      // Buscar neste arquivo específico
      console.log(`   📂 Iniciando busca por código "${codigo}" em "${nome}"...`);
      
      try {
        const resultado = buscarCodigoNoArquivo(arquivo, codigo, nome);
        arquivosProcessados++;
        
        if (resultado.encontrado) {
          console.log(`   🎯 ✅ CÓDIGO ENCONTRADO EM "${nome}"!`);
          console.log(`   📊 Resumo: ${arquivosEncontrados} arquivos encontrados, ${arquivosProcessados} processados`);
          return resultado;
        } else {
          console.log(`   ❌ Código "${codigo}" não encontrado em "${nome}"`);
        }
      } catch (erro) {
        console.error(`   💥 Erro ao processar "${nome}":`, erro.message);
        continue;
      }
    }
    
    console.log(`\n 📊 BUSCA COMPLETA:`);
    console.log(`   - Arquivos encontrados: ${arquivosEncontrados}`);
    console.log(`   - Arquivos processados: ${arquivosProcessados}`);
    console.log(`   - Resultado: Código "${codigo}" NÃO encontrado em nenhum arquivo`);
    
    console.log(' ❌ Código não encontrado em nenhum arquivo da Base CD Principal');
    return { encontrado: false, erro: 'Código não encontrado em nenhuma parte' };

    
  } catch (error) {
    console.error(' Erro ao buscar na Base CD Principal (CSV):', error);
    return { encontrado: false, erro: error.message };
  }
}

/**
 * Função auxiliar para buscar código em um arquivo específico
 * @param {DriveApp.File} arquivo - Arquivo do Drive para buscar
 * @param {string} codigo - Código a buscar
 * @param {string} nomeArquivo - Nome do arquivo (para logs)
 * @returns {Object} {encontrado: boolean, dados: Object}
 */
function buscarCodigoNoArquivo(arquivo, codigo, nomeArquivo) {
  try {
    console.log(`   📖 Lendo conteúdo de "${nomeArquivo}"...`);
    const conteudo = arquivo.getBlob().getDataAsString('UTF-8');
    const linhas = conteudo.split('\n');
    
    console.log(`   📊 Processando ${linhas.length} linhas em "${nomeArquivo}"...`);
    
    // Busca linha por linha
    for (let i = 1; i < linhas.length; i++) { // Pular cabeçalho
      const linha = linhas[i];
      if (!linha || linha.trim() === '') continue;
      
      // Detectar separador (pode ser ; ou ,)
      const temPontoVirgula = linha.includes(';');
      const separador = temPontoVirgula ? ';' : ',';
      
      const colunas = linha.split(separador).map(col => col.replace(/"/g, '').trim());
      
      if (colunas.length >= 2) {
        // Verificar nas primeiras 5 colunas (mais comum: CF, Código, Final, Descrição, Endereço)
        for (let j = 0; j < Math.min(5, colunas.length); j++) {
          const valorCelula = String(colunas[j] || '').trim();
          const codigoLimpo = String(codigo || '').trim();
          
          // Comparação exata primeiro
          if (valorCelula === codigoLimpo) {
            console.log(`   🎯 ✅ MATCH EXATO: "${codigoLimpo}" encontrado em "${nomeArquivo}"`);
            console.log(`   📍 Localização: linha ${i + 1}, coluna ${j + 1} (${separador}-separado)`);
            console.log(`   📝 Dados da linha: [${colunas.slice(0, 5).join(' | ')}]`);
            
            return {
              encontrado: true,
              dados: {
                cf: colunas[0] || '',
                codigo: colunas[1] || codigoLimpo,
                final: colunas[2] || '',
                descricao: colunas[3] || '',
                endereco: colunas[4] || '',
                linhaEncontrada: i + 1,
                colunaEncontrada: j + 1,
                separadorUsado: separador,
                arquivo: nomeArquivo,
                linhaCrua: linha.substring(0, 100) + '...' // Primeiros 100 chars para debug
              }
            };
          }
        }
      }
    }
    
    console.log(`   ❌ Código "${codigo}" não encontrado em "${nomeArquivo}"`);
    return { encontrado: false };
    
  } catch (error) {
    console.error(`   ❌ Erro ao processar "${nomeArquivo}":`, error.message);
    return { encontrado: false, erro: error.message };
  }
}

/**
 * Função de teste para verificar se ambas as partes estão disponíveis
 * @returns {Object} Status dos arquivos
 */
function testarDisponibilidadeArquivos() {
  console.log('\n=== TESTE DE DISPONIBILIDADE DOS ARQUIVOS ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    const arquivosParaTestar = [
      'Base CD Principal - Parte 1.csv',
      'Base CD Principal - Parte 2.csv'
    ];
    
    const status = {};
    
    arquivosParaTestar.forEach(nome => {
      console.log(`\n🔍 Testando: "${nome}"`);
      const arquivos = pasta.getFilesByName(nome);
      
      if (arquivos.hasNext()) {
        const arquivo = arquivos.next();
        const tamanhoMB = arquivo.getSize() / 1024 / 1024;
        const ultimaModificacao = arquivo.getLastUpdated();
        
        status[nome] = {
          existe: true,
          tamanhoMB: tamanhoMB.toFixed(2),
          ultimaModificacao: ultimaModificacao.toLocaleString('pt-BR'),
          processavel: tamanhoMB <= 80
        };
        
        console.log(`✅ ENCONTRADO:`);
        console.log(`   📊 Tamanho: ${tamanhoMB.toFixed(2)} MB`);
        console.log(`   📅 Última modificação: ${ultimaModificacao.toLocaleString('pt-BR')}`);
        console.log(`   ⚙️ Processável: ${tamanhoMB <= 80 ? 'SIM' : 'NÃO (muito grande)'}`);
        
      } else {
        status[nome] = { existe: false };
        console.log(`❌ NÃO ENCONTRADO`);
      }
    });
    
    console.log('\n📊 RESUMO:');
    console.log(JSON.stringify(status, null, 2));
    
    return status;
    
  } catch (error) {
    console.error('❌ Erro no teste:', error.message);
    return { erro: error.message };
  }
}

/**
 * Lista todas as abas disponíveis na planilha (para debug)
 */
function listarAbasDisponiveis() {
  console.log(' === ABAS DISPONÍVEIS NA PLANILHA ===');
  
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets();
  
  abas.forEach((aba, index) => {
    const nome = aba.getName();
    const dados = aba.getDataRange().getValues();
    const totalLinhas = dados.length - 1; // -1 para descontar cabeçalho
    
    console.log(`${index + 1}. "${nome}" - ${totalLinhas} registros`);
  });
  
  console.log(' === FIM DA LISTA ===');
}

/**
 * Busca um produto na Base Smartphones (planilha local)
 * OTIMIZADA - Busca rápida em planilha local
 * @param {string} codigo - Código do produto
 * @returns {Object} {encontrado: boolean, dados: Object}
 */
function buscarProdutoNaBaseSmarphones(codigo) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('Base Smartphones CD Principal');
    
    if (!aba) {
      console.log(' Aba "Base Smartphones CD Principal" não encontrada');
      return { encontrado: false };
    }
    
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) {
      return { encontrado: false };
    }
    
    // Busca otimizada - comparação direta
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const cf = String(linha[0] || '').trim();
      const codigoProduto = String(linha[1] || '').trim();
      
      if (cf === codigo || codigoProduto === codigo) {
        console.log(` Produto encontrado na linha ${i + 1} da Base Smartphones`);
        return {
          encontrado: true,
          dados: {
            cf: cf,
            codigo: codigoProduto,
            final: String(linha[2] || '').trim(),
            descricao: String(linha[3] || '').trim(),
            endereco: String(linha[4] || '').trim()
          }
        };
      }
    }
    
    return { encontrado: false };
    
  } catch (error) {
    console.error(' Erro ao buscar na Base Smartphones:', error);
    return { encontrado: false, erro: error.message };
  }
}

/**
 * Busca um produto na Base CD Principal (CSV otimizado)
 * OTIMIZADA - Usa função existente mais eficientemente
 * @param {string} codigo - Código do produto
 * @returns {Object} {encontrado: boolean, dados: Object}
 */
function buscarProdutoNaBaseCDPrincipal(codigo) {
  try {
    // Usar a função otimizada existente
    const resultado = buscarProdutoBaseCDPrincipal(codigo);
    
    if (resultado && resultado.length > 0) {
      console.log(' Produto encontrado na Base CD Principal via buscarProdutoBaseCDPrincipal');
      return {
        encontrado: true,
        dados: {
          cf: resultado[0] || '',
          codigo: resultado[1] || codigo,
          final: resultado[2] || '',
          descricao: resultado[3] || '',
          endereco: resultado[4] || ''
        }
      };
    }
    
    return { encontrado: false };
    
  } catch (error) {
    console.error(' Erro ao buscar na Base CD Principal:', error);
    return { encontrado: false, erro: error.message };
  }
}

/**
 * Busca um produto na Base RMA (CSV)
 * OTIMIZADA - Busca direta no arquivo mais recente
 * @param {string} codigo - Código do produto
 * @returns {Object} {encontrado: boolean, dados: Object}
 */
function buscarProdutoNaBaseRMA(codigo) {
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    // Encontrar arquivo RMA mais recente
    const arquivos = pasta.getFiles();
    let arquivoRmaMaisRecente = null;
    let dataMaisRecente = null;
    
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nome = arquivo.getName();
      
      if (nome.startsWith('PosicaoEstoqueOpenbox') && nome.endsWith('.csv')) {
        const dataModificacao = arquivo.getLastUpdated();
        if (!arquivoRmaMaisRecente || dataModificacao > dataMaisRecente) {
          arquivoRmaMaisRecente = arquivo;
          dataMaisRecente = dataModificacao;
        }
      }
    }
    
    if (!arquivoRmaMaisRecente) {
      console.log(' Nenhum arquivo RMA encontrado');
      return { encontrado: false };
    }
    
    console.log(` Buscando no arquivo RMA: ${arquivoRmaMaisRecente.getName()}`);
    
    const conteudo = arquivoRmaMaisRecente.getBlob().getDataAsString();
    const linhas = conteudo.split('\n');
    
    if (linhas.length <= 1) {
      return { encontrado: false };
    }
    
    // Busca otimizada linha por linha
    for (let i = 1; i < linhas.length; i++) {
      const linha = linhas[i];
      if (!linha || linha.trim() === '') continue;
      
      const colunas = linha.split(';').map(col => col.replace(/"/g, '').trim());
      
      // Verificar se é o código procurado (posições 0 e 1 geralmente são CF e CODIGO)
      if (colunas.length >= 2) {
        const cf = colunas[0];
        const codigoProduto = colunas[1];
        
        if (cf === codigo || codigoProduto === codigo) {
          console.log(` Produto encontrado na linha ${i + 1} da Base RMA`);
          return {
            encontrado: true,
            dados: {
              cf: cf,
              codigo: codigoProduto,
              final: colunas[2] || '',
              descricao: colunas[3] || '',
              endereco: colunas[4] || ''
            }
          };
        }
      }
    }
    
    return { encontrado: false };
    
  } catch (error) {
    console.error(' Erro ao buscar na Base RMA:', error);
    return { encontrado: false, erro: error.message };
  }
}

/**
 * Salva um registro de auditoria de setados na planilha
 * @param {Object} dados - Dados do registro (data, etiqueta, descricao, situacao, auditor)
 * @returns {string} Mensagem de confirmação
 */
function salvarAuditoriaSetados(dados) {
  console.log(' Salvando auditoria de setados:', dados);
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let aba = planilha.getSheetByName('Auditoria Setados');
    
    // Criar aba se não existir
    if (!aba) {
      console.log(' Criando nova aba "Auditoria Setados"...');
      aba = planilha.insertSheet('Auditoria Setados');
      
      // Adicionar cabeçalho
      aba.getRange(1, 1, 1, 5).setValues([
        ['DATA', 'ETIQUETA', 'DESCRIÇÃO', 'SITUAÇÃO', 'AUDITOR']
      ]);
      
      // Formatar cabeçalho
      const cabecalho = aba.getRange(1, 1, 1, 5);
      cabecalho.setFontWeight('bold');
      cabecalho.setBackground('#4CAF50');
      cabecalho.setFontColor('white');
      aba.setFrozenRows(1);
    }
    
    // Adicionar nova linha
    const proximaLinha = aba.getLastRow() + 1;
    aba.getRange(proximaLinha, 1, 1, 5).setValues([[
      dados.data,
      dados.etiqueta,
      dados.descricao,
      dados.situacao,
      dados.auditor
    ]]);
    
    console.log(` Registro salvo na linha ${proximaLinha}`);
    return `Registro salvo na linha ${proximaLinha} da aba "Auditoria Setados"`;
    
  } catch (error) {
    console.error(' Erro ao salvar auditoria de setados:', error);
    throw new Error('Erro ao salvar na planilha: ' + error.message);
  }
}

/**
 * Função de teste para a verificação de produtos setados
 */
function testarVerificacaoSetados() {
  console.log(' === TESTE VERIFICAÇÃO DE SETADOS ===');
  
  // Códigos de teste - substitua por códigos reais do seu sistema
  const codigosTeste = [
    '12345',  // Código que deve existir
    '67890',  // Outro código que deve existir  
    '99999'   // Código que não deve existir (possível setado)
  ];
  
  for (const codigo of codigosTeste) {
    console.log(`\n Testando código: ${codigo}`);
    
    const inicio = new Date().getTime();
    const resultado = verificarProdutoEmTodasAsBases(codigo);
    const tempo = new Date().getTime() - inicio;
    
    console.log(`⏱️ Tempo: ${tempo}ms`);
    console.log(' Resultado:', resultado);
    
    if (resultado.encontrado) {
      console.log(` Produto encontrado na ${resultado.base}:`);
      console.log(`   CF: ${resultado.dados.cf}`);
      console.log(`   Código: ${resultado.dados.codigo}`);
      console.log(`   Descrição: ${resultado.dados.descricao}`);
      console.log(`   Endereço: ${resultado.dados.endereco}`);
    } else {
      console.log(' Possível produto setado!');
    }
  }
  
  console.log('\n === FIM TESTE VERIFICAÇÃO DE SETADOS ===');
}

/**
 * Teste específico para um código individual
 */
function testarCodigoEspecifico(codigo) {
  console.log(` === TESTE CÓDIGO ESPECÍFICO: ${codigo} ===`);
  
  const inicio = new Date().getTime();
  const resultado = verificarProdutoEmTodasAsBases(codigo);
  const tempo = new Date().getTime() - inicio;
  
  console.log(`⏱️ Tempo total: ${tempo}ms`);
  console.log(' Resultado completo:', JSON.stringify(resultado, null, 2));
  
  return resultado;
}

/**
 * Função de teste FINAL - Valida todo o sistema otimizado
 */
function testarSistemaOtimizadoCompleto() {
  console.log(' === TESTE COMPLETO DO SISTEMA OTIMIZADO ===');
  
  try {
    const inicioTeste = new Date().getTime();
    
    // Teste 1: Performance comparison
    console.log('1️⃣ Comparando performance RMA vs CD...');
    const resultadoPerformance = testarPerformanceRMAvsCD();
    
    // Teste 2: Carregamento completo via doGet
    console.log('2️⃣ Testando integração completa...');
    const enderecosCompletos = extrairTodosEnderecosCsvCompleto();
    console.log(` Endereços completos: ${enderecosCompletos ? enderecosCompletos.length : 'null'}`);
    
    // Teste 3: Busca por endereço específico
    console.log('3️⃣ Testando busca específica...');
    if (enderecosCompletos && enderecosCompletos.length > 0) {
      const enderecoTeste = enderecosCompletos[0];
      console.log(` Testando busca no endereço: ${enderecoTeste}`);
      
      const produtosBusca = buscarProdutosPorEndereco(enderecoTeste, 'principal');
      console.log(` Produtos encontrados: ${produtosBusca ? produtosBusca.length - 1 : 'null'}`);
    }
    
    // Teste 4: Amostra inicial
    console.log('4️⃣ Testando amostra inicial...');
    const amostra = carregarAmostraBaseCDPrincipal();
    console.log(` Amostra carregada: ${amostra ? amostra.length : 'null'} linhas`);
    
    const tempoTotal = new Date().getTime() - inicioTeste;
    
    console.log(' === RESULTADO FINAL ===');
    console.log(`⏱️  Tempo total: ${tempoTotal}ms`);
    console.log(` Performance RMA vs CD: ${resultadoPerformance?.percentualDiferenca || 'N/A'}%`);
    console.log(` Endereços totais: ${enderecosCompletos?.length || 0}`);
    console.log(` Linhas de amostra: ${amostra?.length || 0}`);
    
    // Verificação final
    const sistemaOK = enderecosCompletos && enderecosCompletos.length > 0 && amostra && amostra.length > 0;
    
    if (sistemaOK) {
      console.log(' SISTEMA OTIMIZADO FUNCIONANDO PERFEITAMENTE!');
      console.log(' O carregamento do CD Principal agora é tão rápido quanto o RMA');
    } else {
      console.log('  Sistema precisa de ajustes');
    }
    
    return {
      sistemaOK: sistemaOK,
      tempoTotal: tempoTotal,
      performance: resultadoPerformance,
      enderecosTotal: enderecosCompletos?.length || 0,
      amostraLinhas: amostra?.length || 0
    };
    
  } catch (error) {
    console.error(' Erro no teste completo:', error);
    return { erro: error.message };
  }
}

/**
 * Função de teste para debugar problemas com CSV RMA - versão atualizada
 */
function testeDebugRMA() {
  console.log(' === TESTE DEBUG RMA (v2) ===');
  
  try {
    const folderId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(folderId);
    
    console.log(' Procurando arquivos PosicaoEstoqueOpenbox...');
    const arquivos = pasta.getFiles();
    const arquivosRMA = [];
    
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nome = arquivo.getName();
      
      if (nome.startsWith('PosicaoEstoqueOpenbox') && nome.endsWith('.csv')) {
        arquivosRMA.push({
          nome: nome,
          dataModificacao: arquivo.getLastUpdated(),
          tamanho: (arquivo.getSize()/1024/1024).toFixed(2) + ' MB'
        });
      }
    }
    
    if (arquivosRMA.length === 0) {
      console.error(' Nenhum arquivo PosicaoEstoqueOpenbox encontrado!');
      return;
    }
    
    // Ordena por data de modificação (mais recente primeiro)
    arquivosRMA.sort((a, b) => b.dataModificacao - a.dataModificacao);
    
    console.log(` Encontrados ${arquivosRMA.length} arquivo(s) RMA:`);
    arquivosRMA.forEach((arq, idx) => {
      console.log(`  ${idx + 1}. ${arq.nome} (${arq.tamanho}, ${arq.dataModificacao})`);
    });
    
    const arquivoMaisRecente = arquivosRMA[0];
    console.log(` Usando arquivo mais recente: ${arquivoMaisRecente.nome}`);
    
    // Testa a busca
    console.log(' Testando buscarProdutosPorEndereco com RMA...');
    const resultado = buscarProdutosPorEndereco('200-0-32-1', 'rma');
    
    console.log(' Resultado da busca:');
    console.log(`- Total linhas: ${resultado.length}`);
    if (resultado.length > 1) {
      console.log(`- Cabeçalho: [${resultado[0].join(', ')}]`);
      console.log('- Primeiras 3 linhas de dados:');
      for (let i = 1; i <= Math.min(3, resultado.length - 1); i++) {
        console.log(`  ${i}: [${resultado[i].join(' | ')}]`);
      }
    } else {
      console.log(' Nenhuma linha de dados retornada');
    }
    
  } catch (error) {
    console.error(' Erro no teste:', error);
  }
  
  console.log(' === FIM TESTE DEBUG RMA (v2) ===');
}

/**
 * Função de teste para verificar buscarProdutosPorEndereco com RMA
 */
function testeBuscarRMA() {
  console.log(' === TESTE buscarProdutosPorEndereco RMA ===');
  
  const resultado = buscarProdutosPorEndereco('200-0-32-1', 'rma');
  
  console.log(' Resultado:');
  console.log(`- Total linhas: ${resultado.length}`);
  console.log(`- Cabeçalho: [${resultado[0].join(', ')}]`);
  
  if (resultado.length > 1) {
    console.log('- Primeiras 3 linhas de dados:');
    for (let i = 1; i <= Math.min(3, resultado.length - 1); i++) {
      console.log(`  ${i}: [${resultado[i].join(' | ')}]`);
    }
  } else {
    console.log(' Nenhuma linha de dados retornada');
  }
  
  console.log(' === FIM TESTE buscarProdutosPorEndereco RMA ===');
  return resultado;
}

/**
 * Extrai endereços únicos dos dados da Base CD Principal
 * @param {Array} dados - Dados da Base CD Principal
 * @returns {Array} Lista de endereços únicos ordenados
 */
function extrairEnderecosCDPrincipal(dados) {
  console.log(' Extraindo endereços da Base CD Principal...');
  console.log(' Dados recebidos:', dados ? dados.length : 'null');
  
  if (!dados || dados.length <= 1) {
    console.log(' Dados vazios ou só cabeçalho');
    return [];
  }
  
  try {
    // Log dos primeiros registros para diagnóstico
    console.log(' Cabeçalho (primeiro registro):', dados[0]);
    console.log(' Segundo registro:', dados[1]);
    console.log(' Terceiro registro:', dados[2]);
    
    // Detectar qual coluna tem os endereços
    const cabecalho = dados[0];
    let colunaEndereco = -1;
    
    // Procurar por possíveis nomes de coluna de endereço
    const possiveisNomes = ['ENDERECO', 'ENDEREÇO', 'endereco', 'endereço', 'FINAL', 'final', 'ADDRESS'];
    for (let i = 0; i < cabecalho.length; i++) {
      if (possiveisNomes.includes(cabecalho[i])) {
        colunaEndereco = i;
        console.log(` Coluna de endereço encontrada: ${cabecalho[i]} (índice ${i})`);
        break;
      }
    }
    
    // Se não encontrou, tentar índice 4 (padrão)
    if (colunaEndereco === -1) {
      colunaEndereco = 4;
      console.log(` Usando índice padrão 4 para endereços`);
    }
    
    const enderecos = dados.slice(1) // Pular cabeçalho
      .map(row => row[colunaEndereco]) // Usar coluna detectada
      .filter(endereco => endereco && endereco.toString().trim() !== '')
      .map(endereco => endereco.toString().trim());
    
    console.log(` Endereços brutos extraídos: ${enderecos.length}`);
    console.log(' Primeiros 10 endereços:', enderecos.slice(0, 10));
      
    const enderecosUnicos = [...new Set(enderecos)].sort((a, b) => {
      // Ordenação mais flexível
      try {
        const pa = a.split('-').map(Number);
        const pb = b.split('-').map(Number);
        for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
          const numA = pa[i] || 0;
          const numB = pb[i] || 0;
          if (numA !== numB) return numA - numB;
        }
        return 0;
      } catch {
        return a.localeCompare(b); // Fallback para ordenação alfabética
      }
    });
    
    console.log(` ${enderecosUnicos.length} endereços únicos finais`);
    console.log(' Primeiros 5 endereços únicos:', enderecosUnicos.slice(0, 5));
    
    return enderecosUnicos;
    
  } catch (error) {
    console.error(' Erro ao extrair endereços:', error);
    return [];
  }
}

function doGet(e) {
  console.log(' === INICIANDO doGet() COM DEBUG ===');
  const htmlTemplate = HtmlService.createTemplateFromFile('interface');

  // ========== NOVA IMPLEMENTAÇÃO CSV DIRETO ==========
  console.log(' Iniciando carregamento CSV direto...');
  
  try {
    console.log('  NOVA ESTRATÉGIA: Lista completa de endereços + dados sob demanda');
    
    // ESTRATÉGIA HÍBRIDA:
    // 1. Carregar TODOS os endereços para a lista
    // 2. Carregar dados sob demanda quando endereço é selecionado
    
    console.log(' Extraindo TODOS os endereços do CSV...');
    const todosEnderecos = extrairTodosEnderecosCsvCompleto();
    
    console.log(' Carregando amostra pequena para compatibilidade inicial...');
    const dadosAmostra = carregarAmostraBaseCDPrincipal();
    htmlTemplate.dadosBanco = JSON.stringify(dadosAmostra);
    
    // Usar TODOS os endereços na lista
    if (todosEnderecos && todosEnderecos.length > 0) {
      htmlTemplate.todosEnderecos = JSON.stringify(todosEnderecos);
      console.log(` TODOS os endereços carregados: ${todosEnderecos.length}`);
      console.log(' Primeiros 10 endereços:', todosEnderecos.slice(0, 10));
      console.log(' Últimos 10 endereços:', todosEnderecos.slice(-10));
    } else {
      console.warn(' Falha ao carregar endereços completos, usando fallback');
      const enderecosFallback = gerarEnderecosFallback();
      htmlTemplate.todosEnderecos = JSON.stringify(enderecosFallback);
    }
    
    console.log(` Amostra inicial: ${dadosAmostra.length} registros`);
    console.log(` Total de endereços disponíveis: ${todosEnderecos ? todosEnderecos.length : 'fallback'}`);
    console.log('  Dados completos serão carregados sob demanda ao selecionar endereço');
    
  } catch (error) {
    console.error(' Erro no carregamento otimizado:', error);
    console.error('Stack trace:', error.stack);
    
    // Fallback: dados vazios, mas válidos
    const dadosFallback = [['CF', 'CODIGO', 'FINAL', 'DESCRICAO', 'ENDERECO']];
    htmlTemplate.dadosBanco = JSON.stringify(dadosFallback);
    
    // Usar endereços de fallback mais realistas
    const enderecosFallbackFinal = gerarEnderecosFallback();
    htmlTemplate.todosEnderecos = JSON.stringify(enderecosFallbackFinal);
    
    console.log(' Usando dados de fallback devido ao erro');
  }

  // ========== Planilha local para outras abas ==========
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  // Carrega Dados da aba "Base Smartphones CD Principal" (aba local)
  const sheetSmartphones = planilha.getSheetByName('Base Smartphones CD Principal');
  let dadosSmartphones = [];
  let enderecosSmartphones = [];
  if (sheetSmartphones) {
    dadosSmartphones = sheetSmartphones.getDataRange().getValues();
    // Extrai endereços únicos, ignorando cabeçalho e vazios
    const colunaEnderecoSmart = dadosSmartphones.slice(1).map(row => row[4]); // Supondo endereço na coluna 5
    enderecosSmartphones = Array.from(new Set(colunaEnderecoSmart)).filter(e => e && e !== "ENDEREÇO");
    enderecosSmartphones = enderecosSmartphones.sort((a, b) => {
      const pa = a.split('-').map(Number);
      const pb = b.split('-').map(Number);
      for (let i = 0; i < 4; i++) {
        if (pa[i] !== pb[i]) return pa[i] - pb[i];
      }
      return 0;
    });
  }
  htmlTemplate.dadosSmartphones = JSON.stringify(dadosSmartphones);
  htmlTemplate.todosEnderecosSmartphones = JSON.stringify(enderecosSmartphones);

  // Carrega Dados da aba "Base RMA" 
  // NOVO: Para RMA, também extrair endereços únicos
  const sheetRMA = planilha.getSheetByName('Base RMA');
  let dadosRMA = [];
  let enderecosRMA = [];
  if (sheetRMA) {
    dadosRMA = sheetRMA.getDataRange().getValues();
    // Extrai endereços únicos da Base RMA local se existir
    if (dadosRMA.length > 1) {
      const colunaEnderecoRMA = dadosRMA.slice(1).map(row => row[4]); // Supondo endereço na coluna 5
      enderecosRMA = Array.from(new Set(colunaEnderecoRMA)).filter(e => e && e !== "ENDEREÇO" && e !== "Endereço");
      enderecosRMA = enderecosRMA.sort((a, b) => {
        const pa = a.split('-').map(Number);
        const pb = b.split('-').map(Number);
        for (let i = 0; i < 4; i++) {
          if (pa[i] !== pb[i]) return pa[i] - pb[i];
        }
        return 0;
      });
    }
  }
  
  // Se não tem aba RMA local ou está vazia, carrega endereços do CSV
  if (enderecosRMA.length === 0) {
    console.log(' Carregando endereços RMA do CSV...');
    try {
      enderecosRMA = extrairEnderecosRmaDireto();
    } catch (e) {
      console.warn(' Erro ao carregar endereços RMA do CSV:', e);
      enderecosRMA = [];
    }
  }
  
  htmlTemplate.dadosRMA = JSON.stringify(dadosRMA);
  htmlTemplate.todosEnderecosRMA = JSON.stringify(enderecosRMA);
  console.log(` Endereços RMA carregados: ${enderecosRMA.length}`);

  // Carrega Dados da aba "Base NL"
  const sheetNL = planilha.getSheetByName('Base NL');
  let dadosNL = [];
  if (sheetNL) {
    dadosNL = sheetNL.getDataRange().getValues();
  }
  htmlTemplate.dadosNL = JSON.stringify(dadosNL);

  // Carrega Dados das abas "Base Transitorios - Parte 1" e "Base Transitorios - Parte 2"
  let dadosTransitorios = [];
  
  // Carregar Base Transitorios - Parte 1
  const sheetTransitorios1 = planilha.getSheetByName('Base Transitorios - Parte 1');
  if (sheetTransitorios1) {
    const dados1 = sheetTransitorios1.getDataRange().getValues();
    if (dados1.length > 1) {
      // Adicionar cabeçalho da primeira parte
      if (dadosTransitorios.length === 0) {
        dadosTransitorios.push(dados1[0]); // Cabeçalho
      }
      dadosTransitorios = dadosTransitorios.concat(dados1.slice(1)); // Dados sem cabeçalho
    }
  }
  
  // Carregar Base Transitorios - Parte 2
  const sheetTransitorios2 = planilha.getSheetByName('Base Transitorios - Parte 2');
  if (sheetTransitorios2) {
    const dados2 = sheetTransitorios2.getDataRange().getValues();
    if (dados2.length > 1) {
      dadosTransitorios = dadosTransitorios.concat(dados2.slice(1)); // Adicionar dados sem cabeçalho
    }
  }
  
  console.log(` Base Transitórios carregada: ${dadosTransitorios.length - 1} registros total`);
  htmlTemplate.dadosTransitorios = JSON.stringify(dadosTransitorios);

  return htmlTemplate.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Auditoria de Estoque")
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
}

/**
 * Retorna o e-mail do usuário logado para identificar o operador.
 * Observação: é necessário publicar o web app executando como o usuário e com acesso restrito.
 */
function getOperador() {
  try {
    const email = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '';
    return email || '';
  } catch (e) {
    console.warn(' Não foi possível obter o operador via Session:', e);
    return '';
  }
}

/**
 * Converte um e-mail corporativo para um nome amigável.
 * Ex.: "joao.silva@kabum.com.br" => "Joao Silva"
 */
function extrairNomeOperador(email) {
  try {
    if (!email) return '';
    const local = String(email).split('@')[0] || '';
    if (!local) return '';
    const partes = local.split(/[._-]+/).filter(Boolean);
    if (partes.length === 0) return local;
    const capitalize = (s) => s ? (s.charAt(0).toUpperCase() + s.slice(1).toLowerCase()) : s;
    // Mantém preposições comuns em minúsculas, exceto a primeira palavra
    const preps = new Set(['da','de','do','das','dos','e']);
    return partes.map((p, idx) => {
      const pLower = p.toLowerCase();
      if (idx > 0 && preps.has(pLower)) return pLower;
      return capitalize(p);
    }).join(' ');
  } catch (e) {
    return email;
  }
}


function salvarCodigosInseridos(dados) {
  console.log(' === INICIO salvarCodigosInseridos ===');
  console.log(' Dados recebidos:', dados);
  
  // Suporte a dois formatos de entrada:
  // - Antigo: Array de { codigo, endereco }
  // - Novo: { operador, tipoEstoque, endereco, itens: [{codigo}], uuid }
  let operador = '';
  let tipoEstoque = 'principal';
  let enderecoAtual = '';
  /** @type {{codigo:string}[]} */
  let itens = [];

  if (Array.isArray(dados)) {
    // Formato antigo
    enderecoAtual = (dados && dados.length > 0) ? dados[0].endereco : '';
    itens = dados.map(d => ({ codigo: d.codigo }));
    try { operador = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || ''; } catch {}
  } else if (dados && typeof dados === 'object') {
    operador = dados.operador || '';
    tipoEstoque = dados.tipoEstoque || 'principal';
    enderecoAtual = dados.endereco || '';
    itens = Array.isArray(dados.itens) ? dados.itens.map(it => ({ codigo: it.codigo })) : [];
  }

  console.log(' Parâmetros extraídos:');
  console.log('- Operador:', operador);
  console.log('- Tipo Estoque:', tipoEstoque);
  console.log('- Endereço:', enderecoAtual);
  console.log('- Itens:', itens.length, 'códigos:', itens.map(i => i.codigo));

  // Lock será adquirido apenas no momento da escrita, para não segurar durante leituras pesadas
  let lock = null;

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  // Seleciona a aba de destino conforme o tipo de estoque
  const nomeAbaDestino = (tipoEstoque === 'rma') ? 'RMA Auditados' : 'Estoque Auditados';
  let sheet = planilha.getSheetByName(nomeAbaDestino);
  if (!sheet) {
    // Cria a aba caso não exista
    sheet = planilha.insertSheet(nomeAbaDestino);
    // Cria cabeçalho com 11 colunas na ordem solicitada
    sheet.getRange(1, 1, 1, 11).setValues([[
      'DATA',                 // 1
      'CF',                   // 2
      'Finais Auditados',     // 3
      'Descrição',            // 4
      'Qtd Sis',              // 5
      'Qtd Aud',              // 6
      'Valor Auditado',       // 7
      'Valor Divergencia',    // 8
      'Endereço',             // 9
      'Divergencia',          // 10 (texto)
      'Auditor'               // 11
    ]]);
  }
  // Garante cabeçalho com 11 colunas caso a aba exista mas tenha formato antigo
  try {
    const headerDesejado = [[
      'DATA', 'CF', 'Finais Auditados', 'Descrição', 'Qtd Sis', 'Qtd Aud',
      'Valor Auditado', 'Valor Divergencia', 'Endereço', 'Divergencia', 'Auditor'
    ]];
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() < 11) {
      sheet.getRange(1, 1, 1, 11).setValues(headerDesejado);
    }
  } catch (e) {
    console.warn('Falha ao garantir cabeçalho 11 colunas:', e);
  }

  // ======================== CARREGAMENTO DE DADOS ========================
  
  // Para RMA: forçar uso direto do CSV independente da aba local
  // Para Principal: manter lógica original
  let bancoDados = [];
  
  if (tipoEstoque === 'rma') {
    // RMA: SEMPRE usar CSV direto (igual ao Principal funcionando)
    console.log(' RMA: Usando CSV direto (forçado)');
    try {
      const filtrados = buscarProdutosPorEndereco(enderecoAtual, 'rma');
      if (filtrados && filtrados.length > 1) {
        bancoDados = filtrados.slice(1); // remove cabeçalho
        console.log(` RMA CSV carregou ${bancoDados.length} linhas para endereço ${enderecoAtual}`);
        
        // Log amostra dos dados carregados
        if (bancoDados.length > 0) {
          console.log(' Amostra dados RMA (primeiras 3 linhas):');
          bancoDados.slice(0, 3).forEach((linha, idx) => {
            console.log(`  ${idx}: CF=${linha[0]}, CODIGO=${linha[1]}, FINAL=${linha[2]}, DESC=${linha[3]}, END=${linha[4]}`);
          });
        }
      } else {
        console.warn(' RMA CSV não retornou linhas para o endereço:', enderecoAtual);
      }
    } catch (e) {
      console.error(' Erro ao carregar RMA CSV:', e);
    }
  } else {
    // PRINCIPAL: Lógica original (aba local + fallback CSV)
    console.log(' Principal: Usando lógica original (aba local + fallback)');
    
    const bancoSheet = planilha.getSheetByName('Base CD Principal');
    if (!bancoSheet) {
      console.warn('Aba Base CD Principal não encontrada. Usando fallback CSV.');
    }

    // Carrega dados da base local se disponível
    try {
      if (bancoSheet) {
        const totalRows = bancoSheet.getLastRow();
        const totalCols = bancoSheet.getLastColumn();
        if (totalRows > 1 && totalCols > 0) {
          const header = bancoSheet.getRange(1, 1, 1, totalCols).getValues()[0].map(v => String(v || '').trim().toUpperCase());
          const findIdx = (cands) => {
            const set = new Set(cands.map(c=>c.toUpperCase()));
            for (let i=0;i<header.length;i++){ if (set.has(header[i])) return i; }
            return -1;
          };
          
          const cfCands = ['CF','CD MERCADARIA','CDMERCADORIA','CODIGO','CÓDIGO','COD'];
          const codCands = ['CODIGO','CÓDIGO','CD MERCADARIA','CDMERCADORIA','COD'];
          const descCands = ['DESCRICAO','DESCRIÇÃO','DESCR','DESCRIACA DO PRODUTO','DESCRIACAO DO PRODUTO','DESCRIÇÃO DO PRODUTO','DESCRIACAO'];
          const endCands = ['ENDERECO','ENDEREÇO','END','ENDERECO DO PRODUTO','ENDEREÇO DO PRODUTO'];
          const custoCands = ['CUSTO DO PRODUTO','CUSTO','PRECO','PREÇO','VALOR','VALOR UNITARIO','VALOR UNITÁRIO','CUSTO PRODUTO','CUSTO_UNITARIO','CUSTO UNITARIO','CUSTO UND','CUSTO UNIT'];
          const finalCands = [
            'FINAL','FINAL DO PRODUTO','FINAIS','FINAIS DO PRODUTO','FIN',
            'SÉRIE','SERIE','Nº SÉRIE','Nº SERIE','NO SÉRIE','NO SERIE',
            'NUMERO DE SERIE','NÚMERO DE SÉRIE','NUMERO SERIE','NUMERO SÉRIE',
            'SERIAL','SERIAL NUMBER','IMEI'
          ];

          let idxCF = findIdx(cfCands);
          let idxCOD = findIdx(codCands);
          let idxFINAL = findIdx(finalCands);
          let idxDESC = findIdx(descCands);
          let idxEND = findIdx(endCands);
          let idxCUSTO = findIdx(custoCands);
          // Fallbacks se não achar
          if (idxCF < 0) idxCF = 0;
          if (idxCOD < 0) idxCOD = 1;
          if (idxFINAL < 0) idxFINAL = 2;
          if (idxDESC < 0) idxDESC = 3;
          if (idxEND < 0) idxEND = 4;
          if (idxCUSTO < 0) idxCUSTO = -1;
          
          const dados = bancoSheet.getRange(2, 1, totalRows - 1, totalCols).getValues();
          bancoDados = dados.map(r => [r[idxCF], r[idxCOD], r[idxFINAL], r[idxDESC], r[idxEND], (idxCUSTO>=0 ? r[idxCUSTO] : '')]);
          console.log('🧭 Mapeamento Principal local:', { idxCF, idxCOD, idxFINAL, idxDESC, idxEND, idxCUSTO });
        } else {
          console.warn('Aba Principal sem dados (apenas cabeçalho).');
        }
      }
    } catch (e) {
      console.warn('Falha ao ler aba Principal local:', e);
    }

    // Fallback: se a base local está vazia ou sem dados do endereço
    if (!bancoDados || bancoDados.length === 0) {
      try {
        console.log(' Principal: Fallback CSV por base vazia');
        const filtrados = buscarProdutosPorEndereco(enderecoAtual, 'principal');
        if (filtrados && filtrados.length > 1) {
          bancoDados = filtrados.slice(1);
          console.log(` Principal Fallback CSV carregou ${bancoDados.length} linhas`);
        }
      } catch (e) {
        console.error(' Erro no fallback Principal CSV:', e);
      }
    } else {
      // Verificar se tem registros do endereço
      const temEnderecoNaBaseLocal = bancoDados.some(r => String(r[4]||'').trim() === String(enderecoAtual).trim());
      if (!temEnderecoNaBaseLocal) {
        console.warn(' Principal: Base local sem registros do endereço. Fallback CSV...');
        try {
          const filtrados = buscarProdutosPorEndereco(enderecoAtual, 'principal');
          if (filtrados && filtrados.length > 1) {
            bancoDados = filtrados.slice(1);
            console.log(` Principal Fallback CSV (por ausência) carregou ${bancoDados.length} linhas`);
          }
        } catch (e) {
          console.warn('Falha no fallback Principal CSV:', e);
        }
      }
    }
  }

  
  // ======================== VALIDAÇÃO ========================
  if (!enderecoAtual) {
    throw new Error("Endereço não informado. Dados enviados sem endereço.");
  }

  if (!bancoDados || bancoDados.length === 0) {
    console.error(' Nenhum dado carregado para o endereço:', enderecoAtual);
    throw new Error(`Nenhum produto encontrado para o endereço ${enderecoAtual} na base ${tipoEstoque}`);
  }

  console.log(`  Total de dados carregados: ${bancoDados.length} linhas`);

  // ======================== PROCESSAMENTO ========================
  
  // Mapas para armazenar os finais do sistema e auditados por CF
  const sistemaPorCF = new Map();
  const auditadoPorCF = new Map();
  const cfEnderecoMap = new Map();
  // Mapa auxiliar: CF -> (Final -> {codigo, descricao}) para descrever divergências
  const mapaCFParaFinal = new Map();

  // Construir dados do sistema
  console.log(` Processando dados para endereço: ${enderecoAtual}`);
  let produtosFiltrados = 0;
  const normaliza = (v) => String(v == null ? '' : v).trim();
  const normalizaNum = (v) => normaliza(v).replace(/^0+(\d)/, '$1');

  // Índices por código e CF para acelerar e robustecer matching
  const mapaPorCodigo = new Map(); // codigo -> array de linhas do endereço
  const mapaPorCF = new Map();     // cf -> array de linhas do endereço
  
  const parseMoeda = (v) => {
    const s0 = String(v == null ? '' : v).trim();
    if (!s0) return NaN;
    const s = s0.replace(/[R$\s]/g, '');
    const hasComma = s.includes(',');
    const hasDot = s.includes('.');
    // Caso tenha vírgula e ponto, usa o último separador como decimal
    if (hasComma && hasDot) {
      const lastComma = s.lastIndexOf(',');
      const lastDot = s.lastIndexOf('.');
      if (lastComma > lastDot) {
        // Formato PT-BR: 1.234,56 -> remove milhares '.', troca ',' por '.'
        const noThousands = s.replace(/\./g, '');
        const normalized = noThousands.replace(/,/g, '.');
        const n = parseFloat(normalized);
        return isNaN(n) ? NaN : n;
      } else {
        // Formato EN/US: 1,234.56 -> remove milhares ',' e mantém '.'
        const noThousands = s.replace(/,/g, '');
        const n = parseFloat(noThousands);
        return isNaN(n) ? NaN : n;
      }
    }
    // Apenas vírgula: tratar como decimal
    if (hasComma && !hasDot) {
      const noThousands = s.replace(/\./g, '');
      const normalized = noThousands.replace(/,/g, '.');
      const n = parseFloat(normalized);
      return isNaN(n) ? NaN : n;
    }
    // Apenas ponto: tratar como decimal, remover eventuais vírgulas de milhar
    if (hasDot && !hasComma) {
      const noThousands = s.replace(/,/g, '');
      const n = parseFloat(noThousands);
      return isNaN(n) ? NaN : n;
    }
    // Apenas dígitos
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
  };
  const custoPorCF = new Map();

  bancoDados.forEach((row, index) => {
    const cf = row[0];
    const codigo = row[1];
    const final = row[2];
    const descricao = row[3];
    const endereco = row[4];
    const custoRaw = row.length > 5 ? row[5] : '';
    
    // Log detalhado das primeiras linhas
    if (index < 5) {
      console.log(` Linha ${index}: CF="${cf}", CODIGO="${codigo}", FINAL="${final}", DESC="${descricao}", END="${endereco}"`);
      console.log(`    Normalizado: CF="${normaliza(cf)}", END="${normaliza(endereco)}", TARGET="${normaliza(enderecoAtual)}"`);
      console.log(`    Match endereco: ${normaliza(endereco) === normaliza(enderecoAtual)}`);
    }
    
  if (normaliza(endereco) !== normaliza(enderecoAtual)) return;
    
    produtosFiltrados++;
    const chave = `${cf}|${endereco}`;
    if (!cfEnderecoMap.has(chave)) {
      cfEnderecoMap.set(chave, { qtdSistema: 0, qtdAuditada: 0 });
    }
    cfEnderecoMap.get(chave).qtdSistema++;

    // FINAIS DO SISTEMA - LOG DETALHADO
    if (!sistemaPorCF.has(cf)) sistemaPorCF.set(cf, new Set());
    if (final !== undefined && final !== null && String(final).trim() !== '') {
      const finalNorm = String(final).trim();
      sistemaPorCF.get(cf).add(finalNorm);
      console.log(` Adicionado final do sistema: CF="${cf}" -> FINAL="${finalNorm}"`);
    } else {
      console.log(` Final vazio ou nulo: CF="${cf}", FINAL="${final}"`);
    }

    // Preenche mapa auxiliar para divergências (CF -> Final -> {codigoProduto, codigoEtiqueta, descricao})
    if (!mapaCFParaFinal.has(cf)) mapaCFParaFinal.set(cf, new Map());
    if (final !== undefined && final !== null && String(final).trim() !== '' && !mapaCFParaFinal.get(cf).has(String(final).trim())) {
      mapaCFParaFinal.get(cf).set(String(final).trim(), { codigoProduto: cf, codigoEtiqueta: codigo, descricao });
    }

    // Popular índices
    const codKey = normalizaNum(codigo);
    const cfKey = normalizaNum(cf);
    if (codKey) {
      if (!mapaPorCodigo.has(codKey)) mapaPorCodigo.set(codKey, []);
      mapaPorCodigo.get(codKey).push(row);
    }
    if (cfKey) {
      if (!mapaPorCF.has(cfKey)) mapaPorCF.set(cfKey, []);
      mapaPorCF.get(cfKey).push(row);
    }

    // Capturar custo por CF (primeiro válido encontrado)
    const custoNum = parseMoeda(custoRaw);
    if (!isNaN(custoNum) && custoNum > 0) {
      const chaveCF = String(cf);
      if (!custoPorCF.has(chaveCF)) {
        custoPorCF.set(chaveCF, custoNum);
      }
    }
  });
  
  console.log(` Produtos encontrados no endereço ${enderecoAtual}: ${produtosFiltrados}`);
  console.log(` CFs únicos no endereço: ${cfEnderecoMap.size}`);
  
  // Log do sistemaPorCF
  console.log(' Finais por CF no sistema:');
  sistemaPorCF.forEach((finaisSet, cf) => {
    const finaisArray = Array.from(finaisSet);
    console.log(`  CF "${cf}": ${finaisArray.length} finais -> [${finaisArray.join(', ')}]`);
  });

  // Construir dados auditados
  console.log(' Processando itens auditados...');
  itens.forEach((item, itemIndex) => {
    const codigoBruto = item && item.codigo != null ? item.codigo : '';
    const codigo = normalizaNum(codigoBruto);
    console.log(` Item ${itemIndex}: "${codigoBruto}" -> normalizado: "${codigo}"`);
    
    if (!codigo) {
      console.log(` Código vazio ignorado: ${codigoBruto}`);
      return;
    }
    
    // Tenta por código; se não achar, tenta por CF
    let registros = mapaPorCodigo.get(codigo);
    if (!registros || registros.length === 0) {
      registros = mapaPorCF.get(codigo) || [];
      console.log(` Código "${codigo}" não encontrado por CODIGO, tentando por CF: ${registros.length} registros`);
    } else {
      console.log(` Código "${codigo}" encontrado por CODIGO: ${registros.length} registros`);
    }
    
    // Filtra apenas do endereço atual
    registros = registros.filter(r => normaliza(r[4]) === normaliza(enderecoAtual));
    console.log(` Após filtro de endereço: ${registros.length} registros`);
    
    if (!registros || registros.length === 0) {
      console.log(` Nenhum registro encontrado para código "${codigo}" no endereço "${enderecoAtual}"`);
      return;
    }

    registros.forEach((registro, regIndex) => {
      const [cf, , final, , endereco] = registro;
      console.log(` Registro ${regIndex}: CF="${cf}", FINAL="${final}"`);
      
      const chave = `${cf}|${endereco}`;
      if (cfEnderecoMap.has(chave)) {
        cfEnderecoMap.get(chave).qtdAuditada++;
      }
      
      // FINAIS AUDITADOS - LOG DETALHADO
      if (!auditadoPorCF.has(cf)) auditadoPorCF.set(cf, new Set());
      if (final !== undefined && final !== null && String(final).trim() !== '') {
        const finalNorm = String(final).trim();
        auditadoPorCF.get(cf).add(finalNorm);
        console.log(` Adicionado final auditado: CF="${cf}" -> FINAL="${finalNorm}"`);
      } else {
        console.log(` Final vazio no registro auditado: CF="${cf}", FINAL="${final}"`);
      }
    });
  });
  
  // Log dos auditados
  console.log(' Finais auditados por CF:');
  auditadoPorCF.forEach((finaisSet, cf) => {
    const finaisArray = Array.from(finaisSet);
    console.log(`  CF "${cf}": ${finaisArray.length} finais -> [${finaisArray.join(', ')}]`);
  });

  // ======================== NOVA MONTAGEM DOS RESULTADOS - UMA LINHA POR CF ========================
  console.log(' Montando dados finais - UMA LINHA POR CF...');
  
  // Monta as linhas para salvar conforme solicitado
  // Colunas: DATA / CF / Finais Auditados / Endereço / Qtd Base / Qtd Auditada / Divergencia / Auditor
  const timezone = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const dataStr = Utilities.formatDate(new Date(), timezone, 'dd/MM/yy');
  const operadorNome = extrairNomeOperador(operador || '');

  // NOVA ESTRATÉGIA: Uma linha por CF do endereço
  const cfsEndereco = Array.from(new Set(Array.from(cfEnderecoMap.keys()).map(k => k.split('|')[0])));
  const linhasParaSalvar = [];

  console.log(` Processando ${cfsEndereco.length} CFs no endereço:`);
  
  cfsEndereco.forEach(cf => {
    const chave = `${cf}|${enderecoAtual}`;
    const cfInfo = cfEnderecoMap.get(chave) || { qtdSistema: 0, qtdAuditada: 0 };
    
    // Finais deste CF específico
    const finaisSistema = sistemaPorCF.get(cf) || new Set();
    const finaisAuditados = auditadoPorCF.get(cf) || new Set();
    const finaisNaoAuditados = Array.from(finaisSistema).filter(f => f && !finaisAuditados.has(f));
    
    console.log(` CF "${cf}":`);
    console.log(`  - Sistema: [${Array.from(finaisSistema).join(',')}] (${cfInfo.qtdSistema} itens)`);
    console.log(`  - Auditados: [${Array.from(finaisAuditados).join(',')}] (${cfInfo.qtdAuditada} itens)`);
    console.log(`  - Faltam: [${finaisNaoAuditados.join(',')}]`);
    
    // Finais auditados para este CF
    const finaisAuditadosStr = Array.from(finaisAuditados)
      .sort((a,b) => String(a).localeCompare(String(b)))
      .join(', ');
    
    // Divergências específicas deste CF
    const divergenciasMsgs = [];
    finaisNaoAuditados.forEach(final => {
      const infoFinal = mapaCFParaFinal.get(cf)?.get(final);
      const codVal = (infoFinal?.codigoProduto || infoFinal?.codigoEtiqueta || 'N/D');
      const cod = `COD ${codVal}`;
      const fin = `FINAL ${final}`;
      const mensagem = `${cod} ${fin} NÃO FOI AUDITADO`;
      divergenciasMsgs.push(mensagem);
    });
    
    // Determinar status da divergência
    let divergenciaStr = '';
    if (cfInfo.qtdAuditada === 0) {
      // CF não foi auditado completamente - incluir todos os finais que faltaram
      const finaisFaltando = Array.from(finaisSistema).sort((a,b) => String(a).localeCompare(String(b)));
      if (finaisFaltando.length > 0) {
        divergenciaStr = `CF ${cf} NÃO FOI AUDITADO - FINAIS FALTANDO: ${finaisFaltando.join(', ')} (${cfInfo.qtdSistema} itens)`;
      } else {
        divergenciaStr = `CF ${cf} NÃO FOI AUDITADO (${cfInfo.qtdSistema} itens faltando)`;
      }
    } else if (finaisNaoAuditados.length > 0) {
      // CF parcialmente auditado - mostrar finais específicos que faltaram
      const finaisFaltandoOrdenados = finaisNaoAuditados.sort((a,b) => String(a).localeCompare(String(b)));
      divergenciaStr = `CF ${cf} PARCIALMENTE AUDITADO - FINAIS FALTANDO: ${finaisFaltandoOrdenados.join(', ')}`;
    } else {
      // CF completamente auditado
      divergenciaStr = 'CONFORME';
    }
    
  console.log(`  - Divergência: "${divergenciaStr}"`);
    
    // Descrição representativa do CF neste endereço (primeira não vazia encontrada)
    let descricaoCF = '';
    try {
      for (const row of bancoDados) {
        const [cfRow, , , descRow, endRow] = row;
        if (String(cfRow) === String(cf) && normaliza(endRow) === normaliza(enderecoAtual)) {
          if (String(descRow || '').trim()) { descricaoCF = String(descRow).trim(); break; }
        }
      }
    } catch (e) {
      descricaoCF = '';
    }

    // Cálculo de valores monetários: sem coluna de preço disponível na base atual
    // Mantemos 0 como padrão até existir fonte de preço por CF
  const qtdSis = Number(cfInfo.qtdSistema || 0);
  const qtdAud = Number(cfInfo.qtdAuditada || 0);
  const delta = qtdAud - qtdSis;
  const statusDiverg = delta > 0 ? 'SOBRA' : (delta < 0 ? 'FALTA' : 'SEM DIVERGÊNCIA');
  const divergenciaExibicao = delta !== 0 ? `${divergenciaStr} — ${statusDiverg}` : divergenciaStr;
  const valorUnitario = custoPorCF.has(String(cf)) ? Number(custoPorCF.get(String(cf))) : 0;
  const valorAuditado = valorUnitario * qtdAud;
  const valorDivergencia = valorUnitario * Math.abs(delta);

    // Adicionar linha para este CF na nova ordem solicitada
    linhasParaSalvar.push([
      dataStr,              // DATA
      cf,                   // CF
      finaisAuditadosStr,   // FINAIS AUDITADOS
      descricaoCF,          // DESCRIÇÃO
      qtdSis,               // QTD SIS
      qtdAud,               // QTD AUD
      valorAuditado,        // VALOR AUDITADO
      valorDivergencia,     // VALOR DIVERGENCIA
      enderecoAtual,        // ENDEREÇO
      divergenciaExibicao,  // DIVERGENCIA (texto + status SOBRA/FALTA)
      operadorNome          // AUDITOR (nome)
    ]);
  });

  console.log(` ${linhasParaSalvar.length} linhas serão salvas (uma por CF):`);
  linhasParaSalvar.forEach((linha, index) => {
    // Estrutura: [0 DATA, 1 CF, 2 Finais, 3 Descrição, 4 Qtd Sis, 5 Qtd Aud, 6 Valor Auditado, 7 Valor Divergencia, 8 Endereço, 9 Divergencia (texto), 10 Auditor]
    console.log(`  [${index}] CF: "${linha[1]}", Qtd Sis: ${linha[4]}, Qtd Aud: ${linha[5]}, Divergência: "${linha[9]}"`);
  });
  console.log(' === FIM DOS LOGS salvarCodigosInseridos ===');

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 11).setValues([[
      'DATA',
      'CF',
      'Finais Auditados',
      'Descrição',
      'Qtd Sis',
      'Qtd Aud',
      'Valor Auditado',
      'Valor Divergencia',
      'Endereço',
      'Divergencia',
      'Auditor'
    ]]);
  }

  if (!linhasParaSalvar || linhasParaSalvar.length === 0) {
    console.warn('Nenhuma linha encontrada para salvar.');
    return 'Nenhum CF encontrado para o endereço selecionado. Nada foi salvo.';
  }

  // Escrita com lock + idempotência se houver UUID
  try {
    lock = LockService.getDocumentLock();
    lock.waitLock(20000);

    if (dados && typeof dados === 'object' && dados.uuid) {
      const uuid = String(dados.uuid);
      const controle = planilha.getSheetByName('Controle Auditoria') || planilha.insertSheet('Controle Auditoria');
      if (controle.getLastRow() === 0) {
        controle.getRange(1,1,1,6).setValues([[
          'UUID', 'TIMESTAMP', 'OPERADOR', 'TIPO', 'ENDERECO', 'QTDE_LINHAS']]);
      }
      const total = controle.getLastRow();
      if (total > 1) {
        const uuids = controle.getRange(2,1,total-1,1).getValues().map(r=>String(r[0]||''));
        if (uuids.includes(uuid)) {
          return 'Operação já processada anteriormente (idempotente).';
        }
      }
      // Efetiva gravação com UUID
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, linhasParaSalvar.length, 11).setValues(linhasParaSalvar);
      // Formatar colunas de valores como moeda (R$) com duas casas decimais
      try { sheet.getRange(startRow, 7, linhasParaSalvar.length, 2).setNumberFormat('"R$" #,##0.00'); } catch (e) { console.warn('Falha ao formatar moeda (UUID):', e); }
      controle.appendRow([uuid, new Date(), operador || '', tipoEstoque, enderecoAtual, linhasParaSalvar.length]);
      return `Registros salvos com sucesso: ${linhasParaSalvar.length}`;
    } else {
      // Gravação simples (formato antigo ou sem UUID)
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, linhasParaSalvar.length, 11).setValues(linhasParaSalvar);
      // Formatar colunas de valores como moeda (R$) com duas casas decimais
      try { sheet.getRange(startRow, 7, linhasParaSalvar.length, 2).setNumberFormat('"R$" #,##0.00'); } catch (e) { console.warn('Falha ao formatar moeda (simples):', e); }
      return `Registros salvos com sucesso: ${linhasParaSalvar.length}`;
    }
  } catch (e) {
    throw e;
  } finally {
    try { if (lock) lock.releaseLock(); } catch (_) {}
  }
}


function adicionarNovoRegistro(novoRegistro) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Base CD Principal');

    if (!sheet) {
      throw new Error("Aba 'Base CD Principal' não encontrada!");
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

/**
 * Salva um produto auditado na aba 'Transitorios Auditados'.
 * @param {Object} dados - { etiqueta, codigo, final, descricao, enderecoAuditado, status, data }
 */
function salvarTransitorioAuditado(dados) {
  try {
    console.log('=== DEBUG Backend: salvarTransitorioAuditado chamada ===');
    console.log('Dados recebidos:', dados);
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    console.log('Planilha obtida:', planilha.getName());
    
    let sheet = planilha.getSheetByName('Transitorios Auditados');
    console.log('Aba "Transitorios Auditados" encontrada:', !!sheet);
    
    if (!sheet) {
      console.log('Criando nova aba "Transitorios Auditados"...');
      sheet = planilha.insertSheet('Transitorios Auditados');
      // Cria cabeçalho na nova ordem: DATA / ETIQUETA / CODIGO / FINAL / DESCRICAO / ENDERECO SELECIONADO / STATUS / OPERADOR
      sheet.appendRow(['Data', 'Etiqueta', 'Código', 'Final', 'Descrição', 'Endereço Selecionado', 'Status', 'Operador']);
      console.log('Aba criada com cabeçalho');
    }
    
    // Se a aba está vazia (só cabeçalho ou nada), cria cabeçalho
    const ultimaLinha = sheet.getLastRow();
    console.log('Última linha da aba:', ultimaLinha);
    
    if (ultimaLinha === 0) {
      console.log('Aba vazia, adicionando cabeçalho...');
      sheet.appendRow(['Data', 'Etiqueta', 'Código', 'Final', 'Descrição', 'Endereço Selecionado', 'Status', 'Operador']);
    }
    
    // Lock para evitar colisões
    const lock = LockService.getDocumentLock();
    try { lock.waitLock(5000); } catch (e) { throw new Error('Sistema ocupado. Tente novamente.'); }

    // Operador
  let operador = '';
  try { operador = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || ''; } catch {}
  if (!operador && dados && dados.operador) operador = dados.operador;
  const operadorNome = extrairNomeOperador(operador || '');

    // Dados a serem inseridos na nova ordem
    const linhaDados = [
      dados.data || '',                    // DATA
      dados.etiqueta || '',                // ETIQUETA  
      dados.codigo || '',                  // CODIGO
      dados.final || '',                   // FINAL
      dados.descricao || '',               // DESCRICAO
      dados.enderecoAuditado || '',        // ENDERECO SELECIONADO
      dados.status || '',                  // STATUS
      operadorNome                          // OPERADOR (apenas nome)
    ];
    
    console.log('Inserindo linha:', linhaDados);
    
    // Adiciona o produto auditado
    sheet.appendRow(linhaDados);
    
    lock.releaseLock();

    const novaUltimaLinha = sheet.getLastRow();
    console.log('Nova última linha:', novaUltimaLinha);
    
    const resultado = 'Produto auditado salvo com sucesso na linha ' + novaUltimaLinha;
    console.log('Sucesso:', resultado);
    
    return resultado;
    
  } catch (error) {
    console.error('Erro em salvarTransitorioAuditado:', error);
    console.error('Stack trace:', error.stack);
    throw new Error('Erro ao salvar produto auditado: ' + error.message);
  }
}

/**
 * Remove um produto auditado da aba 'Transitorios Auditados'.
 * @param {string} etiqueta - Etiqueta do produto a ser removido
 */
function removerTransitorioAuditado(etiqueta) {
  try {
    console.log('=== DEBUG: Removendo produto com etiqueta:', etiqueta);
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = planilha.getSheetByName('Transitorios Auditados');
    
    if (!sheet) {
      throw new Error("Aba 'Transitorios Auditados' não encontrada.");
    }
    
    const dados = sheet.getDataRange().getValues();
    console.log('Total de linhas na planilha:', dados.length);
    
    // Procura pela etiqueta e remove a linha (estrutura: DATA / ETIQUETA / CODIGO / ...)
    for (let i = dados.length - 1; i >= 0; i--) {
      const etiquetaNaLinha = String(dados[i][1]).trim(); // Coluna 1 = Etiqueta
      console.log(`Linha ${i}: Etiqueta encontrada: "${etiquetaNaLinha}"`);
      
      if (etiquetaNaLinha === String(etiqueta).trim()) {
        console.log('Etiqueta encontrada, removendo linha:', i + 1);
        sheet.deleteRow(i + 1); // +1 porque as linhas no Sheets começam em 1
        return 'Produto removido com sucesso!';
      }
    }
    
    console.log('Etiqueta não encontrada:', etiqueta);
    throw new Error('Produto não encontrado para remoção.');
    
  } catch (error) {
    console.error('Erro em removerTransitorioAuditado:', error);
    throw error;
  }
}

/**
 * FUNÇÃO DE TESTE - Verificar se as abas estão sendo encontradas corretamente
 * Execute esta função para testar a busca nas abas
 */
function testarVerificacaoSetados() {
  console.log(' === TESTE DE VERIFICAÇÃO DE SETADOS ===');
  
  // Testar se conseguimos acessar as abas
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['Base Smartphones CD Principal', 'Base RMA', 'Base Transitorios - Parte 1', 'Base Transitorios - Parte 2', 'Base NL'];
  
  abas.forEach(nomeAba => {
    const aba = planilha.getSheetByName(nomeAba);
    if (aba) {
      const dados = aba.getDataRange().getValues();
      console.log(` Aba "${nomeAba}": ${dados.length - 1} produtos encontrados`);
      
      // Mostrar primeiro produto como exemplo
      if (dados.length > 1) {
        const primeiroProduto = dados[1];
        console.log(`   Exemplo: CF=${primeiroProduto[0]}, CODIGO=${primeiroProduto[1]}`);
      }
    } else {
      console.log(` Aba "${nomeAba}": NÃO ENCONTRADA`);
    }
  });
  
  console.log(' === FIM DO TESTE ===');
}

/**
 * FUNÇÃO DE TESTE - Testar especificamente a Base CD Principal CSV
 */
function testarBaseCDPrincipalCSV() {
  console.log(' === TESTE BASE CD PRINCIPAL CSV ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    const nomesPossiveis = [
      'Base CD Principal.csv',
      'Base CD Principal - Parte 1.csv',
      'Base CD Principal - Parte1.csv',
      'Base CD Principal - Parte 2.csv', 
      'Base CD Principal - Parte2.csv'
    ];
    
    console.log(' Verificando arquivos na pasta...');
    
    nomesPossiveis.forEach(nome => {
      const arquivos = pasta.getFilesByName(nome);
      if (arquivos.hasNext()) {
        const arquivo = arquivos.next();
        const tamanhoMB = (arquivo.getSize() / 1024 / 1024).toFixed(2);
        const dataModificacao = arquivo.getLastUpdated();
        console.log(` "${nome}" - ${tamanhoMB} MB - Modificado: ${dataModificacao}`);
      } else {
        console.log(` "${nome}" - NÃO ENCONTRADO`);
      }
    });
    
    // Testar a função de busca
    console.log('\n Testando função de busca...');
    const resultado = buscarNaBaseCDPrincipalCSV('TESTE123');
    console.log(' Resultado do teste:', JSON.stringify(resultado, null, 2));
    
  } catch (error) {
    console.error(' Erro no teste:', error);
  }
  
  console.log(' === FIM DO TESTE CSV ===');
}

/**
 * FUNÇÃO DE TESTE - Testar especificamente as Base Transitórios Parte 1 e 2
 */
function testarBaseTransitorios() {
  console.log(' === TESTE BASE TRANSITÓRIOS PARTE 1 E 2 ===');
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const nomesAbas = ['Base Transitorios - Parte 1', 'Base Transitorios - Parte 2'];
    let totalRegistros = 0;
    
    nomesAbas.forEach((nomeAba, index) => {
      console.log(`\n Verificando "${nomeAba}"...`);
      
      const aba = planilha.getSheetByName(nomeAba);
      if (aba) {
        const dados = aba.getDataRange().getValues();
        const registros = dados.length - 1; // -1 para descontar cabeçalho
        totalRegistros += registros;
        
        console.log(` "${nomeAba}": ${registros} registros encontrados`);
        
        // Mostrar primeiro produto como exemplo
        if (dados.length > 1) {
          const primeiroProduto = dados[1];
          console.log(`   Exemplo: CF=${primeiroProduto[0]}, CODIGO=${primeiroProduto[1]}`);
        }
      } else {
        console.log(` "${nomeAba}": NÃO ENCONTRADA`);
      }
    });
    
    console.log(`\n TOTAL DE REGISTROS TRANSITÓRIOS: ${totalRegistros}`);
    
    // Testar busca combinada
    console.log('\n Testando busca combinada...');
    if (totalRegistros > 0) {
      // Pegar primeiro código da Parte 1 para teste
      const aba1 = planilha.getSheetByName('Base Transitorios - Parte 1');
      if (aba1) {
        const dados = aba1.getDataRange().getValues();
        if (dados.length > 1) {
          const codigoTeste = String(dados[1][1] || dados[1][0]).trim(); // CODIGO ou CF
          if (codigoTeste) {
            console.log(` Testando com código real: "${codigoTeste}"`);
            const resultado = verificarProdutoEmTodasAsBases(codigoTeste);
            console.log(' Resultado:', JSON.stringify(resultado, null, 2));
          }
        }
      }
    }
    
  } catch (error) {
    console.error(' Erro no teste:', error);
  }
  
  console.log(' === FIM DO TESTE TRANSITÓRIOS ===');
}

/**
 * FUNÇÃO DE TESTE - Testar verificação de um produto específico
 * @param {string} codigo - Código para testar
 */
function testarProdutoEspecifico(codigo) {
  console.log(`\n === TESTE PRODUTO ESPECÍFICO: ${codigo} ===`);
  
  const inicio = new Date().getTime();
  const resultado = verificarProdutoEmTodasAsBases(codigo);
  const fim = new Date().getTime();
  const tempoMs = fim - inicio;
  
  console.log(`⏱️ Tempo de busca: ${tempoMs}ms`);
  console.log('📋 Resultado:', JSON.stringify(resultado, null, 2));
  
  return resultado;
}

/**
 * Teste específico para Base CD Principal (ambas as partes)
 * @param {string} codigo - Código para testar
 */
function testarApenasBaseCDPrincipal(codigo) {
  console.log(`\n=== TESTE ESPECÍFICO BASE CD PRINCIPAL: ${codigo} ===`);
  
  const inicio = new Date().getTime();
  const resultado = buscarNaBaseCDPrincipalCSV(codigo);
  const fim = new Date().getTime();
  const tempoMs = fim - inicio;
  
  console.log(`\n📊 RESULTADO FINAL:`);
  console.log(`⏱️ Tempo: ${tempoMs}ms`);
  console.log(`🎯 Encontrado: ${resultado.encontrado}`);
  
  if (resultado.encontrado) {
    console.log(`📂 Arquivo: ${resultado.dados.arquivo}`);
    console.log(`📍 Posição: linha ${resultado.dados.linhaEncontrada}, coluna ${resultado.dados.colunaEncontrada}`);
    console.log(`📝 Dados:`, resultado.dados);
  } else {
    console.log(`❌ Motivo: ${resultado.erro || 'Código não encontrado'}`);
  }
  
  return resultado;
}

/**
 * Teste com alguns códigos de amostra das duas partes
 */
function testeAmostrasBaseCDPrincipal() {
  console.log('\n=== TESTE COM AMOSTRAS DE AMBAS AS PARTES ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // Buscar amostras da Parte 1
    console.log('\n📂 Obtendo amostras da Parte 1...');
    const parte1 = pasta.getFilesByName('Base CD Principal - Parte 1.csv');
    if (parte1.hasNext()) {
      const arquivo1 = parte1.next();
      const conteudo1 = arquivo1.getBlob().getDataAsString('UTF-8');
      const linhas1 = conteudo1.split('\n');
      
      // Pegar algumas amostras da Parte 1 (linhas 2, 3, 4)
      for (let i = 1; i <= Math.min(3, linhas1.length - 1); i++) {
        const linha = linhas1[i];
        if (linha && linha.trim()) {
          const colunas = linha.split(',').map(c => c.replace(/"/g, '').trim());
          if (colunas.length >= 2 && colunas[1]) {
            console.log(`   Parte 1 - Linha ${i + 1}: Código "${colunas[1]}" (CF: ${colunas[0] || 'N/A'})`);
          }
        }
      }
    }
    
    // Buscar amostras da Parte 2
    console.log('\n📂 Obtendo amostras da Parte 2...');
    const parte2 = pasta.getFilesByName('Base CD Principal - Parte 2.csv');
    if (parte2.hasNext()) {
      const arquivo2 = parte2.next();
      const conteudo2 = arquivo2.getBlob().getDataAsString('UTF-8');
      const linhas2 = conteudo2.split('\n');
      
      // Pegar algumas amostras da Parte 2 (linhas 2, 3, 4)
      for (let i = 1; i <= Math.min(3, linhas2.length - 1); i++) {
        const linha = linhas2[i];
        if (linha && linha.trim()) {
          const colunas = linha.split(',').map(c => c.replace(/"/g, '').trim());
          if (colunas.length >= 2 && colunas[1]) {
            console.log(`   Parte 2 - Linha ${i + 1}: Código "${colunas[1]}" (CF: ${colunas[0] || 'N/A'})`);
          }
        }
      }
    }
    
    console.log('\n💡 Use um desses códigos para testar: testarApenasBaseCDPrincipal("CODIGO_AQUI")');
    
  } catch (error) {
    console.error('❌ Erro ao obter amostras:', error.message);
  }
}

/**
 * TESTE ESPECÍFICO PARA CACHE - Verifica se o problema é cache do browser
 */
function testarCacheVersusRealidade() {
  console.log('\n=== TESTE: CACHE vs REALIDADE ===');
  
  try {
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // 1. VERIFICAR ARQUIVOS REAIS NO DRIVE
    console.log('1️⃣ ARQUIVOS REAIS NO DRIVE:');
    const arquivos = pasta.getFiles();
    const arquivosEncontrados = [];
    
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nome = arquivo.getName();
      
      if (nome.includes('Base CD Principal') && nome.endsWith('.csv')) {
        arquivosEncontrados.push({
          nome: nome,
          tamanho: (arquivo.getSize() / 1024 / 1024).toFixed(2) + ' MB',
          modificacao: arquivo.getLastUpdated()
        });
      }
    }
    
    arquivosEncontrados.forEach((arq, idx) => {
      console.log(`  ${idx + 1}. "${arq.nome}" (${arq.tamanho}) - ${arq.modificacao}`);
    });
    
    // 2. TESTAR CACHE DO APPS SCRIPT
    console.log('\n2️⃣ TESTANDO CACHE DO APPS SCRIPT:');
    const cache = CacheService.getDocumentCache();
    
    // Limpar qualquer cache relacionado
    const chavesParaLimpar = [
      'principal:29-0-2-1',
      'principal_otimizado:29-0-2-1',
      'enderecos_cd_principal',
      'cd_principal_enderecos'
    ];
    
    chavesParaLimpar.forEach(chave => {
      try {
        cache.remove(chave);
        console.log(`   Cache removido: ${chave}`);
      } catch (e) {
        console.log(`   Cache não encontrado: ${chave}`);
      }
    });
    
    // 3. BUSCAR ENDEREÇOS SEM CACHE (FORÇADO)
    console.log('\n3️⃣ BUSCANDO ENDEREÇOS SEM CACHE (TEMPO REAL):');
    
    // Testar cada arquivo individualmente
    arquivosEncontrados.forEach(arquivoInfo => {
      console.log(`\n   Testando: ${arquivoInfo.nome}`);
      
      const arquivos = pasta.getFilesByName(arquivoInfo.nome);
      if (arquivos.hasNext()) {
        const arquivo = arquivos.next();
        
        try {
          const conteudo = arquivo.getBlob().getDataAsString();
          const linhas = conteudo.split('\n');
          
          console.log(`     Total de linhas: ${linhas.length}`);
          
          // Procurar especificamente por 29-0-2-1
          let encontrou29021 = false;
          let amostraEnderecos = new Set();
          
          for (let i = 1; i < Math.min(1000, linhas.length); i++) {
            const linha = linhas[i];
            if (!linha || linha.trim() === '') continue;
            
            const colunas = linha.split(';').map(c => c.replace(/"/g, '').trim());
            if (colunas.length >= 5) {
              const endereco = colunas[4];
              if (endereco) {
                amostraEnderecos.add(endereco);
                
                if (endereco === '29-0-2-1') {
                  encontrou29021 = true;
                  console.log(`     🎯 ENCONTRADO 29-0-2-1 na linha ${i + 1}!`);
                  console.log(`     Dados: CF="${colunas[0]}", CODIGO="${colunas[1]}", FINAL="${colunas[2]}"`);
                }
              }
            }
          }
          
          console.log(`     Amostra de endereços (primeiros 1000): ${amostraEnderecos.size} únicos`);
          console.log(`     29-0-2-1 encontrado: ${encontrou29021 ? '✅ SIM' : '❌ NÃO'}`);
          
          // Mostrar alguns endereços de exemplo
          const enderecosArray = Array.from(amostraEnderecos);
          const enderecos29 = enderecosArray.filter(e => e.startsWith('29-'));
          if (enderecos29.length > 0) {
            console.log(`     Endereços que começam com "29": [${enderecos29.slice(0, 5).join(', ')}${enderecos29.length > 5 ? '...' : ''}]`);
          }
          
        } catch (erro) {
          console.error(`     ERRO ao processar ${arquivoInfo.nome}:`, erro.message);
        }
      }
    });
    
    // 4. TESTAR FUNÇÃO ATUAL (QUE PODE ESTAR USANDO CACHE)
    console.log('\n4️⃣ TESTANDO FUNÇÃO ATUAL (carregarApenasEnderecosCDPrincipal):');
    const enderecosFuncaoAtual = carregarApenasEnderecosCDPrincipal();
    
    console.log(`   Endereços retornados pela função atual: ${enderecosFuncaoAtual ? enderecosFuncaoAtual.length : 'null'}`);
    
    if (enderecosFuncaoAtual && enderecosFuncaoAtual.length > 0) {
      const tem29021 = enderecosFuncaoAtual.includes('29-0-2-1');
      console.log(`   29-0-2-1 na função atual: ${tem29021 ? '✅ SIM' : '❌ NÃO'}`);
      
      const enderecos29Funcao = enderecosFuncaoAtual.filter(e => e.startsWith('29-'));
      console.log(`   Endereços "29-" na função: ${enderecos29Funcao.length} encontrados`);
      if (enderecos29Funcao.length > 0) {
        console.log(`   Exemplos: [${enderecos29Funcao.slice(0, 5).join(', ')}]`);
      }
    }
    
    // 5. COMPARAR TIMESTAMP DOS ARQUIVOS
    console.log('\n5️⃣ ANÁLISE DE TIMESTAMPS:');
    arquivosEncontrados.forEach(arq => {
      const agora = new Date();
      const diff = agora - arq.modificacao;
      const horas = Math.floor(diff / (1000 * 60 * 60));
      const minutos = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
      
      console.log(`   "${arq.nome}": modificado há ${horas}h${minutos}m`);
    });
    
    console.log('\n=== CONCLUSÃO DO TESTE DE CACHE ===');
    console.log('Se os arquivos reais têm 29-0-2-1 mas a função não retorna,');
    console.log('então é definitivamente um problema de cache ou lógica da função.');
    
    return {
      arquivosReais: arquivosEncontrados.length,
      cacheRemovido: chavesParaLimpar.length,
      funcaoAtual: enderecosFuncaoAtual ? enderecosFuncaoAtual.length : 0
    };
    
  } catch (error) {
    console.error('❌ Erro no teste de cache:', error);
    return { erro: error.message };
  }
}

/**
 * FUNÇÃO PARA FORÇAR RECARGA SEM CACHE
 */
function forcarRecargaSemCache() {
  console.log('\n=== FORÇANDO RECARGA SEM CACHE ===');
  
  try {
    // 1. Limpar TODO o cache
    console.log('1️⃣ Limpando cache...');
    const cache = CacheService.getDocumentCache();
    
    try {
      cache.removeAll();
      console.log('   ✅ Cache documento limpo');
    } catch (e) {
      console.log('   ⚠️ Erro ao limpar cache:', e.message);
    }
    
    // 2. Forçar nova leitura dos arquivos
    console.log('2️⃣ Forçando nova leitura...');
    
    const pastaId = '1cWU8Vy3lSvQQvXgLWo25FZdzEA-UNueE';
    const pasta = DriveApp.getFolderById(pastaId);
    
    // Buscar Parte 1 e 2 com timestamp atual
    const arquivos = ['Base CD Principal - Parte 1.csv', 'Base CD Principal - Parte 2.csv'];
    const enderecos = new Set();
    
    for (const nomeArquivo of arquivos) {
      console.log(`   Processando: ${nomeArquivo}`);
      
      const arquivos = pasta.getFilesByName(nomeArquivo);
      if (arquivos.hasNext()) {
        const arquivo = arquivos.next();
        
        // Forçar nova leitura com timestamp
        const agora = new Date().getTime();
        console.log(`     Lendo arquivo em: ${new Date(agora)}`);
        
        const conteudo = arquivo.getBlob().getDataAsString();
        const linhas = conteudo.split('\n');
        
        console.log(`     Linhas lidas: ${linhas.length}`);
        
        // Processar com logs detalhados
        let enderecosEsteArquivo = 0;
        let encontrou29021 = false;
        
        for (let i = 1; i < linhas.length; i++) {
          const linha = linhas[i];
          if (!linha || linha.trim() === '') continue;
          
          const colunas = linha.split(';').map(c => c.replace(/"/g, '').trim());
          if (colunas.length >= 5 && colunas[4]) {
            const endereco = colunas[4];
            if (/^\d{1,3}-\d{1,2}-\d{1,3}-\d{1,2}$/.test(endereco)) {
              enderecos.add(endereco);
              enderecosEsteArquivo++;
              
              if (endereco === '29-0-2-1') {
                encontrou29021 = true;
                console.log(`     🎯 29-0-2-1 ENCONTRADO na linha ${i + 1}`);
              }
            }
          }
          
          // Log a cada 10000 linhas
          if (i % 10000 === 0) {
            console.log(`       Processadas ${i} linhas...`);
          }
        }
        
        console.log(`     Endereços únicos neste arquivo: ${enderecosEsteArquivo}`);
        console.log(`     29-0-2-1 encontrado: ${encontrou29021 ? '✅ SIM' : '❌ NÃO'}`);
      }
    }
    
    const enderecosArray = Array.from(enderecos).sort();
    
    console.log('\n3️⃣ RESULTADO FINAL:');
    console.log(`   Total de endereços únicos: ${enderecosArray.length}`);
    console.log(`   29-0-2-1 no resultado final: ${enderecosArray.includes('29-0-2-1') ? '✅ SIM' : '❌ NÃO'}`);
    
    if (enderecosArray.length > 0) {
      console.log(`   Primeiros 10: [${enderecosArray.slice(0, 10).join(', ')}]`);
      console.log(`   Últimos 10: [${enderecosArray.slice(-10).join(', ')}]`);
      
      const enderecos29 = enderecosArray.filter(e => e.startsWith('29-'));
      console.log(`   Endereços "29-": ${enderecos29.length} encontrados`);
      if (enderecos29.length > 0) {
        console.log(`     Exemplos "29-": [${enderecos29.slice(0, 10).join(', ')}]`);
      }
    }
    
    return enderecosArray;
    
  } catch (error) {
    console.error('❌ Erro na recarga forçada:', error);
    return [];
  }
}
