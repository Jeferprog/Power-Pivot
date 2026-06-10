/**
 * Power Pivot — Atualização diária da Base de Agências
 * =====================================================
 *
 * Este script lê o arquivo da base no Google Drive (CSV separado por ";" ou
 * Planilha Google, cabeçalho: nome_completo; conta; cpf_cnpj; data_nascimento;
 * m_carteira_responsavel; nm_carteira; nm_agencia) e copia, UMA VEZ POR DIA,
 * somente os campos que o dashboard precisa — Conta, Nome Agência e Gerente
 * de Contas — para uma planilha Google ("planilha base"). O dashboard lê essa
 * planilha base — assim os dados sensíveis (CPF/CNPJ, data de nascimento)
 * nunca chegam ao navegador.
 *
 * COMO INSTALAR (uma única vez):
 *   1. Acesse https://script.google.com e crie um novo projeto.
 *   2. Cole todo este código no editor e salve.
 *   3. Selecione a função "atualizarBase" e clique em "Executar" para testar
 *      (autorize as permissões na primeira execução). Veja no log a URL da
 *      planilha base criada.
 *   4. Selecione a função "instalarGatilhoDiario" e clique em "Executar".
 *      Pronto: a base será atualizada automaticamente todos os dias.
 *   5. Abra a planilha base (URL no log), compartilhe como "Qualquer pessoa
 *      com o link" (Leitor) e cole o link dela no campo "Atualização
 *      Automática (Google Drive)" do dashboard Power Pivot.
 *
 * EM CASO DE ERRO: execute a função "diagnosticarArquivo" e veja o log
 * (Ctrl+Enter) — ela mostra o tipo do arquivo, o cabeçalho e o início do
 * conteúdo lido.
 */

var CONFIG = {
  // ID do arquivo da base no Google Drive — ou de uma PASTA: nesse caso o
  // script usa o arquivo mais recente dentro dela (ex.: exportações do Qlik)
  CSV_FILE_ID: '1XoQamjbYeu0xdIDpj-Y9X4ngqGrCE4Ku',
  // Quando o ID acima é de uma pasta, considera apenas arquivos cujo nome
  // contém este texto (deixe '' para considerar todos). Ex.: '.csv'
  FILTRO_NOME: '',
  // Codificação preferida do arquivo CSV ('UTF-8' ou 'ISO-8859-1')
  ENCODING: 'UTF-8',
  // Nome da planilha base criada automaticamente
  NOME_PLANILHA_BASE: 'Power Pivot - Base de Agências (Automática)',
  // Nome da aba dentro da planilha base
  NOME_ABA: 'base',
  // Hora do dia (0-23) em que o gatilho diário roda
  HORA_GATILHO: 6,
  // Mapeamento: coluna de origem no CSV → cabeçalho na planilha base.
  // Em "origem" pode-se listar mais de um nome de coluna (usa o primeiro que existir).
  CAMPOS: [
    { origem: ['conta'], destino: 'Conta' },
    { origem: ['nm_agencia'], destino: 'Nome Agência' },
    { origem: ['m_carteira_responsavel', 'nm_carteira_responsavel'], destino: 'Gerente de Contas' }
  ]
};

/**
 * Função principal: lê a base (CSV ou Planilha Google) e grava os campos
 * necessários na planilha base. É esta função que o gatilho diário executa.
 */
function atualizarBase() {
  var linhas = lerLinhasDaOrigem();

  if (!linhas || linhas.length < 2) {
    throw new Error('Base vazia ou sem linhas de dados. Execute a função "diagnosticarArquivo" e veja o log para entender o conteúdo do arquivo.');
  }

  // Localiza as colunas desejadas pelo cabeçalho (sem diferenciar maiúsculas)
  var cabecalho = linhas[0].map(function (c) {
    return String(c).replace(/^\uFEFF/, '').trim().toLowerCase();
  });
  var indices = CONFIG.CAMPOS.map(function (campo) {
    var candidatos = campo.origem.map(function (c) { return c.toLowerCase(); });
    for (var j = 0; j < candidatos.length; j++) {
      var idx = cabecalho.indexOf(candidatos[j]);
      if (idx !== -1) return idx;
    }
    throw new Error("Coluna '" + campo.origem.join("' ou '") + "' não encontrada na base. Cabeçalho lido: " + cabecalho.join(' | '));
  });

  // Monta as linhas de saída, ignorando registros sem conta
  var saida = [CONFIG.CAMPOS.map(function (campo) { return campo.destino; })];
  for (var i = 1; i < linhas.length; i++) {
    var valores = indices.map(function (idx) {
      return linhas[i][idx] !== undefined ? String(linhas[i][idx]).trim() : '';
    });
    if (valores[0] !== '') saida.push(valores);
  }

  // Grava tudo de uma vez na planilha base
  var planilha = obterPlanilhaBase();
  var aba = planilha.getSheetByName(CONFIG.NOME_ABA) || planilha.insertSheet(CONFIG.NOME_ABA);
  aba.clearContents();
  aba.getRange(1, 1, saida.length, CONFIG.CAMPOS.length).setValues(saida);

  // Registra a data/hora da atualização em uma aba de controle
  var abaInfo = planilha.getSheetByName('info') || planilha.insertSheet('info');
  abaInfo.clearContents();
  abaInfo.getRange(1, 1, 2, 2).setValues([
    ['ultima_atualizacao', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')],
    ['registros', String(saida.length - 1)]
  ]);

  Logger.log('Base atualizada: ' + (saida.length - 1) + ' registros.');
  Logger.log('Planilha base: ' + planilha.getUrl());
  Logger.log('Cole este link no dashboard: ' + planilha.getUrl());
}

/**
 * Lê a origem e devolve uma matriz de linhas (a primeira é o cabeçalho).
 * Aceita arquivo CSV, arquivo já convertido em Planilha Google ou uma pasta
 * (usa o arquivo mais recente dentro dela).
 */
function lerLinhasDaOrigem() {
  var arquivo = obterArquivoBase();
  var mime = arquivo.getMimeType();

  // Caso 1: o arquivo foi convertido para Planilha Google ao subir no Drive
  if (mime === MimeType.GOOGLE_SHEETS) {
    var abaOrigem = SpreadsheetApp.openById(arquivo.getId()).getSheets()[0];
    return abaOrigem.getDataRange().getValues();
  }

  // Caso 2: arquivos Excel não são suportados nesta leitura direta
  if (mime === MimeType.MICROSOFT_EXCEL || mime === MimeType.MICROSOFT_EXCEL_LEGACY) {
    throw new Error('O arquivo da base é um Excel (.xls/.xlsx). Converta-o para CSV ou abra-o no Drive e salve como Planilha Google.');
  }

  // Caso 3: arquivo de texto/CSV
  var texto = lerTextoCsv(arquivo);
  var separador = detectarSeparador(texto);
  return Utilities.parseCsv(texto, separador);
}

/**
 * Resolve o arquivo da base. Se CONFIG.CSV_FILE_ID apontar para uma pasta,
 * devolve o arquivo modificado mais recentemente dentro dela (respeitando
 * CONFIG.FILTRO_NOME); caso contrário devolve o próprio arquivo.
 */
function obterArquivoBase() {
  var item = DriveApp.getFileById(CONFIG.CSV_FILE_ID);
  if (item.getMimeType() !== MimeType.FOLDER) return item;

  var pasta = DriveApp.getFolderById(CONFIG.CSV_FILE_ID);
  var arquivos = pasta.getFiles();
  var escolhido = null;
  while (arquivos.hasNext()) {
    var f = arquivos.next();
    if (CONFIG.FILTRO_NOME && f.getName().toLowerCase().indexOf(CONFIG.FILTRO_NOME.toLowerCase()) === -1) continue;
    if (!escolhido || f.getLastUpdated() > escolhido.getLastUpdated()) escolhido = f;
  }
  if (!escolhido) {
    throw new Error('Nenhum arquivo encontrado na pasta "' + pasta.getName() + '"' +
      (CONFIG.FILTRO_NOME ? ' com nome contendo "' + CONFIG.FILTRO_NOME + '"' : '') + '.');
  }
  Logger.log('Pasta "' + pasta.getName() + '": usando o arquivo mais recente — "' + escolhido.getName() + '" (' + escolhido.getLastUpdated() + ').');
  return escolhido;
}

/**
 * Decodifica o CSV tratando codificação (UTF-8 / Latin-1 / UTF-16), BOM e
 * quebras de linha no formato antigo (\r).
 */
function lerTextoCsv(arquivo) {
  var blob = arquivo.getBlob();
  var texto = blob.getDataAsString(CONFIG.ENCODING);

  // Bytes nulos indicam UTF-16 (comum em exportações do Windows)
  if (texto.indexOf('\u0000') !== -1) {
    texto = blob.getDataAsString('UTF-16');
  }
  // Caractere de substituição indica codificação errada — tenta Latin-1
  if (texto.indexOf('�') !== -1) {
    texto = blob.getDataAsString('ISO-8859-1');
  }

  return texto.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
}

/**
 * Detecta o separador do CSV pela primeira linha (";", "," ou tabulação).
 */
function detectarSeparador(texto) {
  var fimLinha = texto.indexOf('\n');
  var primeiraLinha = fimLinha === -1 ? texto : texto.slice(0, fimLinha);
  var contagens = [
    { sep: ';', qtd: primeiraLinha.split(';').length },
    { sep: ',', qtd: primeiraLinha.split(',').length },
    { sep: '\t', qtd: primeiraLinha.split('\t').length }
  ];
  contagens.sort(function (a, b) { return b.qtd - a.qtd; });
  return contagens[0].sep;
}

/**
 * Mostra no log informações do arquivo da base, para investigar problemas
 * de leitura. Execute manualmente e veja o log.
 */
function diagnosticarArquivo() {
  var item = DriveApp.getFileById(CONFIG.CSV_FILE_ID);

  // Se o ID for de uma pasta, lista o conteúdo e diagnostica o arquivo escolhido
  if (item.getMimeType() === MimeType.FOLDER) {
    var pasta = DriveApp.getFolderById(CONFIG.CSV_FILE_ID);
    Logger.log('O ID configurado é uma PASTA: "' + pasta.getName() + '". Conteúdo:');
    var lista = pasta.getFiles();
    while (lista.hasNext()) {
      var f = lista.next();
      Logger.log('  - "' + f.getName() + '" | ' + f.getMimeType() + ' | ' + f.getSize() + ' bytes | modificado em ' + f.getLastUpdated());
    }
  }

  var arquivo = obterArquivoBase();
  var mime = arquivo.getMimeType();
  Logger.log('Arquivo analisado: ' + arquivo.getName());
  Logger.log('Tipo (MIME): ' + mime);
  Logger.log('Tamanho: ' + arquivo.getSize() + ' bytes');

  if (mime === MimeType.GOOGLE_SHEETS) {
    var aba = SpreadsheetApp.openById(arquivo.getId()).getSheets()[0];
    Logger.log('É uma Planilha Google. Aba "' + aba.getName() + '": ' + aba.getLastRow() + ' linhas x ' + aba.getLastColumn() + ' colunas.');
    if (aba.getLastRow() > 0) {
      Logger.log('Cabeçalho: ' + aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].join(' | '));
    }
    return;
  }

  var texto = lerTextoCsv(arquivo);
  var separador = detectarSeparador(texto);
  var linhas = Utilities.parseCsv(texto, separador);
  Logger.log('Separador detectado: "' + (separador === '\t' ? 'TAB' : separador) + '"');
  Logger.log('Total de caracteres lidos: ' + texto.length);
  Logger.log('Linhas interpretadas: ' + linhas.length);
  Logger.log('Primeiros 300 caracteres do conteúdo:');
  Logger.log(texto.slice(0, 300));
}

/**
 * Retorna a planilha base, criando-a na primeira execução e guardando o ID
 * nas propriedades do script para reutilizar nas execuções seguintes.
 */
function obterPlanilhaBase() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PLANILHA_BASE_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // Planilha foi excluída — cria outra abaixo
    }
  }
  var planilha = SpreadsheetApp.create(CONFIG.NOME_PLANILHA_BASE);
  props.setProperty('PLANILHA_BASE_ID', planilha.getId());
  return planilha;
}

/**
 * Instala o gatilho que executa atualizarBase() uma vez por dia.
 * Execute esta função UMA única vez (reexecutar apenas substitui o gatilho).
 */
function instalarGatilhoDiario() {
  // Remove gatilhos antigos desta função para não duplicar
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'atualizarBase') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('atualizarBase')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.HORA_GATILHO)
    .create();

  Logger.log('Gatilho diário instalado: atualizarBase() roda todo dia por volta das ' + CONFIG.HORA_GATILHO + 'h.');
}

/**
 * Remove o gatilho diário, caso queira desativar a atualização automática.
 */
function removerGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'atualizarBase') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Gatilho diário removido.');
}
