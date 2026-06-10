/**
 * Power Pivot — Atualização diária da Base de Agências
 * =====================================================
 *
 * Este script lê o arquivo CSV da base no Google Drive (separado por ";",
 * cabeçalho: nome_completo; conta; cpf_cnpj; data_nascimento;
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
 */

var CONFIG = {
  // ID do arquivo CSV da base no Google Drive
  CSV_FILE_ID: '1XoQamjbYeu0xdIDpj-Y9X4ngqGrCE4Ku',
  // Separador de campos do CSV
  SEPARADOR: ';',
  // Codificação do arquivo CSV ('UTF-8' ou 'ISO-8859-1')
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
 * Função principal: lê o CSV e grava os campos necessários na planilha base.
 * É esta função que o gatilho diário executa.
 */
function atualizarBase() {
  var arquivo = DriveApp.getFileById(CONFIG.CSV_FILE_ID);
  var texto = arquivo.getBlob().getDataAsString(CONFIG.ENCODING);

  // Se a decodificação UTF-8 gerou caracteres inválidos, tenta Latin-1
  if (texto.indexOf('�') !== -1 && CONFIG.ENCODING === 'UTF-8') {
    texto = arquivo.getBlob().getDataAsString('ISO-8859-1');
  }
  texto = texto.replace(/^\uFEFF/, '');

  var linhas = Utilities.parseCsv(texto, CONFIG.SEPARADOR);
  if (linhas.length < 2) {
    throw new Error('CSV vazio ou sem linhas de dados.');
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
    throw new Error("Coluna '" + campo.origem.join("' ou '") + "' não encontrada no CSV. Cabeçalho lido: " + cabecalho.join(' | '));
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
