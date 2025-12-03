/**
 * Aplicativo Web para Controle e Lançamento de Processos
 * Backend - Google Apps Script
 */

/**
 * Função principal executada quando o URL do aplicativo web é acessado
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Controle de Processos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Função otimizada para buscar todos os dados iniciais dos menus suspensos
 * @return {Object} Objeto contendo todos os dados dos dropdowns
 */
function getDropdownData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Buscar dados das abas de configuração
    const configSistemas = ss.getSheetByName('Config Sistemas');
    const configUGOrigem = ss.getSheetByName('Config UGOrigem');
    const configAssuntos = ss.getSheetByName('Config Assuntos');
    const configSetor = ss.getSheetByName('Config Setor');
    const configAssimetria = ss.getSheetByName('Config Assimetria');
    
    // Obter dados das colunas A (a partir da linha 2)
    const sistemas = configSistemas.getRange('A2:A').getValues()
      .filter(row => row[0] !== '')
      .map(row => row[0]);
    
    const ugOrigem = configUGOrigem.getRange('A2:A').getValues()
      .filter(row => row[0] !== '')
      .map(row => row[0]);
    
    const assuntosData = configAssuntos.getRange('A2:A').getValues()
      .filter(row => row[0] !== '')
      .map(row => row[0]);
    
    // Remover duplicatas dos assuntos
    const assuntos = [...new Set(assuntosData)];
    
    const setores = configSetor.getRange('A2:A').getValues()
      .filter(row => row[0] !== '')
      .map(row => row[0]);
    
    // Buscar situações da aba Config Assimetria (valores únicos da coluna A)
    const situacoesData = configAssimetria.getRange('A2:A').getValues()
      .filter(row => row[0] !== '')
      .map(row => row[0]);
    
    // Remover duplicatas das situações
    const situacoes = [...new Set(situacoesData)];
    
    return {
      sistemas: sistemas,
      ugOrigem: ugOrigem,
      assuntos: assuntos,
      setores: setores,
      situacoes: situacoes
    };
    
  } catch (error) {
    console.error('Erro ao buscar dados dos dropdowns:', error);
    return {
      sistemas: [],
      ugOrigem: [],
      assuntos: [],
      setores: [],
      situacoes: []
    };
  }
}

/**
 * Busca os sub-assuntos relacionados a um assunto específico
 * @param {string} assuntoSelecionado - O assunto selecionado
 * @return {Array} Array de sub-assuntos
 */
function getSubAssuntos(assuntoSelecionado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configAssuntos = ss.getSheetByName('Config Assuntos');
    
    // Obter dados das colunas A e B
    const data = configAssuntos.getRange('A2:B').getValues();
    
    // Filtrar sub-assuntos relacionados ao assunto selecionado
    const subAssuntos = data
      .filter(row => row[0] === assuntoSelecionado && row[1] !== '')
      .map(row => row[1]);
    
    return subAssuntos;
    
  } catch (error) {
    console.error('Erro ao buscar sub-assuntos:', error);
    return [];
  }
}

/**
 * Busca as assimetrias relacionadas a uma situação específica
 * @param {string} situacaoSelecionada - A situação selecionada
 * @return {Array} Array de assimetrias
 */
function getAssimetrias(situacaoSelecionada) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configAssimetria = ss.getSheetByName('Config Assimetria');
    
    // Obter dados das colunas A e B
    const data = configAssimetria.getRange('A2:B').getValues();
    
    // Filtrar assimetrias relacionadas à situação selecionada
    const assimetrias = data
      .filter(row => row[0] === situacaoSelecionada && row[1] !== '')
      .map(row => row[1]);
    
    return assimetrias;
    
  } catch (error) {
    console.error('Erro ao buscar assimetrias:', error);
    return [];
  }
}

/**
 * Obtém o nome do usuário logado
 * @return {string} Nome do usuário ou email se não encontrado
 */
function getUsuarioLogado() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configUsuarios = ss.getSheetByName('Config Usuarios');
    
    // Buscar dados das colunas A e B
    const data = configUsuarios.getRange('A2:B').getValues();
    
    // Procurar o email na coluna A e retornar o nome da coluna B
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === email) {
        return data[i][1] || email;
      }
    }
    
    // Se não encontrar, retornar o próprio email
    return email;
    
  } catch (error) {
    console.error('Erro ao obter usuário logado:', error);
    return '';
  }
}

/**
 * Adiciona um novo registro na aba "Dados"
 * @param {Object} formData - Dados do formulário
 * @return {Object} Objeto com status da operação
 */
function addNewRecord(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dadosSheet = ss.getSheetByName('Dados');
    
    // Ordem dos dados conforme especificado:
    // [Sistema, Núm.PAE/SAJ, Interessado, Entrada, Situação, Assimetria, Observação, UG Origem, Assunto, Sub Assunto, ACI Responsável, Destino, Saída]
    const novaLinha = [
      formData.sistema || '',
      formData.numPAESAJ || '',
      formData.interessado || '',
      formData.entrada || '',
      formData.situacao || '',
      formData.assimetria || '',
      formData.observacao || '',
      formData.ugOrigem || '',
      formData.assunto || '',
      formData.subAssunto || '',
      formData.aciResponsavel || '',
      formData.destino || '',
      formData.saida || ''
    ];
    
    // Adicionar nova linha
    dadosSheet.appendRow(novaLinha);
    
    return {
      status: 'success',
      message: 'Registro salvo com sucesso!'
    };
    
  } catch (error) {
    console.error('Erro ao salvar registro:', error);
    return {
      status: 'error',
      message: 'Erro ao salvar registro: ' + error.toString()
    };
  }
}

/**
 * Função de teste para verificar a estrutura da planilha
 * @return {Object} Informações sobre as abas e dados
 */
function testSheetStructure() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    const info = {
      totalSheets: sheets.length,
      sheetNames: sheets.map(sheet => sheet.getName()),
      dadosSheetExists: false,
      dadosSheetRows: 0,
      dadosSheetCols: 0
    };
    
    // Verificar aba "Dados"
    const dadosSheet = ss.getSheetByName('Dados');
    if (dadosSheet) {
      info.dadosSheetExists = true;
      info.dadosSheetRows = dadosSheet.getLastRow();
      info.dadosSheetCols = dadosSheet.getLastColumn();
      
      // Se há dados, pegar uma amostra
      if (info.dadosSheetRows > 1) {
        const sampleData = dadosSheet.getRange(1, 1, Math.min(3, info.dadosSheetRows), info.dadosSheetCols).getValues();
        info.sampleData = sampleData;
      }
    }
    
    return info;
    
  } catch (error) {
    return {
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Função de teste simples para verificar dados da aba "Dados"
 * @return {Object} Dados de teste
 */
function testGetDados() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dadosSheet = ss.getSheetByName('Dados');
    
    if (!dadosSheet) {
      return { error: 'Aba "Dados" não encontrada' };
    }
    
    const lastRow = dadosSheet.getLastRow();
    const lastCol = dadosSheet.getLastColumn();
    
    console.log('Teste - Última linha:', lastRow, 'Última coluna:', lastCol);
    
    if (lastRow <= 1) {
      return { 
        message: 'Apenas cabeçalho encontrado',
        lastRow: lastRow,
        lastCol: lastCol
      };
    }
    
    // Pegar os primeiros 3 registros para teste
    const numRows = Math.min(3, lastRow - 1);
    const data = dadosSheet.getRange(2, 1, numRows, lastCol).getValues();
    
    return {
      lastRow: lastRow,
      lastCol: lastCol,
      sampleData: data,
      message: 'Dados encontrados com sucesso'
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Busca os últimos 20 registros da aba "Dados"
 * @return {Array} Array de objetos representando os registros
 */
function getRecentRecords() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Tentando obter a aba com o nome: Dados para leitura de registros recentes.');
    var aba = planilha.getSheetByName('Dados');

    if (aba == null) {
      Logger.log('A aba "Dados" não foi encontrada ao ler registros recentes. Retornando array vazio.');
      return [];
    }
    Logger.log('Aba "Dados" encontrada para leitura de registros recentes.');

    var ultimaLinha = aba.getLastRow();
    var ultimaColuna = aba.getLastColumn();
    Logger.log('Última linha da aba "Dados": ' + ultimaLinha + ', Última coluna: ' + ultimaColuna);

    // Ajustar o número de colunas esperadas
    var numColunasEsperadas = 13;
    if (ultimaColuna < numColunasEsperadas) {
        Logger.log('Atenção: A planilha tem menos colunas do que o esperado. Lendo até a última coluna disponível.');
        ultimaColuna = numColunasEsperadas;
    }

    if (ultimaLinha < 2 || ultimaColuna === 0) {
      Logger.log('Planilha "Dados" está vazia, contém apenas cabeçalho ou não há colunas. Retornando array vazio.');
      return [];
    }

    // Calcular quantas linhas buscar (máximo 20)
    var startRow = Math.max(2, ultimaLinha - 19);
    var numRows = ultimaLinha - startRow + 1;
    
    Logger.log('Buscando registros da linha ' + startRow + ' até ' + ultimaLinha + ' (' + numRows + ' registros)');

    // Obtém os dados da planilha, exceto a primeira linha (cabeçalho)
    var dadosBrutos = aba.getRange(startRow, 1, numRows, ultimaColuna).getValues();
    Logger.log('Dados brutos obtidos: ' + dadosBrutos.length + ' linhas');

    // Processa os dados para garantir que as datas sejam formatadas corretamente
    var dadosProcessados = dadosBrutos.map(function(row, index) {
      // Log da primeira linha para debug
      if (index === 0) {
        Logger.log('Primeira linha bruta: ' + JSON.stringify(row));
      }
      
      var registro = {
        sistema: row[0] ? row[0].toString() : '',
        numPAESAJ: row[1] ? row[1].toString() : '',
        interessado: row[2] ? row[2].toString() : '',
        entrada: row[3] instanceof Date ? formatDate(row[3]) : (row[3] ? row[3].toString() : ''),
        situacao: row[4] ? row[4].toString() : '',
        assimetria: row[5] ? row[5].toString() : '',
        observacao: row[6] ? row[6].toString() : '',
        ugOrigem: row[7] ? row[7].toString() : '',
        assunto: row[8] ? row[8].toString() : '',
        subAssunto: row[9] ? row[9].toString() : '',
        aciResponsavel: row[10] ? row[10].toString() : '',
        destino: row[11] ? row[11].toString() : '',
        saida: row[12] instanceof Date ? formatDate(row[12]) : (row[12] ? row[12].toString() : '')
      };
      
      // Log do primeiro registro processado
      if (index === 0) {
        Logger.log('Primeiro registro processado: ' + JSON.stringify(registro));
      }
      
      return registro;
    });

    // Retornar em ordem reversa (mais recentes primeiro)
    var registrosReversed = dadosProcessados.reverse();
    Logger.log('Total de registros retornados: ' + registrosReversed.length);

    return registrosReversed;

  } catch (error) {
    Logger.log('Erro ao buscar registros recentes: ' + error.toString());
    return [];
  }
}

/**
 * Função auxiliar para formatar datas
 * @param {Date} date - Data a ser formatada
 * @return {string} Data formatada como dd/mm/aaaa
 */
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  // Adiciona 1 dia conforme solicitado pelo usuário (corrige deslocamento observado)
  var adjusted = new Date(date.getTime());
  adjusted.setDate(adjusted.getDate() + 1);

  var dia = adjusted.getDate().toString().padStart(2, '0');
  var mes = (adjusted.getMonth() + 1).toString().padStart(2, '0');
  var ano = adjusted.getFullYear();
  return dia + '/' + mes + '/' + ano;
}

/**
 * Retorna contagem de registros por UG Origem entre duas datas (inclusive).
 * startDate e endDate podem ser strings no formato 'yyyy-mm-dd' ou objetos Date.
 * Retorna um array de objetos: [{ ugOrigem: 'UG X', count: 10 }, ...]
 */
function getCountsByUGOrigem(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName('Dados');
    if (!aba) return [];

    // Converter parâmetros para Date
    var sDate = (startDate instanceof Date) ? startDate : new Date(startDate);
    var eDate = (endDate instanceof Date) ? endDate : new Date(endDate);
    // Normalizar horas para incluir o dia inteiro
    sDate.setHours(0,0,0,0);
    eDate.setHours(23,59,59,999);

    var ultimaLinha = aba.getLastRow();
    var ultimaColuna = aba.getLastColumn();
    if (ultimaLinha < 2) return [];

    var dados = aba.getRange(2,1,ultimaLinha-1, Math.max(13, ultimaColuna)).getValues();

    var counts = {};
    dados.forEach(function(row){
      // Coluna UG Origem é a 8ª na ordem usada (índice 7)
      var ug = row[7] ? row[7].toString() : '';
      var entradaRaw = row[3];

      var entradaDate = null;
      if (entradaRaw instanceof Date) entradaDate = entradaRaw;
      else if (entradaRaw) entradaDate = new Date(entradaRaw);

      if (!entradaDate || isNaN(entradaDate.getTime())) return;

      // comparar no intervalo
      if (entradaDate >= sDate && entradaDate <= eDate) {
        counts[ug] = (counts[ug] || 0) + 1;
      }
    });

    // Transformar em array ordenado por count desc
    var result = Object.keys(counts).map(function(k){ return { ugOrigem: k, count: counts[k] }; });
    result.sort(function(a,b){ return b.count - a.count; });
    return result;
  } catch (error) {
    Logger.log('Erro em getCountsByUGOrigem: ' + error.toString());
    return [];
  }
}
