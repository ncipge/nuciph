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
 * Obtém informações completas do usuário logado (nome e perfil)
 * @return {Object} Objeto com nome, email e perfil do usuário
 */
function getUsuarioLogado() {
  try {
    var email = Session.getActiveUser().getEmail();
    Logger.log('Email do usuário ativo: ' + email);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configUsuarios = ss.getSheetByName('Config Usuarios');
    
    if (!configUsuarios) {
      Logger.log('Aba "Config Usuarios" não encontrada. Retornando dados básicos do usuário.');
      return {
        nome: email,
        email: email,
        perfil: 'Usuário'
      };
    }
    
    // Buscar dados das colunas A, B e C
    var data = configUsuarios.getRange('A2:C').getValues();
    Logger.log('Dados da aba Config Usuarios carregados: ' + data.length + ' linhas');
    
    // Procurar o email na coluna A e retornar nome (coluna B) e perfil (coluna C)
    for (var i = 0; i < data.length; i++) {
      var emailConfig = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
      var nomeConfig = data[i][1] ? data[i][1].toString().trim() : '';
      var perfilConfig = data[i][2] ? data[i][2].toString().trim() : 'Usuário';
      
      if (emailConfig === email.toLowerCase()) {
        Logger.log('Usuário encontrado: ' + email + ' -> ' + nomeConfig + ' (' + perfilConfig + ')');
        return {
          nome: nomeConfig || email,
          email: email,
          perfil: perfilConfig
        };
      }
    }
    
    // Se não encontrar, retornar dados básicos
    Logger.log('Usuário não encontrado na configuração. Retornando dados básicos: ' + email);
    return {
      nome: email,
      email: email,
      perfil: 'Usuário'
    };
    
  } catch (error) {
    Logger.log('Erro ao obter usuário logado: ' + error.toString());
    return {
      nome: 'Usuário não identificado',
      email: '',
      perfil: 'Usuário'
    };
  }
}

/**
 * Verifica se o usuário logado é administrador
 * @return {boolean} True se for administrador, false caso contrário
 */
function isUserAdmin() {
  try {
    var userInfo = getUsuarioLogado();
    var isAdmin = userInfo.perfil && userInfo.perfil.toLowerCase() === 'administrador';
    Logger.log('Verificação de admin para ' + userInfo.email + ': ' + isAdmin);
    return isAdmin;
  } catch (error) {
    Logger.log('Erro ao verificar se usuário é admin: ' + error.toString());
    return false;
  }
}

/**
 * Obtém apenas o nome do usuário logado (para compatibilidade)
 * @return {string} Nome do usuário
 */
function getNomeUsuarioLogado() {
  try {
    var userInfo = getUsuarioLogado();
    return userInfo.nome;
  } catch (error) {
    Logger.log('Erro ao obter nome do usuário: ' + error.toString());
    return 'Usuário não identificado';
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
 * Busca TODOS os registros da aba "Dados" (para a página de registros completa)
 * @return {Array} Array de objetos representando todos os registros
 */
function getAllRecords() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Tentando obter a aba com o nome: Dados para leitura de TODOS os registros.');
    var aba = planilha.getSheetByName('Dados');

    if (aba == null) {
      Logger.log('A aba "Dados" não foi encontrada ao ler todos os registros. Retornando array vazio.');
      return [];
    }
    Logger.log('Aba "Dados" encontrada para leitura de todos os registros.');

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

    // Obtém TODOS os dados da planilha, exceto a primeira linha (cabeçalho)
    var dadosBrutos = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
    Logger.log('Dados brutos obtidos: ' + dadosBrutos.length + ' linhas (TODOS os registros)');

    // Processa os dados para garantir que as datas sejam formatadas corretamente
    var dadosProcessados = dadosBrutos.map(function(row, index) {
      // Log a cada 100 registros para acompanhar o progresso
      if (index % 100 === 0) {
        Logger.log('Processando registro ' + (index + 1) + ' de ' + dadosBrutos.length);
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
      
      return registro;
    });

    // Retornar em ordem reversa (mais recentes primeiro)
    var registrosReversed = dadosProcessados.reverse();
    Logger.log('Total de TODOS os registros retornados: ' + registrosReversed.length);

    return registrosReversed;

  } catch (error) {
    Logger.log('Erro ao buscar TODOS os registros: ' + error.toString());
    return [];
  }
}

/**
 * Exclui um registro da aba "Dados"
 * @param {number} rowIndex - Índice da linha na planilha (base 1) a ser excluída
 * @return {Object} Objeto com status da operação
 */
function deleteRecord(rowIndex) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName('Dados');
    
    if (!aba) {
      Logger.log('Aba "Dados" não encontrada para exclusão.');
      return {
        status: 'error',
        message: 'Aba "Dados" não encontrada.'
      };
    }
    
    var ultimaLinha = aba.getLastRow();
    
    // Verificar se o índice da linha é válido
    if (rowIndex < 2 || rowIndex > ultimaLinha) {
      Logger.log('Índice de linha inválido para exclusão: ' + rowIndex);
      return {
        status: 'error',
        message: 'Índice de linha inválido. Linha: ' + rowIndex
      };
    }
    
    // Excluir a linha
    aba.deleteRow(rowIndex);
    
    Logger.log('Registro excluído com sucesso da linha: ' + rowIndex);
    return {
      status: 'success',
      message: 'Registro excluído com sucesso!'
    };
    
  } catch (error) {
    Logger.log('Erro ao excluir registro: ' + error.toString());
    return {
      status: 'error',
      message: 'Erro ao excluir registro: ' + error.toString()
    };
  }
}

/**
 * Obtém todos os usuários da aba "Config Usuarios"
 * @return {Array} Array com todos os usuários
 */
function getAllUsers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configUsuarios = ss.getSheetByName('Config Usuarios');
    
    if (!configUsuarios) {
      Logger.log('Aba "Config Usuarios" não encontrada.');
      return [];
    }
    
    var data = configUsuarios.getRange('A2:C').getValues();
    var users = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) { // Se há email na linha
        users.push({
          email: data[i][0].toString().trim(),
          nome: data[i][1] ? data[i][1].toString().trim() : '',
          perfil: data[i][2] ? data[i][2].toString().trim() : 'Usuário'
        });
      }
    }
    
    Logger.log('Total de usuários encontrados: ' + users.length);
    return users;
    
  } catch (error) {
    Logger.log('Erro ao buscar usuários: ' + error.toString());
    return [];
  }
}

/**
 * Obtém dados de um usuário específico por linha
 * @param {number} rowIndex - Índice da linha (base 1)
 * @return {Object} Dados do usuário
 */
function getUserByRow(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configUsuarios = ss.getSheetByName('Config Usuarios');
    
    if (!configUsuarios) {
      throw new Error('Aba "Config Usuarios" não encontrada.');
    }
    
    var data = configUsuarios.getRange(rowIndex, 1, 1, 3).getValues()[0];
    
    return {
      email: data[0] ? data[0].toString().trim() : '',
      nome: data[1] ? data[1].toString().trim() : '',
      perfil: data[2] ? data[2].toString().trim() : 'Usuário'
    };
    
  } catch (error) {
    Logger.log('Erro ao buscar usuário por linha: ' + error.toString());
    throw error;
  }
}

/**
 * Adiciona um novo usuário
 * @param {Object} userData - Dados do usuário
 * @return {Object} Resultado da operação
 */
function addUser(userData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configUsuarios = ss.getSheetByName('Config Usuarios');
    
    if (!configUsuarios) {
      return {
        status: 'error',
        message: 'Aba "Config Usuarios" não encontrada.'
      };
    }
    
    // Verificar se o email já existe
    var existingUsers = getAllUsers();
    for (var i = 0; i < existingUsers.length; i++) {
      if (existingUsers[i].email.toLowerCase() === userData.email.toLowerCase()) {
        return {
          status: 'error',
          message: 'Este email já está cadastrado no sistema.'
        };
      }
    }
    
    // Adicionar nova linha
    var nextRow = configUsuarios.getLastRow() + 1;
    configUsuarios.getRange(nextRow, 1, 1, 3).setValues([[
      userData.email,
      userData.nome,
      userData.perfil
    ]]);
    
    Logger.log('Usuário adicionado: ' + userData.email);
    return {
      status: 'success',
      message: 'Usuário adicionado com sucesso!'
    };
    
  } catch (error) {
    Logger.log('Erro ao adicionar usuário: ' + error.toString());
    return {
      status: 'error',
      message: 'Erro ao adicionar usuário: ' + error.toString()
    };
  }
}

/**
 * Atualiza dados de um usuário
 * @param {number} rowIndex - Índice da linha (base 1)
 * @param {Object} userData - Novos dados do usuário
 * @return {Object} Resultado da operação
 */
function updateUser(rowIndex, userData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configUsuarios = ss.getSheetByName('Config Usuarios');
    
    if (!configUsuarios) {
      return {
        status: 'error',
        message: 'Aba "Config Usuarios" não encontrada.'
      };
    }
    
    // Verificar se o email já existe em outra linha
    var existingUsers = getAllUsers();
    for (var i = 0; i < existingUsers.length; i++) {
      var currentRowIndex = i + 2; // +2 porque começamos da linha 2
      if (currentRowIndex !== rowIndex && 
          existingUsers[i].email.toLowerCase() === userData.email.toLowerCase()) {
        return {
          status: 'error',
          message: 'Este email já está cadastrado em outro usuário.'
        };
      }
    }
    
    // Atualizar dados
    configUsuarios.getRange(rowIndex, 1, 1, 3).setValues([[
      userData.email,
      userData.nome,
      userData.perfil
    ]]);
    
    Logger.log('Usuário atualizado: ' + userData.email);
    return {
      status: 'success',
      message: 'Usuário atualizado com sucesso!'
    };
    
  } catch (error) {
    Logger.log('Erro ao atualizar usuário: ' + error.toString());
    return {
      status: 'error',
      message: 'Erro ao atualizar usuário: ' + error.toString()
    };
  }
}

/**
 * Exclui um usuário
 * @param {number} rowIndex - Índice da linha (base 1)
 * @return {Object} Resultado da operação
 */
function deleteUser(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configUsuarios = ss.getSheetByName('Config Usuarios');
    
    if (!configUsuarios) {
      return {
        status: 'error',
        message: 'Aba "Config Usuarios" não encontrada.'
      };
    }
    
    // Verificar se não é o último administrador
    var userData = getUserByRow(rowIndex);
    if (userData.perfil === 'Administrador') {
      var allUsers = getAllUsers();
      var adminCount = allUsers.filter(user => user.perfil === 'Administrador').length;
      
      if (adminCount <= 1) {
        return {
          status: 'error',
          message: 'Não é possível excluir o último administrador do sistema.'
        };
      }
    }
    
    // Excluir linha
    configUsuarios.deleteRow(rowIndex);
    
    Logger.log('Usuário excluído da linha: ' + rowIndex);
    return {
      status: 'success',
      message: 'Usuário excluído com sucesso!'
    };
    
  } catch (error) {
    Logger.log('Erro ao excluir usuário: ' + error.toString());
    return {
      status: 'error',
      message: 'Erro ao excluir usuário: ' + error.toString()
    };
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
  
  var dia = date.getDate().toString().padStart(2, '0');
  var mes = (date.getMonth() + 1).toString().padStart(2, '0');
  var ano = date.getFullYear();
  return dia + '/' + mes + '/' + ano;
}
