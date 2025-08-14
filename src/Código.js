/**
 * @fileoverview Este script demonstra operações básicas de leitura e escrita
 * em uma planilha Google usando Google Apps Script, e agora inclui
 * funcionalidade para servir uma página web para inserção e leitura de dados
 * com base na estrutura de colunas fornecida.
 * Esta versão inclui suporte para paginação, filtragem dos dados, o novo campo "Sistema",
 * uma função para ler todos os dados para o dashboard, e foi atualizada para o campo "Assimetria".
 * Além disso, foi adicionado o campo "Sub Assunto" e uma função para carregar o mapeamento
 * de Assunto/Sub Assunto, e agora inclui a funcionalidade de edição de registros.
 * A ordem dos campos na gravação, leitura e atualização foi ajustada para refletir
 * a ordem final da planilha: Sub Assunto, ACI Responsável, Destino, Saída.
 * Os campos "Resp." e "Ato Responsável" foram removidos da estrutura principal.
 * Adicionado controle de acesso baseado em e-mail do usuário e exibição do e-mail na interface.
 */

/**
 * Função especial que é executada automaticamente quando a planilha é aberta.
 * Ela cria um menu personalizado para o script e para o web app.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi(); // Obtém a interface do usuário da planilha ativa
  ui.createMenu('Manipulação de Dados') // Cria um novo menu chamado 'Manipulação de Dados'
      .addItem('Abrir Formulário Web', 'abrirFormularioWeb') // Adiciona item para abrir o web app
      .addToUi(); // Adiciona o menu à interface do usuário da planilha
}

/**
 * Função para abrir o web app.
 * Utiliza HtmlService para exibir o index.html como um diálogo na planilha.
 */
function abrirFormularioWeb() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(1000) // Aumenta a largura para acomodar mais campos
      .setHeight(800); // Aumenta a altura para acomodar mais campos
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gerenciador de Dados da Planilha'); // Mostra o diálogo
}

/**
 * Função especial que é executada quando o script é acessado como um web app.
 * Ela verifica as permissões do usuário antes de servir o index.html.
 * @return {HtmlOutput} O objeto HtmlOutput contendo o conteúdo do web app ou uma mensagem de acesso negado.
 */
function doGet() {
  var userEmail = Session.getActiveUser().getEmail(); // Obtém o e-mail do usuário ativo
  Logger.log('Tentativa de acesso de: ' + userEmail);

  var authorizedUsers = getAuthorizedUsers(); // Obtém a lista de usuários autorizados

  // Verifica se o e-mail do usuário está na lista de autorizados
  if (authorizedUsers.has(userEmail)) {
    Logger.log('Acesso concedido para: ' + userEmail);
    // Serve o index.html e passa o e-mail do usuário como um objeto de dados
    var template = HtmlService.createTemplateFromFile('index');
    template.userEmail = userEmail; // Passa o e-mail do usuário para o template
    return template.evaluate()
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Permite que o web app seja incorporado em um iframe
  } else {
    Logger.log('Acesso negado para: ' + userEmail);
    // Retorna uma página HTML simples com a mensagem de acesso negado
    return HtmlService.createHtmlOutput('<h1 style="color: #ef4444; text-align: center; font-family: \'Inter\', sans-serif; margin-top: 50px;">Acesso Negado</h1><p style="text-align: center; font-family: \'Inter\', sans-serif;">Seu e-mail: <strong>' + userEmail + '</strong> não está autorizado a acessar esta aplicação. Por favor, contate o administrador da planilha.</p>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Lê a lista de e-mails de usuários autorizados da aba 'Config Usuarios'.
 * @return {Set<string>} Um conjunto contendo os e-mails dos usuários autorizados (em minúsculas).
 */
function getAuthorizedUsers() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Tentando obter a aba com o nome: Config Usuarios.');
  var abaConfig = planilha.getSheetByName('Config Usuarios');

  var authorizedEmails = new Set();

  if (abaConfig == null) {
    Logger.log('A aba "Config Usuarios" não foi encontrada. Nenhum usuário será autorizado.');
    return authorizedEmails; // Retorna um conjunto vazio se a aba não existir
  }
  Logger.log('Aba "Config Usuarios" encontrada.');

  var lastRow = abaConfig.getLastRow();
  if (lastRow < 1) {
    Logger.log('A aba "Config Usuarios" está vazia.');
    return authorizedEmails;
  }

  // Assume que os e-mails estão na primeira coluna (Coluna A)
  var emailsRange = abaConfig.getRange(1, 1, lastRow, 1).getValues();

  emailsRange.forEach(function(row) {
    var email = row[0].toString().trim();
    if (email) { // Adiciona apenas e-mails não vazios
      authorizedEmails.add(email.toLowerCase()); // Armazena em minúsculas para comparação sem distinção de maiúsculas e minúsculas
    }
  });

  Logger.log('Usuários autorizados carregados: ' + Array.from(authorizedEmails).join(', '));
  return authorizedEmails;
}

/**
 * Retorna o e-mail do usuário logado para o cliente (HTML).
 * @return {string} O e-mail do usuário ativo.
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}


/**
 * Grava um novo registro de dados na aba 'Dados'.
 * Os dados são recebidos como um objeto do lado do cliente com os novos campos.
 * @param {Object} dados Um objeto contendo os dados a serem gravados.
 */
function gravarDadosWeb(dados) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Tentando obter a aba com o nome: Dados para gravação.');
    var aba = planilha.getSheetByName('DadosTeste');

    if (aba == null) {
      Logger.log('A aba "Dados" não foi encontrada. Criando uma nova aba com esse nome.');
      // Se a aba não existe, cria uma nova
      aba = planilha.insertSheet('DadosTeste');
      // Adiciona cabeçalhos se for uma nova aba (IMPORTANTE: Mantenha a ordem dos campos aqui)
      // Ordem final dos campos
      aba.appendRow([
        'Status', 'Sistema', 'Num. PAEJ S.A.J.', 'Interessado', 'Entrada', 'Situacao',
        'Assimetria', 'Observação', 'UG Origem', 'Assunto', 'Sub Assunto',
        'ACI Responsável', 'Destino', 'Saída'
      ]);
      Logger.log('Aba "Dados" criada com cabeçalhos.');
    } else {
      Logger.log('Aba "Dados" encontrada para gravação.');
    }

    // Cria a nova linha com os dados na ordem correta, correspondendo aos cabeçalhos
    // Converte strings de data (YYYY-MM-DD do HTML input type="date") para objetos Date do Apps Script
    var novaLinha = [
      dados.status,
      dados.sistema,
      dados.numPaejSaj,
      dados.interessado,
      dados.entrada ? new Date(dados.entrada + 'T00:00:00') : '',
      dados.situacao,
      dados.assimetria,
      dados.observacao,
      dados.ugOrigem,
      dados.assunto,
      dados.subAssunto,
      dados.aciResponsavel,
      dados.destino,
      dados.saida ? new Date(dados.saida + 'T00:00:00') : ''
    ];
    aba.appendRow(novaLinha);

    Logger.log('Dados gravados com sucesso: ' + JSON.stringify(dados));
    return { sucesso: true, mensagem: 'Dados gravados com sucesso!' };
  } catch (e) {
    Logger.log('Erro ao gravar dados: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro ao gravar dados: ' + e.toString() };
  }
}

/**
 * Atualiza um registro existente na aba 'Dados'.
 * Os dados são recebidos como um objeto do lado do cliente com os campos atualizados.
 * O rowIndex é o número da linha na planilha (base 1) a ser atualizada.
 * @param {number} rowIndex O índice da linha na planilha (base 1) a ser atualizada.
 * @param {Object} dados Um objeto contendo os dados a serem atualizados.
 * @return {Object} Um objeto indicando sucesso ou falha da operação.
 */
function atualizarDadosWeb(rowIndex, dados) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Tentando obter a aba com o nome: Dados para atualização.');
    var aba = planilha.getSheetByName('DadosTeste');

    if (aba == null) {
      Logger.log('A aba "Dados" não foi encontrada para atualização.');
      return { sucesso: false, mensagem: 'Erro: Aba "Dados" não encontrada.' };
    }
    Logger.log('Aba "Dados" encontrada para atualização.');

    // Mapeamento dos nomes dos campos do formulário para os índices das colunas na planilha
    // (A ordem deve corresponder à ordem dos cabeçalhos na planilha)
    // ÍNDICES AQUI SÃO BASE 1 (coluna A é 1, B é 2, etc.) para setValues
    var columnIndexes = {
      status: 1,
      sistema: 2,
      numPaejSaj: 3,
      interessado: 4,
      entrada: 5,
      situacao: 6,
      assimetria: 7,
      observacao: 8,
      ugOrigem: 9,
      assunto: 10,
      subAssunto: 11,
      aciResponsavel: 12,
      destino: 13,
      saida: 14
    };

    // Obter a linha existente para garantir que não sobrescrevemos dados indesejados
    var numColunasTotais = 14; // Agora são 14 colunas no total
    var existingRow = aba.getRange(rowIndex, 1, 1, numColunasTotais).getValues()[0];

    // Iterar sobre os dados recebidos e atualizar a linha existente
    for (var fieldName in dados) {
      if (dados.hasOwnProperty(fieldName) && columnIndexes.hasOwnProperty(fieldName)) {
        var colIndex = columnIndexes[fieldName] - 1; // Ajusta para índice base 0 do array
        var valueToUpdate = dados[fieldName];

        // Tratamento especial para campos de data
        if (fieldName === 'entrada' || fieldName === 'saida') {
          if (valueToUpdate) {
            valueToUpdate = new Date(valueToUpdate + 'T00:00:00'); // Garante que seja um objeto Date
          } else {
            valueToUpdate = ''; // Limpa se o valor for vazio
          }
        }
        existingRow[colIndex] = valueToUpdate;
      }
    }

    // Escrever a linha atualizada de volta na planilha
    aba.getRange(rowIndex, 1, 1, existingRow.length).setValues([existingRow]);

    Logger.log('Dados atualizados com sucesso na linha ' + rowIndex + ': ' + JSON.stringify(dados));
    return { sucesso: true, mensagem: 'Registro atualizado com sucesso!' };
  } catch (e) {
    Logger.log('Erro ao atualizar dados na linha ' + rowIndex + ': ' + e.toString());
    return { sucesso: false, mensagem: 'Erro ao atualizar dados: ' + e.toString() };
  }
}


/**
 * Lê os dados da aba 'Dados' com suporte a paginação e filtragem.
 * Retorna um subconjunto de dados com base no número da página e tamanho da página,
 * juntamente com o número total de registros filtrados.
 *
 * @param {number} pageNumber O número da página a ser lida (base 1).
 * @param {number} pageSize O número de registros por página.
 * @param {Object} filters Um objeto contendo os filtros a serem aplicados (ex: {interessado: "Luciano"}).
 * @return {Object} Um objeto contendo 'data' (array de arrays) e 'totalRecords' (número total de registros filtrados).
 */
function lerTodosOsDadosWeb(pageNumber, pageSize, filters) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Tentando obter a aba com o nome: Dados para leitura paginada e filtrada.');
    var aba = planilha.getSheetByName('DadosTeste');

    if (aba == null) {
      Logger.log('A aba "Dados" não foi encontrada ao ler dados. Retornando array vazio e 0 registros.');
      return { data: [], totalRecords: 0 };
    }
    Logger.log('Aba "Dados" encontrada para leitura paginada e filtrada.');

    var ultimaLinha = aba.getLastRow();
    var ultimaColuna = aba.getLastColumn();
    Logger.log('Última linha da aba "Dados": ' + ultimaLinha + ', Última coluna: ' + ultimaColuna);

    // Ajustar o número de colunas para 14 (assumindo que a planilha tem 14 colunas agora)
    var numColunasEsperadas = 14;
    if (ultimaColuna < numColunasEsperadas) {
        Logger.log('Atenção: A planilha tem menos colunas do que o esperado. Lendo até a última coluna disponível.');
        ultimaColuna = numColunasEsperadas; // Garante que o range lido tenha o número correto de colunas
    }

    if (ultimaLinha < 2 || ultimaColuna === 0) { // Se só tem cabeçalho ou está vazia ou não há colunas
      Logger.log('Planilha "Dados" está vazia, contém apenas cabeçalho ou não há colunas. Retornando array vazio e 0 registros.');
      return { data: [], totalRecords: 0 };
    }

    // Obtém todos os valores da planilha, exceto a primeira linha (cabeçalho)
    var dadosBrutos = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();

    // Mapeamento dos nomes dos campos do filtro para os índices das colunas na planilha
    // (A ordem deve corresponder à ordem dos cabeçalhos na planilha)
    // ÍNDICES AQUI SÃO BASE 0 (para acesso ao array de dadosBrutos)
    var columnIndexes = {
      status: 0,
      sistema: 1,
      numPaejSaj: 2,
      interessado: 3,
      entrada: 4,
      situacao: 5,
      assimetria: 6,
      observacao: 7,
      ugOrigem: 8,
      assunto: 9,
      subAssunto: 10,
      aciResponsavel: 11,
      destino: 12,
      saida: 13
    };

    // Aplica os filtros aos dados lidos
    var dadosFiltrados = dadosBrutos.filter(function(row) {
      var match = true;
      for (var filterKey in filters) {
        if (filters.hasOwnProperty(filterKey)) {
          var filterValue = filters[filterKey].toString().toLowerCase(); // Valor do filtro em minúsculas
          var columnIndex = columnIndexes[filterKey];

          // Verifica se o columnIndex é válido e se a célula existe na linha
          if (columnIndex !== undefined && row.length > columnIndex && row[columnIndex] !== undefined && row[columnIndex] !== null) {
            var cellValue = row[columnIndex].toString().toLowerCase(); // Valor da célula em minúsculas

            // Verifica se o valor da célula contém o valor do filtro
            if (cellValue.indexOf(filterValue) === -1) {
              match = false; // Não houve correspondência para este filtro
              break; // Sai do loop de filtros
            }
          } else if (filterValue !== '') { // Se o filtro tem valor, mas a coluna/célula não existe, não há correspondência
            match = false;
            break;
          }
        }
      }
      return match;
    });

    Logger.log('Número total de registros antes da paginação (filtrados): ' + dadosFiltrados.length);
    var totalRecordsFiltrados = dadosFiltrados.length;

    // Aplica a paginação aos dados filtrados
    var startIndex = (pageNumber - 1) * pageSize;
    var endIndex = startIndex + pageSize;
    var dadosPaginados = dadosFiltrados.slice(startIndex, endIndex);

    // Processa os dados paginados para garantir que as datas sejam strings ISO 8601
    var dadosProcessados = dadosPaginados.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) {
          // Adiciona 1 dia e formata como dd/mm/aaaa
          var novaData = new Date(cell.getTime());
          novaData.setDate(novaData.getDate() + 1);
          var dia = novaData.getDate().toString().padStart(2, '0');
          var mes = (novaData.getMonth() + 1).toString().padStart(2, '0');
          var ano = novaData.getFullYear();
          return dia + '/' + mes + '/' + ano;
        }
        return cell;
      });
    });

    Logger.log('Dados lidos, filtrados e processados com sucesso. Número de linhas paginadas: ' + dadosProcessados.length);

    return { data: dadosProcessados, totalRecords: totalRecordsFiltrados }; // Retorna os dados processados e o total de registros filtrados
  } catch (e) {
    Logger.log('Erro ao ler e filtrar dados: ' + e.toString());
    return { data: [], totalRecords: 0 };
  }
}

/**
 * Lê todos os dados da aba 'Dados' sem paginação ou filtros.
 * Esta função é usada especificamente para popular o dashboard.
 * @return {Object} Um objeto contendo 'data' (array de arrays) e 'totalRecords' (número total de registros).
 */
function lerTodosOsDadosWebSemPaginacao() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Tentando obter a aba com o nome: Dados para dashboard (sem paginação/filtros).');
    var aba = planilha.getSheetByName('DadosTeste');

    if (aba == null) {
      Logger.log('A aba "Dados" não foi encontrada. Retornando 0 para todos os totais do dashboard.');
      return {
        totalRecords: 0,
        statusAnalisado: 0,
        recentRecords: 0
      };
    }
    Logger.log('Aba "Dados" encontrada para dashboard.');

    var ultimaLinha = aba.getLastRow();
    var ultimaColuna = aba.getLastColumn();
    var numColunasEsperadas = 14;
    if (ultimaColuna < numColunasEsperadas) {
        Logger.log('Atenção: A planilha tem menos colunas do que o esperado. Lendo até a última coluna disponível.');
        ultimaColuna = numColunasEsperadas;
    }
    if (ultimaLinha < 2 || ultimaColuna === 0) {
      Logger.log('Planilha "Dados" está vazia ou contém apenas cabeçalho ou não há colunas. Retornando 0 para todos os totais do dashboard.');
      return {
        totalRecords: 0,
        statusAnalisado: 0,
        recentRecords: 0
      };
    }

    // Obtém todos os valores da planilha, exceto a primeira linha (cabeçalho)
    var dadosBrutos = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
    var totalRecords = dadosBrutos.length;

    // Contar quantos registros têm status "Analisado"
    var statusColIndex = 0; // Coluna A = Status
    var statusAnalisado = dadosBrutos.filter(function(row) {
      return row[statusColIndex] && row[statusColIndex].toString().toLowerCase() === 'analisado';
    }).length;

    // Contar quantos registros têm data de entrada nos últimos 30 dias
    var entradaColIndex = 4; // Coluna E = Entrada
    var hoje = new Date();
    var trintaDiasAtras = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
    var recentRecords = dadosBrutos.filter(function(row) {
      var dataEntrada = row[entradaColIndex];
      if (dataEntrada instanceof Date) {
        return dataEntrada >= trintaDiasAtras && dataEntrada <= hoje;
      }
      return false;
    }).length;

    Logger.log('Dashboard: totalRecords=' + totalRecords + ', statusAnalisado=' + statusAnalisado + ', recentRecords=' + recentRecords);
    return {
      totalRecords: totalRecords,
      statusAnalisado: statusAnalisado,
      recentRecords: recentRecords
    };
  } catch (e) {
    Logger.log('Erro ao ler dados para dashboard: ' + e.toString());
    return {
      totalRecords: 0,
      statusAnalisado: 0,
      recentRecords: 0
    };
  }
}

/**
 * Lê o mapeamento de Assuntos e Sub Assuntos da aba 'Config Assuntos'.
 * A estrutura esperada na planilha é: Coluna A = Assunto, Coluna B = Sub Assunto.
 * Se um Assunto tiver múltiplos Sub Assuntos, eles devem ser listados em linhas separadas.
 * Exemplo na planilha:
 * Assunto | Sub Assunto
 * --------|------------
 * Contrato| Inexigibilidade
 * Contrato| Pagamento
 * Contrato| Termo Aditivo
 * Suprimento| Aprovação
 * Suprimento| Baixa
 *
 * @return {Object} Um objeto onde as chaves são Assuntos e os valores são arrays de Sub Assuntos.
 * Ex: { "Contrato": ["Inexigibilidade", "Pagamento"], "Pagamento": ["Aprovação"] }
 */
function getAssuntoSubAssuntoMap() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaConfig = planilha.getSheetByName('Config Assuntos'); // Nome da aba para configuração

  var assuntoSubAssuntoMap = {};

  if (abaConfig == null) {
    Logger.log('A aba "Config Assuntos" não foi encontrada. Retornando mapa de assuntos vazio.');
    return assuntoSubAssuntoMap;
  }

  // Obter todos os dados da aba de configuração, assumindo que Assunto está na coluna A e Sub Assunto na B
  var range = abaConfig.getDataRange();
  var values = range.getValues();

  // Ignorar a linha do cabeçalho se existir (assumindo que a primeira linha é o cabeçalho)
  if (values.length > 0 && values[0][0] && values[0][0].toString().trim().toLowerCase() === 'assunto' && values[0][1] && values[0][1].toString().trim().toLowerCase() === 'sub assunto') {
    values.shift(); // Remove a primeira linha (cabeçalho)
  }

  values.forEach(function(row) {
    var assunto = row[0] ? row[0].toString().trim() : '';
    var subAssunto = row[1] ? row[1].toString().trim() : '';

    if (assunto) { // Apenas processa se houver um assunto
      if (!assuntoSubAssuntoMap[assunto]) {
        assuntoSubAssuntoMap[assunto] = [];
      }
      if (subAssunto) { // Adiciona sub assunto apenas se não for vazio
        assuntoSubAssuntoMap[assunto].push(subAssunto);
      }
    }
  });

  // Opcional: Para ter certeza de que cada sub assunto aparece apenas uma vez por assunto
  for (var key in assuntoSubAssuntoMap) {
    if (assuntoSubAssuntoMap.hasOwnProperty(key)) {
      assuntoSubAssuntoMap[key] = Array.from(new Set(assuntoSubAssuntoMap[key]));
    }
    // Se não há sub-assuntos definidos explicitamente para um Assunto, mas o Assunto existe,
    // garanta que ele ainda seja uma opção válida (com um array vazio de sub-assuntos)
    if (assuntoSubAssuntoMap[key].length === 0 && !values.some(row => row[0].toString().trim() === key && row[1].toString().trim() === '')) {
      // Remove o assunto se ele não tiver subassuntos explícitos E não for uma entrada de assunto sozinho
      delete assuntoSubAssuntoMap[key];
    }
  }

  // Ordenar os assuntos para que sejam exibidos em ordem alfabética no dropdown
  var sortedAssuntos = Object.keys(assuntoSubAssuntoMap).sort();
  var sortedAssuntoSubAssuntoMap = {};
  sortedAssuntos.forEach(function(assunto) {
    sortedAssuntoSubAssuntoMap[assunto] = assuntoSubAssuntoMap[assunto];
  });
  assuntoSubAssuntoMap = sortedAssuntoSubAssuntoMap;


  Logger.log('Mapeamento de Assunto/Sub Assunto carregado: ' + JSON.stringify(assuntoSubAssuntoMap));
  return assuntoSubAssuntoMap;
}

/**
 * Deleta um registro da aba 'Dados' pelo índice da linha (base 1).
 * @param {number} rowIndex O índice da linha na planilha (base 1) a ser deletada.
 * @return {Object} Um objeto indicando sucesso ou falha da operação.
 */
function deletarRegistroWeb(rowIndex) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName('DadosTeste');
    if (!aba) {
      return { sucesso: false, mensagem: 'Aba "Dados" não encontrada.' };
    }
    var ultimaLinha = aba.getLastRow();
    if (rowIndex < 2 || rowIndex > ultimaLinha) {
      // Linha 1 é o cabeçalho, não pode ser deletada
      return { sucesso: false, mensagem: 'Índice de linha inválido.' };
    }
    aba.deleteRow(rowIndex);
    return { sucesso: true, mensagem: 'Registro deletado com sucesso!' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao deletar registro: ' + e.toString() };
  }
}
