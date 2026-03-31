//***************************************************************************************************
// Função para processar o status das transportadoras em três planilhas diferentes
function processTransportadoraStatus() {
  // ID da planilha que contém os dados
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU';
  
  // Nomes das planilhas (sheets) a serem acessadas
  var baseEmailSheetName = 'Base_e-mail';
  var enviarParaSheetName = 'EnviarPara';
  var oQueEnviarSheetName = 'o que enviar';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa as sheets especificadas
  var baseEmailSheet = spreadsheet.getSheetByName(baseEmailSheetName);
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  var oQueEnviarSheet = spreadsheet.getSheetByName(oQueEnviarSheetName);
  
  // Verifica se todas as sheets foram encontradas, lança erro se não
  if (!baseEmailSheet) throw new Error('Sheet "' + baseEmailSheetName + '" não encontrada.');
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  if (!oQueEnviarSheet) throw new Error('Sheet "' + oQueEnviarSheetName + '" não encontrada.');
  
  // Obtém os dados das sheets
  var baseEmailData = baseEmailSheet.getDataRange().getValues();
  var oQueEnviarData = oQueEnviarSheet.getDataRange().getValues();
  
  var colunaIIndex = 8;  // Índice da Coluna I
  var colunaKIndex = 10; // Índice da Coluna K
  
  // Coleção para armazenar valores únicos da Coluna A da sheet "o que enviar"
  var oQueEnviarSet = new Set();
  for (var i = 1; i < oQueEnviarData.length; i++) {
    oQueEnviarSet.add(oQueEnviarData[i][0]);
  }
  
  // Conjunto para armazenar combinações únicas de valores das Colunas I e K
  var uniqueCombinations = new Set();
  
  // Coleta combinações únicas das Colunas I e K
  for (var i = 1; i < baseEmailData.length; i++) {
    var colunaIValue = baseEmailData[i][colunaIIndex];
    var colunaKValue = baseEmailData[i][colunaKIndex];
    if (oQueEnviarSet.has(colunaKValue)) {
      var combination = colunaIValue + '|' + colunaKValue;
      uniqueCombinations.add(combination);
    }
  }
  
  // Limpa a sheet "EnviarPara", exceto a linha de cabeçalho
  var lastRow = enviarParaSheet.getLastRow();
  if (lastRow > 1) {
    enviarParaSheet.getRange(2, 1, lastRow - 1, enviarParaSheet.getLastColumn()).clearContent();
  }
  
  // Cola as combinações únicas encontradas na sheet "EnviarPara"
  var rowIndex = 2; // A partir da linha 2
  uniqueCombinations.forEach(function(combination) {
    var parts = combination.split('|');
    enviarParaSheet.getRange(rowIndex, 1).setValue(parts[0]); // Coloca na Coluna A
    enviarParaSheet.getRange(rowIndex, 2).setValue(parts[1]); // Coloca na Coluna B
    rowIndex++;
  });
  
  // Ordena os dados na sheet "EnviarPara" pela Coluna A
  lastRow = enviarParaSheet.getLastRow();
  if (lastRow > 1) {
    enviarParaSheet.getRange(2, 1, lastRow - 1, 2).sort(1);
  }
  
  // Chama a função para atualizar os dados na sheet "EnviarPara"
  updateEnviarParaWithCadastroEmail();
}

//***************************************************************************************************
// Atualiza a planilha "EnviarPara" com informações da planilha "Cadastro_e-mail"
function updateEnviarParaWithCadastroEmail() {
  // ID da planilha que contém os dados
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU';
  
  // Nomes das sheets a serem acessadas
  var enviarParaSheetName = 'EnviarPara';
  var cadastroEmailSheetName = 'Cadastro_e-mail';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa as sheets especificadas
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  var cadastroEmailSheet = spreadsheet.getSheetByName(cadastroEmailSheetName);
  
  // Verifica se todas as sheets foram encontradas, lança erro se não
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  if (!cadastroEmailSheet) throw new Error('Sheet "' + cadastroEmailSheetName + '" não encontrada.');
  
  // Obtém os dados das sheets
  var enviarParaData = enviarParaSheet.getDataRange().getValues();
  var cadastroEmailData = cadastroEmailSheet.getDataRange().getValues();
  
  // Mapeia os dados da sheet Cadastro_e-mail para facilitar a busca
  var cadastroEmailMap = new Map();
  for (var i = 1; i < cadastroEmailData.length; i++) {
    var key = cadastroEmailData[i][0]; // Valor da Coluna A
    var valueB = cadastroEmailData[i][1]; // Valor da Coluna B
    var valueC = cadastroEmailData[i][2]; // Valor da Coluna C
    cadastroEmailMap.set(key, [valueB, valueC]);
  }
  
  // Conjunto para rastrear as chaves já processadas
  var processedKeys = new Set();
  
  // Atualiza a sheet "EnviarPara" com dados da sheet "Cadastro_e-mail"
  for (var j = 1; j < enviarParaData.length; j++) {
    var key = enviarParaData[j][0]; // Coluna A
    if (cadastroEmailMap.has(key) && !processedKeys.has(key)) {
      var values = cadastroEmailMap.get(key);
      enviarParaSheet.getRange(j + 1, 5).setValue(values[0]); // Atualiza Coluna E
      enviarParaSheet.getRange(j + 1, 6).setValue(values[1]); // Atualiza Coluna F
      processedKeys.add(key); // Marca a chave como processada
    }
  }
  
  // Mensagem de confirmação de atualização
  SpreadsheetApp.getUi().alert('Dados atualizados com sucesso na sheet "EnviarPara".');
}

//***************************************************************************************************
// Processa e copia dados entre diferentes sheets com base em condições específicas
function processAndCopyData() {
  // IDs das planilhas
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU'; // ID da planilha de dados
  
  // Nomes das sheets a serem acessadas
  var baseEmailSheetName = 'Base_e-mail';
  var enviarParaSheetName = 'EnviarPara';
  var pendenciaSheetName = 'Pendência';
  var oQueEnviarSheetName = 'o que enviar';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa as sheets especificadas
  var baseEmailSheet = spreadsheet.getSheetByName(baseEmailSheetName);
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  var pendenciaSheet = spreadsheet.getSheetByName(pendenciaSheetName);
  var oQueEnviarSheet = spreadsheet.getSheetByName(oQueEnviarSheetName);
  
  // Verifica se todas as sheets foram encontradas, lança erro se não
  if (!baseEmailSheet) throw new Error('Sheet "' + baseEmailSheetName + '" não encontrada.');
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  if (!pendenciaSheet) throw new Error('Sheet "' + pendenciaSheetName + '" não encontrada.');
  if (!oQueEnviarSheet) throw new Error('Sheet "' + oQueEnviarSheetName + '" não encontrada.');
  
  // Obtém os dados da sheet "EnviarPara"
  var enviarParaData = enviarParaSheet.getDataRange().getValues();
  var transportadoraIndexEnviarPara = 0; // Índice Coluna A
  var statusIndexEnviarPara = 1; // Índice Coluna B
  var checkIndexEnviarPara = 2; // Índice Coluna C
  
  // Verifica se há algum valor marcado na coluna C
  var hasValueInColumnC = enviarParaData.slice(1).some(row => row[checkIndexEnviarPara] !== '');
  if (!hasValueInColumnC) {
    SpreadsheetApp.getUi().alert('Nenhuma linha está marcada na coluna C. Corrija antes de continuar.');
    return;
  }
  
  // Verifica se os valores na Coluna A são consistentes quando relacionados à Coluna C
  var uniqueTransportadoras = new Set();
  for (var i = 1; i < enviarParaData.length; i++) {
    if (enviarParaData[i][checkIndexEnviarPara] !== '') {
      uniqueTransportadoras.add(enviarParaData[i][transportadoraIndexEnviarPara]);
    }
  }
  
  if (uniqueTransportadoras.size > 1) {
    SpreadsheetApp.getUi().alert('Você marcou mais do que uma Transportadora, corrija!');
    return;
  }
  
  // Limpa a sheet "Pendência" a partir da linha 2
  clearSheet(pendenciaSheet);
  
  // Obtém os dados da sheet "Base_e-mail"
  var baseEmailData = baseEmailSheet.getDataRange().getValues();
  var transportadoraIndexBaseEmail = 8; // Índice Coluna I
  var statusIndexBaseEmail = 10; // Índice Coluna K
  
  // Mapeia dados da sheet "o que enviar"
  var oQueEnviarData = oQueEnviarSheet.getDataRange().getValues();
  var oQueEnviarMap = new Map();
  for (var i = 1; i < oQueEnviarData.length; i++) {
    var key = oQueEnviarData[i][3]; // Coluna D
    var value = oQueEnviarData[i][4]; // Coluna E
    oQueEnviarMap.set(key, value);
  }
  
  // Array para armazenar dados que serão copiados
  var dataToCopy = [];
  
  // Processa cada linha da sheet "EnviarPara"
  for (var i = 1; i < enviarParaData.length; i++) {
    if (enviarParaData[i][checkIndexEnviarPara] !== '') {
      var transportadora = enviarParaData[i][transportadoraIndexEnviarPara];
      var status = enviarParaData[i][statusIndexEnviarPara];
      
      // Verifica cada linha da sheet "Base-e-mail"
      for (var j = 1; j < baseEmailData.length; j++) {
        if (baseEmailData[j][transportadoraIndexBaseEmail] === transportadora &&
            baseEmailData[j][statusIndexBaseEmail] === status) {
          // Copia dados das colunas A até N
          var rowData = baseEmailData[j].slice(0, 14);
          
          // Adiciona os dados ao array para copiar posteriormente
          dataToCopy.push(rowData);
        }
      }
    }
  }
  
  // Cola os dados copiados na sheet "Pendência" a partir da linha 2
  if (dataToCopy.length > 0) {
    pendenciaSheet.getRange(2, 1, dataToCopy.length, dataToCopy[0].length).setValues(dataToCopy);
  }
  
  // Atualiza informações na coluna K da sheet "Pendência"
  var pendenciaData = pendenciaSheet.getRange(2, 1, pendenciaSheet.getLastRow() - 1, pendenciaSheet.getLastColumn()).getValues();
  for (var i = 0; i < pendenciaData.length; i++) {
    var key = pendenciaData[i][10]; // Coluna K
    if (oQueEnviarMap.has(key)) {
      pendenciaData[i][10] = oQueEnviarMap.get(key);
    }
  }
  
  // Cola de volta os dados atualizados na sheet "Pendência"
  pendenciaSheet.getRange(2, 1, pendenciaData.length, pendenciaData[0].length).setValues(pendenciaData);
  
  // Mensagem de sucesso no processamento
  SpreadsheetApp.getUi().alert('Dados copiados e atualizados com sucesso para a sheet "Pendência".');
  
  // Marca a sheet como "BaseOK" para indicar que o processamento foi concluído
  markBaseOk();
}

// Função que limpa o conteúdo de uma sheet a partir da linha 2
function clearSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

//***************************************************************************************************
// Marca a planilha "EnviarPara" com o status "BaseOK" na célula G1
function markBaseOk() {
  // IDs das planilhas
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU'; // ID da planilha de dados
  
  // Nome da sheet a ser acessada
  var enviarParaSheetName = 'EnviarPara';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa a sheet especificada
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  
  // Verifica se a sheet foi encontrada, lança erro se não
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  
  // Define o valor "BaseOK" na célula G1 para sinalizar que a base está correta
  enviarParaSheet.getRange('G1').setValue('BaseOK');
}

//***************************************************************************************************
// Verifica se a célula G1 na planilha "EnviarPara" contém 'BaseOK' e a limpa se necessário
function verifyAndClearBaseOk() {
  // IDs das planilhas
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU'; // ID da planilha de dados
  
  // Nome da sheet a ser acessada
  var enviarParaSheetName = 'EnviarPara';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa a sheet especificada
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  
  // Verifica se a sheet foi encontrada, lança erro se não
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  
  // Obtém o valor da célula G1
  var cellValue = enviarParaSheet.getRange('G1').getValue();
  
  // Se o valor for "BaseOK", limpe a célula e permita a execução
  if (cellValue === 'BaseOK') {
    enviarParaSheet.getRange('G1').clearContent();
    return true;  // Continuação permitida
  } else {
    // Alerta o usuário caso a execução não deva continuar
    SpreadsheetApp.getUi().alert('Execute primeiro [1 - Base Pendência]');
    return false; // Impede a continuação
  }
}

//***************************************************************************************************
// Envia email com planilha anexada para o primeiro email definido, validando se a send pode prosseguir
function sendEmailForFirstNonEmptyEmail2() {
  if (!verifyAndClearBaseOk()) {
    return; // Interrompe a execução se a célula G1 não conter "BaseOK"
  }
  
  // IDs das planilhas
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU'; // ID da planilha de dados
  
  // Nomes das sheets a serem acessadas
  var enviarParaSheetName = 'EnviarPara';
  var pendenciaSheetName = 'Pendência';
  var m1SheetName = 'M1';
  var m2SheetName = 'M2';
  var m3SheetName = 'M3';
  var historicoSheetName = 'Histórico';
  var cadastroEmailSheetName = 'Cadastro_e-mail';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa as sheets especificadas
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  var pendenciaSheet = spreadsheet.getSheetByName(pendenciaSheetName);
  var m1Sheet = spreadsheet.getSheetByName(m1SheetName);
  var m2Sheet = spreadsheet.getSheetByName(m2SheetName);
  var m3Sheet = spreadsheet.getSheetByName(m3SheetName);
  var historicoSheet = spreadsheet.getSheetByName(historicoSheetName);
  var cadastroEmailSheet = spreadsheet.getSheetByName(cadastroEmailSheetName);
  
  // Verifica se todas as sheets foram encontradas, lança erro se não
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  if (!pendenciaSheet) throw new Error('Sheet "' + pendenciaSheetName + '" não encontrada.');
  if (!m1Sheet) throw new Error('Sheet "' + m1SheetName + '" não encontrada.');
  if (!m2Sheet) throw new Error('Sheet "' + m2SheetName + '" não encontrada.');
  if (!m3Sheet) throw new Error('Sheet "' + m3SheetName + '" não encontrada.');
  if (!historicoSheet) throw new Error('Sheet "' + historicoSheetName + '" não encontrada.');
  if (!cadastroEmailSheet) throw new Error('Sheet "' + cadastroEmailSheetName + '" não encontrada.');
  
  // Obtém os dados das sheets
  var enviarParaData = enviarParaSheet.getDataRange().getValues();
  var cadastroEmailData = cadastroEmailSheet.getDataRange().getValues();
  
  // Percorre a sheet "EnviarPara" da linha 2 em diante
  for (var i = 1; i < enviarParaData.length; i++) {
    var email = enviarParaData[i][4];  // Endereço de e-mail na Coluna E
    var ccEmails = enviarParaData[i][5];  // CC de e-mails na Coluna F
    var sheetIndicator = enviarParaData[i][3];  // Indicador de sheet na Coluna D
    var identifier = enviarParaData[i][0];  // Identificador na Coluna A
    var emailBody = ''; // Inicializa o corpo do e-mail
    
    // Verifica se a coluna C possui valor
    if (enviarParaData[i][2] !== '') { 
      
      // Caso o indicador de sheet esteja vazio, envia alerta
      if (sheetIndicator === '') {
        SpreadsheetApp.getUi().alert('Informe qual mensagem será enviada: M?');
        return;
      }
      
      // Caso o endereço de e-mail esteja vazio, envia alerta
      if (email === '') {
        SpreadsheetApp.getUi().alert('Necessário informar o e-mail.');
        return;
      }
      
      // Seleciona o corpo do e-mail com base no indicador de sheet
      if (sheetIndicator === 'M1') {
        emailBody = m1Sheet.getRange('A1').getValue();
      } else if (sheetIndicator === 'M2') {
        emailBody = m2Sheet.getRange('A1').getValue();
      } else if (sheetIndicator === 'M3') {
        emailBody = m3Sheet.getRange('A1').getValue();
      }
      
      // Busca o nome da transportadora a partir do identificador
      var transportadora = '';
      for (var j = 1; j < cadastroEmailData.length; j++) {
        if (cadastroEmailData[j][0] === identifier) {
          transportadora = cadastroEmailData[j][3];
          break;
        }
      }
      
      // Verifica se a transportadora foi encontrada
      if (transportadora === '') {
        SpreadsheetApp.getUi().alert('Transportadora não encontrada para o identificador: ' + identifier);
        return;
      }
      
      // Obtém a data atual
      var date = new Date();
      var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      
      // Define o assunto do e-mail
      var emailSubject = 'Colgate - Devolução (Pendências/Ações) - ' + transportadora + ' - ' + formattedDate;
      
      // Cria uma nova planilha temporária para enviar como anexo
      var newSpreadsheet = SpreadsheetApp.create('TempSpreadsheet');
      var newSpreadsheetId = newSpreadsheet.getId();
      var newSpreadsheetFile = DriveApp.getFileById(newSpreadsheetId);
      
      // Copia a sheet "Pendência" para a nova planilha
      pendenciaSheet.copyTo(newSpreadsheet).setName('Pendência');
      newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]); // Remove a sheet padrão
      
      // Converte a planilha para o formato XLSX
      var xlsxBlob = convertSpreadsheetToXlsx(newSpreadsheetId);
      xlsxBlob.setName('Pendencia.xlsx');
      
      // Envia o e-mail com o anexo
      MailApp.sendEmail({
        to: email,
        cc: ccEmails,
        subject: emailSubject,
        body: emailBody,
        attachments: [xlsxBlob]
      });
      
      // Deleta a planilha temporária
      newSpreadsheetFile.setTrashed(true);
      
      // Marca a linha como "enviado" na Coluna G
      enviarParaSheet.getRange(i + 1, 7).setValue('enviado');
      
      // Limpa as linhas correspondentes na coluna C com o mesmo valor na coluna A
      var valueA = enviarParaData[i][0];
      for (var j = 1; j < enviarParaData.length; j++) {
        if (enviarParaData[j][0] === valueA) {
          enviarParaSheet.getRange(j + 1, 3).clearContent();
        }
      }
      
      // Adiciona o envio ao histórico
      var formattedTime = Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss');
      var lastRowHistorico = historicoSheet.getLastRow() + 1;
      historicoSheet.getRange(lastRowHistorico, 1).setValue(formattedDate);
      historicoSheet.getRange(lastRowHistorico, 2).setValue(formattedTime);
      historicoSheet.getRange(lastRowHistorico, 3).setValue(valueA);
      
      // Sai do loop após enviar o e-mail
      break;
    }
  }
  
  // Mensagem de confirmação de envio
  SpreadsheetApp.getUi().alert('E-mail enviado com sucesso.');
}

// Função para converter uma planilha em blob do tipo XLSX
function convertSpreadsheetToXlsx(spreadsheetId) {
  var url = 'https://docs.google.com/feeds/download/spreadsheets/Export?key=' + spreadsheetId + '&exportFormat=xlsx';
  var params = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };
  var response = UrlFetchApp.fetch(url, params);
  return response.getBlob();
}

//***************************************************************************************************
// Função para enviar um e-mail para a primeira linha não vazia na coluna C após verificar condições
function sendEmailForFirstNonEmptyEmail3() {
  if (!verifyAndClearBaseOk()) {
    return; // Interrompe a execução se a célula G1 não conter "BaseOK"
  }
  
  // IDs das planilhas
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU'; // ID da planilha de dados
  
  // Nomes das sheets
  var enviarParaSheetName = 'EnviarPara';
  var pendenciaSheetName = 'Pendência';
  var m1SheetName = 'M1';
  var m2SheetName = 'M2';
  var m3SheetName = 'M3';
  var historicoSheetName = 'Histórico';
  var cadastroEmailSheetName = 'Cadastro_e-mail';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  var pendenciaSheet = spreadsheet.getSheetByName(pendenciaSheetName);
  var m1Sheet = spreadsheet.getSheetByName(m1SheetName);
  var m2Sheet = spreadsheet.getSheetByName(m2SheetName);
  var m3Sheet = spreadsheet.getSheetByName(m3SheetName);
  var historicoSheet = spreadsheet.getSheetByName(historicoSheetName);
  var cadastroEmailSheet = spreadsheet.getSheetByName(cadastroEmailSheetName);
  
  // Verifica se todas as sheets foram encontradas, lança erro se não
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  if (!pendenciaSheet) throw new Error('Sheet "' + pendenciaSheetName + '" não encontrada.');
  if (!m1Sheet) throw new Error('Sheet "' + m1SheetName + '" não encontrada.');
  if (!m2Sheet) throw new Error('Sheet "' + m2SheetName + '" não encontrada.');
  if (!m3Sheet) throw new Error('Sheet "' + m3SheetName + '" não encontrada.');
  if (!historicoSheet) throw new Error('Sheet "' + historicoSheetName + '" não encontrada.');
  if (!cadastroEmailSheet) throw new Error('Sheet "' + cadastroEmailSheetName + '" não encontrada.');
  
  // Obtém os dados das sheets
  var enviarParaData = enviarParaSheet.getDataRange().getValues();
  var cadastroEmailData = cadastroEmailSheet.getDataRange().getValues();
  
  // Percorre a sheet "EnviarPara" da linha 2 em diante
  for (var i = 1; i < enviarParaData.length; i++) {
    var email = enviarParaData[i][4];  // Endereço de e-mail na Coluna E
    var ccEmails = enviarParaData[i][5];  // CC de e-mails na Coluna F
    var sheetIndicator = enviarParaData[i][3];  // Indicador de sheet na Coluna D
    var identifier = enviarParaData[i][0];  // Identificador na Coluna A
    var emailBody = ''; // Inicializa o corpo do e-mail
    
    // Verifica se a coluna C possui valor
    if (enviarParaData[i][2] !== '') {
      if (sheetIndicator === '') {
        SpreadsheetApp.getUi().alert('Informe qual mensagem será enviada: M?');
        return;
      }
      
      // Armazena os conteúdos da coluna B das linhas onde a coluna C não está vazia
      var conteudoColunaB = enviarParaData
                            .slice(1)  // Ignora a primeira linha (cabeçalho)
                            .filter(row => row[2] !== '')  // Filtra as linhas onde a coluna C não está vazia
                            .map(row => row[1])  // Mapeia apenas os valores da coluna B dessas linhas
                            .join('; ');  // Concatena os valores com "; "
      console.log(conteudoColunaB);
      
      if (email === '') {
        SpreadsheetApp.getUi().alert('Necessário informar o e-mail.');
        return;
      }
      
      if (sheetIndicator === 'M1') {
        emailBody = m1Sheet.getRange('A1').getValue();
      } else if (sheetIndicator === 'M2') {
        emailBody = m2Sheet.getRange('A1').getValue();
      } else if (sheetIndicator === 'M3') {
        emailBody = m3Sheet.getRange('A1').getValue();
      }
      
      var transportadora = '';
      for (var j = 1; j < cadastroEmailData.length; j++) {
        if (cadastroEmailData[j][0] === identifier) {
          transportadora = cadastroEmailData[j][3];
          break;
        }
      }
      
      if (transportadora === '') {
        SpreadsheetApp.getUi().alert('Transportadora não encontrada para o identificador: ' + identifier);
        return;
      }

      var date = new Date();
      var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      var emailSubject = 'Colgate - Devolução (Pendências/Ações) - ' + transportadora + ' - ' + formattedDate;
      
      var newSpreadsheet = SpreadsheetApp.create('TempSpreadsheet');
      var newSpreadsheetId = newSpreadsheet.getId();
      var newSpreadsheetFile = DriveApp.getFileById(newSpreadsheetId);
      
      pendenciaSheet.copyTo(newSpreadsheet).setName('Pendência');
      newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);
      
      var xlsxBlob = convertSpreadsheetToXlsx(newSpreadsheetId);
      xlsxBlob.setName('Pendencia.xlsx');
      
      MailApp.sendEmail({
        to: email,
        cc: ccEmails,
        subject: emailSubject,
        body: emailBody,
        attachments: [xlsxBlob]
      });
      
      newSpreadsheetFile.setTrashed(true);
      enviarParaSheet.getRange(i + 1, 7).setValue('enviado');
      
      var valueA = enviarParaData[i][0];
      for (var j = 1; j < enviarParaData.length; j++) {
        if (enviarParaData[j][0] === valueA) {
          enviarParaSheet.getRange(j + 1, 3).clearContent();
        }
      }
      
      var formattedTime = Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss');
      var lastRowHistorico = historicoSheet.getLastRow() + 1;
      historicoSheet.getRange(lastRowHistorico, 1).setValue(formattedDate);
      historicoSheet.getRange(lastRowHistorico, 2).setValue(formattedTime);
      historicoSheet.getRange(lastRowHistorico, 3).setValue(valueA);
      
      // Adiciona na coluna D do histórico a concatenação dos itens da coluna B das linhas cuja coluna C não está vazia.
      historicoSheet.getRange(lastRowHistorico, 4).setValue(conteudoColunaB);
      
      break;
    }
  }
  SpreadsheetApp.getUi().alert('E-mail enviado com sucesso.');
}

//***************************************************************************************************
// Copia dados da Coluna D para a Coluna A na mesma sheet, preservando a correspondência
function copyDataFromColumnDToColumnA() {
  // Abre a planilha desejada na aba "o que enviar"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("o que enviar");

  // Limpa a coluna A a partir da linha 2
  sheet.getRange('A2:A' + sheet.getLastRow()).clearContent();
  
  // Obtém todos os valores das colunas C e D a partir da linha 2
  var range = sheet.getRange('C2:D' + sheet.getLastRow());
  var values = range.getValues();

  // Cria um array para manter os valores que serão copiados para a coluna A
  var valuesToCopyToColumnA = [];

  // Loop pelos valores nas colunas C e D
  for (var i = 0; i < values.length; i++) {
    // Verifica se o valor na coluna C (índice 0) não está vazio
    if (values[i][0] !== '') {
      // Adiciona o valor correspondente da coluna D (índice 1) ao array
      var valuesOfColumnD = values[i][1];
      valuesToCopyToColumnA.push([valuesOfColumnD]);
    } else {
      // Adiciona um valor vazio para manter a correspondência das linhas
      valuesToCopyToColumnA.push(['']);
    }
  }

  // Escreve os valores obtidos na coluna A, a partir da linha 2
  if (valuesToCopyToColumnA.length > 0) {
    sheet.getRange(2, 1, valuesToCopyToColumnA.length, 1).setValues(valuesToCopyToColumnA);
  }
}

//***************************************************************************************************
// Copia dados relevantes da sheet "Base_e-mail" para a sheet "Pendência" com base em campos correspondentes
function copyDataFromBaseEmailToPendencia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Abre a planilha "Base_e-mail"
  var baseEmailSheet = ss.getSheetByName('Base_e-mail');
  // Abre a planilha "Pendência"
  var pendenciaSheet = ss.getSheetByName('Pendência');
  
  // Obtém o intervalo de dados em "Base_e-mail"
  var baseEmailRange = baseEmailSheet.getDataRange();
  var baseEmailValues = baseEmailRange.getValues();
  
  // Cria um dicionário para mapear endereços de e-mails da coluna N aos seus dados na planilha "Base_e-mail"
  var baseEmailDict = {};
  var headers = baseEmailValues[0];  // A primeira linha contém os cabeçalhos
  for (var i = 1; i < baseEmailValues.length; i++) {
    var row = baseEmailValues[i];
    var emailKey = row[13];  // Coluna N tem índice 13
    baseEmailDict[emailKey] = row;  // Mapeia o email inteiro para a linha de dados correspondente
  }
  
  // Obtém os cabeçalhos em "Pendências" a partir da coluna O
  var pendenciaHeaders = pendenciaSheet.getRange(1, 15, 1, pendenciaSheet.getLastColumn() - 14).getValues()[0];
  
  // Itera pelas linhas de dados em "Pendência"
  var pendenciaDataRange = pendenciaSheet.getRange(2, 14, pendenciaSheet.getLastRow() - 1, pendenciaSheet.getLastColumn() - 13);
  var pendenciaValues = pendenciaDataRange.getValues();
  
  for (var j = 0; j < pendenciaValues.length; j++) {
    var pendenciaRow = pendenciaValues[j];
    var pendenciaEmail = pendenciaRow[0];  // Coluna N tem índice 0 aqui
    if (pendenciaEmail in baseEmailDict) {
      // Se o email na sheet "Pendência" corresponde a um na "Base_e-mail"
      for (var k = 0; k < pendenciaHeaders.length; k++) {
        var header = pendenciaHeaders[k];
        var columnIndex = headers.indexOf(header);  // Encontra o índice do cabeçalho na Base de e-mails
        if (columnIndex !== -1) {
          // Se o cabeçalho existe, copia o dado correspondente para "Pendência"
          pendenciaSheet.getRange(j + 2, 15 + k).setValue(baseEmailDict[pendenciaEmail][columnIndex]);
        }
      }
    }
  }
}

//***************************************************************************************************
// Processa os dados e copia para folhas diferentes dependendo do status de devolução
function processAndCopyData_v4() {
  // IDs das planilhas
  var spreadsheetId = '14IyGceCIa8ntvBea6YTgT7-awO0kLCkspq5HGA9javU'; // ID da planilha de dados
  
  // Nomes das sheets a serem acessadas
  var baseEmailSheetName = 'Base_e-mail';
  var enviarParaSheetName = 'EnviarPara';
  var confirmarDevolucaoSheetName = 'ConfirmarDevolução';
  var agendarSheetName = 'Agendar';
  var anuenciaSheetName = 'Anuência';
  var debitoSheetName = 'Débito';
  
  // Abre a planilha usando o ID fornecido
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Acessa as sheets especificadas
  var baseEmailSheet = spreadsheet.getSheetByName(baseEmailSheetName);
  var enviarParaSheet = spreadsheet.getSheetByName(enviarParaSheetName);
  var confirmarDevolucaoSheet = spreadsheet.getSheetByName(confirmarDevolucaoSheetName);
  var agendarSheet = spreadsheet.getSheetByName(agendarSheetName);
  var anuenciaSheet = spreadsheet.getSheetByName(anuenciaSheetName);
  var debitoSheet = spreadsheet.getSheetByName(debitoSheetName);
  
  // Verifica se todas as sheets foram encontradas, lança erro se não
  if (!baseEmailSheet) throw new Error('Sheet "' + baseEmailSheetName + '" não encontrada.');
  if (!enviarParaSheet) throw new Error('Sheet "' + enviarParaSheetName + '" não encontrada.');
  if (!confirmarDevolucaoSheet) throw new Error('Sheet "' + confirmarDevolucaoSheetName + '" não encontrada.');
  if (!agendarSheet) throw new Error('Sheet "' + agendarSheetName + '" não encontrada.');
  if (!anuenciaSheet) throw new Error('Sheet "' + anuenciaSheetName + '" não encontrada.');
  if (!debitoSheet) throw new Error('Sheet "' + debitoSheetName + '" não encontrada.');
  
  // Limpa as sheets de destino a partir da linha 2
  clearSheet(confirmarDevolucaoSheet);
  clearSheet(agendarSheet);
  clearSheet(anuenciaSheet);
  clearSheet(debitoSheet);
  
  // Obtém os dados da sheet "EnviarPara"
  var enviarParaData = enviarParaSheet.getDataRange().getValues();
  var transportadoraIndexEnviarPara = 0; // Índice Coluna A
  var statusIndexEnviarPara = 1; // Índice Coluna B
  var checkIndexEnviarPara = 2; // Índice Coluna C
  
  // Obtém os dados da sheet "Base-e-mail"
  var baseEmailData = baseEmailSheet.getDataRange().getValues();
  var transportadoraIndexBaseEmail = 8; // Índice Coluna I
  var statusIndexBaseEmail = 10; // Índice Coluna K
  
  // Arrays para armazenar os dados a serem copiados para cada status
  var dataToCopyConfirmarDevolucao = [];
  var dataToCopyAgendar = [];
  var dataToCopyAnuencia = [];
  var dataToCopyDebito = [];
  
  // Processa cada linha da sheet "EnviarPara"
  for (var i = 1; i < enviarParaData.length; i++) {
    if (enviarParaData[i][checkIndexEnviarPara] !== '') {
      var transportadora = enviarParaData[i][transportadoraIndexEnviarPara];
      var status = enviarParaData[i][statusIndexEnviarPara];
      
      // Verifica cada linha da sheet "Base-e-mail"
      for (var j = 1; j < baseEmailData.length; j++) {
        if (baseEmailData[j][transportadoraIndexBaseEmail] === transportadora &&
            baseEmailData[j][statusIndexBaseEmail] === status) {
          // Copia os dados das colunas A à J
          var rowData = baseEmailData[j].slice(0, 10); 
          
          // Adiciona os dados ao array correspondente com base no status
          switch (status) {
            case 'Checar com WH e Transportadora se a nota foi entregue e Checar':
              dataToCopyConfirmarDevolucao.push(rowData);
              break;
            case 'Cobrar Agendamento Urgente ao Transportador':
              dataToCopyAgendar.push(rowData);
              break;
            case 'Cobrar Anuência da Transportadora com Urgência':
              dataToCopyAnuencia.push(rowData);
              break;
            case 'Informar Débito ao Transportador':
              dataToCopyDebito.push(rowData);
              break;
          }
        }
      }
    }
  }
  
  // Cola os dados nas sheets de destino: ConfirmarDevolucao, Agendar, Anuencia, Debito, conforme o status
  if (dataToCopyConfirmarDevolucao.length > 0) {
    confirmarDevolucaoSheet.getRange(2, 1, dataToCopyConfirmarDevolucao.length, dataToCopyConfirmarDevolucao[0].length).setValues(dataToCopyConfirmarDevolucao);
  }
  if (dataToCopyAgendar.length > 0) {
    agendarSheet.getRange(2, 1, dataToCopyAgendar.length, dataToCopyAgendar[0].length).setValues(dataToCopyAgendar);
  }
  if (dataToCopyAnuencia.length > 0) {
    anuenciaSheet.getRange(2, 1, dataToCopyAnuencia.length, dataToCopyAnuencia[0].length).setValues(dataToCopyAnuencia);
  }
  if (dataToCopyDebito.length > 0) {
    debitoSheet.getRange(2, 1, dataToCopyDebito.length, dataToCopyDebito[0].length).setValues(dataToCopyDebito);
  }
  
  // Mensagem de sucesso para o usuário
  SpreadsheetApp.getUi().alert('Dados copiados com sucesso para as sheets de destino.');
}