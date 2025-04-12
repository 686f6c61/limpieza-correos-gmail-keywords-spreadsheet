/**
 * Email Unsubscribe Manager
 * 
 * Script para identificar y gestionar correos de suscripción en Gmail.
 * Permite buscar, listar y eliminar correos no deseados de forma selectiva o masiva.
 * 
 * Autor: 686f6c61
 * Fecha: 12/04/2025
 * v0.2
 */

// Configuración global - Ajustar estos valores si necesitas personalizar el script
const CONFIG = {
  MAX_EMAILS: 500,        // Máx correos a procesar (500 = límite de Gmail API)
  SHEET_NAME: "Correos de Suscripción", // Nombre hoja resultados
  CONFIG_SHEET_NAME: "Configuración",   // Nombre hoja config
  KEYWORDS_RANGE: "B2:B",  // Rango para keywords personalizadas
  CONFIRM_CELL: "B1",      // Celda confirmación para borrado masivo
  CONFIRM_TEXT: "CONFIRMAR" // Texto de confirmación (case sensitive)
};

/**
 * Crea el menú en la UI cuando se abre la hoja
 * @return {void}
 */
function onOpen() {
  try {
    // Creamos el menú en la UI
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Gestor de Suscripciones')
      .addItem('Buscar correos de suscripción', 'setupAndFindSubscriptionEmails') // Función principal
      .addItem('Eliminar correos seleccionados', 'deleteSelectedEmails')           // Borrado selectivo
      .addItem('Eliminar todos los correos listados', 'deleteAllEmails')           // Borrado masivo
      .addToUi();
  } catch (error) {
    // Log del error sin romper ejecución
    Logger.log('Error al crear menú: ' + error.toString()); // Probablemente ejecutando desde editor
  }
}

/**
 * Handler para la instalación del complemento
 * @param {Object} e - Evento de instalación
 * @return {void}
 */
function onInstall() {
  onOpen();
}

/**
 * Prepara el entorno y ejecuta la búsqueda de correos
 * Punto de entrada principal para la aplicación
 * @return {void}
 */
function setupAndFindSubscriptionEmails() {
  try {
    // Get spreadsheet activa o crear nueva
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Setup hoja principal (crear o limpiar si existe)
    let mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!mainSheet) {
      mainSheet = ss.insertSheet(CONFIG.SHEET_NAME); // Nueva hoja
    } else {
      mainSheet.clear(); // Reset si ya existe
    }
    
    // Setup hoja config (crear si no existe)
    let configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG.CONFIG_SHEET_NAME);
      setupConfigSheet(configSheet); // Inicializar con defaults
    }
    
    // Lanzar búsqueda principal
    findSubscriptionEmails(mainSheet, configSheet); // Función core
    
  } catch (error) {
    Logger.log('Error en setupAndFindSubscriptionEmails: ' + error.toString());
    try {
      SpreadsheetApp.getUi().alert('Error al configurar: ' + error.toString());
    } catch (uiError) {
      Logger.log('No se pudo mostrar alerta: ' + uiError.toString());
    }
  }
}

/**
 * Inicializa la hoja de configuración con valores por defecto
 * @param {Sheet} sheet - Hoja de configuración a inicializar
 * @return {void}
 */
function setupConfigSheet(sheet) {
  // Establecer encabezados y valores predeterminados
  sheet.getRange("A1").setValue("Confirmar eliminación masiva (escriba CONFIRMAR):");
  sheet.getRange("B1").setValue("");
  
  sheet.getRange("A2").setValue("Palabras clave predefinidas:");
  sheet.getRange("B2").setValue("unsubscribe");
  sheet.getRange("B3").setValue("darse de baja");
  sheet.getRange("B4").setValue("cancelar suscripción");
  sheet.getRange("B5").setValue("cancel subscription");
  sheet.getRange("B6").setValue("opt-out");
  sheet.getRange("B7").setValue("unsuscribe");
  
  // Instrucciones
  sheet.getRange("A9:B9").merge();
  sheet.getRange("A9").setValue("Añada palabras clave adicionales en las celdas B8 en adelante");
  
  // Formatear la hoja
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
}

/**
 * Core del script: busca correos con keywords y los muestra en la hoja
 * @param {Sheet} mainSheet - Hoja principal donde mostrar resultados
 * @param {Sheet} configSheet - Hoja de configuración con keywords
 * @return {void}
 */
function findSubscriptionEmails(mainSheet, configSheet) {
  try {
    // Configurar encabezados en la hoja principal
    const headers = [
      "Seleccionar", "Fecha", "Asunto", "Remitente", "Nombre Remitente", 
      "Dominio", "Palabra Clave", "Fragmento", "ID del Mensaje"
    ];
    
    mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    mainSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    
    // Obtener palabras clave personalizadas
    const keywordsRange = configSheet.getRange(CONFIG.KEYWORDS_RANGE);
    const keywordsValues = keywordsRange.getValues();
    const keywords = keywordsValues.flat().filter(String);
    
    if (keywords.length === 0) {
      try {
        SpreadsheetApp.getUi().alert('No se encontraron palabras clave. Por favor, añada al menos una palabra clave en la hoja de Configuración.');
      } catch (uiError) {
        Logger.log('No se pudo mostrar alerta: ' + uiError.toString());
      }
      return;
    }
    
    // Construir la consulta de búsqueda para Gmail
    // Usar un enfoque diferente para evitar el error de límite
    const searchQuery = keywords.map(keyword => `"${keyword}"`).join(" OR ");
    
    // Limitar estrictamente a 500 hilos (límite de la API)
    let threads;
    try {
      // Intentar con el método search
      threads = GmailApp.search(searchQuery, 0, 500);
    } catch (searchError) {
      // Si falla, usar una alternativa más segura
      Logger.log('Error en búsqueda estándar: ' + searchError.toString());
      threads = GmailApp.search(searchQuery);
      // Limitar manualmente el número de hilos
      if (threads.length > 500) {
        threads = threads.slice(0, 500);
      }
    }
    
    // Procesar los resultados
    let row = 2; // Comenzar después de los encabezados
    let resultsData = [];
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        // Extraer información del mensaje
        const date = message.getDate();
        const subject = message.getSubject();
        const from = message.getFrom();
        const body = message.getPlainBody();
        
        // Extraer nombre y dominio del remitente
        let senderName = "";
        let senderEmail = from;
        let senderDomain = "";
        
        const emailMatch = from.match(/<([^>]+)>/);
        if (emailMatch) {
          senderEmail = emailMatch[1];
          senderName = from.split('<')[0].trim();
        }
        
        const domainMatch = senderEmail.match(/@([^>]+)/);
        if (domainMatch) {
          senderDomain = domainMatch[1];
        }
        
        // Encontrar qué palabra clave coincidió y un fragmento relevante
        let matchedKeyword = "";
        let snippet = "";
        
        for (const keyword of keywords) {
          if (body.toLowerCase().includes(keyword.toLowerCase())) {
            matchedKeyword = keyword;
            
            // Extraer un fragmento de texto alrededor de la palabra clave
            const keywordIndex = body.toLowerCase().indexOf(keyword.toLowerCase());
            const startIndex = Math.max(0, keywordIndex - 50);
            const endIndex = Math.min(body.length, keywordIndex + keyword.length + 50);
            snippet = "..." + body.substring(startIndex, endIndex).replace(/\n/g, " ") + "...";
            
            break;
          }
        }
        
        // Si se encontró una coincidencia, añadir a los resultados
        if (matchedKeyword) {
          resultsData.push([
            false, // Checkbox para selección
            date,
            subject,
            senderEmail,
            senderName,
            senderDomain,
            matchedKeyword,
            snippet,
            message.getId() // Guardar ID para poder eliminar después
          ]);
        }
      }
    }
    
    // Si hay resultados, añadirlos a la hoja
    if (resultsData.length > 0) {
      // Crear casillas de verificación en la primera columna
      const checkboxRange = mainSheet.getRange(2, 1, resultsData.length, 1);
      checkboxRange.insertCheckboxes();
      
      // Escribir los datos
      mainSheet.getRange(2, 2, resultsData.length, resultsData[0].length - 1)
               .setValues(resultsData.map(row => row.slice(1)));
      
      // Añadir IDs de mensajes en la última columna (oculta)
      const idColumn = headers.length;
      mainSheet.getRange(2, idColumn, resultsData.length, 1)
               .setValues(resultsData.map(row => [row[row.length - 1]]));
      mainSheet.hideColumns(idColumn);
      
      // Formatear la hoja
      mainSheet.autoResizeColumns(2, headers.length - 2);
      
      // Añadir botones de acción
      addActionButtons(mainSheet, resultsData.length + 3);
      
      try {
        SpreadsheetApp.getUi().alert(`Se encontraron ${resultsData.length} correos de suscripción.`);
      } catch (uiError) {
        Logger.log(`Se encontraron ${resultsData.length} correos de suscripción, pero no se pudo mostrar alerta: ${uiError}`);
      }
    } else {
      mainSheet.getRange(2, 2).setValue("No se encontraron correos con las palabras clave especificadas.");
      try {
        SpreadsheetApp.getUi().alert('No se encontraron correos con las palabras clave especificadas.');
      } catch (uiError) {
        Logger.log('No se encontraron correos con las palabras clave especificadas, pero no se pudo mostrar alerta: ' + uiError);
      }
    }
    
  } catch (error) {
    Logger.log('Error en findSubscriptionEmails: ' + error.toString());
    try {
      SpreadsheetApp.getUi().alert('Error al buscar correos: ' + error.toString());
    } catch (uiError) {
      Logger.log('Error al buscar correos y no se pudo mostrar alerta: ' + error.toString());
    }
  }
}

/**
 * Crea los botones de acción en la hoja de resultados
 * @param {Sheet} sheet - Hoja donde añadir los botones
 * @param {Number} startRow - Fila donde colocar los botones
 * @return {void}
 */
function addActionButtons(sheet, startRow) {
  // Crear botones usando dibujos insertados
  sheet.getRange(startRow, 2).setValue("Acciones:");
  
  // Botón para eliminar seleccionados
  sheet.getRange(startRow, 3).setValue("Eliminar Seleccionados")
       .setBackground("#f1c232")
       .setFontWeight("bold")
       .setBorder(true, true, true, true, true, true);
  
  // Botón para eliminar todos
  sheet.getRange(startRow, 4).setValue("Eliminar Todos")
       .setBackground("#ea9999")
       .setFontWeight("bold")
       .setBorder(true, true, true, true, true, true);
  
  // Asignar scripts a los botones
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Crear una nota explicativa para los botones
  sheet.getRange(startRow, 5).setValue("Haga clic en las celdas de color para ejecutar las acciones");
}

/**
 * Elimina solo los correos marcados con checkbox
 * @return {void}
 */
function deleteSelectedEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('No se encontró la hoja principal. Por favor, ejecute primero la búsqueda de correos.');
      return;
    }
    
    // Obtener el rango de datos con correos
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      SpreadsheetApp.getUi().alert('No hay correos para eliminar.');
      return;
    }
    
    // Encontrar la columna del ID del mensaje
    const idColumnIndex = values[0].indexOf("ID del Mensaje");
    if (idColumnIndex === -1) {
      SpreadsheetApp.getUi().alert('No se encontró la columna de ID del mensaje. Por favor, ejecute nuevamente la búsqueda.');
      return;
    }
    
    // Contar cuántos correos están seleccionados
    let selectedCount = 0;
    let selectedEmails = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === true) { // Si el checkbox está marcado
        selectedCount++;
        selectedEmails.push(values[i][idColumnIndex]);
      }
    }
    
    if (selectedCount === 0) {
      SpreadsheetApp.getUi().alert('No hay correos seleccionados para eliminar.');
      return;
    }
    
    // Confirmar antes de eliminar
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Confirmar eliminación',
      `¿Está seguro de que desea eliminar ${selectedCount} correos seleccionados?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Eliminar los correos seleccionados
    let deletedCount = 0;
    
    for (const messageId of selectedEmails) {
      try {
        const message = GmailApp.getMessageById(messageId);
        if (message) {
          message.moveToTrash();
          deletedCount++;
        }
      } catch (e) {
        Logger.log('Error al eliminar mensaje ' + messageId + ': ' + e.toString());
      }
    }
    
    // Actualizar la hoja después de eliminar
    if (deletedCount > 0) {
      SpreadsheetApp.getUi().alert(`Se eliminaron ${deletedCount} correos correctamente.`);
      // Volver a ejecutar la búsqueda para actualizar la lista
      setupAndFindSubscriptionEmails();
    } else {
      SpreadsheetApp.getUi().alert('No se pudo eliminar ningún correo. Verifique los permisos.');
    }
    
  } catch (error) {
    Logger.log('Error en deleteSelectedEmails: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error al eliminar correos: ' + error.toString());
  }
}

/**
 * Elimina todos los correos de la lista (requiere confirmación)
 * @return {void}
 */
function deleteAllEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!sheet || !configSheet) {
      SpreadsheetApp.getUi().alert('No se encontraron las hojas necesarias. Por favor, ejecute primero la búsqueda de correos.');
      return;
    }
    
    // Verificar la confirmación en la hoja de configuración
    const confirmCell = configSheet.getRange(CONFIG.CONFIRM_CELL);
    const confirmValue = confirmCell.getValue();
    
    if (confirmValue !== CONFIG.CONFIRM_TEXT) {
      SpreadsheetApp.getUi().alert(
        `Para eliminar todos los correos, debe escribir "${CONFIG.CONFIRM_TEXT}" en la celda ${CONFIG.CONFIRM_CELL} de la hoja de Configuración.`
      );
      return;
    }
    
    // Obtener el rango de datos con correos
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      SpreadsheetApp.getUi().alert('No hay correos para eliminar.');
      return;
    }
    
    // Encontrar la columna del ID del mensaje
    const idColumnIndex = values[0].indexOf("ID del Mensaje");
    if (idColumnIndex === -1) {
      SpreadsheetApp.getUi().alert('No se encontró la columna de ID del mensaje. Por favor, ejecute nuevamente la búsqueda.');
      return;
    }
    
    // Contar cuántos correos hay para eliminar
    const emailCount = values.length - 1; // Restar la fila de encabezados
    
    // Confirmar antes de eliminar
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Confirmar eliminación masiva',
      `¿Está seguro de que desea eliminar TODOS los ${emailCount} correos listados?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Eliminar todos los correos
    let deletedCount = 0;
    
    for (let i = 1; i < values.length; i++) {
      try {
        const messageId = values[i][idColumnIndex];
        const message = GmailApp.getMessageById(messageId);
        if (message) {
          message.moveToTrash();
          deletedCount++;
        }
      } catch (e) {
        Logger.log('Error al eliminar mensaje ' + i + ': ' + e.toString());
      }
    }
    
    // Resetear la celda de confirmación
    confirmCell.setValue("");
    
    // Actualizar la hoja después de eliminar
    if (deletedCount > 0) {
      SpreadsheetApp.getUi().alert(`Se eliminaron ${deletedCount} correos correctamente.`);
      // Volver a ejecutar la búsqueda para actualizar la lista
      setupAndFindSubscriptionEmails();
    } else {
      SpreadsheetApp.getUi().alert('No se pudo eliminar ningún correo. Verifique los permisos.');
    }
    
  } catch (error) {
    Logger.log('Error en deleteAllEmails: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error al eliminar correos: ' + error.toString());
  }
}

/**
 * Event handler para los clics en la hoja (maneja los botones)
 * @param {Object} e - Evento de edición
 * @return {void}
 */
function onEdit(e) {
  // Verificar si el clic fue en uno de los botones
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    return;
  }
  
  // Buscar la fila de los botones (después de los datos)
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  let buttonRow = -1;
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][1] === "Acciones:") {
      buttonRow = i + 1; // +1 porque los índices de array empiezan en 0
      break;
    }
  }
  
  if (buttonRow === -1 || range.getRow() !== buttonRow) {
    return;
  }
  
  // Verificar qué botón se presionó
  if (range.getColumn() === 3) { // Botón "Eliminar Seleccionados"
    deleteSelectedEmails();
  } else if (range.getColumn() === 4) { // Botón "Eliminar Todos"
    deleteAllEmails();
  }
}
