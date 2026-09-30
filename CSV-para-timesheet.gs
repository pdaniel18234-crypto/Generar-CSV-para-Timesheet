function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Amazon Timesheet') // <-- Aquí añadimos la cajita al menú principal
      .addItem('📥 Generar CSV por Rango', 'solicitarRangoYGenerarCsv') // <-- Y un icono de descarga al botón
      .addToUi();
}

function solicitarRangoYGenerarCsv() {
  var ui = SpreadsheetApp.getUi();
  
  var respuesta = ui.prompt(
    'Exportar por Rango de Columnas', 
    'Introduce el rango total de la semana (ejemplo: MG:NL):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuesta.getSelectedButton() == ui.Button.OK) {
    var rangoTexto = respuesta.getResponseText().trim();
    if (rangoTexto === "" || rangoTexto.indexOf(":") === -1) {
      ui.alert("Por favor, introduce un rango válido (ejemplo: MG:NL).");
      return;
    }
    generarCsvTimesheetInteligente(rangoTexto);
  }
}

function generarCsvTimesheetInteligente(rangoColumnasTexto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var zonaHorariaContable = ss.getSpreadsheetTimeZone();
  
  try {
    var partes = rangoColumnasTexto.split(':');
    var letraInicio = partes[0].trim().toUpperCase();
    var letraFin = partes[1].trim().toUpperCase();
    
    // Función para convertir letras de columna a números indexados
    var letraANumero = function(letra) {
      var num = 0;
      for (var i = 0; i < letra.length; i++) {
        num = num * 26 + (letra.charCodeAt(i) - 64);
      }
      return num;
    };
    
    var colInicio = letraANumero(letraInicio);
    var colFin = letraANumero(letraFin);
    
    if (colInicio <= 0 || colFin <= 0 || colInicio > colFin) {
      ui.alert("El rango introducido no es válido. Asegúrate de que la primera columna sea menor que la segunda.");
      return;
    }
    
    // Configuración estricta de filas para conductores (Fila 8 a 98)
    var filaInicio = 8;
    var filaFin = 98;
    var totalConductores = (filaFin - filaInicio) + 1; 
    
    var codigosConductores = sheet.getRange(filaInicio, 3, totalConductores, 1).getValues();
    var estadosValidos = ["ST", "L1", "L2", "L3", "L4", "BU", "EXTRA", "SP"];
    
    var csvContenido = "emp_number;job_number;task_number;start_date;end_date;event_information;punch_in_reminder;punch_out_reminder;event_type;payroll_name\r\n";
    var lineasAgregadas = 0;
    
    // RECORRIDO INTELIGENTE: Analizamos columna por columna en el rango especificado
    for (var col = colInicio; col <= colFin; col++) {
      
      var valorCeldaFecha = sheet.getRange(6, col).getValue();
      
      // Si la celda actual no contiene una fecha real, pasamos a la siguiente columna
      // Esto se salta automáticamente las columnas de separación vacías (ya sean de 2 o 3 espacios)
      if (!valorCeldaFecha || !(valorCeldaFecha instanceof Date)) {
        continue;
      }
      
      // Formateamos la fecha blindándola contra desfases de zona horaria
      var fechaFormateada = Utilities.formatDate(valorCeldaFecha, zonaHorariaContable, "d/M/yyyy");
      
      // Absorber los estados de los conductores para esta columna con fecha válida
      var estadosConductoresDia = sheet.getRange(filaInicio, col, totalConductores, 1).getValues();
      
      // Procesar cada conductor verticalmente
      for (var i = 0; i < totalConductores; i++) {
        var empNumber = codigosConductores[i][0]; 
        var estadoRaw = estadosConductoresDia[i][0]; 
        
        if (!empNumber || empNumber.toString().trim() === "") continue;
        if (!estadoRaw) continue;
        
        var estadoDia = estadoRaw.toString().trim().toUpperCase();
        
        if (estadosValidos.indexOf(estadoDia) !== -1) {
          var horaInicio, horaFin;
          
          // REGLA HORARIA PARA SOPORTE (SP): Turno de 8 horas (13:00 a 21:00)
          if (estadoDia === "SP") {
            horaInicio = fechaFormateada + " 13:00:00";
            horaFin = fechaFormateada + " 21:00:00";
          } else {
            // Horario estándar para ST, L1, L2, L3, L4, BU, EXTRA (10:00 a 19:00)
            horaInicio = fechaFormateada + " 10:00:00";
            horaFin = fechaFormateada + " 19:00:00";
          }
          
          csvContenido += empNumber + ";1;1;" + horaInicio + ";" + horaFin + ";1;1;1;0;\r\n";
          lineasAgregadas++;
        }
      }
    }
    
    // Generar la descarga automática del archivo final
    if (lineasAgregadas > 0) {
      var nombreArchivo = "TimeSheet_Completo_" + letraInicio + "_" + letraFin + ".csv";
      
      var blob = Utilities.newBlob(csvContenido, 'text/csv;charset=utf-8', nombreArchivo);
      var archivoBase64 = Utilities.base64Encode(blob.getBytes());
      
      var htmlOutput = HtmlService.createHtmlOutput(
        '<html><body>' +
        '<p>CSV generado con éxito para el rango <b>' + rangoColumnasTexto.toUpperCase() + '</b> (' + lineasAgregadas + ' turnos procesados).</p>' +
        '<a id="download_link" download="' + nombreArchivo + '" href="data:text/csv;charset=utf-8;base64,' + archivoBase64 + '" style="padding:10px 20px; background-color:#ff9900; color:white; text-decoration:none; font-weight:bold; border-radius:5px; display:inline-block; margin-top:10px;">Descargar CSV</a>' +
        '<script>document.getElementById("download_link").click();</script>' +
        '</body></html>'
      ).setWidth(420).setHeight(160);
      
      ui.showModalDialog(htmlOutput, 'Descargar CSV');
    } else {
      ui.alert("No se encontraron turnos válidos en las columnas indicadas.");
    }
    
  } catch(error) {
    ui.alert("Ocurrió un error inesperado:\n" + error.toString());
  }
}
