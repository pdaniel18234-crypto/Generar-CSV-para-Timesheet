function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Amazon Kenjo')
      .addItem('📧 1. Importar/Actualizar Correos', 'importarCorreosDesdeHojaDeRuta')
      .addItem('📊 2. Generar Excel (shiftplan.xlsx)', 'solicitarRangoYGenerarExcelKenjo')
      .addToUi();
}

/**
 * 1. Importar automáticamente correos desde Hoja de Ruta (Drivers)
 */
function importarCorreosDesdeHojaDeRuta() {
  var ssActual = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRotaciones = ssActual.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var ID_HOJA_RUTA = "1bS8YNCH_1HHLr8gUnMmXjIrmqEhHt9vg8etQ86ab7a0"; 

  try {
    var ssHojaRuta = SpreadsheetApp.openById(ID_HOJA_RUTA);
    var sheetDrivers = ssHojaRuta.getSheetByName("Drivers");

    if (!sheetDrivers) {
      ui.alert("Error: No se encontró la pestaña 'Drivers' en la Hoja de Ruta.");
      return;
    }

    var ultFilaDrivers = sheetDrivers.getLastRow();
    if (ultFilaDrivers < 11) {
      ui.alert("No hay suficientes datos en la pestaña Drivers.");
      return;
    }

    var datosDrivers = sheetDrivers.getRange(11, 8, ultFilaDrivers - 10, 2).getValues(); 
    
    var mapaCorreos = {};
    for (var d = 0; d < datosDrivers.length; d++) {
      var nombreDriver = datosDrivers[d][0] ? datosDrivers[d][0].toString().trim().toUpperCase() : "";
      var emailDriver = datosDrivers[d][1] ? datosDrivers[d][1].toString().trim().toLowerCase() : "";

      if (nombreDriver && emailDriver) {
        mapaCorreos[nombreDriver] = emailDriver;
      }
    }

    var filaInicio = 8;
    var filaFin = 98;
    var totalConductores = (filaFin - filaInicio) + 1;

    var nombresRotaciones = sheetRotaciones.getRange(filaInicio, 2, totalConductores, 1).getValues(); 
    var matrizCorreosSalida = [];
    var correosEncontrados = 0;

    for (var i = 0; i < totalConductores; i++) {
      var nombreConductor = nombresRotaciones[i][0] ? nombresRotaciones[i][0].toString().trim().toUpperCase() : "";

      if (nombreConductor && mapaCorreos[nombreConductor]) {
        matrizCorreosSalida.push([mapaCorreos[nombreConductor]]);
        correosEncontrados++;
      } else {
        matrizCorreosSalida.push([""]); 
      }
    }

    sheetRotaciones.getRange(filaInicio, 4, totalConductores, 1).setValues(matrizCorreosSalida);

    ui.alert("✅ Correos actualizados", "Se han importado " + correosEncontrados + " correos en la Columna D.", ui.ButtonSet.OK);

  } catch (error) {
    ui.alert("Ocurrió un error al conectar con la Hoja de Ruta:\n" + error.toString());
  }
}

/**
 * 2. Solicitar rango de columnas
 */
function solicitarRangoYGenerarExcelKenjo() {
  var ui = SpreadsheetApp.getUi();
  
  var respuesta = ui.prompt(
    'Generar Plantilla Excel Kenjo', 
    'Introduce el rango de columnas de la semana (ejemplo: H:N o MG:NL):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuesta.getSelectedButton() == ui.Button.OK) {
    var rangoTexto = respuesta.getResponseText().trim();
    if (rangoTexto === "" || rangoTexto.indexOf(":") === -1) {
      ui.alert("Por favor, introduce un rango válido (ejemplo: H:N).");
      return;
    }
    generarExcelKenjo(rangoTexto);
  }
}

/**
 * 3. Generar archivo .XLSX real con los campos ajustados para Kenjo
 */
function generarExcelKenjo(rangoColumnasTexto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var zonaHorariaContable = ss.getSpreadsheetTimeZone();
  
  try {
    var partes = rangoColumnasTexto.split(':');
    var letraInicio = partes[0].trim().toUpperCase();
    var letraFin = partes[1].trim().toUpperCase();
    
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
      ui.alert("El rango introducido no es válido.");
      return;
    }
    
    var filaInicio = 8;
    var filaFin = 98;
    var totalConductores = (filaFin - filaInicio) + 1; 
    
    // Leer Columna C (ID) y Columna D (Email)
    var datosConductores = sheet.getRange(filaInicio, 3, totalConductores, 2).getValues(); 
    var estadosValidos = ["ST", "L1", "L2", "L3", "L4", "BU", "EXTRA", "SP"];
    
    // Encabezados exactos de shiftplan.xlsx
    var filasExcel = [
      [
        "Email de trabajo (Se requiere al menos el email o el ID de empleado)",
        "ID de empleado (Se requiere al menos el email o el ID de empleado)",
        "Fecha * (YYYY-MM-DD)",
        "Hora de inicio * (HH:mm)",
        "Hora de fin * (HH:mm)",
        "Tiempo de descanso (minutos)",
        "Ubicación *",
        "Rol *",
        "Área de trabajo",
        "Etiqueta",
        "Notas"
      ]
    ];
    
    var lineasAgregadas = 0;
    
    for (var col = colInicio; col <= colFin; col++) {
      var valorCeldaFecha = sheet.getRange(6, col).getValue();
      
      if (!valorCeldaFecha || !(valorCeldaFecha instanceof Date) || isNaN(valorCeldaFecha.getTime())) {
        continue;
      }
      
      var fechaFormateada = Utilities.formatDate(valorCeldaFecha, zonaHorariaContable, "yyyy-MM-dd");
      var estadosConductoresDia = sheet.getRange(filaInicio, col, totalConductores, 1).getValues();
      
      for (var i = 0; i < totalConductores; i++) {
        var emailEmp = datosConductores[i][1] ? datosConductores[i][1].toString().trim() : ""; 
        var estadoRaw = estadosConductoresDia[i][0]; 
        
        if (!emailEmp) continue; // Si no hay email en la columna D, omitimos la fila
        if (!estadoRaw) continue;
        
        var estadoDia = estadoRaw.toString().trim().toUpperCase();
        
        if (estadosValidos.indexOf(estadoDia) !== -1) {
          var horaInicio, horaFin;
          
          if (estadoDia === "SP") {
            horaInicio = "13:00";
            horaFin = "21:00";
          } else {
            horaInicio = "10:00";
            horaFin = "19:00";
          }
          
          filasExcel.push([
            emailEmp,       // Email de trabajo
            "",             // ID de empleado (Vacío)
            fechaFormateada,// Fecha * (YYYY-MM-DD)
            horaInicio,     // Hora de inicio * (HH:mm)
            horaFin,        // Hora de fin * (HH:mm)
            0,              // Tiempo de descanso (minutos)
            "DMZ4",         // Ubicación *
            "Conductor",    // Rol *
            "",             // Área de trabajo (Vacío)
            "",             // Etiqueta (Vacío)
            ""              // Notas (Vacío)
          ]);
          
          lineasAgregadas++;
        }
      }
    }
    
    if (lineasAgregadas > 0) {
      var nombreArchivo = "shiftplan_Kenjo_" + letraInicio + "_" + letraFin + ".xlsx";
      var datosJson = JSON.stringify(filasExcel);
      
      var htmlContent = '<!DOCTYPE html><html><head>' +
        '<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>' +
        '</head><body>' +
        '<p>Generando archivo Excel (<b>.xlsx</b>)...</p>' +
        '<script>' +
        '  var data = ' + datosJson + ';' +
        '  var ws = XLSX.utils.aoa_to_sheet(data);' +
        '  var wb = XLSX.utils.book_new();' +
        '  XLSX.utils.book_append_sheet(wb, ws, "shiftplan");' +
        '  XLSX.writeFile(wb, "' + nombreArchivo + '");' +
        '  setTimeout(function(){ google.script.host.close(); }, 1500);' +
        '</script>' +
        '</body></html>';
      
      var htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(380).setHeight(130);
      ui.showModalDialog(htmlOutput, 'Descargando Excel');
    } else {
      ui.alert("No se encontraron turnos válidos en las columnas indicadas.");
    }
    
  } catch(error) {
    ui.alert("Ocurrió un error inesperado:\n" + error.toString());
  }
}
