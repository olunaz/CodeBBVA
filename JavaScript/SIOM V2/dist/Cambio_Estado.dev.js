"use strict";

function aceptacion_masiva() {
  var correo_propietario = "luis.luna.cruz@bbva.com"; //Definicion de la hoja

  var id = SpreadsheetApp.openById("154qpAC9tk2WzbIaBPf-Cn2sCFqssExqrdhh1KObFDk4");
  var ss = id.getSheetByName("WF");
  var datawf = ss.getRange(2, 1, ss.getRange("A10").getDataRegion().getLastRow(), ss.getLastColumn()).getValues();
  var sscc = id.getSheetByName("MAQUETA LOURDES");
  var data = sscc.getRange(2, 1, sscc.getRange("A10").getDataRegion().getLastRow(), sscc.getLastColumn()).getValues(); //Lectura de la base

  for (var i = 0; i < data.length; i++) {
    var peticion = data[i][0];
    var cod_central = data[i][1];
    var cliente = data[i][2];
    var grupo = data[i][3];
    var tipo_estado = data[i][4];
    var detalle_balance = data[i][5];
    var detalle_resultados = data[i][6];
    var archivos = data[i][7];
    var numero_trabajadores = data[i][8];
    var oficina = data[i][9];
    var cod_gestor = data[i][10];
    var nombre_gestor = data[i][11];
    var correo_gestor = data[i][12];
    var correo_ingresante = data[i][13];
    var analista_asignado = data[i][14];
    var correo_analista = data[i][15];
    var giro_negocio = data[i][16];
    var tipo_metodizacion = data[i][17];
    var ventas = data[i][18];
    var periodo = data[i][19];
    var fecha_checklist = data[i][20];
    var registro_fecha = data[i][21];
    var fecha_desestimacion = data[i][22];
    var fecha_revision = data[i][23];
    var fecha_cierre = data[i][24];
    var tipo_asignacion = data[i][25];
    var tipo_sancion = data[i][26];
    var tipo_estadov2 = data[i][27];
    var cliente_priorizado = data[i][28];
    var correo_gestorv2 = data[i][29];
    var correo_ingresantev2 = data[i][30];
    var correo_analistav2 = data[i][31]; //var fc = new Date(marca_temporal);
    //var text_fc = Utilities.formatDate(fc, "GMT-5", "yyyy-MM-dd' 'HH:mm:ss' '");
    //var marca_temporal = text_fc.substring(8, 10) +'/'+ text_fc.substring(5, 7) +'/'+ 
    //text_fc.substring(0, 4) +' '+ text_fc.substring(11, 19);

    /*const conditionsArray = [
        analista_asignado == "LOURDES DEL PILAR GUTIERREZ MANCHEGO", 
        tipo_sancion == ""
    ]*/

    if (analista_asignado == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" && cliente_priorizado == "SÍ") {
      var recipient = correo_gestorv2;
      var subject = "(PRUEBA MEJORA ACEPTACIÓN MASIVA) SIOM: La Peticion de Metodizacion con Codigo " + peticion + " (" + cliente + ") Ha Sido Aceptada";
      var body = "Segmento Empresas ha aceptado la peticion con codigo: " + peticion + "\n" + "\nInformacion general sobre la operacion: " + "\nCodigo Central: " + cod_central + "\nGestor: " + nombre_gestor + "\nCliente: " + cliente + "\nGrupo : " + grupo + "\nTipo de Estado: " + tipo_estado + "\nAnalista Asignado: " + analista_asignado + "\n\nDetalle de Balance General: " + detalle_balance + "\n\nDetalle de Estado de Resultados: " + detalle_resultados + "\nArchivo Adjunto: " + archivos + "\n" + "\n#CampañaEEFF2022" + "\n#MONETIZACIÓNBEC" + "\n\n\nPor el momento, no requiere realizar ninguna accion adicional en cuanto a esta peticion.";
      var options = {
        cc: correo_ingresantev2 + ",lourdes.gutierrez@bbva.com, wendy.amaya.tobon@bbva.com, siom@bbva.com"
      };
      MailApp.sendEmail(recipient, subject, body, options);
      /*
       var rpeticion = ss2.getRange(a,1);
       var rcod_central = ss2.getRange(a,2);
       var rcliente = ss2.getRange(a,3);
       var rgrupo = ss2.getRange(a,4);
       var restado = ss2.getRange(a,5);
       var rdetalle_balance = ss2.getRange(a,6);
       var rdetalle_resultados = ss2.getRange(a,7);
       
       rpeticion.setValue(peticion);
       rcod_central.setValue(cod_central);
       rcliente.setValue(cliente);
       rgrupo.setValue(grupo);
       restado.setValue(tipo_estado);
       rdetalle_balance.setValue(tipo_estado);
       rdetalle_resultados.setValue(tipo_resultados);
         a = a + 1;
       */
    }

    if (analista_asignado == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" && cliente_priorizado == "NO") {
      var recipient = correo_gestorv2;
      var subject = "(PRUEBA MEJORA ACEPTACIÓN MASIVA) SIOM: La Peticion de Metodizacion con Codigo " + peticion + " (" + cliente + ") Ha Sido Desestimada";
      var body = "Segmento Empresas ha desestimado la peticion con codigo: " + peticion + "\n" + "\nInformacion general sobre la operacion: " + "\nCodigo Central: " + cod_central + "\nGestor: " + nombre_gestor + "\nCliente: " + cliente + "\nGrupo : " + grupo + "\nTipo de Estado: " + tipo_estado + "\nAnalista Asignado: " + analista_asignado + "\n\nDetalle de Balance General: " + detalle_balance + "\n\nDetalle de Estado de Resultados: " + detalle_resultados + "\nArchivo Adjunto: " + archivos + "\n" + "\nPor el momento, no se estan aceptando peticiones que se encuentren fuera de la lista de Priorización de EEFF." + "\n#CampañaEEFF2022" + "\n#MONETIZACIÓNBEC";
      var options = {
        cc: correo_ingresantev2 + ",lourdes.gutierrez@bbva.com, wendy.amaya.tobon@bbva.com, siom@bbva.com"
      };
      MailApp.sendEmail(recipient, subject, body, options);
      /*
       var rpeticion = ss2.getRange(a,1);
       var rcod_central = ss2.getRange(a,2);
       var rcliente = ss2.getRange(a,3);
       var rgrupo = ss2.getRange(a,4);
       var restado = ss2.getRange(a,5);
       var rdetalle_balance = ss2.getRange(a,6);
       var rdetalle_resultados = ss2.getRange(a,7);
       
       rpeticion.setValue(peticion);
       rcod_central.setValue(cod_central);
       rcliente.setValue(cliente);
       rgrupo.setValue(grupo);
       restado.setValue(tipo_estado);
       rdetalle_balance.setValue(tipo_estado);
       rdetalle_resultados.setValue(tipo_resultados);
         a = a + 1;
       */
    }
  }
  /*
    for (var j = 0; j < datawf.length; j++) {
  
      var peticionwf    = data[j][0]; 
      var cod_centralwf    = data[j][1];   
      var rucwf    = data[j][2];  
      var tipo_documentowf        = data[j][3];   
      var cod_oficinawf       = data[j][4];
      var nombre_oficinawf = data[j][5];
      var cod_gestorwf = data[j][6];
      var nombre_gestorwf = data[j][7];
      var procedenciawf = data[j][8];
      var fecha_basewf = data[j][9];
      var fecha_finalwf = data[j][10];
      var cod_territoriowf = data[j][11];
      var nombre_territoriowf = data[j][12];
      var clientewf = data[j][13];
      var grupofwf = data[j][14];
      var archivowf = data[j][15];
      var cod_documentowf = data[j][16];
      var tipo_estadowf = data[j][17];
      var detalle_balancewf = data[j][18];
      var detallle_resultadoswf = data[j][19];
      var analista_asignadowf = data[j][20];
      var fecha_checklistwf = data[j][21];
      var primer_ingreso_checklistwf = data[j][22];
      var devuelve_siomwf = data [j][23];
      var fecha_revisionwf = data[j][25];
      var fecha_consulta_bewf = data[j][26];
      var fecha_devolucionwf = data[j][27];
      var fecha_cierrewf = data[j][30];
      var tipo_asignacionwf = data[j][31];
      var tipo_sancionwf = data[j][32];
      var motivo_devolucionwf = data[j][33];
      var giro_negociowf = data[j][34];
      var nro_trabajadoreswf = data[j][35];
      var tipo_metodizacionwf = data[j][36];
      var respuestas_checklistwf = data[j][37];
      var respustas_consultaswf = data[j][38];
      var notas_adicionaleswf = data[j][39];
      var fecha_desestimadowf = data[j][40];
      var monto_ventaswf = data[j][41];
      var ventaswf = data[j][42];
      var delegacionwf = data[j][43];
      var herramienta_sancionwf = data[j][44];
  
      if ( analista_asignadowf == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" &&  (tipo_sancionwf == "" && fecha_revisionwf == "" ) && devuelve_siomwf == ""  ) {
  
        var today = new Date();
        var dd = String(today.getDate()).padStart(2, '0');
        var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
        var yyyy = today.getFullYear();
  
        today = dd + '/' + mm + '/' + yyyy;
  
        fecha_revisionwf.setvalue(today)
  
      }
     
  
    }
  
    */

}