function seleccion_masiva () {

  var correo_propietario = "luis.luna.cruz@bbva.com";

  //Definicion de la hoja

  var id = SpreadsheetApp.getActiveSpreadsheet();
  var sheetWF = id.getSheetByName("WF");
  var sheetForm = id.getSheetByName("Ingreso Metodización")
  //var sheetWF = id.getSheetByName("WF");
  var sheetEO = id.getSheetByName("MN");
  var finicioWF = 10 //Inicio de la BD del WF
  var lRowWF = sheetWF.getLastRow();
  var lColumnWF = sheetWF.getLastColumn();
  var lRowEO = sheetEO.getLastRow();
  var lColumnEO = sheetEO.getLastColumn();
  var lRowForm = sheetForm.getLastRow();
  var lColumnForm = sheetForm.getLastColumn();
  var cCodSol = 1;
  var cCheckList = 22;
  var c1erCheckList = 23;
  var cDevuelveCheckList = 24;
  var cArchivoDevuelve = 25
  var cRev = 26;
  var cCBC = 27;
  var cDev = 28;
  var cNotasAdicionales = 40;
  var cDesestimado = 41;
  var notasAdicionales = "LISTA PRIORIZADA";

  var fechaHoy = new Date();
  fechaHoy.setHours(0,0,0,0);
  var fecha = fechaHoy; //Obtiene fecha

  //var datawf = ss.getRange(2, 1, ss.getRange("A10").getDataRegion().getLastRow(), ss.getLastColumn()).getValues();
  
  var sscc = id.getSheetByName("MAQUETA LOURDES");
  var data = sscc.getRange(2, 1, sscc.getRange("A10").getDataRegion().getLastRow(), sscc.getLastColumn()).getValues();


  //Lectura de la base de Lourdes

  for (var i = 0; i < data.length; i++) {

    var peticion    = data[i][0]; 
    var cod_central    = data[i][1];  
    var cliente    = data[i][2];
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
    var correo_analistav2 = data[i][31];

    

    //var fc = new Date(marca_temporal);
    //var text_fc = Utilities.formatDate(fc, "GMT-5", "yyyy-MM-dd' 'HH:mm:ss' '");
    //var marca_temporal = text_fc.substring(8, 10) +'/'+ text_fc.substring(5, 7) +'/'+ 
                          //text_fc.substring(0, 4) +' '+ text_fc.substring(11, 19);


    /*const conditionsArray = [
        analista_asignado == "LOURDES DEL PILAR GUTIERREZ MANCHEGO", 
        tipo_sancion == ""
    ]*/

    if ( analista_asignado == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" &&  cliente_priorizado == "SÍ") {


      //Loop que encuentra el código de solicitud

      var arrayWF = sheetWF.getRange(finicioWF-1,1,lRowWF - finicioWF+2,lColumnWF).getValues() //lRowWF - fInicioWF +1

      for(var j = 1; j <= lRowWF - finicioWF +1; j++){ //Acá empieza en 1 porque se incluyeron los Headings
        var codSolWF = arrayWF[j][cCodSol-1];
        if(codSolWF === peticion){           //Compara códigos de solicitud
          var fEncontrada = j + finicioWF -1  //Guarda el valor de la fila del código de solicitud encontrado dentro de la base WF.
        }
      }
   
      var recipient = correo_gestorv2

      var subject = "SIOM: La Peticion de Metodizacion con Codigo " + peticion + " (" + cliente + ") Ha Sido Aceptada"
      var body = "Segmento Empresas ha aceptado la peticion con codigo: " + peticion
      + "\n"
      + "\nInformacion general sobre la operacion: "
      + "\nCodigo Central: " + cod_central
      + "\nGestor: " + nombre_gestor
      + "\nCliente: " + cliente
      + "\nGrupo : " + grupo 
      + "\nTipo de Estado: " + tipo_estado
      +  "\nAnalista Asignado: " + analista_asignado
      + "\n\nDetalle de Balance General: " + detalle_balance
      + "\n\nDetalle de Estado de Resultados: " + detalle_resultados
      + "\nArchivo Adjunto: " + archivos
      + "\n"
      + "\n#CampañaEEFF2022"
      + "\n#IMPLEMENTACIÓNBEC"
      + "\n\n\nPor el momento, no requiere realizar ninguna accion adicional en cuanto a esta peticion."


      var options = {cc: correo_ingresantev2 + ",lourdes.gutierrez@bbva.com, wendy.amaya.tobon@bbva.com, siom@bbva.com"}
      MailApp.sendEmail(recipient,subject,body,options);

      sheetWF.getRange(fEncontrada,cRev).setValue(fecha) //Escribe la fecha.

       
    } else if ( analista_asignado == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" &&  cliente_priorizado == "NO") {


      //Loop que encuentra el código de solicitud

      var arrayWF = sheetWF.getRange(finicioWF-1,1,lRowWF - finicioWF+2,lColumnWF).getValues() //lRowWF - fInicioWF +1

      for(var j = 1; j <= lRowWF - finicioWF +1; j++){ //Acá empieza en 1 porque se incluyeron los Headings
        var codSolWF = arrayWF[j][cCodSol-1];
        if(codSolWF === peticion){           //Compara códigos de solicitud
          var fEncontrada = j + finicioWF -1  //Guarda el valor de la fila del código de solicitud encontrado dentro de la base WF.
        }
      }

      var recipient = correo_gestorv2

      var subject = "SIOM: La Peticion de Metodizacion con Codigo " + peticion + " (" + cliente + ") Ha Sido Desestimada"
      var body = "Segmento Empresas ha desestimado la peticion con codigo: " + peticion
      + "\n"
      + "\nInformacion general sobre la operacion: "
      + "\n"
      + "\nCodigo Central: " + cod_central
      + "\nGestor: " + nombre_gestor
      + "\nCliente: " + cliente
      + "\nGrupo : " + grupo 
      + "\nTipo de Estado: " + tipo_estado
      +  "\nAnalista Asignado: " + analista_asignado
      + "\n\nDetalle de Balance General: " + detalle_balance
      + "\n\nDetalle de Estado de Resultados: " + detalle_resultados
      + "\nArchivo Adjunto: " + archivos
      + "\n"
      + "\nPor el momento, no se estan aceptando peticiones que se encuentren fuera de la lista de Priorización de EEFF."
      + "\n"
      + "\n#CampañaEEFF2022"
      + "\n#IMPLEMENTACIÓNBEC"
     

      var options = {cc: correo_ingresantev2 + ",lourdes.gutierrez@bbva.com, wendy.amaya.tobon@bbva.com, siom@bbva.com"}
      MailApp.sendEmail(recipient,subject,body,options);

      sheetWF.getRange(fEncontrada,cCheckList).setValue("");
      sheetWF.getRange(fEncontrada,c1erCheckList).setValue(fecha);
      sheetWF.getRange(fEncontrada,cDevuelveCheckList).setValue(fecha)
      //sheetWF.getRange(fEncontrada,cArchivoDevuelve).setValue("N/A")
      //sheetWF.getRange(fEncontrada,cRev).setValue(fecha)
      //sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
      //sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
  
      sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
      sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
      
    }

  }


}