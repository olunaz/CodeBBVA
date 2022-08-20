function lectura_esctritura () {
    var correo_propietario = "luis.luna.cruz@bbva.com";
  
    var id = SpreadsheetApp.openById("1-3drbuD2egCuGu_M-U7tdgKXLHd-dsMzutwPvlZNXaU");
    var ss = id.getSheetByName("WF");
    var id2 = SpreadsheetApp.openById("1-3drbuD2egCuGu_M-U7tdgKXLHd-dsMzutwPvlZNXaU");
    var ss2 = id.getSheetByName("BASE PRUEBA");
    var data = ss.getRange(1, 1, ss.getRange("A10").getDataRegion().getLastRow(), ss.getLastColumn()).getValues();
    
    var a = 21;

    for (var i = 0; i < data.length; i++) {
  
      
      var peticion    = data[i][0]; 
      var cod_central    = data[i][1];   
      var ruc    = data[i][2];  
      var tipo_documento        = data[i][3];   
      var cod_oficina       = data[i][4];
      var nombre_oficina = data[i][5];
      var cod_gestor = data[i][6];
      var nombre_gestor = data[i][7];
      var procedencia = data[i][8];
      var fecha_base = data[i][9];
      var fecha_final = data[i][10];
      var cod_territorio = data[i][11];
      var nombre_territorio = data[i][12];
      var cliente = data[i][13];
      var grupo = data[i][14];
      var arhivo = data[i][15];
      var cod_documento = data[i][16];
      var tipo_estado = data[i][17];
      var detalle_balance = data[i][18];
      var detallle_resultados = data[i][19];
      var analista_siom = data[i][20];
      var fecha_checklist = data[i][21];
      var fecha_revision = data[i][25];
      var fecha_consulta_be = data[i][26];
      var fecha_devolucion = data[i][27];
      var fecha_cierre = data[i][30];
      var tipo_asignacion = data[i][31];
      var tipo_sancion = data[i][32];
      var giro_negocio = data[i][34];
      var nro_trabajadores = data[i][35];
      var tipo_metodizacion = data[i][36];
      var notas_adicionales = data[i][39];
      var fecha_desestimado = data[i][40];
      var monto_ventas = data[i][41];
      var ventas = data[i][42];
      var delegacion = data[i][43];
      var herramienta_sancion = data[i][44];
    
      //var fc = new Date(marca_temporal);
      //var text_fc = Utilities.formatDate(fc, "GMT-5", "yyyy-MM-dd' 'HH:mm:ss' '");
      //var marca_temporal = text_fc.substring(8, 10) +'/'+ text_fc.substring(5, 7) +'/'+ 
                            //text_fc.substring(0, 4) +' '+ text_fc.substring(11, 19);


      const conditionsArray = [
          analista_siom == "LOURDES DEL PILAR GUTIERREZ MANCHEGO", 
          tipo_sancion == ""
      ]

      if ( analista_siom == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" &&  tipo_sancion == "" && fecha_revision == "") {
       
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
          
      }                      
  
    }
  

}
  

  