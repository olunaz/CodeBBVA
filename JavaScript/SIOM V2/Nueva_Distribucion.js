function lectura_esctritura_LG () {
    var correo_propietario = "luis.luna.cruz@bbva.com";
  
    var id = SpreadsheetApp.openById("154qpAC9tk2WzbIaBPf-Cn2sCFqssExqrdhh1KObFDk4");
    var ss = id.getSheetByName("WF");
    var lg = id.getSheetByName("MAQUETA LOURDES");
    var cc = id.getSheetByName("MAQUETA CRISTOPHER");

    var id2 = SpreadsheetApp.openById("154qpAC9tk2WzbIaBPf-Cn2sCFqssExqrdhh1KObFDk4");
    var ss2 = id.getSheetByName("MAQUETA LOURDES");
    var data = ss.getRange(2, 1, ss.getRange("A10").getDataRegion().getLastRow(), ss.getLastColumn()).getValues();

    var op_meto1 = lg.getRange(10, 1, lg.getRange("A10").getDataRegion().getLastRow(), 28);
    var op_meto2 = cc.getRange(2, 1, cc.getRange("A10").getDataRegion().getLastRow(), cc.getLastColumn());

    op_meto1.clearContent();


    var a = 10;

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
      var archivo = data[i][15];
      var cod_documento = data[i][16];
      var tipo_estado = data[i][17];
      var detalle_balance = data[i][18];
      var detallle_resultados = data[i][19];
      var analista_siom = data[i][20];
      var fecha_checklist = data[i][21];
      var primer_ingreso_checklist = data[i][22];
      var devuelve_siom = data [i][23];
      var fecha_revision = data[i][25];
      var fecha_consulta_be = data[i][26];
      var fecha_devolucion = data[i][27];
      var fecha_cierre = data[i][30];
      var tipo_asignacion = data[i][31];
      var tipo_sancion = data[i][32];
      var motivo_devolucion = data[i][33];
      var giro_negocio = data[i][34];
      var nro_trabajadores = data[i][35];
      var tipo_metodizacion = data[i][36];
      var respuestas_checklist = data[i][37];
      var respustas_consultas = data[i][38];
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


      //const conditionsArray = [
          analista_siom == "LOURDES DEL PILAR GUTIERREZ MANCHEGO", 
          tipo_sancion == ""
      //]

      
      if ( analista_siom == "LOURDES DEL PILAR GUTIERREZ MANCHEGO" &&  (tipo_sancion == "" && fecha_revision == "" ) && devuelve_siom == ""  ) {
       
        var rpeticion = ss2.getRange(a,1);
        var rcod_central = ss2.getRange(a,2);
        var rcliente = ss2.getRange(a,3);
        var rgrupo = ss2.getRange(a,4);
        var restado = ss2.getRange(a,5);
        var rdetalle_balance = ss2.getRange(a,6);
        var rdetalle_resultados = ss2.getRange(a,7);
        var rarchivo = ss2.getRange(a,8);
        var rnro_trabajadores = ss2.getRange(a,9);
        var rnombre_oficina = ss2.getRange(a,10);
        var rcod_gestor = ss2.getRange(a,11);
        var rnombre_gestor = ss2.getRange(a,12);
        var rcorreo_gestor = ss2.getRange(a,13);
        var rcorreo_ingresante = ss2.getRange(a,14);
        var ranalista_siom = ss2.getRange(a,15);
        var rcorreoanalista = ss2.getRange(a,16);
        var rgiro_negocio = ss2.getRange(a,17);
        var rtipo_metodizacion = ss2.getRange(a,18);
        var rmonto_ventas = ss2.getRange(a,19);
        var rfecha_base  = ss2.getRange(a,20);
        var rfecha_checklist = ss2.getRange(a,21);




        rpeticion.setValue(peticion);
        rcod_central.setValue(cod_central);
        rcliente.setValue(cliente);
        rgrupo.setValue(grupo);
        restado.setValue(tipo_estado);
        rdetalle_balance.setValue(detalle_balance);
        rdetalle_resultados.setValue(detallle_resultados);
        rarchivo.setValue(archivo);
        rnro_trabajadores.setValue(nro_trabajadores);
        rnombre_oficina.setValue(nombre_oficina);
        rcod_gestor.setValue(cod_gestor);
        rnombre_gestor.setValue(nombre_gestor);
        ranalista_siom.setValue(analista_siom);
        rgiro_negocio.setValue(giro_negocio);
        rtipo_metodizacion.setValue(tipo_metodizacion);
        rmonto_ventas.setValue(monto_ventas);
        rfecha_base.setValue(fecha_base);
        rfecha_checklist.setValue(fecha_checklist);
                       
        
        a = a + 1;
      
      }


  
  }
  

}



function lectura_esctritura_CC () {
    var correo_propietario = "luis.luna.cruz@bbva.com";
  
    var id = SpreadsheetApp.openById("154qpAC9tk2WzbIaBPf-Cn2sCFqssExqrdhh1KObFDk4");
    var ss = id.getSheetByName("WF");
    var lg = id.getSheetByName("MAQUETA LOURDES");
    var cc = id.getSheetByName("MAQUETA CRISTOPHER");

    var id2 = SpreadsheetApp.openById("154qpAC9tk2WzbIaBPf-Cn2sCFqssExqrdhh1KObFDk4");
    var ss2 = id.getSheetByName("MAQUETA CRISTOPHER");
    var data = ss.getRange(2, 1, ss.getRange("A10").getDataRegion().getLastRow(), ss.getLastColumn()).getValues();

    var op_meto2 = cc.getRange(10, 1, cc.getRange("A10").getDataRegion().getLastRow(),28);

    op_meto2.clearContent();

   
    var a = 10;

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
      var archivo = data[i][15];
      var cod_documento = data[i][16];
      var tipo_estado = data[i][17];
      var detalle_balance = data[i][18];
      var detallle_resultados = data[i][19];
      var analista_siom = data[i][20];
      var fecha_checklist = data[i][21];
      var primer_ingreso_checklist = data[i][22];
      var devuelve_siom = data [i][23];
      var fecha_revision = data[i][25];
      var fecha_consulta_be = data[i][26];
      var fecha_devolucion = data[i][27];
      var fecha_cierre = data[i][30];
      var tipo_asignacion = data[i][31];
      var tipo_sancion = data[i][32];
      var motivo_devolucion = data[i][33];
      var giro_negocio = data[i][34];
      var nro_trabajadores = data[i][35];
      var tipo_metodizacion = data[i][36];
      var respuestas_checklist = data[i][37];
      var respustas_consultas = data[i][38];
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


      //const conditionsArray = [
          analista_siom == "CRISTOPHER CERRON BAQUERIZO", 
          tipo_sancion == ""
      //]

      

      if ( analista_siom == "CRISTOPHER CERRON BAQUERIZO" &&  (tipo_sancion == "" && fecha_revision == "" ) && devuelve_siom == ""  ) {
       
        var rpeticion = ss2.getRange(a,1);
        var rcod_central = ss2.getRange(a,2);
        var rcliente = ss2.getRange(a,3);
        var rgrupo = ss2.getRange(a,4);
        var restado = ss2.getRange(a,5);
        var rdetalle_balance = ss2.getRange(a,6);
        var rdetalle_resultados = ss2.getRange(a,7);
        var rarchivo = ss2.getRange(a,8);
        var rnro_trabajadores = ss2.getRange(a,9);
        var rnombre_oficina = ss2.getRange(a,10);
        var rcod_gestor = ss2.getRange(a,11);
        var rnombre_gestor = ss2.getRange(a,12);
        var rcorreo_gestor = ss2.getRange(a,13);
        var rcorreo_ingresante = ss2.getRange(a,14);
        var ranalista_siom = ss2.getRange(a,15);
        var rcorreoanalista = ss2.getRange(a,16);
        var rgiro_negocio = ss2.getRange(a,17);
        var rtipo_metodizacion = ss2.getRange(a,18);
        var rmonto_ventas = ss2.getRange(a,19);
        var rfecha_base  = ss2.getRange(a,20);
        var rfecha_checklist = ss2.getRange(a,21);




        rpeticion.setValue(peticion);
        rcod_central.setValue(cod_central);
        rcliente.setValue(cliente);
        rgrupo.setValue(grupo);
        restado.setValue(tipo_estado);
        rdetalle_balance.setValue(detalle_balance);
        rdetalle_resultados.setValue(detallle_resultados);
        rarchivo.setValue(archivo);
        rnro_trabajadores.setValue(nro_trabajadores);
        rnombre_oficina.setValue(nombre_oficina);
        rcod_gestor.setValue(cod_gestor);
        rnombre_gestor.setValue(nombre_gestor);
        ranalista_siom.setValue(analista_siom);
        rgiro_negocio.setValue(giro_negocio);
        rtipo_metodizacion.setValue(tipo_metodizacion);
        rmonto_ventas.setValue(monto_ventas);
        rfecha_base.setValue(fecha_base);
        rfecha_checklist.setValue(fecha_checklist);
                      
        
        a = a + 1;
      
      }


  
  }
  

}