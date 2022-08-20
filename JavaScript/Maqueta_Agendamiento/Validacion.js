function validar_ingreso () {
  var correo_propietario = "luis.luna.cruz@bbva.com";
  
  var id = SpreadsheetApp.openById("1bYOA1-2k3-SnQYq970ImBzR-tyaJ4q6M0s6cQiyNfas");
  var ss = id.getSheetByName("INGRESAR SOLICITUD");
  var wf = id.getSheetByName("WF");
  var data = wf.getRange(1, 1, wf.getRange("A2").getDataRegion().getLastRow(),wf.getLastColumn()).getValues();

  var area_prime = ss.getRange("G8").getValue();
  var tema_prime = ss.getRange("G9").getValue();
  var responsable_prime = ss.getRange("G17").getValue();
  var tipo_prime = ss.getRange("G19").getValue(); 
  var perfil_prime = ss.getRange("G21").getValue();
  var oficina_prime = ss.getRange("G23").getValue();
  var fecha_prime = ss.getRange("G25").getValue();
  var hinicio_prime = ss.getRange("G27").getValue();
  var hinicio_ext = hinicio_prime.toString().toString().substring(0,2);
  var hinicio_number = Number(hinicio_ext)
  var hfin_prime = ss.getRange("J27").getValue();
  var hfin_ext = hfin_prime.toString().substring(0,2);
  var hfin_number = Number(hfin_ext)


  var horadif = hfin_number - hinicio_number;

  var fc = new Date(fecha_prime);
  var fc_prime = Utilities.formatDate(fc, "GMT-5", "dd-MM-yyyy");

  //Creacion de llaves validadoras
  


  if((horadif) > 1){

    var key_prime1 = fc_prime + (hinicio_number) + perfil_prime;
    var key_prime2 = fc_prime + (hinicio_number + 1) + perfil_prime;
    var key_prime3 = fc_prime + (hinicio_number + 2) + perfil_prime;
    var key_prime4 = fc_prime + (hinicio_number + 3) + perfil_prime;

 
  } else {

    var key_prime1 = fc_prime + (hinicio_number) + perfil_prime;

  }

    
  



  //Validacion de fechas existentes

  for (var i = 1; i < data.length; i++) {

    var areabase    = data[i][0]; 
    var temabase    = data[i][1]; 
    var tipobase = data[i][2]; 
    var perfilbase       = data[i][3];
    var oficinabase       = data[i][4];
    var fechabase    = data[i][5];   
    var hora_inicio    = data[i][6]; 
    var hora_inicibase = hora_inicio.toString().substring(0,2);
    var hora_fin        = data[i][7];   
    var hora_finbase= hora_fin.toString().substring(0,2);
    var id1 = data[i][8];
    var id2 = data[i][9];
    var id3 = data[i][10];
    var id4 = data[i][11];
     
    var fc2 = new Date(fecha);
    var fc_prime2 = Utilities.formatDate(fc2, "GMT-5", "yyyy-MM-dd' 'HH:mm:ss' '");


    if ((oficina_prime == oficina) && (fc_prime == fc_prime2) && (hinicio_prime == hora_inicioplus)) {
      
      pos6.setValue("SE CRUZAN LAS FECHAS"); 
      var v = 1;

    }

  }

  //-----------------------------------------
  /*
  var pos1 = ss.getRange("M5");
  var pos2 = ss.getRange("M6");
  var pos3 = ss.getRange("M7");
  var pos4 = ss.getRange("M8");
  var pos5 = ss.getRange("M9");

  var pos6 = ss.getRange("M20");
  var pos7 = ss.getRange("M12");
  var pos8 = ss.getRange("M13");
  var pos9 = ss.getRange("M14");
  var pos10 = ss.getRange("M15");
  var pos11 = ss.getRange("M16");

  */

  if (oficina_prime == "TODAS LAS OFICINAS") {

   var oficina_total = ["OF. BANCA INSTITUCIONAL",
    "BANCA CORPORATIVA LOCAL",
    "BE CENTRAL",
    "BE LIMA 1",
    "BE CHICLAYO",
    "BE NORTE CHICO",
    "BE MIRAFLORES",
    "BE PIURA",
    "BE REPÚBLICA DE PANAMÁ",
    "BE TRUJILLO",
    "BE AREQUIPA",
    "BE CHACARILLA",
    "BE SUR CHICO",
    "BE CUSCO",
    "BE LA MOLINA",
    "BE LIMA ORIENTE",
    "BE LIMA OESTE 2",
    "BE SANTA CRUZ",
    "BE LIMA OESTE 1",
    "BE DOS DE MAYO",
    "BE HUANCAYO",
    "BE LIMA CENTRO",
    "BE SAN ISIDRO"];

    for (var k = 1; k < oficina_total.length; k++){

      //IMPRESION DE OFICINAS MULTIPLES

      oficina_prime = oficina_total[k];

      var pos1 = ss.getRange(k+4,13);
      var pos2 = ss.getRange(k+4,14);
      var pos3 = ss.getRange(k+4,15);
      var pos4 = ss.getRange(k+4,16);
      var pos5 = ss.getRange(k+4,17);
      var pos6 = ss.getRange(k+4,18);
      var pos7 = ss.getRange(k+4,19);
      var pos8 = ss.getRange(k+4,20);
      var pos9 = ss.getRange(k+4,21);
      var pos10 = ss.getRange(k+4,22);
      var pos11 = ss.getRange(k+4,23);

      pos1.setValue(perfil_prime);
      pos2.setValue(oficina_prime);
      pos3.setValue(fc_prime);
      pos4.setValue(hinicio_number);
      pos5.setValue(hfin_number);

      var key_prime1 = "";
      var key_prime2 = "";
      var key_prime3 = "";
      var key_prime4 = "";
    

      if((horadif) > 1){

      var key_prime1 = fc_prime + (hinicio_number) + perfil_prime;
      var key_prime2 = fc_prime + (hinicio_number + 1) + perfil_prime;
      var key_prime3 = fc_prime + (hinicio_number + 2) + perfil_prime;
      var key_prime4 = fc_prime + (hinicio_number + 3) + perfil_prime;

      pos7.setValue(key_prime1)
      pos8.setValue(key_prime2)
      pos9.setValue(key_prime3)
      pos10.setValue(key_prime4)

      } else {

        var key_prime1 = fc_prime + (hinicio_number) + perfil_prime;

        pos7.setValue(key_prime1)

      }
    }
  } else {
  
    //IMPRESION DE OFICINA UNICA


    var pos1 = ss.getRange("M5");
    var pos2 = ss.getRange("N5");
    var pos3 = ss.getRange("O5");
    var pos4 = ss.getRange("P5");
    var pos5 = ss.getRange("Q5");
  
    var pos6 = ss.getRange("R5");
    var pos7 = ss.getRange("S5");
    var pos8 = ss.getRange("T5");
    var pos9 = ss.getRange("U5");
    var pos10 = ss.getRange("V5");
    var pos11 = ss.getRange("W5");

    pos1.setValue(perfil_prime);
    pos2.setValue(oficina_prime);
    pos3.setValue(fc_prime);
    pos4.setValue(hinicio_number);
    pos5.setValue(hfin_number);

    var key_prime1 = "";
    var key_prime2 = "";
    var key_prime3 = "";
    var key_prime4 = "";
    

    if((horadif) > 1){

    var key_prime1 = fc_prime + (hinicio_number) + perfil_prime;
    var key_prime2 = fc_prime + (hinicio_number + 1) + perfil_prime;
    var key_prime3 = fc_prime + (hinicio_number + 2) + perfil_prime;
    var key_prime4 = fc_prime + (hinicio_number + 3) + perfil_prime;

    pos7.setValue(key_prime1)
    pos8.setValue(key_prime2)
    pos9.setValue(key_prime3)
    pos10.setValue(key_prime4)

    } else {

      var key_prime1 = fc_prime + (hinicio_number) + perfil_prime;

      pos7.setValue(key_prime1)

    }
  }


  //Aca empieza el resgistro en Base
  
  /*
  var lastrow = wf.getLastRow()+1;

  var rn_area = wf.getRange(lastrow,1);
  var rn_tema = wf.getRange(lastrow,2);
  var rn_responsable = wf.getRange(lastrow,3);
  var rn_tipo = wf.getRange(lastrow,4);
  var rn_perfil = wf.getRange(lastrow,5);
  var rn_oficina = wf.getRange(lastrow,6);
  var rn_fc_pprime = wf.getRange(lastrow,7);
  var rn_hincio = wf.getRange(lastrow,8);
  var rn_hfin = wf.getRange(lastrow,9);
  var rn_id1 = wf.getRange(lastrow,10);
  var rn_id2 = wf.getRange(lastrow,11);
  var rn_id3 = wf.getRange(lastrow,12);
  var rn_id4 = wf.getRange(lastrow,13);

  rn_area.setValue(area_prime); 
  rn_tema.setValue(tema_prime); 
  rn_responsable.setValue(responsable_prime); 
  rn_tipo.setValue(tipo_prime); 
  rn_perfil.setValue(perfil_prime); 
  rn_oficina.setValue(oficina_prime); 
  rn_fc_pprime.setValue(fc_prime); 
  rn_hincio.setValue(hinicio_prime); 
  rn_hfin.setValue(hfin_prime); 

  rn_id1.setValue(key_prime1); 
  rn_id2.setValue(key_prime2); 
  rn_id3.setValue(key_prime3); 
  rn_id4.setValue(key_prime4); 

    
  var a = 2;
    
  
  for (var i = 1; i < data.length; i++) {

    var areabase    = data[i][0]; 
    var temabase    = data[i][1]; 
    var tipobase = data[i][2]; 
    var perfilbase       = data[i][3];
    var oficinabase       = data[i][4];
    var fechabase    = data[i][5];   
    var hora_inicio    = data[i][6]; 
    var hora_inicibase = hora_inicio.toString().substring(0,2);
    var hora_fin        = data[i][7];   
    var hora_finbase= hora_fin.toString().substring(0,2);
    var id1 = data[i][8];
    var id2 = data[i][9];
    var id3 = data[i][10];
    var id4 = data[i][11];
     
    var fc2 = new Date(fecha);
    var fc_prime2 = Utilities.formatDate(fc2, "GMT-5", "yyyy-MM-dd' 'HH:mm:ss' '");


    if ((oficina_prime == oficina) && (fc_prime == fc_prime2) && (hinicio_prime == hora_inicioplus)) {
      
      pos6.setValue("SE CRUZAN LAS FECHAS"); 
      var v = 1;

    }

  }  

  */

  
  
}



  

  