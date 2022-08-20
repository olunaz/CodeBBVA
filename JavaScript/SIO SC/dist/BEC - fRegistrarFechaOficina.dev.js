"use strict";

//¡¡¡¡¡¡¡SE HAN AGREGADO LOS CORREOS DE LEASING PARA OPERACIONES LP!!!!!!!!!
function ejecRegistrarFechaRiesgos(visor) {
  //Registra cambios en fechas
  //Compara los headings de la sheetEO con los de la base WF. Ambos tienen que ser IDÉNTICOS para que funcione. 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEO = ss.getSheetByName(visor);

  if (visor === "ER1" || visor === "ER2" || visor === "ER3" || visor === "ER4" || visor === "ER5") {} else {
    return;
  }

  var finicioEO = 14; //Inicio de la revisión de operaciones

  var finicioWF = 10; //Inicio de la BD del WF

  var lRowEO = sheetEO.getLastRow();
  var lColumnEO = sheetEO.getLastColumn();
  var cDatosE1 = 2;
  var cHeadingsE1 = 1;
  var cDatosE2 = 4;
  var cHeadingsE2 = 3;
  var cDatosE3 = 6;
  var cHeadingsE3 = 5;
  var fHeadingsWF = 9;
  var valcellcodsol = "H8";
  var rNDEV = "B99";
  var valcellEstEst = "B98";
  var verCodSol = sheetEO.getRange("C3").getValue();
  var fInicioBC = 2;
  var datosCodSol = sheetEO.getRange("B1:B10").getValues();
  var CodSol = datosCodSol[0][0].toUpperCase();
  var CodCentral = datosCodSol[1][0];
  var Ejecutivo = datosCodSol[2][0];
  var Cliente = datosCodSol[3][0];
  var Grupo = datosCodSol[4][0];
  var tipoOp = datosCodSol[5][0];
  var Oper = datosCodSol[5][0];
  var MontoProp = datosCodSol[6][0];
  var Correo = datosCodSol[8][0]; //  var celdaCodSol = sheetEO.getRange("B1")
  //  var CodSol = celdaCodSol.getValue().toUpperCase()
  //  var CodCentral = sheetEO.getRange("B2").getValue()
  //  var Ejecutivo = sheetEO.getRange("B3").getValue()
  //  var Cliente = sheetEO.getRange("B4").getValue()
  //  var Grupo = sheetEO.getRange("B5").getValue()
  //  var tipoOp = sheetEO.getRange("B6").getValue()
  //  var Oper = sheetEO.getRange("B6").getValue()
  //  var MontoProp = sheetEO.getRange("B7").getValue()
  //  var Correo = sheetEO.getRange("B9").getValue()

  var condicMitig = "";
  var codGE = sheetEO.getRange("C2").getValue();
  var cNuevaFV = 10;
  var fInicioPF = 3;
  var cCodSolPF = 43;
  var strOfReg = "OFICINA: REGISTRE FECHA";
  var strRiReg = "RIESGOS: CONFIRMAR O DEVOLVER";
  var strRiReg2 = "RIESGOS: REGISTRE FECHA";
  var strRiReg3 = "RIESGOS: CONFIRMAR O CONSULTAR"; // Columnas del Drive

  var cCodSol = 1;
  var cCodCentral = 2;
  var cEjecutivo = 3;
  var cCliente = 4;
  var cGrupo = 5;
  var cOperacion = 6;
  var cCheckList = 7;
  var cGeneracion1erRVGL = 8;
  var c1erController = 9;
  var cDevuelveController = 10;
  var cIngresoController = 11;
  var cAsignacionRVGL = 12;
  var cCRC = 13;
  var cCBC = 14;
  var cDev = 15;
  var cRcR1 = 16;
  var c2daCRC = 17;
  var c2daCBC = 18;
  var c2daDev = 19;
  var cRcR2 = 20;
  var cIAn = 21;
  var cFEv = 22;
  var cFS = 23;
  var cAmb = 24;
  var cTS = 25;
  var cAsEv = 26;
  var c3raCRC = 27;
  var c3raCBC = 28;
  var c3raDev = 29;
  var cRcR3 = 30;
  var cFev2 = 31;
  var cFS2 = 32;
  var cAmb2 = 33;
  var cTS2 = 34;
  var cMontSol = 35;
  var cMontSan = 36;
  var cAnAsig = 37;
  var cCas = 38;
  var cMotDev = 39;
  var cProducto = 40;
  var cCondicion = 41;
  var cRating = 42;
  var cEEFF = 43;
  var cHerramienta = 44;
  var cBuro = 45;
  var cMitig = 46;
  var cMarcaPuntual = 51;
  var cLink = 52;
  var cTipodePF = 53;
  var cRT = 54;
  var cBuroBase = 55;
  var cRatingBase = 56;
  var cEstMes = 57;
  var cEstMesAnterior = 58;
  var cTop = 59;
  var cPerfil = 60;
  var cEstratSanc = 61;
  var cCodRelacionado = 62;
  var cTipoCliente = 64;
  var cAmbSancTrans = 65; // Columnas del Drive

  var fechaHoy = new Date();
  fechaHoy.setHours(0, 0, 0, 0);
  var fecha = fechaHoy; //Obtiene fecha

  /*  var date = new Date();
    //date.setHours(date.getHours() + 6)
    var hour = date.getHours();
    if(hour >= 17 && date.getDay() >= 1 && date.getDay() <= 5){
      if (date.getDay() === 5){
        fecha.setDate(fecha.getDate() + 3);
      }
      else{
        fecha.setDate(fecha.getDate() + 1);
      }
    }
    else if (date.getDay() === 6){
      fecha.setDate(fecha.getDate() + 2);
    }
    else if (date.getDay() === 7){
      fecha.setDate(fecha.getDate() + 1);
    }*/

  if (CodSol === "") {
    return;
  }

  if (CodSol != verCodSol) {
    Browser.msgBox("¡Alerta!", "La fecha NO ha sido registrada debido a que la revisión ha expirado. Actualice las estaciones nuevamente.", Browser.Buttons.OK);
    ss.toast("Antes de registrar una fecha, por favor presione el botón 'Actualizar Estaciones'.", "Recordatorio", 10);
    return;
  }

  var id = CodSol.substring(0, 2).toUpperCase();
  var cEstEst = sheetEO.getRange(valcellEstEst);
  var valEstEst = cEstEst.getValue();
  var ofic = CodSol.substring(0, 2);

  if (ofic === "MR") {
    var ssOfic = SpreadsheetApp.openById("1QWGnqKN3XHMn1OKhbM0RISH9ea4enfJCldPP45umSvY");
  } //TEST: 1yRDf3zTFkXBMfdjMbbfEUYSuy_uMKRGy21yU2nvjPvE
  //PROD: 1QWGnqKN3XHMn1OKhbM0RISH9ea4enfJCldPP45umSvY
  else if (ofic === "RP") {
      var ssOfic = SpreadsheetApp.openById("1IzZYXPPurDd9CO4tFmavwucgwzNDJbHk7kc0XWNaGsw");
    } else if (ofic === "LE") {
      var ssOfic = SpreadsheetApp.openById("1trTnjyg4YtgAsdSUdXBehlrYAOQCr8fA46alhRTtF-0");
    } else if (ofic === "CY") {
      var ssOfic = SpreadsheetApp.openById("1ZmoIc_hSBKdSxVprJA434IKxVuD8QFEZmXy3pc6Oj_8");
    } else if (ofic === "TR") {
      var ssOfic = SpreadsheetApp.openById("1gGn9lnFWwgzhT7CLEi8o9d9pD6mwUWVuYSG0q74fCLg");
    } else if (ofic === "PR") {
      var ssOfic = SpreadsheetApp.openById("1cMWBO-dkHqXrNFblAvHCEaRCrDMbFWNu4n8R9RhT3OQ");
    } else if (ofic === "SC") {
      var ssOfic = SpreadsheetApp.openById("1BgJ5KoPjJlVrLVJfdYbbjmXDqF4B5n6qlZ8yj4XZCyo");
    } else if (ofic === "CL") {
      var ssOfic = SpreadsheetApp.openById("14VHSYaQGmHbpNVM7hJRwhcMbiICBQkv5EL7o6rA8db8");
    } else if (ofic === "LM") {
      var ssOfic = SpreadsheetApp.openById("1aEqDmDy5RpLqCGm8w9pGfkybfS7lk1iqzVQ4KB_RYPE");
    } else if (ofic === "AR") {
      var ssOfic = SpreadsheetApp.openById("17YKpkqo5xNqf6OAfZM3Cz5IUVaeb3_MygvKGMNACFVY");
    } else if (ofic === "CU") {
      var ssOfic = SpreadsheetApp.openById("1blF00dZieGN8r93bvVxd1hnJP6aeF4jKVGzRTnz0QtI");
    } else if (ofic === "CN") {
      var ssOfic = SpreadsheetApp.openById("1NlUCAwOLC2azNLEzuzDgxro89tBb_MEhuNqEbJl2q_Q");
    } else if (ofic === "SD") {
      var ssOfic = SpreadsheetApp.openById("1dgiYDX1k6Bi83ScQBrYgM3rnlRdh_qr-k7E4ZwmEweg");
    } else if (ofic === "DM") {
      var ssOfic = SpreadsheetApp.openById("1yMxgwU95F6aSY7MAEgt0kFwesGXXrEZ2OV-c2Uahack");
    } else if (ofic === "LA") {
      var ssOfic = SpreadsheetApp.openById("1lv0Rjvpv-U9HXHq9ZuzibtFCCgUUdr6HFrHwIxuhq_I");
    } else if (ofic === "CA") {
      var ssOfic = SpreadsheetApp.openById("1iZmMBmOI_g0H22Orfrph_L7pxNu7UpSVaVqpfvquNuo");
    } else if (ofic === "CC") {
      var ssOfic = SpreadsheetApp.openById("1YvtvCRSYCzAKBEWR_BglKD75q680c0uNCIgpdRdCLyg");
    } else if (ofic === "QQ") {
      var ssOfic = SpreadsheetApp.openById("1HwWn4XML-o0vtc8pj4YUDhMgcskEMZ6_hJ6nHHsT67s");
    } else if (ofic === "HU") {
      var ssOfic = SpreadsheetApp.openById("1YOUajnXjUhmrbUMLUuDua6Jiigpa9DIYV-h7Asb8gnY");
    } else if (ofic === "TP") {
      var ssOfic = SpreadsheetApp.openById("1svbDX5R9Xn0X2PHF5gGhv_oT5E9BsyldCXlElfSF4KQ");
    } else if (ofic === "BC") {
      var ssOfic = SpreadsheetApp.openById("1DodVsAiUnWB-FlLO6MPzJ4rL9IflTc_QQgxn9lmEZLM");
    } else if (ofic === "BN") {
      var ssOfic = SpreadsheetApp.openById("1hA4j5y-XcvbnRWkwWfghWpBgA3xDhiGmPEFIQsCzDjo");
    } else if (ofic === "LL") {
      var ssOfic = SpreadsheetApp.openById("1rJFKRxR6vtwWxPP1trXQol9x6cC0FYsknHwOobOmXSY");
    } else if (ofic === "LK") {
      var ssOfic = SpreadsheetApp.openById("1y03Z9ko33533fHvMXZ58MqmlgaeETUso-QOIyNsRjs0");
    } else if (ofic === "ZT") {
      var ssOfic = SpreadsheetApp.openById("1yRDf3zTFkXBMfdjMbbfEUYSuy_uMKRGy21yU2nvjPvE");
    } else {
      Browser.msgBox("Código no válido.");
      return;
    }

  if (ofic === "MR") {
    var ssL = SpreadsheetApp.openById("10b_4NbsI-ww5S9-DAs4V41p2BqX9c6ZeHeEFowLHpRM");
  } //TEST: 1hIJ-lbhJdVQY8kl7B1EyEB5E6uIolOffPxbxOWvo07M
  //PROD: 10b_4NbsI-ww5S9-DAs4V41p2BqX9c6ZeHeEFowLHpRM
  else if (ofic === "RP") {
      var ssL = SpreadsheetApp.openById("1t9Exs5PTwRXqwVoP2A56Me0ffwkqyVsWMsXCQoCLh1k");
    } else if (ofic === "LE") {
      var ssL = SpreadsheetApp.openById("1QgOLgln4PEhjGn0cK-umAmswgZInPsIvN-aWlxzn32A");
    } else if (ofic === "CY") {
      var ssL = SpreadsheetApp.openById("1zMdHwrtrJ1d27y8weQ0e_hxIfasi-x6oTzaKJMOyHrM");
    } else if (ofic === "TR") {
      var ssL = SpreadsheetApp.openById("1Uyyhr-GhXtZF8ZUyxBu8M3qSn9OzfhK_-G74PHh51Z0");
    } else if (ofic === "PR") {
      var ssL = SpreadsheetApp.openById("1XZ8MmKHJ9tzAHXSvfhledE-d3Rnu5C-8yz0RADaRoxs");
    } else if (ofic === "SC") {
      var ssL = SpreadsheetApp.openById("1UhwqfoVSQdIWRa6xcs6K4q5bMUTKJXcCYxokA5nEf0k");
    } else if (ofic === "CL") {
      var ssL = SpreadsheetApp.openById("1F4AP2XEUR_2qIYKJ32izqd36-vLITM3eyWtQZPyIVY8");
    } else if (ofic === "LM") {
      var ssL = SpreadsheetApp.openById("1oNyNwefKWbHGPp7AD-OOf3cXP624vVUecXrEHJzLIoU");
    } else if (ofic === "AR") {
      var ssL = SpreadsheetApp.openById("1B4tdZIfl6XhvuNOwAKsgJxu1kXxJYV4rYkVwM3-QvyI");
    } else if (ofic === "CU") {
      var ssL = SpreadsheetApp.openById("1plOAymsMgkL3_d_mU7ONYx08a2gHFVU9yvUkLWwQKto");
    } else if (ofic === "CN") {
      var ssL = SpreadsheetApp.openById("1DG3dwbeOShflYecazX6E6XFVDeYKtHpW5CdRHbPWyww");
    } else if (ofic === "SD") {
      var ssL = SpreadsheetApp.openById("17LceV-VXt9jFUBypQxLrV4WY-7VBGyAKxSGUcyfKoPQ");
    } else if (ofic === "DM") {
      var ssL = SpreadsheetApp.openById("12yzFhT22NNcTo8zbLLYpFU6v8IzCOVgrvmEawauS0pw");
    } else if (ofic === "LA") {
      var ssL = SpreadsheetApp.openById("1h_frotn83ZdDGIvm9cqocIboHhlYoaWkZN7HFlbqJUI");
    } else if (ofic === "CA") {
      var ssL = SpreadsheetApp.openById("1mpuhHCuzWjDHJQqUeFXnH3nud9Qbxr1-8HPNzHAlzaQ");
    } else if (ofic === "CC") {
      var ssL = SpreadsheetApp.openById("1hnyyr0kL8VDcHigeV8e3N94pD-OfmdwZma0paVpzpxs");
    } else if (ofic === "QQ") {
      var ssL = SpreadsheetApp.openById("1Sz_t6dFqeaVvzbZ2Mn2mKMqDJfu7j3GIn02LOiPpKPU");
    } else if (ofic === "HU") {
      var ssL = SpreadsheetApp.openById("1wrXlELjDZJxkU8DnyjM99W-ED7adTILBuC5e4-dmmhU");
    } else if (ofic === "TP") {
      var ssL = SpreadsheetApp.openById("1EE3j6o1qXu2_yySxolXTyF5oCr-7D0a2BzSZhygzUe0");
    } else if (ofic === "BC") {
      var ssL = SpreadsheetApp.openById("1shRa6Yji_wItZwOJz8msw2ZK4bzTfRNKLOsSHtx_g10");
    } else if (ofic === "BN") {
      var ssL = SpreadsheetApp.openById("1OYfbP32WOxJoeHC6TbnH2dqXXBO55D77F66DvMCG13I");
    } else if (ofic === "LL") {
      var ssL = SpreadsheetApp.openById("1sV-jqOblYxc3eFcDHAeAdzlkuzVUjGdg6jHgi3d706I");
    } else if (ofic === "LK") {
      var ssL = SpreadsheetApp.openById("1594XbzW0FYhJskHlpv1Rs4VinlW0ATyOuywIigODFfw");
    } else if (ofic === "ZT") {
      var ssL = SpreadsheetApp.openById("1hIJ-lbhJdVQY8kl7B1EyEB5E6uIolOffPxbxOWvo07M");
    } else {
      Browser.msgBox("Código no válido.");
      return;
    }

  var sheetPF = ssL.getSheetByName("Líneas");
  var Avals = sheetPF.getRange("A1:A").getValues();
  var lRowPF = Avals.filter(String).length;
  var lColumnPF = sheetPF.getLastColumn();
  var fInicioPF = 2;
  var cCodCentPF = 1;
  var cGrupoPF = 2;
  var cEjecPF = 3;
  var cFechaOG = 4;
  var cGrupoL = 6;
  var cComentariosPF = 8;
  var cEstRiesgos = 9;
  var cTraPF = 10;
  var cCodSolPF = 11;
  var cTipoOpPF = 12;
  var cMontoPF = 14;
  var cMontoSanc = 15;
  var cFechaSanc = 18;
  var cUCodSolPF = 22;
  var cFSPF = 23;
  var cTipoClientePF = 7;
  var sheetBC = ssOfic.getSheetByName("Base Clientes");
  var nombreGOF = sheetBC.getRange("J4").getValue();
  var correoGOF = sheetBC.getRange("J2").getValue();
  var sheetAssets = ssOfic.getSheetByName("Base Clientes Carterizados");
  var lRowAssets = sheetAssets.getLastRow();
  var lColAssets = sheetAssets.getLastColumn();
  var cCCGC = 1;
  var cGrupoSect = 2;
  var cAnPrim = 5;
  var cAnSec = 6;
  var cCorrAnPrim = 3;
  var cCorrAnSec = 4;
  var AvalsAsist = sheetBC.getRange("I1:I").getValues();
  var lRowAsist = AvalsAsist.filter(String).length;
  var correoAsistEncontrado = false;

  if (Ejecutivo != "SIN EJECUTIVO") {
    var myEjec = sheetBC.getRange(fInicioBC, 7, lRowAsist - fInicioBC + 1, 3).getValues();

    for (var i = 0; i <= lRowAsist - fInicioBC; i++) {
      var codGestMatriz = myEjec[i][0]; //Busca código de ejecutivos.

      if (Ejecutivo === codGestMatriz) {
        correoAsistEncontrado = true;
        var correoAsist = myEjec[i][2];
        break;
      }
    }
  } else if (Ejecutivo === "SIN EJECUTIVO" || Ejecutivo === "") {
    correoAsistEncontrado = true;
    var correoAsist = "siom@bbva.com";
  }

  if (correoAsistEncontrado === false) {
    correoAsist = "siom@bbva.com";
  }

  var valEncontrado = false;
  var sheetWF = ssOfic.getSheetByName("WF");
  var lRowWF = sheetWF.getLastRow();
  var lColumnWF = sheetWF.getLastColumn();
  var sheetBC = ssOfic.getSheetByName("Base Clientes");
  var lRowWF = sheetWF.getLastRow();
  var lColumnWF = sheetWF.getLastColumn();
  var arrayWF = sheetWF.getRange(finicioWF - 1, 1, lRowWF - finicioWF + 2, lColumnWF).getValues(); //lRowWF - fInicioWF +1

  for (var i = 1; i <= lRowWF - finicioWF + 1; i++) {
    var codSolWF = arrayWF[i][cCodSol - 1];

    if (codSolWF === CodSol) {
      valEncontrado = true;
    }
  }

  if (valEncontrado === false) {
    Browser.msgBox("Código no encontrado.");
    return;
  } //Loop que encuentra el código de solicitud


  for (var i = 1; i <= lRowWF - finicioWF + 1; i++) {
    //Acá empieza en 1 porque se incluyeron los Headings

    /*for(var i = finicioWF; i <= lRowWF; i++){*/
    var codSolWF = arrayWF[i][cCodSol - 1];

    if (codSolWF === CodSol) {
      //Compara códigos de solicitud
      var fEncontrada = i + finicioWF - 1; //Guarda el valor de la fila del código de solicitud encontrado dentro de la base WF.
    }
  } //Solo para la estación 1


  if (valEstEst === "I1") {
    for (var i = finicioEO; i <= lRowEO; i++) {
      var celda = sheetEO.getRange(i, cDatosE1);
      var val = celda.getValue();

      if (val === strRiReg || val === strRiReg2 || val === strRiReg3) {
        //Encuentra la instancia en la que riesgos puede ingresar una fecha.
        var celdaEO = sheetEO.getRange(i, cHeadingsE1);
        var headingEO = celdaEO.getValue(); //Obtiene el heading de la hoja de estación.

        if (headingEO === "Asignación RVGL" || headingEO === "Consulta de Riesgos a Cliente" || headingEO === "Consulta de BE a Cliente" || headingEO === "Devolución" || headingEO === "Consulta de Riesgos a Cliente (2)" || headingEO === "Consulta de BE a Cliente (2)" || headingEO === "Devolución (2)" || headingEO === "Fin Evaluación (VB Jefe)" || headingEO === "Fecha Sanción" || headingEO === "Asignación Evaluación II" || headingEO === "Consulta de Riesgos a Cliente (3)" || headingEO === "Consulta de BE a Cliente (3)" || headingEO === "Devolución (3)" || headingEO === "Fin Evaluación (VB Jefe) (2)" || headingEO === "Fecha Sanción (2)") {
          /*Continuar*/
        } else {
          Browser.msgBox("Riesgos", "El registro de fecha de esta operación le pertenece a la Oficina.", Browser.Buttons.OK);
          ss.toast("Cuando pueda registrar una fecha, aparecerá una celda de color amarillo (o celeste).", "Tip", 5);
        }

        var celdavaldev = sheetEO.getRange(rNDEV);
        var valdev = celdavaldev.getValue();

        if (valdev != "NDEV") {
          var respuesta = Browser.msgBox("Riesgos", "¿Desea devolver la operación a la oficina? Si presiona no, entonces se procederá a registrar la fecha de la Asignación RVGL.", Browser.Buttons.YES_NO);

          if (respuesta === "no") {
            var decRi = "no";
          } else if (respuesta === "cancel") {
            return;
          } else {
            var decRi = "yes";
          }
        } else {
          var decRi = "no";
        }

        for (var j = 1; j <= lColumnWF; j++) {
          //Recorre las columnas de la base WF.
          var celdaWF = sheetWF.getRange(fHeadingsWF, j);
          var headingWF = celdaWF.getValue(); //Obtiene el heading de la base WF.

          if (headingWF === headingEO) {
            //Encuentra la instancia en la que se igualan ambos valores de los headings.
            if (headingWF === "Asignación RVGL" && decRi === "yes") {
              var valDevCL = false;
              var valCasImp = false;
              var valCas = false;

              while (valCas === false) {
                var tMotDev = "J"; //Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. CLIENTE CON COVENANTS\\nB. CLIENTES CON ALERTAS EN EL SISTEMA\\nC. DESESTIMADA A SOLICITUD DE OFICINA\\nD. DESESTIMADO POR ANTC. NEGATIVOS CREDIT. Y/O PERFIL\\nE. DESESTIMADO POR CONTRASTE\\nF. DESESTIMADO POR NIVEL DE ENDEUDA Y/O CAPAC DE PAGO\\nG. DESESTIMADO POR VETO DE OFICINA\\nH. DEVOLUCION PARA MEJORAS DE CONDICIONES\\nI. FALTA ANTECEDENTES CREDITICIOS EN OTRAS ENT FINANC\\nJ. FALTA CHECK LIST\\nK. FALTA DE CONDICIONES PUNTUALES DE LA OPERACION\\nL. FALTAN DATOS BASICOS\\nM. RATING MAL ELABORADO\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                if (tMotDev === "cancel") {
                  return;
                }

                tMotDev = tMotDev.toUpperCase();
                var split_str = tMotDev.split("+");

                for (var iStr = 0; iStr < split_str.length; iStr++) {
                  if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L" || split_str[iStr] === "M") {
                    if (split_str[iStr] === "A") {
                      split_str[iStr] = "CLIENTE CON COVENANTS";
                    } else if (split_str[iStr] === "B") {
                      split_str[iStr] = "CLIENTES CON ALERTAS EN EL SISTEMA";
                    } else if (split_str[iStr] === "C") {
                      valCasImp = true;
                      split_str[iStr] = "DESESTIMADA A SOLICITUD DE OFICINA";
                    } else if (split_str[iStr] === "D") {
                      split_str[iStr] = "DESESTIMADO POR ANTC. NEGATIVOS CREDIT. Y/O PERFIL";
                    } else if (split_str[iStr] === "E") {
                      split_str[iStr] = "DESESTIMADO POR CONTRASTE";
                    } else if (split_str[iStr] === "F") {
                      split_str[iStr] = "DESESTIMADO POR NIVEL DE ENDEUDA Y/O CAPAC DE PAGO";
                    } else if (split_str[iStr] === "G") {
                      split_str[iStr] = "DESESTIMADO POR VETO DE OFICINA";
                    } else if (split_str[iStr] === "H") {
                      split_str[iStr] = "DEVOLUCION PARA MEJORAS DE CONDICIONES";
                    } else if (split_str[iStr] === "I") {
                      split_str[iStr] = "FALTA ANTECEDENTES CREDITICIOS EN OTRAS ENT FINANC";
                    } else if (split_str[iStr] === "J") {
                      split_str[iStr] = "FALTA CHECK LIST";
                      valDevCL = true;
                    } else if (split_str[iStr] === "K") {
                      split_str[iStr] = "FALTA DE CONDICIONES PUNTUALES DE LA OPERACION";
                    } else if (split_str[iStr] === "L") {
                      split_str[iStr] = "FALTAN DATOS BASICOS";
                    } else if (split_str[iStr] === "M") {
                      split_str[iStr] = "RATING MAL ELABORADO";
                    }

                    valCas = true;
                  } else {
                    Browser.msgBox("Error", "Motivo de devolución no válido.", Browser.Buttons.OK);
                    valCas = false;
                  }
                }
              }

              tMotDev = split_str.join("+");
              tMotDev = tMotDev.toUpperCase();

              if (valDevCL === true) {
                var valLoopCL = false;

                while (valLoopCL === false) {
                  var tMotDevCL = Browser.inputBox("Falta Check List", "Digite la opción que corresponde a los elementos que falten del Check List de la siguiente lista:\\nA. FALTA INFORMACIÓN COMERCIAL\\nB. FALTA DECLARACIÓN PATRIMONIAL ACCIONISTA\\nC. EEFF NO ACTUALIZADOS\\nD. R.C ANTERIORES\\nE. FALTA FLUJO DE CAJA\\nF. FALTA BACK LOG (PIPELINE, ACTUAL, HISTÓRICO)\\nG. FALTA PF GLOBAL\\nH. FALTA VOBO GESTOR GLOBAL\\nI. FALTA INF. PERITO\\nJ. FALTA TASACIÓN\\nK. FALTA INFORME COMERCIAL / INCOMPLETO\\nL. FALTA PFA\\nM. RIESGO EQUIVALENTE\\nN. OPERACIÓN EN DELEGACIÓN DE OFICINA\\nO. SOBREGIRO\\nP. ESTRATEGIA\\nQ. FEN\\nR. OTROS\\nS. VALIJA DIGITAL: ESTRUCTURA DE CARPETAS INCORRECTA / SIN ACCESO / RUTA INCORRECTA\\nT. VALIJA DIGITAL: DISTINTOS ARCHIVOS CONSOLIDADOS\\nU. VALIJA DIGITAL: OTROS MOTIVOS DE DEVOLUCIÓN\\n\\nV. SIO CON ERRORES DE PRODUCTO/IMPORTE/PLAZO/BANCO\\nW. RC WEB SIN FIRMAS O INCOMPLETAS\\nX. ÚLTIMA VALORIZACIÓN /ESTADO DE ARBITRAJE / AVANCE DE OBRA ACTUALIZADO (firmados por el beneficiario y/o supervisor)\\nY. EEFF SITUACIÓN INCOMPLENTOS (SIN DETALLE Y/O SIN FIRMA DEL CONTADOR COLEGIADO)\\nZ. COTIZACIÓN ASOCIADA A LA PROPUESTA LEASING / MP\\nAA. POSICIÓN CLIENTE (EXCEL)\\nAB. ESTATUS DE CARTAS FIANZAS EMITIDAS VIGENTES POR BBVA\\nAC. BACK LOG (PIPELINE, ACTUAL, HISTÓRICO) INCOMPLETO\\nAD. PFA INCOMPLETO / DESACTUALIZADO\\nAE. INFORME DE PERITO\\nAF. PERMISOS ACTUALIZADOS DE MINERÍA (EIA / CONCESIÓN / EXTRACCIÓN / TRANSPORTE / OTROS)\\nAG. CUOTA DE PESCA / LICENCIA DE PROCESAMIENTO (información actualizada y vigente)\\nAH. INFORME COMERCIAL INCOMPLETO / DESACTUALIZADO Y/O SIN FIRMAS\\nAI. DECLARACIÓN PATRIMONIAL ACCIONISTA INCOMPLETO / DESACTUALIZADO Y/O SIN FIRMAS\\nAJ. LICENCIAMIENTO (Universidades)\\nAK. PERMISO DE EXTRACCIÓN MADERERA\\nAL. PRESUPUESTO DE EJECUCIÓN DE OBRA (firmado por el ingeniero / arquitecto)\\nAM. FLUJO DE CAJA DE OBRA (evolución de avance/hitos)\\nAN. COPIA DEL CONTRATO VINCULADO AL FINANCIAMIENTO SOLICITADO\\nAO. INFORMACIÓN DE CONSORCIOS (EEFF / ACTA DE CONSTITUCIÓN / ADJUDICACIÓN)\\nAP. POSICIÓN DEUDORA (WEB)\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                  if (tMotDevCL === "cancel") {
                    return;
                  }

                  tMotDevCL = tMotDevCL.toUpperCase();
                  var split_str = tMotDevCL.split("+");

                  for (var iStr = 0; iStr < split_str.length; iStr++) {
                    if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L" || split_str[iStr] === "M" || split_str[iStr] === "N" || split_str[iStr] === "O" || split_str[iStr] === "P" || split_str[iStr] === "Q" || split_str[iStr] === "R" || split_str[iStr] === "S" || split_str[iStr] === "T" || split_str[iStr] === "U" || split_str[iStr] === "V" || split_str[iStr] === "W" || split_str[iStr] === "X" || split_str[iStr] === "Y" || split_str[iStr] === "Z" || split_str[iStr] === "AA" || split_str[iStr] === "AB" || split_str[iStr] === "AC" || split_str[iStr] === "AD" || split_str[iStr] === "AE" || split_str[iStr] === "AF" || split_str[iStr] === "AG" || split_str[iStr] === "AH" || split_str[iStr] === "AI" || split_str[iStr] === "AJ" || split_str[iStr] === "AK" || split_str[iStr] === "AL" || split_str[iStr] === "AM" || split_str[iStr] === "AN" || split_str[iStr] === "AO" || split_str[iStr] === "AP") {
                      if (split_str[iStr] === "A") {
                        split_str[iStr] = "FALTA INFORMACIÓN COMERCIAL";
                      } else if (split_str[iStr] === "B") {
                        split_str[iStr] = "FALTA DECLARACIÓN PATRIMONIAL ACCIONISTA";
                      } else if (split_str[iStr] === "C") {
                        valCasImp = true;
                        split_str[iStr] = "EEFF NO ACTUALIZADOS";
                      } else if (split_str[iStr] === "D") {
                        split_str[iStr] = "R.C ANTERIORES";
                      } else if (split_str[iStr] === "E") {
                        split_str[iStr] = "FALTA FLUJO DE CAJA";
                      } else if (split_str[iStr] === "F") {
                        split_str[iStr] = "FALTA BACK LOG (PIPELINE, ACTUAL, HISTÓRICO)";
                      } else if (split_str[iStr] === "G") {
                        split_str[iStr] = "FALTA PF GLOBAL";
                      } else if (split_str[iStr] === "H") {
                        split_str[iStr] = "FALTA VOBO GESTOR GLOBAL";
                      } else if (split_str[iStr] === "I") {
                        split_str[iStr] = "FALTA INF. PERITO";
                      } else if (split_str[iStr] === "J") {
                        split_str[iStr] = "FALTA TASACIÓN";
                      } else if (split_str[iStr] === "K") {
                        split_str[iStr] = "FALTA INFORME COMERCIAL / INCOMPLETO";
                      } else if (split_str[iStr] === "L") {
                        split_str[iStr] = "FALTA PFA";
                      } else if (split_str[iStr] === "M") {
                        split_str[iStr] = "RIESGO EQUIVALENTE";
                      } else if (split_str[iStr] === "N") {
                        split_str[iStr] = "OPERACIÓN EN DELEGACIÓN DE OFICINA";
                      } else if (split_str[iStr] === "O") {
                        split_str[iStr] = "SOBREGIRO";
                      } else if (split_str[iStr] === "P") {
                        split_str[iStr] = "ESTRATEGIA";
                      } else if (split_str[iStr] === "Q") {
                        split_str[iStr] = "FEN";
                      } else if (split_str[iStr] === "R") {
                        split_str[iStr] = "OTROS";
                      } else if (split_str[iStr] === "S") {
                        split_str[iStr] = "VALIJA DIGITAL: ESTRUCTURA DE CARPETAS INCORRECTA / SIN ACCESO / RUTA INCORRECTA";
                      } else if (split_str[iStr] === "T") {
                        split_str[iStr] = "VALIJA DIGITAL: DISTINTOS ARCHIVOS CONSOLIDADOS";
                      } else if (split_str[iStr] === "U") {
                        split_str[iStr] = "VALIJA DIGITAL: OTROS MOTIVOS DE DEVOLUCIÓN";
                      } else if (split_str[iStr] === "V") {
                        split_str[iStr] = "SIO CON ERRORES DE PRODUCTO/IMPORTE/PLAZO/BANCO";
                      } else if (split_str[iStr] === "W") {
                        split_str[iStr] = "RC WEB SIN FIRMAS O INCOMPLETAS";
                      } else if (split_str[iStr] === "X") {
                        split_str[iStr] = "ÚLTIMA VALORIZACIÓN /ESTADO DE ARBITRAJE / AVANCE DE OBRA ACTUALIZADO (firmados por el beneficiario y/o supervisor)";
                      } else if (split_str[iStr] === "Y") {
                        split_str[iStr] = "EEFF SITUACIÓN INCOMPLENTOS (SIN DETALLE Y/O SIN FIRMA DEL CONTADOR COLEGIADO)";
                      } else if (split_str[iStr] === "Z") {
                        split_str[iStr] = "COTIZACIÓN ASOCIADA A LA PROPUESTA LEASING / MP";
                      } else if (split_str[iStr] === "AA") {
                        split_str[iStr] = "POSICIÓN CLIENTE (EXCEL)";
                      } else if (split_str[iStr] === "AB") {
                        split_str[iStr] = "ESTATUS DE CARTAS FIANZAS EMITIDAS VIGENTES POR BBVA";
                      } else if (split_str[iStr] === "AC") {
                        split_str[iStr] = "BACK LOG (PIPELINE, ACTUAL, HISTÓRICO) INCOMPLETO";
                      } else if (split_str[iStr] === "AD") {
                        split_str[iStr] = "PFA INCOMPLETO / DESACTUALIZADO";
                      } else if (split_str[iStr] === "AE") {
                        split_str[iStr] = "INFORME DE PERITO";
                      } else if (split_str[iStr] === "AF") {
                        split_str[iStr] = "PERMISOS ACTUALIZADOS DE MINERÍA (EIA / CONCESIÓN / EXTRACCIÓN / TRANSPORTE / OTROS)";
                      } else if (split_str[iStr] === "AG") {
                        split_str[iStr] = "CUOTA DE PESCA / LICENCIA DE PROCESAMIENTO (información actualizada y vigente)";
                      } else if (split_str[iStr] === "AH") {
                        split_str[iStr] = "INFORME COMERCIAL INCOMPLETO / DESACTUALIZADO Y/O SIN FIRMAS";
                      } else if (split_str[iStr] === "AI") {
                        split_str[iStr] = "DECLARACIÓN PATRIMONIAL ACCIONISTA INCOMPLETO / DESACTUALIZADO Y/O SIN FIRMAS";
                      } else if (split_str[iStr] === "AJ") {
                        split_str[iStr] = "LICENCIAMIENTO (Universidades)";
                      } else if (split_str[iStr] === "AK") {
                        split_str[iStr] = "PERMISO DE EXTRACCIÓN MADERERA";
                      } else if (split_str[iStr] === "AL") {
                        split_str[iStr] = "PRESUPUESTO DE EJECUCIÓN DE OBRA (firmado por el ingeniero / arquitecto)";
                      } else if (split_str[iStr] === "AM") {
                        split_str[iStr] = "FLUJO DE CAJA DE OBRA (evolución de avance/hitos)";
                      } else if (split_str[iStr] === "AN") {
                        split_str[iStr] = "COPIA DEL CONTRATO VINCULADO AL FINANCIAMIENTO SOLICITADO";
                      } else if (split_str[iStr] === "AO") {
                        split_str[iStr] = "INFORMACIÓN DE CONSORCIOS (EEFF / ACTA DE CONSTITUCIÓN / ADJUDICACIÓN)";
                      } else if (split_str[iStr] === "AP") {
                        split_str[iStr] = "POSICIÓN DEUDORA (WEB)";
                      }

                      valLoopCL = true;
                    } else {
                      Browser.msgBox("Error", "Motivo de devolución de Check List no válido.", Browser.Buttons.OK);
                      valLoopCL = false;
                    }
                  }
                }

                tMotDevCL = split_str.join("+");
                tMotDevCL = tMotDevCL.toUpperCase();
              }

              var respuestamail = "yes"; //Browser.msgBox("Correo de Devolución", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

              if (respuestamail === "yes") {
                if (Correo != "SIN CORREO") {
                  var recipient = Correo;
                } else if (Correo === "SIN CORREO") {
                  var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                }

                if (recipient === "cancel") {
                  return;
                }

                if (Cliente === "GRUPO ECONÓMICO") {
                  var subject = "SIO OC: " + Grupo + " - Devolución de la Operación " + CodSol;
                } else {
                  var subject = "SIO OC: " + Cliente + " - Devolución de la Operación " + CodSol;
                }

                var body = "Riesgos ha devuelto la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                var tCasCorreo = tMotDev.split('+').join(" || ");
                body = body + "\n\nMotivo(s) de Devolución: " + tCasCorreo;

                if (valDevCL === true) {
                  var tMotDevCorreo = tMotDevCL.split('+').join(" || ");
                  body = body + "\n\nMotivo(s) de Devolución por Check List: " + tMotDevCorreo;
                }

                var options = {
                  cc: correoGOF + ", " + correoAsist + ", " + "siom@bbva.com"
                };
                MailApp.sendEmail(recipient, subject, body, options); //HERE DWP Devuelve Controller

                if (ofic === "MR" || ofic === "RP" || ofic === "LE" || ofic === "CY" || ofic === "TR" || ofic === "PR") {
                  var jefeEquipo = "LUIS ARIAS";
                } else if (ofic === "SC" || ofic === "CL" || ofic === "LM" || ofic === "AR" || ofic === "CU" || ofic === "CN") {
                  var jefeEquipo = "CHRISTIAN BARAYBAR";
                } else if (ofic === "SD" || ofic === "DM" || ofic === "LA" || ofic === "CA" || ofic === "CC" || ofic === "QQ" || ofic === "HU") {
                  var jefeEquipo = "SANDRA MIANI";
                }

                var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com"; //comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com 

                var subjectDWP = CodSol + "##Devuelve Controller##" + jefeEquipo + "##" + fecha + "##" + CodCentral;
                var bodyDWP = Oper;
                MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
              }

              var fechainicial = sheetWF.getRange(fEncontrada, cIngresoController).getValue();
              sheetWF.getRange(fEncontrada, c1erController).setValue(fechainicial);
              sheetWF.getRange(fEncontrada, cIngresoController).setValue("");
              sheetWF.getRange(fEncontrada, cDevuelveController).setValue(fecha);

              if (valDevCL === true) {
                sheetWF.getRange(fEncontrada, cMotDev).setValue(tMotDev + " // " + tMotDevCL);
              } else {
                sheetWF.getRange(fEncontrada, cMotDev).setValue(tMotDev);
              }

              Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
              ejecfWorkflowRiesgos(visor);
              return;
            } else if (headingWF === "Asignación RVGL" && decRi === "no") {
              var respuestamail = "yes"; //Browser.msgBox("Correo de Asignación de RVGL", "Esta acción notificará al gestor del registro de fecha. ¿Desea continuar?", Browser.Buttons.YES_NO)      

              if (respuestamail === "yes") {
                var recipient = Correo;

                if (recipient === "cancel") {
                  return;
                }

                if (Cliente === "GRUPO ECONÓMICO") {
                  var subject = "SIO OC: " + Grupo + " - Analista Asignado Para la Operación " + CodSol;
                } else {
                  var subject = "SIO OC: " + Cliente + " - Analista Asignado Para la Operación " + CodSol;
                }

                var body = "Riesgos ha registrado el analista para la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
              } else {
                return;
              }

              var tipoCliente = sheetWF.getRange(fEncontrada, cTipoCliente).getValue();
              var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
              var valElecAnAsig = 0;

              if (tipoCliente === "ESTANDAR" || tipoCliente === "STAGE 3") {
                valElecAnAsig = 1;
              } else {
                var arrayAssets = sheetAssets.getRange(1, 1, lRowAssets - 1 + 1, lColAssets).getValues();
                var fEncontradaAssets = 0;

                for (var iAssets = 1; iAssets < lRowAssets - 1 + 1; iAssets++) {
                  var CodCentralAssets = arrayAssets[iAssets][cCCGC - 1];

                  if (CodCentral === CodCentralAssets) {
                    fEncontradaAssets = iAssets + 1;
                    var nomAnPrim = arrayAssets[iAssets][cAnPrim - 1].trim();
                    var nomAnSec = arrayAssets[iAssets][cAnSec - 1].trim();
                    var corrAnPrim = arrayAssets[iAssets][cCorrAnPrim - 1].trim();
                    var corrAnSec = arrayAssets[iAssets][cCorrAnSec - 1].trim();
                    var grupoSect = arrayAssets[iAssets][cGrupoSect - 1].trim();
                    break;
                  }
                }

                if (fEncontradaAssets === 0 || corrAnPrim === "") {
                  valElecAnAsig = 1;
                } else {
                  var decAnAsig = Browser.msgBox("Analista de Riesgos", "Se ha encontrado a " + nomAnPrim + " como analista primario. \\nSe ha encontrado a " + nomAnSec + " como analista secundario.\\n\\n ¿Desea asignar al analista primario? En caso elegir la opción 'No', podrá elegir manualmente a un analista.", Browser.Buttons.YES_NO_CANCEL);

                  if (decAnAsig === "cancel" || decAnAsig === "") {
                    return;
                  }

                  if (decAnAsig === "yes") {
                    var AnAsig = nomAnPrim;
                  }

                  if (decAnAsig === "no") {
                    valElecAnAsig = 1;
                  }
                }
              }

              if (valElecAnAsig === 1) {
                var listaAnalistas = "Digite la letra correspondiente al analista de riesgos asignado:" + "\\nA. ANDRÉS  LÓPEZ ARGUEDAS " + "\\n\\nSECTORIZADO 1:" + "\\nB. CARLOS CASADIEGO JAIMES" + "\\nC. CHRISTIAN ALEJANDRO BRAVO MEDINA" + "\\nD. CHRISTIAN BARAYBAR" + "\\nE. ANDRÉS MARTÍN PAREDES MEDINA" + "\\nF. CLAUDIA PORTUGAL DEL CARPIO" + "\\nG. EDITH MARLENY CANO GABRIEL" + "\\nH. GIULIANA XIMENA BILIBIO ARAGONES" + "\\nI. JUAN CARLOS ROMERO UNOCC" + "\\nJ. SUSAN MARIELA ARRIETA ALFARO" + "\\nK. PATRICIA PARDO DELGADO" + "\\nL. FIORELLA DEL ROSARIO SALAZAR CHIRINOS" + "\\nM. MELISSA DEL PILAR ANDERSON MARTINEZ" + "\\nN. PEDRO NICSSON SOTO MUNIVE" + "\\nO. ROGELIO LEON GUZMAN" + "\\nP. ANGELO PIGNANO BRAVO" + "\\n\\nSECTORIZADO 2:" + "\\nQ. VICTORIA COSSIO" + "\\nR. ALEJANDRO GIAMBRONI" + "\\nS. ANTONIO ALONSO GARCIA ROSADIO" + "\\nT. CRISTIAN PAREDES PARIONA" + "\\nU. GIULIANA VANESSA OÑA ALVAREZ" + "\\nV. LIA ADRIANZEN RODRIGUEZ" + "\\nW. FLAVIA DEL CARMEN POLAR CASTILLO" + "\\nX. DAFNA HEITNER JARA" + "\\nY. BRENDA CAROLINA FELIPPE AGUILAR" + "\\nZ. ERIKA PALACIOS ROJAS" + "\\nAA. GIANFRANCO ERNESTO ALVARADO CAVERO" + "\\nAB.CESAR AUGUSTO CASTILLO AGUIRRE" + "\\nAC.MARIA ISABEL BOCANEGRA CONTRERAS" + "\\n\\nSECTORIZADO 3:" + "\\nAD. ALICIA ALEJANDRA ALAVE CALANI" + "\\nAE. FELIPE CHONN" + "\\nAF. MARIA GRAZIA MESIAS VALDIVIA" + "\\nAG. LUCERO DEL SOCORRO BURGOS DIAZ" + "\\nAH. PERCY CAYCHO ARCE" + "\\nAI. ROSA VICTORIA OLCESE FARROMEQUE" + "\\nAJ. MARIA GABRIELA MARKY VARHEN" + "\\nAK. LEANDRO VILLANUEVA UGAS" + "\\nAL. JENNY CATHERY VELEZ DIEZ CANSECO" + "\\n\\nSECTORIZADO 4:" + "\\nAM. ROSA ANGELINA LOAYZA" + "\\nAN. TANIA FRANCESCA MONZON ASTENGO" + "\\nAO. ALEJANDRO REMAR RUIZ DE SOMOCURCIO" + "\\nAP. ALONSO TEJEDA ZAMORA" + "\\nAQ. MARCEL CONTRERAS" + "\\nAR. MARIEL GARCIA NARANJO" + "\\nAS. LUIS ANTONIO LANDAURO ELERA" + "\\nAT. LUIS VALLE MORENO" + "\\nAU. OSCAR BERNARDO ADRIANZEN HERRERA" + "\\nAV. CLAUDIA CHAMORRO" + "\\nAW. KATHERINE COTERA" + "\\n\\nSTAGE 2:" + "\\nAX. JONATHAN RIVERO" + "\\nAY. LISBETH CALLE" + "\\nAZ. BLANCA VERONICA DIAZ FALCON" + "\\nBA. CESAR ALFREDO PORTAL LOO" + "\\nBB. CYNTHIA ELIZABETH HIDALGO PEREZ" + "\\nBC. ROCIO DEL MILAGRO SIERRA LLENQUE" + "\\nBD. RENZO MARTIN DENEGRI CROVETTO" + "\\nBE. EVELIN ROCIO DE LA CRUZ MARIN" + "\\nBF. ALVARO GARLAND REYES" + "\\n\\nSTAGE 3:" + "\\nBG. CARLOS FIDEL LORA FRISANCHO" + "\\nBH. ANA GIANNINA MENDOZA QUIROGA" + "\\nBI. URSULA PATRICIA NOAIN MORENO" + "\\nBJ. PAMELA MOLINARI MOLINARI ARROYO" + "\\n\\nESTANDAR:" + "\\nBK. CHRISTIAN BARAYBAR" + "\\nBL. GIULIANA VANESSA OÑA ALVAREZ" + "\\nBM. MARIA GABRIELA MARKY VARHEN" + "\\nBN. MARCEL CONTRERAS";
                var AnAsig = Browser.inputBox("Analista de Riesgos", listaAnalistas, Browser.Buttons.OK_CANCEL);

                if (AnAsig === "cancel") {
                  return;
                }

                AnAsig = AnAsig.toUpperCase();

                if (AnAsig === "A") {
                  AnAsig = "ANDRÉS  LÓPEZ ARGUEDAS";
                } else if (AnAsig === "B") {
                  AnAsig = "CARLOS CASADIEGO JAIMES";
                } else if (AnAsig === "C") {
                  AnAsig = "CHRISTIAN ALEJANDRO BRAVO MEDINA";
                } else if (AnAsig === "D") {
                  AnAsig = "CHRISTIAN BARAYBAR";
                } else if (AnAsig === "E") {
                  AnAsig = "ANDRÉS MARTÍN PAREDES MEDINA";
                } else if (AnAsig === "F") {
                  AnAsig = "CLAUDIA PORTUGAL DEL CARPIO";
                } else if (AnAsig === "G") {
                  AnAsig = "EDITH MARLENY CANO GABRIEL";
                } else if (AnAsig === "H") {
                  AnAsig = "GIULIANA XIMENA BILIBIO ARAGONES";
                } else if (AnAsig === "I") {
                  AnAsig = "JUAN CARLOS ROMERO UNOCC";
                } else if (AnAsig === "J") {
                  AnAsig = "SUSAN MARIELA ARRIETA ALFARO";
                } else if (AnAsig === "K") {
                  AnAsig = "PATRICIA PARDO DELGADO";
                } else if (AnAsig === "L") {
                  AnAsig = "FIORELLA DEL ROSARIO SALAZAR CHIRINOS";
                } else if (AnAsig === "M") {
                  AnAsig = "MELISSA DEL PILAR ANDERSON MARTINEZ";
                } else if (AnAsig === "N") {
                  AnAsig = "PEDRO NICSSON SOTO MUNIVE";
                } else if (AnAsig === "O") {
                  AnAsig = "ROGELIO LEON GUZMAN";
                } else if (AnAsig === "P") {
                  AnAsig = "ANGELO PIGNANO BRAVO";
                } else if (AnAsig === "Q") {
                  AnAsig = "VICTORIA COSSIO";
                } else if (AnAsig === "R") {
                  AnAsig = "ALEJANDRO GIAMBRONI";
                } else if (AnAsig === "S") {
                  AnAsig = "ANTONIO ALONSO GARCIA ROSADIO";
                } else if (AnAsig === "T") {
                  AnAsig = "CRISTIAN PAREDES PARIONA";
                } else if (AnAsig === "U") {
                  AnAsig = "GIULIANA VANESSA OÑA ALVAREZ";
                } else if (AnAsig === "V") {
                  AnAsig = "LIA ADRIANZEN RODRIGUEZ";
                } else if (AnAsig === "W") {
                  AnAsig = "FLAVIA DEL CARMEN POLAR CASTILLO";
                } else if (AnAsig === "X") {
                  AnAsig = "DAFNA HEITNER JARA";
                } else if (AnAsig === "Y") {
                  AnAsig = "BRENDA CAROLINA FELIPPE AGUILAR";
                } else if (AnAsig === "Z") {
                  AnAsig = "ERIKA PALACIOS ROJAS";
                } else if (AnAsig === "AA") {
                  AnAsig = "GIANFRANCO ERNESTO ALVARADO CAVERO";
                } else if (AnAsig === "AB") {
                  AnAsig = "CESAR AUGUSTO CASTILLO AGUIRRE";
                } else if (AnAsig === "AC") {
                  AnAsig = "MARIA ISABEL BOCANEGRA CONTRERAS";
                } else if (AnAsig === "AD") {
                  AnAsig = "ALICIA ALEJANDRA ALAVE CALANI";
                } else if (AnAsig === "AE") {
                  AnAsig = "FELIPE CHONN";
                } else if (AnAsig === "AF") {
                  AnAsig = "MARIA GRAZIA MESIAS VALDIVIA";
                } else if (AnAsig === "AG") {
                  AnAsig = "LUCERO DEL SOCORRO BURGOS DIAZ";
                } else if (AnAsig === "AH") {
                  AnAsig = "PERCY CAYCHO ARCE";
                } else if (AnAsig === "AI") {
                  AnAsig = "ROSA VICTORIA OLCESE FARROMEQUE";
                } else if (AnAsig === "AJ") {
                  AnAsig = "MARIA GABRIELA MARKY VARHEN";
                } else if (AnAsig === "AK") {
                  AnAsig = "LEANDRO VILLANUEVA UGAS";
                } else if (AnAsig === "AL") {
                  AnAsig = "JENNY CATHERY VELEZ DIEZ CANSECO";
                } else if (AnAsig === "AM") {
                  AnAsig = "ROSA ANGELINA LOAYZA";
                } else if (AnAsig === "AN") {
                  AnAsig = "TANIA FRANCESCA MONZON ASTENGO";
                } else if (AnAsig === "AO") {
                  AnAsig = "ALEJANDRO REMAR RUIZ DE SOMOCURCIO";
                } else if (AnAsig === "AP") {
                  AnAsig = "ALONSO TEJEDA ZAMORA";
                } else if (AnAsig === "AQ") {
                  AnAsig = "MARCEL CONTRERAS";
                } else if (AnAsig === "AR") {
                  AnAsig = "MARIEL GARCIA NARANJO";
                } else if (AnAsig === "AS") {
                  AnAsig = "LUIS ANTONIO LANDAURO ELERA";
                } else if (AnAsig === "AT") {
                  AnAsig = "LUIS VALLE MORENO";
                } else if (AnAsig === "AU") {
                  AnAsig = "OSCAR BERNARDO ADRIANZEN HERRERA";
                } else if (AnAsig === "AV") {
                  AnAsig = "CLAUDIA CHAMORRO";
                } else if (AnAsig === "AW") {
                  AnAsig = "KATHERINE COTERA";
                } else if (AnAsig === "AX") {
                  AnAsig = "JONATHAN RIVERO";
                } else if (AnAsig === "AY") {
                  AnAsig = "LISBETH CALLE";
                } else if (AnAsig === "AZ") {
                  AnAsig = "BLANCA VERONICA DIAZ FALCON";
                } else if (AnAsig === "BA") {
                  AnAsig = "CESAR ALFREDO PORTAL LOO";
                } else if (AnAsig === "BB") {
                  AnAsig = "CYNTHIA ELIZABETH HIDALGO PEREZ";
                } else if (AnAsig === "BC") {
                  AnAsig = "ROCIO DEL MILAGRO SIERRA LLENQUE";
                } else if (AnAsig === "BD") {
                  AnAsig = "RENZO MARTIN DENEGRI CROVETTO";
                } else if (AnAsig === "BE") {
                  AnAsig = "EVELIN ROCIO DE LA CRUZ MARIN";
                } else if (AnAsig === "BF") {
                  AnAsig = "ALVARO GARLAND REYES";
                } else if (AnAsig === "BG") {
                  AnAsig = "CARLOS FIDEL LORA FRISANCHO";
                } else if (AnAsig === "BH") {
                  AnAsig = "ANA GIANNINA MENDOZA QUIROGA";
                } else if (AnAsig === "BI") {
                  AnAsig = "URSULA PATRICIA NOAIN MORENO";
                } else if (AnAsig === "BJ") {
                  AnAsig = "PAMELA MOLINARI MOLINARI ARROYO";
                } else if (AnAsig === "BK") {
                  AnAsig = "CHRISTIAN BARAYBAR";
                } else if (AnAsig === "BL") {
                  AnAsig = "GIULIANA VANESSA OÑA ALVAREZ";
                } else if (AnAsig === "BM") {
                  AnAsig = "MARIA GABRIELA MARKY VARHEN";
                } else if (AnAsig === "BN") {
                  AnAsig = "MARCEL CONTRERAS";
                } else {
                  Browser.msgBox("Error", "Digite una letra válida de la lista.", Browser.Buttons.OK);
                  return;
                }
              }

              var valAmb = false;

              while (valAmb === false) {
                valAmb = true;
                var amb = Browser.inputBox("Registrar Nivel de Aprobación en Tránsito", "Digite la letra correspondiente a un nivel de aprobación en tránsito de la siguiente lista:\\nA. ANALISTA\\nB. JEFE DE GRUPO\\nC. JEFE DE EQUIPO\\nD. SUBGERENTE\\nE. GERENTE DE UNIDAD\\nF. GERENTE DE ÁREA\\nG. COMITÉ DE CONTRASTE \\nH. CTO \\nI. CEC\\nJ. WCRMC\\nK. GCRMC\\nL. RECONDUCCIÓN\\nM. ECCMWO", Browser.Buttons.OK_CANCEL);
                var tempAmb = amb;

                if (amb === "") {
                  return;
                }

                if (amb === "cancel") {
                  return;
                }

                amb = amb.toUpperCase();

                if (amb === "A") {
                  amb = "ANALISTA";
                } else if (amb === "B") {
                  amb = "JEFE DE GRUPO";
                } else if (amb === "C") {
                  amb = "JEFE DE EQUIPO";
                } else if (amb === "D") {
                  amb = "SUBGERENTE";
                } else if (amb === "E") {
                  amb = "GERENTE DE UNIDAD";
                } else if (amb === "F") {
                  amb = "GERENTE DE ÁREA";
                } else if (amb === "G") {
                  amb = "COMITÉ DE CONTRASTE";
                } else if (amb === "H") {
                  amb = "CTO";
                } else if (amb === "I") {
                  amb = "CEC";
                } else if (amb === "J") {
                  amb = "WCRMC";
                } else if (amb === "K") {
                  amb = "GCRMC";
                } else if (amb === "L") {
                  amb = "RECONDUCCIÓN";
                } else if (amb === "M") {
                  amb = "ECCMWO";
                } else {
                  valAmb = false;
                  Browser.msgBox("Error", "No se ingresó una letra válida de la lista de nivel de aprobación.", Browser.Buttons.OK);
                  return;
                }
              }

              amb = amb.toUpperCase();
              var ambitoTrans = amb;
              var valAmb = false;

              while (valAmb === false) {
                valAmb = true;
                var amb = Browser.inputBox("Registrar Ámbito", "Digite la letra correspondiente al ámbito de la siguiente lista:\\nA. LOCAL\\nB. GCR ARGENTINA\\nC. GCR BRASIL\\nD. GCR COLOMBIA\\nE. GCR COMPASS\\nF. GCR MEXICO\\nG. GCR NY\\nH. GCR PARAGUAY\\nI. GCR PANAMA\\nJ. GCR URUGUAY\\nK. RPM ALEMANIA\\nL. RPM BELGICA\\nM. RPM COREA\\nN. RPM ESPAÑA\\nO. RPM FRANCIA\\nP. RPM HK\\nQ. RPM ITALIA\\nR. RPM JAPON\\nS. RPM SINGAPUR\\nT. RPM UK", Browser.Buttons.OK_CANCEL);
                var tempAmb = amb;

                if (amb === "") {
                  return;
                }

                if (amb === "cancel") {
                  return;
                }

                amb = amb.toUpperCase();

                if (amb === "A") {
                  amb = "LOCAL";
                } else if (amb === "B") {
                  amb = "GCR ARGENTINA";
                } else if (amb === "C") {
                  amb = "GCR BRASIL";
                } else if (amb === "D") {
                  amb = "GCR COLOMBIA";
                } else if (amb === "E") {
                  amb = "GCR COMPASS";
                } else if (amb === "F") {
                  amb = "GCR MEXICO";
                } else if (amb === "G") {
                  amb = "GCR NY";
                } else if (amb === "H") {
                  amb = "GCR PARAGUAY";
                } else if (amb === "I") {
                  amb = "GCR PANAMA";
                } else if (amb === "J") {
                  amb = "GCR URUGUAY";
                } else if (amb === "K") {
                  amb = "RPM ALEMANIA";
                } else if (amb === "L") {
                  amb = "RPM BELGICA";
                } else if (amb === "M") {
                  amb = "RPM COREA";
                } else if (amb === "N") {
                  amb = "RPM ESPAÑA";
                } else if (amb === "O") {
                  amb = "RPM FRANCIA";
                } else if (amb === "P") {
                  amb = "RPM HK";
                } else if (amb === "Q") {
                  amb = "RPM ITALIA";
                } else if (amb === "R") {
                  amb = "RPM JAPON";
                } else if (amb === "S") {
                  amb = "RPM SINGAPUR";
                } else if (amb === "T") {
                  amb = "RPM UK";
                } else {
                  valAmb = false;
                  Browser.msgBox("Error", "No se ingresó una letra válida de la lista de ámbitos.", Browser.Buttons.OK);
                  return;
                }
              }

              amb = amb.toUpperCase();
              ambitoTrans = ambitoTrans + " // " + amb;
              var AvalsAnAsig = sheetBC.getRange("K1:K").getValues();
              var lRowAnAsig = AvalsAnAsig.filter(String).length;
              var myAnAsig = sheetBC.getRange(fInicioBC, 11, lRowAnAsig - fInicioBC + 1, 4).getValues();

              for (var i = 0; i <= lRowAnAsig - fInicioBC; i++) {
                var codGestMatriz = myAnAsig[i][0]; //Busca código de ejecutivos.

                if (AnAsig === codGestMatriz) {
                  var correoAnAsig = myAnAsig[i][3];
                  break;
                }
              }

              sheetWF.getRange(fEncontrada, cAsignacionRVGL).setValue(fecha); //Escribe la fecha.

              sheetWF.getRange(fEncontrada, cAnAsig).setValue(AnAsig);
              sheetWF.getRange(fEncontrada, cAmbSancTrans).setValue(ambitoTrans);
              celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

              celda.setBackgroundRGB(85, 199, 104);

              if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                var valCodCentEnc = false;
                var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                  var codCentPF = arrayPF[i][cCodCentPF - 1];
                  var codSolPF = arrayPF[i][cCodSolPF - 1];

                  if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                    valCodCentEnc = true;
                    sheetPF.getRange(i + fInicioPF, cTraPF).setValue("SÍ");
                    sheetPF.getRange(i + fInicioPF, cEstRiesgos).setValue(fecha);
                    break;
                  }
                }

                if (valCodCentEnc === false) {
                  Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                  var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                  var subjectE = "Operación no encontrada en la base de Líneas.";
                  var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                  MailApp.sendEmail(recipientE, subjectE, bodyE);
                }
              }

              var options = {
                cc: correoGOF + ", " + correoAsist + ", " + correoAnAsig + ", " + "siom@bbva.com"
              };
              var body = body + "\nAnalista Asignado: " + AnAsig;
              var body = body + "\nÁmbito de Sanción en Tránsito: " + ambitoTrans;
              var link = sheetWF.getRange(fEncontrada, cLink).getValue();

              if (link != "") {
                body = body + "\nLink de archivo: " + link;
              }

              MailApp.sendEmail(recipient, subject, body, options); //HERE DWP Asignación RVGL

              var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com"; //comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com 

              var subjectDWP = CodSol + "##Asignación RVGL##" + AnAsig + "##" + fecha + "##" + CodCentral;
              var bodyDWP = Oper;
              MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
              Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
              ejecfWorkflowRiesgos(visor);
              return;
            }
          }
        }

        break;
      }
    }
  } //Solo para la estación 1
  //Solo para la estación 2
  else if (valEstEst === "I2") {
      for (var i = finicioEO; i <= lRowEO; i++) {
        var celda = sheetEO.getRange(i, cDatosE2);
        var val = celda.getValue();

        if (val === strRiReg || val === strRiReg2 || val === strRiReg3) {
          //Encuentra la instancia en la que riesgos puede ingresar una fecha.
          var celdavaldev = sheetEO.getRange(rNDEV);
          var valdev = celdavaldev.getValue();
          var celdaEO = sheetEO.getRange(i, cHeadingsE2);
          var headingEO = celdaEO.getValue(); //Obtiene el heading de la hoja de estación.

          if (headingEO === "Asignación RVGL" || headingEO === "Consulta de Riesgos a Cliente" || headingEO === "Consulta de BE a Cliente" || headingEO === "Devolución" || headingEO === "Consulta de Riesgos a Cliente (2)" || headingEO === "Consulta de BE a Cliente (2)" || headingEO === "Devolución (2)" || headingEO === "Fin Evaluación (VB Jefe)" || headingEO === "Fecha Sanción" || headingEO === "Asignación Evaluación II" || headingEO === "Consulta de Riesgos a Cliente (3)" || headingEO === "Consulta de BE a Cliente (3)" || headingEO === "Devolución (3)" || headingEO === "Fin Evaluación (VB Jefe) (2)" || headingEO === "Fecha Sanción (2)") {
            /*Continuar*/
          } else {
            Browser.msgBox("Riesgos", "El registro de fecha de esta operación le pertenece a la Oficina.", Browser.Buttons.OK);
            ss.toast("Cuando pueda registrar una fecha, aparecerá una celda de color amarillo (o celeste).", "Tip", 5);
          }

          if (headingEO === "Fin Evaluación (VB Jefe)") {
            if (valdev != "NDEV") {
              var respuesta = Browser.msgBox("Riesgos", "¿La evaluación ha concluido? Presione 'Sí' solo si concluyó. En caso contrario presione 'No' para registrar la fecha de las consultas enviadas a la Oficina.", Browser.Buttons.YES_NO);

              if (respuesta == "no") {
                var decRi = "no";
                var respuestamail = "yes"; //Browser.msgBox("Correo de Solicitud", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                if (respuestamail === "yes") {
                  if (Correo != "SIN CORREO") {
                    var recipient = Correo;
                  } else if (Correo === "SIN CORREO") {
                    var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                  }

                  if (recipient === "cancel") {
                    return;
                  }

                  if (Cliente === "GRUPO ECONÓMICO") {
                    var subject = "SIO OC: " + Grupo + " - Solicitud de Ingreso de Respuestas Para la Operación " + CodSol;
                  } else {
                    var subject = "SIO OC: " + Cliente + " - Solicitud de Ingreso de Respuestas Para la Operación " + CodSol;
                  }

                  var body = "Se solicita su intervención para el ingreso de respuestas para la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                  var options = {
                    cc: correoGOF + ", " + correoAsist + ", siom@bbva.com"
                  };
                  MailApp.sendEmail(recipient, subject, body, options);
                }
              } else if (respuesta === "cancel") {
                return;
              } else {
                var decRi = "yes";
              }
            } else {
              var decRi = "no";
            }
          } else if (headingEO === "Devolución") {} else if (headingEO === "Reingreso con Respuestas (1)") {} else if (headingEO === "Devolución (2)") {} else if (headingEO === "Reingreso con Respuestas (2)") {} else if (headingEO === "Fecha Sanción") {} else {
            var respuesta = Browser.msgBox("Riesgos", "¿Desea devolver la operación?", Browser.Buttons.YES_NO);

            if (respuesta == "no") {
              var decRi = "no";
            } else if (respuesta === "cancel") {
              return;
            } else {
              var decRi = "yes";
            }
          }

          for (var j = 1; j <= lColumnWF; j++) {
            //Recorre las columnas de la base WF.
            var celdaWF = sheetWF.getRange(fHeadingsWF, j);
            var headingWF = celdaWF.getValue(); //Obtiene el heading de la base WF.

            if (headingWF === headingEO) {
              //Encuentra la instancia en la que se igualan ambos valores de los headings.
              var celdaHEOCRC = sheetEO.getRange(finicioEO + 1, cHeadingsE2);
              var HEOCRC = celdaHEOCRC.getValue();
              var celdanDEV = sheetEO.getRange(rNDEV);
              var valNDEV = celdanDEV.getValue();

              if (headingWF === "Fin Evaluación (VB Jefe)" && decRi === "no" && HEOCRC != "Consulta de BE a Cliente" && valNDEV === "") {
                //HERE DWP Consulta de BE a Cliente
                var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com"; //comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com 

                var subjectDWP = CodSol + "##Consulta de BE a Cliente##" + AnAsig + "##" + fecha + "##" + CodCentral;
                var bodyDWP = Oper;
                MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                sheetWF.getRange(fEncontrada, cCBC).setValue(fecha);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                ejecfWorkflowRiesgos(visor);
                return;
              } else if (headingWF === "Fin Evaluación (VB Jefe)" && decRi === "no" && HEOCRC === "Consulta de BE a Cliente (2)" && valNDEV === "") {
                //HERE DWP Consulta de BE a Cliente (2)
                var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                var subjectDWP = CodSol + "##Consulta de BE a Cliente (2)##" + AnAsig + "##" + fecha + "##" + CodCentral;
                var bodyDWP = Oper;
                MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                sheetWF.getRange(fEncontrada, c2daCBC).setValue(fecha);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                ejecfWorkflowRiesgos(visor);
                return;
              } else if (headingWF === "Fin Evaluación (VB Jefe)" && decRi === "yes" && valNDEV === "") {
                //CORREO VB JEFE puede ir aquí.           
                sheetWF.getRange(fEncontrada, cFEv).setValue(fecha); //Escribe la fecha.

                celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                celda.setBackgroundRGB(85, 199, 104);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                ejecfWorkflowRiesgos(visor);
                return;
              } else if (headingWF === "Fin Evaluación (VB Jefe)" && valNDEV === "NDEV") {
                //CORREO VB JEFE puede ir aquí.
                sheetWF.getRange(fEncontrada, cFEv).setValue(fecha);
                celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                celda.setBackgroundRGB(85, 199, 104);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                ejecfWorkflowRiesgos(visor);
                return;
              } else if (headingWF === "Fin Evaluación (VB Jefe)" && valNDEV === "") {
                var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                var subjectDWP = CodSol + "##Consulta de BE a Cliente (2)##" + AnAsig + "##" + fecha + "##" + CodCentral;
                var bodyDWP = Oper;
                MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                sheetWF.getRange(fEncontrada, c2daCBC).setValue(fecha); //ALERTA ESTO SE CAMBIÓ

                celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                celda.setBackgroundRGB(85, 199, 104);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                ejecfWorkflowRiesgos(visor);
                return;
              } //////////Modificar esto para que acepte la segunda ronda de preguntas                
              /////////Primera ronda de preguntas                


              if (headingWF === "Devolución") {
                var respuesta = Browser.msgBox("Riesgos", "¿El cliente le ha contestado a la oficina? Si elige no, entonces se procederá a devolver la operación hasta que la oficina reingrese las respuestas.", Browser.Buttons.YES_NO);

                if (respuesta === "cancel") {
                  return;
                } else if (respuesta === "yes") {
                  Browser.msgBox("Riesgos", "Espere hasta que la oficina registre la fecha de respuestas.", Browser.Buttons.OK);
                  return;
                }

                var valCas = false;

                while (valCas === false) {
                  var tMotDev = Browser.inputBox("Casuística de Devolución", "Digite la opción de la casuística de la siguiente lista:\\nA. DEVUELTA A SOLICITUD DE LA OFICINA\\nB. DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO\\nC. DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA\\nD. DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING\\nE. DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR\\nF. DEVUELTA POR FALTA DE INFORMACION EN GENERAL\\nG. OTROS\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                  if (tMotDev === "cancel") {
                    return;
                  }

                  tMotDev = tMotDev.toUpperCase();
                  var split_str = tMotDev.split("+");

                  for (var iStr = 0; iStr < split_str.length; iStr++) {
                    if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G") {
                      if (split_str[iStr] === "A") {
                        split_str[iStr] = "DEVUELTA A SOLICITUD DE LA OFICINA";
                      } else if (split_str[iStr] === "B") {
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO";
                      } else if (split_str[iStr] === "C") {
                        valCasImp = true;
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA";
                      } else if (split_str[iStr] === "D") {
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING";
                      } else if (split_str[iStr] === "E") {
                        split_str[iStr] = "DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR";
                      } else if (split_str[iStr] === "F") {
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION EN GENERAL";
                      } else if (split_str[iStr] === "G") {
                        split_str[iStr] = "OTROS";
                      }

                      valCas = true;
                    } else {
                      Browser.msgBox("Error", "Motivo de devolución no válido.", Browser.Buttons.OK);
                      valCas = false;
                    }
                  }
                }

                tMotDev = split_str.join("+");
                tMotDev = tMotDev.toUpperCase();
                var respuestamail = "yes"; //Browser.msgBox("Correo de Devolución", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                if (respuestamail === "yes") {
                  if (Correo != "SIN CORREO") {
                    var recipient = Correo;
                  } else if (Correo === "SIN CORREO") {
                    var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                  }

                  if (recipient === "cancel") {
                    return;
                  }

                  if (Cliente === "GRUPO ECONÓMICO") {
                    var subject = "SIO OC: " + Grupo + " - Devolución de la Operación " + CodSol;
                  } else {
                    var subject = "SIO OC: " + Cliente + " - Devolución de la Operación " + CodSol;
                  }

                  var body = "Riesgos ha devuelto la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                  var tCasCorreo = tMotDev.split('+').join(" || ");
                  body = body + "\nCasuística de Devolución: " + tCasCorreo;
                  var options = {
                    cc: correoGOF + ", " + correoAsist + ", siom@bbva.com"
                  };
                  MailApp.sendEmail(recipient, subject, body, options); //HERE DWP Devolución

                  var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                  var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                  var subjectDWP = CodSol + "##Devolución##" + AnAsig + "##" + fecha + "##" + CodCentral;
                  var bodyDWP = Oper;
                  MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                  sheetWF.getRange(fEncontrada, cCas).setValue(tMotDev);
                }

                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                  var valCodCentEnc = false;
                  var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                  for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                    var codCentPF = arrayPF[i][cCodCentPF - 1];
                    var codSolPF = arrayPF[i][cCodSolPF - 1]; //var nuevaFV = "DEVUELTO"
                    //var tMontSan = "DEVUELTO"

                    if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                      valCodCentEnc = true; //if(tMontSan === "DENEGADO" || tMontSan === "DEVUELTO"){nuevaFV === tMontSan}
                      //sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)

                      sheetPF.getRange(i + fInicioPF, cTraPF).setValue("NO"); //sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)

                      break;
                    }
                  }

                  if (valCodCentEnc === false) {
                    Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                    var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                    var subjectE = "Operación no encontrada en la base de Líneas.";
                    var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                    MailApp.sendEmail(recipientE, subjectE, bodyE);
                  }
                }

                sheetWF.getRange(fEncontrada, cDev).setValue(fecha);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                celda.setBackgroundRGB(85, 199, 104);
                ejecfWorkflowRiesgos(visor);
                return;
              } /////////Primera ronda de preguntas
              /////////Segunda ronda de preguntas                


              if (headingWF === "Devolución (2)") {
                var respuesta = Browser.msgBox("Riesgos", "¿El cliente le ha contestado a la oficina? Si elige no, entonces se procederá a devolver la operación hasta que la oficina reingrese las respuestas.", Browser.Buttons.YES_NO);

                if (respuesta === "cancel") {
                  return;
                } else if (respuesta === "yes") {
                  Browser.msgBox("Riesgos", "Espere hasta que la oficina registre la fecha de respuestas.", Browser.Buttons.OK);
                  return;
                }

                var valCas = false;

                while (valCas === false) {
                  var tMotDev = Browser.inputBox("Casuística de Devolución", "Digite la opción de la casuística de la siguiente lista:\\nA. DEVUELTA A SOLICITUD DE LA OFICINA\\nB. DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO\\nC. DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA\\nD. DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING\\nE. DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR\\nF. DEVUELTA POR FALTA DE INFORMACION EN GENERAL\\nG. OTROS\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                  if (tMotDev === "cancel") {
                    return;
                  }

                  tMotDev = tMotDev.toUpperCase();
                  var split_str = tMotDev.split("+");

                  for (var iStr = 0; iStr < split_str.length; iStr++) {
                    if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G") {
                      if (split_str[iStr] === "A") {
                        split_str[iStr] = "DEVUELTA A SOLICITUD DE LA OFICINA";
                      } else if (split_str[iStr] === "B") {
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO";
                      } else if (split_str[iStr] === "C") {
                        valCasImp = true;
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA";
                      } else if (split_str[iStr] === "D") {
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING";
                      } else if (split_str[iStr] === "E") {
                        split_str[iStr] = "DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR";
                      } else if (split_str[iStr] === "F") {
                        split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION EN GENERAL";
                      } else if (split_str[iStr] === "G") {
                        split_str[iStr] = "OTROS";
                      }

                      valCas = true;
                    } else {
                      Browser.msgBox("Error", "Motivo de devolución no válido.", Browser.Buttons.OK);
                      valCas = false;
                    }
                  }
                }

                tMotDev = split_str.join("+");
                tMotDev = tMotDev.toUpperCase();
                var respuestamail = "yes"; //Browser.msgBox("Correo de Devolución", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                if (respuestamail === "yes") {
                  if (Correo != "SIN CORREO") {
                    var recipient = Correo;
                  } else if (Correo === "SIN CORREO") {
                    var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                  }

                  if (recipient === "cancel") {
                    return;
                  }

                  if (Cliente === "GRUPO ECONÓMICO") {
                    var subject = "SIO OC: " + Grupo + " - Devolución de la Operación " + CodSol;
                  } else {
                    var subject = "SIO OC: " + Cliente + " - Devolución de la Operación " + CodSol;
                  }

                  var body = "Riesgos ha devuelto la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                  var tCasCorreo = tMotDev.split('+').join(" || ");
                  body = body + "\nCasuística de Devolución: " + tCasCorreo;
                  var options = {
                    cc: correoGOF + ", " + correoAsist + ", siom@bbva.com"
                  };
                  MailApp.sendEmail(recipient, subject, body, options); //HERE DWP Devolución (2)

                  var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                  var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com"; //comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com 

                  var subjectDWP = CodSol + "##Devolución (2)##" + AnAsig + "##" + fecha + "##" + CodCentral;
                  var bodyDWP = Oper;
                  MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                  sheetWF.getRange(fEncontrada, cCas).setValue(tMotDev);
                }

                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                  var valCodCentEnc = false;
                  var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                  for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                    var codCentPF = arrayPF[i][cCodCentPF - 1];
                    var codSolPF = arrayPF[i][cCodSolPF - 1]; //var nuevaFV = "DEVUELTO"
                    //var tMontSan = "DEVUELTO"

                    if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                      valCodCentEnc = true; //if(tMontSan === "DENEGADO" || tMontSan === "DEVUELTO"){nuevaFV === tMontSan}
                      //sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)

                      sheetPF.getRange(i + fInicioPF, cTraPF).setValue("NO"); //sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)

                      break;
                    }
                  }

                  if (valCodCentEnc === false) {
                    Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                    var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                    var subjectE = "Operación no encontrada en la base de Líneas.";
                    var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                    MailApp.sendEmail(recipientE, subjectE, bodyE);
                  }
                }

                sheetWF.getRange(fEncontrada, c2daDev).setValue(fecha);
                Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                celda.setBackgroundRGB(85, 199, 104);
                ejecfWorkflowRiesgos(visor);
                return;
              } /////////Segunda ronda de preguntas


              if (headingWF === "Fecha Sanción") {
                /*if(Cliente === "GRUPO ECONÓMICO"){
                  var valCodRelacionado = false
                  while(valCodRelacionado === false){
                    valCodRelacionado = true
                    
                    var codRelacionado = Browser.inputBox("Registrar clientes del grupo y sus respectivos montos sancionados de líneas", "Digite los códigos centrales relacionados al grupo y su respectivo monto sancionado de línas en miles de USD. Si fuera una devolución o denegación, el monto sería 0. \\n\\nDeben estar separados con la siguiente estructura: Código Central 1//Monto Sancionado 1;Código Central 2//Monto Sancionado 2. \\n\\nEjemplo: 00399280//10000;20728211//8000.", Browser.Buttons.OK_CANCEL)
                    codRelacionado = codRelacionado.toString()
                    if(codRelacionado === "" || codRelacionado === "cancel"){
                      return;
                    }
                    
                    
                    var valFaltaSep1 = codRelacionado.indexOf("//");
                    if(valFaltaSep1 < 0){
                      valCodRelacionado = false
                    }
                    
                    var valFaltaSep2 = codRelacionado.indexOf(";");
                    if(valFaltaSep2 < 0){
                      valCodRelacionado = false
                    }
                    
                    var valFaltaSep3 = codRelacionado.indexOf(",");
                    if(valFaltaSep3 > -1){
                      valCodRelacionado = false
                    }
                    
                    if(valCodRelacionado === false){
                      Browser.msgBox("No se ha encontrado el separador // o el separador ;. \\n\\nRecuerde que la estructura es: Código Central 1 // Monto Sancionado 1; Código Central 2 // Monto Sancionado 2")
                    }
                  }
                }*/
                var valAmb = false;

                while (valAmb === false) {
                  valAmb = true;
                  var amb = Browser.inputBox("Registrar Nivel de Aprobación", "Digite la letra correspondiente a un nivel de aprobación de la siguiente lista:\\nA. ANALISTA\\nB. JEFE DE GRUPO\\nC. JEFE DE EQUIPO\\nD. SUBGERENTE\\nE. GERENTE DE UNIDAD\\nF. GERENTE DE ÁREA\\nG. COMITÉ DE CONTRASTE \\nH. CTO \\nI. CEC\\nJ. WCRMC\\nK. GCRMC\\nL. RECONDUCCIÓN\\nM. ECCMWO", Browser.Buttons.OK_CANCEL);
                  var tempAmb = amb;

                  if (amb === "") {
                    return;
                  }

                  if (amb === "cancel") {
                    return;
                  }

                  amb = amb.toUpperCase();

                  if (amb === "A") {
                    amb = "ANALISTA";
                  } else if (amb === "B") {
                    amb = "JEFE DE GRUPO";
                  } else if (amb === "C") {
                    amb = "JEFE DE EQUIPO";
                  } else if (amb === "D") {
                    amb = "SUBGERENTE";
                  } else if (amb === "E") {
                    amb = "GERENTE DE UNIDAD";
                  } else if (amb === "F") {
                    amb = "GERENTE DE ÁREA";
                  } else if (amb === "G") {
                    amb = "COMITÉ DE CONTRASTE";
                  } else if (amb === "H") {
                    amb = "CTO";
                  } else if (amb === "I") {
                    amb = "CEC";
                  } else if (amb === "J") {
                    amb = "WCRMC";
                  } else if (amb === "K") {
                    amb = "GCRMC";
                  } else if (amb === "L") {
                    amb = "RECONDUCCIÓN";
                  } else if (amb === "M") {
                    amb = "ECCMWO";
                  } else {
                    valAmb = false;
                    Browser.msgBox("Error", "No se ingresó una letra válida de la lista de nivel de aprobación.", Browser.Buttons.OK);
                    return;
                  }
                }

                amb = amb.toUpperCase();
                var ambito = amb;
                var ambitoDWP = ambito;
                var valAmb = false;

                while (valAmb === false) {
                  valAmb = true;
                  var amb = Browser.inputBox("Registrar Ámbito", "Digite la letra correspondiente al ámbito de la siguiente lista:\\nA. LOCAL\\nB. GCR ARGENTINA\\nC. GCR BRASIL\\nD. GCR COLOMBIA\\nE. GCR COMPASS\\nF. GCR MEXICO\\nG. GCR NY\\nH. GCR PARAGUAY\\nI. GCR PANAMA\\nJ. GCR URUGUAY\\nK. RPM ALEMANIA\\nL. RPM BELGICA\\nM. RPM COREA\\nN. RPM ESPAÑA\\nO. RPM FRANCIA\\nP. RPM HK\\nQ. RPM ITALIA\\nR. RPM JAPON\\nS. RPM SINGAPUR\\nT. RPM UK", Browser.Buttons.OK_CANCEL);
                  var tempAmb = amb;

                  if (amb === "") {
                    return;
                  }

                  if (amb === "cancel") {
                    return;
                  }

                  amb = amb.toUpperCase();

                  if (amb === "A") {
                    amb = "LOCAL";
                  } else if (amb === "B") {
                    amb = "GCR ARGENTINA";
                  } else if (amb === "C") {
                    amb = "GCR BRASIL";
                  } else if (amb === "D") {
                    amb = "GCR COLOMBIA";
                  } else if (amb === "E") {
                    amb = "GCR COMPASS";
                  } else if (amb === "F") {
                    amb = "GCR MEXICO";
                  } else if (amb === "G") {
                    amb = "GCR NY";
                  } else if (amb === "H") {
                    amb = "GCR PARAGUAY";
                  } else if (amb === "I") {
                    amb = "GCR PANAMA";
                  } else if (amb === "J") {
                    amb = "GCR URUGUAY";
                  } else if (amb === "K") {
                    amb = "RPM ALEMANIA";
                  } else if (amb === "L") {
                    amb = "RPM BELGICA";
                  } else if (amb === "M") {
                    amb = "RPM COREA";
                  } else if (amb === "N") {
                    amb = "RPM ESPAÑA";
                  } else if (amb === "O") {
                    amb = "RPM FRANCIA";
                  } else if (amb === "P") {
                    amb = "RPM HK";
                  } else if (amb === "Q") {
                    amb = "RPM ITALIA";
                  } else if (amb === "R") {
                    amb = "RPM JAPON";
                  } else if (amb === "S") {
                    amb = "RPM SINGAPUR";
                  } else if (amb === "T") {
                    amb = "RPM UK";
                  } else {
                    valAmb = false;
                    Browser.msgBox("Error", "No se ingresó una letra válida de la lista de ámbitos.", Browser.Buttons.OK);
                    return;
                  }
                }

                amb = amb.toUpperCase();
                ambito = ambito + " // " + amb; //var valRating = false

                var tRating = "(vacío)";
                /*while(valRating === false){
                  var tRating = Browser.inputBox("Registrar el Rating", "Digite el rating. Ej. AAA, AA+, BBB+1, etc. \\nSi no aplicara, digite SIN RATING. ", Browser.Buttons.OK_CANCEL);
                  if(tRating === "cancel"){
                    return
                  }
                  else if(tRating === ""){
                    return
                  }
                  valRating = true
                  tRating = tRating.toUpperCase()
                  if (tRating === "SIN RATING" || tRating === "AAA" || tRating === "AA+" || tRating === "AA" || tRating === "AA-" || tRating === "A+" || tRating === "A" || tRating === "A-" || tRating === "BBB+1" || tRating === "BBB+2" || tRating === "BBB1" || tRating === "BBB2" || tRating === "BBB-1" || tRating === "BBB-2" || tRating === "BB+1" || tRating === "BB+2" || tRating === "BB1" || tRating === "BB2" || tRating === "BB-1" || tRating === "BB-2" || tRating === "B+1" || tRating === "B+2" || tRating === "B+3" || tRating === "B1" || tRating === "B2" || tRating === "B3" || tRating === "B-1" || tRating === "B-2" || tRating === "B-3" || tRating === "CCC+" || tRating === "CCC" || tRating === "CCC-" || tRating === "CC+" || tRating === "CC" || tRating === "CC-" ){/*SIGUE A LA SIGUIENTE*/

                /*}
                else{  
                valRating = false
                Browser.msgBox("Error", "Digite un rating válido.", Browser.Buttons.OK)
                }   
                }*/

                var yEEFF = "(vacío)";
                var herramienta = "(vacío)";
                /*if(tRating === "SIN RATING"){
                  var yEEFF = "SIN RATING"
                  var herramienta = "SIN RATING"
                  }
                else{
                  var valYEEFF = false
                  while(valYEEFF === false){
                    var yEEFF = Browser.inputBox("Registrar el año del EEFF", "Digite el año del EEFF. Ej. 2015, 2016, etc.", Browser.Buttons.OK_CANCEL);
                    if(yEEFF === "cancel"){
                      return
                    }
                    else if(yEEFF === ""){
                      return
                    }
                    valYEEFF = true
                    if (isNaN(yEEFF) != true){
                      if (yEEFF >= 2015 && yEEFF <= 2020){/*SIGUE A LA SIGUIENTE*/

                /*}
                else{  
                valYEEFF = false
                Browser.msgBox("Error", "Digite un año válido.", Browser.Buttons.OK)
                }   
                }
                else{
                valYEEFF = false
                Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                }   
                }
                var valHerramienta = false
                while(valHerramienta === false){
                valHerramienta = true
                var herramienta = Browser.inputBox("Registrar Herramienta", "Digite la letra correspondiente a la herramienta de la siguiente lista:\\nA. NACAR\\nB. RA", Browser.Buttons.OK_CANCEL)
                if(herramienta === ""){return}
                if(herramienta === "cancel"){return}
                herramienta = herramienta.toUpperCase()
                if(herramienta === "A"){herramienta = "NACAR"}
                else if(herramienta === "B"){herramienta = "RA"}
                else{
                valHerramienta = false
                Browser.msgBox("Error", "No se ingresó una letra válida de la lista de herramientas.", Browser.Buttons.OK)
                }
                }
                }*/

                var buro = "(vacío)";
                /*var valBuro = false
                while(valBuro === false){
                  valBuro = true
                  var buro = Browser.inputBox("Registrar Buró", "Digite la letra correspondiente al buró de la siguiente lista:\\nA. BURO G1\\nB. BURO G2\\nC. BURO G3\\nD. BURO G4\\nE. BURO G5\\nF. BURO G6\\nG. BURO G7 \\nH. BURO G8 \\nI. NO BANCARIZADO", Browser.Buttons.OK_CANCEL);
                  
                  if(buro === ""){return}
                  if(buro === "cancel"){return}
                  
                  buro = buro.toUpperCase()
                  if(buro === "A"){buro = "BURO G1"}
                  else if(buro === "B"){buro = "BURO G2"}
                  else if(buro === "C"){buro = "BURO G3"}
                  else if(buro === "D"){buro = "BURO G4"}
                  else if(buro === "E"){buro = "BURO G5"}
                  else if(buro === "F"){buro = "BURO G6"}
                  else if(buro === "G"){buro = "BURO G7"}
                  else if(buro === "H"){buro = "BURO G8"}
                  else if(buro === "I"){buro = "NO BANCARIZADO"}
                  else{
                    valBuro = false
                    Browser.msgBox("Error", "No se ingresó una letra válida de la lista de burós.", Browser.Buttons.OK)
                    return;
                  }
                }*/

                var estrategia = "(vacío)";
                /*var valEstrategia = false
                while(valEstrategia === false){
                  valEstrategia = true
                  var estrategia = Browser.inputBox("Registrar Estrategia", "Digite la letra correspondiente a la estrategia de la siguiente lista:\\nA. Liderar\\nB. Crecer\\nC. Vigilar\\nD. Reducir\\nE. Extinguir\\nF. A Potenciar\\nG. No Sugerido \\nH. Sin Estrategia", Browser.Buttons.OK_CANCEL);
                  
                  if(estrategia === ""){return}
                  if(estrategia === "cancel"){return}
                  
                  estrategia = estrategia.toUpperCase()
                  if(estrategia === "A"){estrategia = "Liderar"}
                  else if(estrategia === "B"){estrategia = "Crecer"}
                  else if(estrategia === "C"){estrategia = "Vigilar"}
                  else if(estrategia === "D"){estrategia = "Reducir"}
                  else if(estrategia === "E"){estrategia = "Extinguir"}
                  else if(estrategia === "F"){estrategia = "A potenciar"}
                  else if(estrategia === "G"){estrategia = "No Sugerido"}
                  else if(estrategia === "H"){estrategia = "(en blanco)"}
                  else{
                    valEstrategia = false
                    Browser.msgBox("Error", "No se ingresó una letra válida de la lista de estrategias.", Browser.Buttons.OK)
                    return;
                  }
                }*/

                var valTSan = false;

                while (valTSan === false) {
                  var tSan = Browser.inputBox("Registrar el Tipo de Sanción", "Digite el número correspondiente a un tipo de sanción de la siguiente lista:\\n1. Aprobado Sin Modificación\\n2. Denegado\\n3. Devuelto\\n4. Aprobado Con Modificación", Browser.Buttons.OK_CANCEL);
                  valTSan = true;
                  var temptSan = tSan;

                  if (tSan === "") {
                    return;
                  }

                  if (tSan === "cancel") {
                    return;
                  }

                  if (tSan == 1) {
                    tSan = "Aprobado SM";
                  } else if (tSan == 2) {
                    tSan = "Denegado";
                  } else if (tSan == 3) {
                    tSan = "Devuelto";
                  } else if (tSan == 4) {
                    tSan = "Aprobado CM";
                  } else {
                    Browser.msgBox("Error", "No se digitó un número válido de la lista de tipos de sanción.", Browser.Buttons.OK);
                    valTSan = false;
                  }
                }

                var tSanDWP = tSan;
                tSan = tSan.toUpperCase();

                switch (tSan) {
                  case "APROBADO SM":
                    var tipoSan = tSan;
                    var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue();
                    break;

                  case "APROBADO CM":
                    var tipoSan = tSan;
                    var valCasImp = false;
                    var valCas = false;

                    while (valCas === false) {
                      var tCas = Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. Plazo\\nB. Garantía\\nC. Importe\\nD. Condicionantes Previas al Desembolso\\nE. Otros\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                      if (tCas === "cancel") {
                        return;
                      }

                      tCas = tCas.toUpperCase();
                      var split_str = tCas.split("+");

                      for (var iStr = 0; iStr < split_str.length; iStr++) {
                        if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E") {
                          if (split_str[iStr] === "A") {
                            split_str[iStr] = "Plazo";
                          } else if (split_str[iStr] === "B") {
                            split_str[iStr] = "Garantía";
                          } else if (split_str[iStr] === "C") {
                            valCasImp = true;
                            split_str[iStr] = "Importe";
                          } else if (split_str[iStr] === "D") {
                            split_str[iStr] = "Condicionantes Previas al Desembolso";
                          } else if (split_str[iStr] === "E") {
                            split_str[iStr] = "Otros";
                          }

                          valCas = true;
                        } else {
                          Browser.msgBox("Error", "Casuística no válida.", Browser.Buttons.OK);
                          valCas = false;
                          break;
                        }
                      }
                    }

                    tCas = split_str.join("+");
                    tCas = tCas.toUpperCase();
                    var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue();

                    if (valCasImp === true) {
                      var valTMont = false;

                      while (valTMont === false) {
                        var tMontSan = Browser.inputBox("Registrar el Monto Sancionado", "Registre el monto sancionado de esta operación en miles de US$.", Browser.Buttons.OK_CANCEL);

                        if (tMontSan === "cancel") {
                          return;
                        } else if (tMontSan === "") {
                          return;
                        }

                        valTMont = true;

                        if (isNaN(tMontSan) != false) {
                          valTMont = false;
                          Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK);
                        }
                      }
                    } else {
                      var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue();
                    } //sheetWF.getRange(fEncontrada,cCas).setValue(tCas)


                    break;

                  case "DEVUELTO":
                    var tipoSan = tSan;
                    var tMontSan = "DEVUELTO";
                    var valCas = false;

                    while (valCas === false) {
                      var tMotDev = Browser.inputBox("Casuística de Devolución", "Digite la opción de la casuística de la siguiente lista:\\nA. DEVUELTA A SOLICITUD DE LA OFICINA\\nB. DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO\\nC. DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA\\nD. DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING\\nE. DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR\\nF. DEVUELTA POR FALTA DE INFORMACION EN GENERAL\\nG. OTROS\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                      if (tMotDev === "cancel") {
                        return;
                      }

                      tMotDev = tMotDev.toUpperCase();
                      var split_str = tMotDev.split("+");

                      for (var iStr = 0; iStr < split_str.length; iStr++) {
                        if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G") {
                          if (split_str[iStr] === "A") {
                            split_str[iStr] = "DEVUELTA A SOLICITUD DE LA OFICINA";
                          } else if (split_str[iStr] === "B") {
                            split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO";
                          } else if (split_str[iStr] === "C") {
                            valCasImp = true;
                            split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA";
                          } else if (split_str[iStr] === "D") {
                            split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING";
                          } else if (split_str[iStr] === "E") {
                            split_str[iStr] = "DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR";
                          } else if (split_str[iStr] === "F") {
                            split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION EN GENERAL";
                          } else if (split_str[iStr] === "G") {
                            split_str[iStr] = "OTROS";
                          }

                          valCas = true;
                        } else {
                          Browser.msgBox("Error", "Motivo de devolución no válido.", Browser.Buttons.OK);
                          valCas = false;
                        }
                      }
                    }

                    tMotDev = split_str.join("+");
                    tMotDev = tMotDev.toUpperCase();
                    break;

                  case "DENEGADO":
                    var tipoSan = tSan;
                    var tMontSan = "DENEGADO";
                    var valCasDeneg = false;

                    while (valCasDeneg === false) {
                      var tCasDeneg = Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. ESTRUCTURA ECONÓMICA FINANCIERA HISTORICA NO FAVORABLE\\nB. ALTO NIVEL DE ENDEUDAMIENTO/NO HAY CAPACIDAD DE PAGO PARA EL RIESGO PROPUESTO\\nC. DETERIORO DE LAS CIFRAS DE SITUACIÓN\\nD. ALTO NIVEL DE ENDEUDAMIENTO Y DEUDA ESTRUCTURAL ACTUAL DESCALZADA\\nE. ALERTAS EN EL COMPORTAMIENTO DE PAGO/ALERTAS EN LOS INDICES DE GESTION\\nF. EMPRESAS VINCULADAS /ACCIONSITAS CON ALERTAS NEGATIVAS.\\nG. INVERSIÓN DE ACTIVO FIJO NO CORRESPONDE AL CORE BUSSINES DEL NEGOCIO\\nH. ALERTAS EN EL SECTOR\\nI. ESTRUCTURA DE LA OPERACIÓN (PRODUCTO / PLAZO / CONDICIONES) SUPERA LA DIMENSIÓN DEL NEGOCIO\\nJ. PFA RECIENTEMENTE SANCIONADO Y/O CLIENTE CON TECHO DE RIESGO\\nK. SOLICITUD DE CP NO CORRESPONDE AL DIMENSIONAMIENTO DE LAS NECESIDADES OPERATIVAS\\nL. NO CUENTA CON EEFF DE CIERRE DE EVIDENCIEN GENERACIÓN DE CAJA SUFICIENTE\\nM. NO CUMPLE CON CONVENANTS/CONDICIONES ESTABLECIDAS EN PF/SANCION ANTERIOR\\nN. OTROS\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                      if (tCasDeneg === "cancel") {
                        return;
                      }

                      tCasDeneg = tCasDeneg.toUpperCase();
                      var split_str = tCasDeneg.split("+");

                      for (var iStr = 0; iStr < split_str.length; iStr++) {
                        if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L" || split_str[iStr] === "M" || split_str[iStr] === "N") {
                          if (split_str[iStr] === "A") {
                            split_str[iStr] = "ESTRUCTURA ECONÓMICA FINANCIERA HISTORICA NO FAVORABLE";
                          } else if (split_str[iStr] === "B") {
                            split_str[iStr] = "ALTO NIVEL DE ENDEUDAMIENTO/NO HAY CAPACIDAD DE PAGO PARA EL RIESGO PROPUESTO";
                          } else if (split_str[iStr] === "C") {
                            valCasImp = true;
                            split_str[iStr] = "DETERIORO DE LAS CIFRAS DE SITUACIÓN";
                          } else if (split_str[iStr] === "D") {
                            split_str[iStr] = "ALTO NIVEL DE ENDEUDAMIENTO Y DEUDA ESTRUCTURAL ACTUAL DESCALZADA";
                          } else if (split_str[iStr] === "E") {
                            split_str[iStr] = "ALERTAS EN EL COMPORTAMIENTO DE PAGO/ALERTAS EN LOS INDICES DE GESTION";
                          } else if (split_str[iStr] === "F") {
                            split_str[iStr] = "EMPRESAS VINCULADAS /ACCIONSITAS CON ALERTAS NEGATIVAS";
                          } else if (split_str[iStr] === "G") {
                            split_str[iStr] = "INVERSIÓN DE ACTIVO FIJO NO CORRESPONDE AL CORE BUSSINES DEL NEGOCIO";
                          } else if (split_str[iStr] === "H") {
                            split_str[iStr] = "ALERTAS EN EL SECTOR";
                          } else if (split_str[iStr] === "I") {
                            split_str[iStr] = "ESTRUCTURA DE LA OPERACIÓN (PRODUCTO / PLAZO / CONDICIONES) SUPERA LA DIMENSIÓN DEL NEGOCIO";
                          } else if (split_str[iStr] === "J") {
                            split_str[iStr] = "PFA RECIENTEMENTE SANCIONADO Y/O CLIENTE CON TECHO DE RIESGO";
                          } else if (split_str[iStr] === "K") {
                            split_str[iStr] = "SOLICITUD DE CP NO CORRESPONDE AL DIMENSIONAMIENTO DE LAS NECESIDADES OPERATIVAS";
                          } else if (split_str[iStr] === "L") {
                            split_str[iStr] = "NO CUENTA CON EEFF DE CIERRE DE EVIDENCIEN GENERACIÓN DE CAJA SUFICIENTE";
                          } else if (split_str[iStr] === "M") {
                            split_str[iStr] = "NO CUMPLE CON CONVENANTS/CONDICIONES ESTABLECIDAS EN PF/SANCION ANTERIOR";
                          } else if (split_str[iStr] === "N") {
                            split_str[iStr] = "OTROS";
                          }

                          valCasDeneg = true;
                        } else {
                          Browser.msgBox("Error", "Casuística no válida.", Browser.Buttons.OK);
                          valCasDeneg = false;
                          break;
                        }
                      }
                    }

                    tCasDeneg = split_str.join("+");
                    tCasDeneg = tCasDeneg.toUpperCase();
                    break;

                  default:
                    Browser.msgBox("Error", "No se ingresó un número válido de la lista de tipos de sanción.", Browser.Buttons.OK);
                    ss.toast("Digite el número correspondiente para el tipo de sanción; por ejemplo, para Aprobado Sin Modificación coloque un 1.", "Tip", 8);
                    return;
                } //CONDICMITIG    


                var condicMitig = "(vacío)"; //              if(tipoSan === "APROBADO CM" || tipoSan === "APROBADO SM"){
                //                var condic = sheetWF.getRange(fEncontrada,cCondicion).getValue()
                //                
                //                if (tRating === "AAA" || tRating === "AA+" || tRating === "AA" || tRating === "AA-" || tRating === "A+" || tRating === "A" || tRating === "A-" || tRating === "BBB+1" || tRating === "BBB+2" || tRating === "BBB1" || tRating === "BBB2" || tRating === "BBB-1" || tRating === "BBB-2" || tRating === "BB+1" || tRating === "BB+2" || tRating === "BB1" || tRating === "BB2" || tRating === "BB-1" || tRating === "BB-2"){
                //                  var condicMitig = "SIN MITIGANTES"
                //                  }
                //                else if (tRating === "B+1" || tRating === "B+2" || tRating === "B+3" || tRating === "B1" || tRating === "B2" || tRating === "B3" || tRating === "B-1" || tRating === "B-2" || tRating === "B-3" || tRating === "CCC+" || tRating === "CCC" || tRating === "CCC-" || tRating === "CC+" || tRating === "CC" || tRating === "CC-" ){
                //                  
                //                  var valCodCentralOpc1 = false
                //                  
                //                  var valCondicOpc1 = false
                //                  var valCondicOpc2 = false
                //                  var valCondicOpc3 = false
                //                  var valCondicOpc4 = false
                //                  var valCondicOpc5 = false
                //                  var valCondicOpc6 = false
                //                  
                //                  var valCondicMitig = false
                //                  while(valCondicMitig === false){
                //                    var condicMitig = Browser.inputBox("Registrar Mitigantes", "Digite la opción de la mitigante de la siguiente lista:\\nA. <=B+1 (CESION DE FLUJOS COMO MEDIO DE PAGO)\\nB. <=B+1 (CON ARRENDATARIO SUTITUTO)\\nC. <=B+1 (CON POLIZA ENDOSADA AL BANCO)\\nD. <=B+1 (CON FIANZA SOLIDARIA)\\nE. <=B+1 (GAR. DEPÓSITO)\\nF. <=B+1 (GAR. FIDEICOMISO DE ACTIVOS Y/O FLUJOS)\\nG. <=B+1 (GAR. HIPOTECARIA EN TRAMITE)\\nH. <=B+1 (GAR. SBLC)\\nI. <=B+1 (GAR. WARRANT)\\nJ. <=B+1 (GAR.HIPOTECARIA)\\nK. <=B+1 (MEJORA DE GARANTÍAS)\\nL. <=B+1 (REPERFILAMIENTO DE DEUDA)\\nM. RIESGO EN BLANCO\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);
                //                    if(condicMitig === "cancel"){return}
                //                    condicMitig = condicMitig.toUpperCase()
                //                    
                //                    var split_str = condicMitig.split("+");
                //                    
                //                    valCondicMitig = true
                //                    for(var iStr = 0; iStr < split_str.length; iStr++){
                //                      if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L"|| split_str[iStr] === "M"){/*No pasa nada*/}
                //                      else{
                //                        Browser.msgBox("Error", "Mitigante no válida.", Browser.Buttons.OK)
                //                        valCondicMitig = false
                //                        break;
                //                      }
                //                    }
                //                    
                //                    if(valCondicMitig === true){
                //                      for(var iStr = 0; iStr < split_str.length; iStr++){
                //                        if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L" || split_str[iStr] === "M" ){
                //                          if(split_str[iStr] === "A"){
                //                            split_str[iStr] = "<=B+1 (CESION DE FLUJOS COMO MEDIO DE PAGO)"
                //                          }
                //                          else if(split_str[iStr] === "B"){
                //                            split_str[iStr] = "<=B+1 (CON ARRENDATARIO SUTITUTO)"
                //                          }
                //                          else if(split_str[iStr] === "C"){
                //                            split_str[iStr] = "<=B+1 (CON POLIZA ENDOSADA AL BANCO)"
                //                          }
                //                          else if(split_str[iStr] === "D"){
                //                            valCodCentralOpc1 = true
                //                            if(valCodCentralOpc1 === true){
                //                              var valTCodCentralOpc1 = false
                //                              while(valTCodCentralOpc1 === false){
                //                                var codCentralOpc1 = Browser.inputBox("Registrar el código central", "Registre el código central relacionado a esta opción de mitigante (incluyendo 0s, por ejemplo: 00399280).", Browser.Buttons.OK_CANCEL);
                //                                if(codCentralOpc1 === "cancel"){
                //                                  return
                //                                }
                //                                else if(codCentralOpc1 === ""){
                //                                  return
                //                                }
                //                                valTCodCentralOpc1 = true
                //                                if (codCentralOpc1.length != 8){
                //                                  valTCodCentralOpc1 = false
                //                                  Browser.msgBox("Error", "Digite un código central válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (CON FIANZA SOLIDARIA), " + codCentralOpc1
                //                          }
                //                          else if(split_str[iStr] === "E"){
                //                            valCondicOpc1 = true
                //                            if(valCondicOpc1 === true){
                //                              var valTCondicOpc1 = false
                //                              while(valTCondicOpc1 === false){
                //                                var condicOpc1 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. DEPÓSITO)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                //                                if(condicOpc1 === "cancel"){
                //                                  return
                //                                }
                //                                else if(condicOpc1 === ""){
                //                                  return
                //                                }
                //                                valTCondicOpc1 = true
                //                                if (isNaN(condicOpc1) != false){
                //                                  valTCondicOpc1 = false
                //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (GAR. DEPÓSITO), " + condicOpc1 + "%"
                //                          }
                //                          else if(split_str[iStr] === "F"){
                //                            valCondicOpc2 = true
                //                            if(valCondicOpc2 === true){
                //                              var valTCondicOpc2 = false
                //                              while(valTCondicOpc2 === false){
                //                                var condicOpc2 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. FIDEICOMISO DE ACTIVOS Y/O FLUJOS)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                //                                if(condicOpc2 === "cancel"){
                //                                  return
                //                                }
                //                                else if(condicOpc2 === ""){
                //                                  return
                //                                }
                //                                valTCondicOpc2 = true
                //                                if (isNaN(condicOpc2) != false){
                //                                  valTCondicOpc2 = false
                //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (GAR. FIDEICOMISO DE ACTIVOS Y/O FLUJOS), " + condicOpc2 + "%"
                //                          }
                //                          else if(split_str[iStr] === "G"){
                //                            valCondicOpc3 = true
                //                            if(valCondicOpc3 === true){
                //                              var valTCondicOpc3 = false
                //                              while(valTCondicOpc3 === false){
                //                                var condicOpc3 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. HIPOTECARIA EN TRAMITE)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                //                                if(condicOpc3 === "cancel"){
                //                                  return
                //                                }
                //                                else if(condicOpc3 === ""){
                //                                  return
                //                                }
                //                                valTCondicOpc3 = true
                //                                if (isNaN(condicOpc3) != false){
                //                                  valTCondicOpc3 = false
                //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (GAR. HIPOTECARIA EN TRAMITE), " + condicOpc3 + "%"
                //                          }
                //                          else if(split_str[iStr] === "H"){
                //                            valCondicOpc4 = true
                //                            if(valCondicOpc4 === true){
                //                              var valTCondicOpc4 = false
                //                              while(valTCondicOpc4 === false){
                //                                var condicOpc4 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. SBLC)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                //                                if(condicOpc4 === "cancel"){
                //                                  return
                //                                }
                //                                else if(condicOpc4 === ""){
                //                                  return
                //                                }
                //                                valTCondicOpc4 = true
                //                                if (isNaN(condicOpc4) != false){
                //                                  valTCondicOpc4 = false
                //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (GAR. SBLC), " + condicOpc4 + "%"
                //                          }
                //                          else if(split_str[iStr] === "I"){
                //                            valCondicOpc5 = true
                //                            if(valCondicOpc5 === true){
                //                              var valTCondicOpc5 = false
                //                              while(valTCondicOpc5 === false){
                //                                var condicOpc5 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. WARRANT)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                //                                if(condicOpc5 === "cancel"){
                //                                  return
                //                                }
                //                                else if(condicOpc5 === ""){
                //                                  return
                //                                }
                //                                valTCondicOpc5 = true
                //                                if (isNaN(condicOpc5) != false){
                //                                  valTCondicOpc5 = false
                //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (GAR. WARRANT), " + condicOpc5 + "%"
                //                          }
                //                          else if(split_str[iStr] === "J"){
                //                            valCondicOpc6 = true
                //                            if(valCondicOpc6 === true){
                //                              var valTCondicOpc6 = false
                //                              while(valTCondicOpc6 === false){
                //                                var condicOpc6 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR.HIPOTECARIA)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                //                                if(condicOpc6 === "cancel"){
                //                                  return
                //                                }
                //                                else if(condicOpc6 === ""){
                //                                  return
                //                                }
                //                                valTCondicOpc6 = true
                //                                if (isNaN(condicOpc6) != false){
                //                                  valTCondicOpc6 = false
                //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            
                //                            split_str[iStr] = "<=B+1 (GAR.HIPOTECARIA), " + condicOpc6 + "%"
                //                          }
                //                          else if(split_str[iStr] === "K"){
                //                            split_str[iStr] = "<=B+1 (MEJORA DE GARANTÍAS)"
                //                          }
                //                          else if(split_str[iStr] === "L"){
                //                            split_str[iStr] = "<=B+1 (REPERFILAMIENTO DE DEUDA)"
                //                          }
                //                          else if(split_str[iStr] === "M"){
                //                            split_str[iStr] = "RIESGO EN BLANCO"
                //                          }
                //                          
                //                          valCondicMitig = true   
                //                        }
                //                      }
                //                    }
                //                  }
                //                  
                //                  condicMitig = split_str.join(" // ")                  
                //                  condicMitig = condicMitig.toUpperCase()
                //                }
                //                else if(tRating === "SIN RATING"){    
                //                  
                //                  var valCodCentralOpc1 = false
                //                  
                //                  var valCondicMitig = false
                //                  while(valCondicMitig === false){
                //                    var condicMitig = Browser.inputBox("Registrar Mitigantes", "Digite la opción de la mitigante de la siguiente lista:\\nA. SIN.RAT 100% GARANTIZADO\\nB. SIN.RAT FIANZA SOLIDARIA\\nC. SIN.RAT PROJECT FINANCE\\nD. SIN.RAT Ventas < 2.8MM\\nE. RIESGO EN BLANCO\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);
                //                    if(condicMitig === "cancel"){return}
                //                    condicMitig = condicMitig.toUpperCase()
                //                    
                //                    var split_str = condicMitig.split("+");
                //                    
                //                    valCondicMitig = true
                //                    for(var iStr = 0; iStr < split_str.length; iStr++){
                //                      if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D"|| split_str[iStr] === "E"){/*No pasa nada*/}
                //                      else{
                //                        Browser.msgBox("Error", "Mitigante no válida.", Browser.Buttons.OK)
                //                        valCondicMitig = false
                //                        break;
                //                      }
                //                    }
                //                    
                //                    if(valCondicMitig === true){
                //                      for(var iStr = 0; iStr < split_str.length; iStr++){
                //                        if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D"|| split_str[iStr] === "E"){
                //                          if(split_str[iStr] === "A"){
                //                            split_str[iStr] = "SIN.RAT 100% GARANTIZADO"
                //                          }
                //                          else if(split_str[iStr] === "B"){
                //                            valCodCentralOpc1 = true
                //                            if(valCodCentralOpc1 === true){
                //                              var valTCodCentralOpc1 = false
                //                              while(valTCodCentralOpc1 === false){
                //                                var codCentralOpc1 = Browser.inputBox("Registrar el código central", "Registre el código central relacionado a esta opción de mitigante (incluyendo 0s, por ejemplo: 00399280).", Browser.Buttons.OK_CANCEL);
                //                                if(codCentralOpc1 === "cancel"){
                //                                  return
                //                                }
                //                                else if(codCentralOpc1 === ""){
                //                                  return
                //                                }
                //                                valTCodCentralOpc1 = true
                //                                if (codCentralOpc1.length != 8){
                //                                  valTCodCentralOpc1 = false
                //                                  Browser.msgBox("Error", "Digite un código central válido.", Browser.Buttons.OK)
                //                                }
                //                              }
                //                            }
                //                            split_str[iStr] = "SIN.RAT FIANZA SOLIDARIA, " + codCentralOpc1
                //                          }
                //                          else if(split_str[iStr] === "C"){
                //                            split_str[iStr] = "SIN.RAT PROJECT FINANCE"
                //                          }
                //                          else if(split_str[iStr] === "D"){
                //                            split_str[iStr] = "SIN.RAT Ventas < 2.8MM"
                //                          }
                //                          else if(split_str[iStr] === "E"){
                //                            split_str[iStr] = "RIESGO EN BLANCO"
                //                          }
                //                          
                //                          valCondicMitig = true   
                //                        }
                //                      }
                //                    }
                //                  }
                //                  
                //                  condicMitig = split_str.join(" // ")                  
                //                  condicMitig = condicMitig.toUpperCase()
                //                }
                //                
                //                var split_str = condic.split("+");
                //                var fen = false
                //                for(var iStr = 0; iStr < split_str.length; iStr++){
                //                  if(split_str[iStr] === "FENOMENO DEL NIÑO"){
                //                    fen = true
                //                  }
                //                }
                //                
                //                if(fen === true){
                //                  if(condicMitig === "SIN MITIGANTES"){
                //                    condicMitig = "SIN EEFF SUNAT 2016 FEN"
                //                  }
                //                  else{
                //                    condicMitig = condicMitig + " // SIN EEFF SUNAT 2016 FEN"
                //                  }
                //                }
                //              }
                //CONDICMITIG

                var respuestamail = "yes"; //Browser.msgBox("Correo de Sanción", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                if (respuestamail === "yes") {
                  if (Correo != "SIN CORREO") {
                    var recipient = Correo;
                  } else if (Correo === "SIN CORREO") {
                    var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                  }

                  if (recipient === "cancel") {
                    return;
                  }

                  if (Cliente === "GRUPO ECONÓMICO") {
                    var subject = "SIO OC: " + Grupo + " - Operación " + CodSol + " Sancionada";
                  } else {
                    var subject = "SIO OC: " + Cliente + " - Operación " + CodSol + " Sancionada";
                  }

                  var body = "Riesgos ha sancionado la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp + "\nMonto de Sanción (Miles de US$): " + tMontSan + "\nÁmbito de Sanción: " + ambito + "\nTipo de Sanción: " + tipoSan;

                  if (tipoSan === "APROBADO CM") {
                    var tCasCorreo = tCas.split('+').join(" || ");
                    body = body + "\nCasuística de Modificación: " + tCasCorreo;
                  } else if (tipoSan === "DENEGADO") {
                    var tCasCorreo = tCasDeneg.split('+').join(" || ");
                    body = body + "\nCasuística de Denegación: " + tCasCorreo;
                  } else if (tipoSan === "DEVUELTO") {
                    var tCasCorreo = tMotDev.split('+').join(" || ");
                    body = body + "\nCasuística de Devolución: " + tCasCorreo;
                  }

                  if (tipoOp === "LP") {
                    body = body + "\nProducto: " + sheetWF.getRange(fEncontrada, cProducto).getValue();
                    var options = {
                      cc: correoGOF + ", " + correoAsist + ", " + "siom@bbva.com, vcardena@bbva.com, aazcoytia@bbva.com, ucastillo@bbva.com, kgonzalesp@bbva.com, christian.blanch@bbva.com"
                    };
                  } else {
                    var options = {
                      cc: correoGOF + ", " + correoAsist + ", " + "siom@bbva.com"
                    };
                  } //MailApp.sendEmail(recipient,subject,body,options); Pasa después de la sanción de la línea

                }

                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                  //HERE DIGITAR LOS CÓDIGOS CENTRALES DEL GRUPO
                  if (tipoSan === "DEVUELTO" || tipoSan === "DENEGADO") {
                    var nuevaFV = tipoSan;
                  } else {
                    var valTMontLin = false;

                    while (valTMontLin === false) {
                      var tMontSanLin = tMontSan; //Browser.inputBox("Registrar el Monto Sancionado de la LÍNEA", "Registre el monto sancionado para la línea del grupo " + Grupo + " en miles de US$.", Browser.Buttons.OK_CANCEL);

                      if (tMontSanLin === "cancel") {
                        return;
                      } else if (tMontSanLin === "") {
                        return;
                      }

                      valTMontLin = true;

                      if (isNaN(tMontSanLin) != false) {
                        valTMontLin = false;
                        Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK);
                      }
                    }

                    var valFecha = false;

                    while (valFecha === false) {
                      var nuevaFV = Browser.inputBox("Fecha de Vencimiento", "Ingrese una fecha de vencimiento. Este cambio se reflejará en la base de Líneas. La fecha debe estar en formato 'dd/mm/aaaa'. Por ejemplo, 31/12/2017.", Browser.Buttons.OK_CANCEL);

                      if (nuevaFV === "cancel" || nuevaFV === "") {
                        Browser.msgBox("Ha decidido no registrar una fecha de vencimiento. La operación no será registrada.");
                        return;
                      }

                      var valFecha = isValidDate(nuevaFV);

                      if (valFecha === false) {
                        Browser.msgBox("Fecha no válida. Por favor intente de nuevo.");
                      }
                    }
                  }
                }

                if (tipoOp === "CP") {
                  var valCodCentEnc2 = false;
                  var lRowPF = Avals.filter(String).length;
                  var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                  for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                    var codCentPF = arrayPF[i][cCodCentPF - 1];
                    var codSolPF = arrayPF[i][cCodSolPF - 1];

                    if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                      valCodCentEnc2 = true;
                      var fechaVencLin = arrayPF[i][cFechaSanc]; //No se le ha puesto -1 a la columna porque estoy usando el último día

                      if (fechaVencLin === "") {
                        valCodCentEnc2 = false;
                        break;
                      } else {
                        if (fechaVencLin != "Vencidas") {
                          fechaVencLin = new Date(fechaVencLin.getFullYear(), fechaVencLin.getMonth(), fechaVencLin.getDate());
                          fechaVencLin.setMonth(fechaVencLin.getMonth() - 6);
                          var valFechaVencLin = fechaVencLin.valueOf();
                          var valFecha = fecha.valueOf();
                          break;
                        }
                      }
                    }
                  }

                  if (valCodCentEnc2 === true) {
                    if (fechaVencLin === "Vencidas") {
                      sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(2);
                    } else {
                      if (valFechaVencLin < valFecha) {
                        sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(4);
                      } else {
                        sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(3);
                      }
                    }
                  } else {
                    sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(1);
                  }
                }

                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA") {
                  var valCodCentEnc2 = false;
                  var lRowPF = Avals.filter(String).length;
                  var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                  for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                    var codCentPF = arrayPF[i][cCodCentPF - 1];
                    var codSolPF = arrayPF[i][cCodSolPF - 1];

                    if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                      valCodCentEnc2 = true;
                      var fechaVencLin = arrayPF[i][cFechaSanc - 1];

                      if (fechaVencLin === "") {
                        valCodCentEnc2 = false;
                        break;
                      } else {
                        if (fechaVencLin != "Vencidas") {
                          valCodCentEnc2 = true;
                          break;
                        }
                      }
                    }
                  }

                  if (valCodCentEnc2 === true) {
                    sheetWF.getRange(fEncontrada, cTipodePF).setValue("RENOVACIÓN");
                  } else {
                    sheetWF.getRange(fEncontrada, cTipodePF).setValue("NUEVO");
                  }
                }

                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                  if (nuevaFV === "cancel" || nuevaFV === "") {
                    return;
                  }

                  var valCodCentEnc = false;
                  var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                  for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                    var codCentPF = arrayPF[i][cCodCentPF - 1];
                    var codSolPF = arrayPF[i][cCodSolPF - 1];

                    if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                      valCodCentEnc = true;

                      if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                        nuevaFV = tipoSan;
                      } else {
                        sheetPF.getRange(i + fInicioPF, cFechaSanc).setValue(nuevaFV);
                      } //HERE STAGE 3 MOD


                      tipoCliente = sheetWF.getRange(fEncontrada, cTipoCliente).getValue();
                      sheetPF.getRange(i + fInicioPF, cTipoClientePF).setValue(tipoCliente);
                      sheetPF.getRange(i + fInicioPF, cComentariosPF).setValue("");
                      sheetPF.getRange(i + fInicioPF, cUCodSolPF).setValue(CodSol);
                      sheetPF.getRange(i + fInicioPF, cFSPF).setValue(fecha);

                      if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                        /*NO PASA NADA*/
                      } else {
                        sheetPF.getRange(i + fInicioPF, cFechaOG).setValue(nuevaFV);
                      }

                      sheetPF.getRange(i + fInicioPF, cTraPF).setValue("NO");

                      if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                        /*NO PASA NADA*/
                      } else {
                        sheetPF.getRange(i + fInicioPF, cMontoSanc).setValue(tMontSanLin);
                      }

                      break;
                    }
                  }

                  if (valCodCentEnc === false) {
                    Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                    var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                    var subjectE = "Operación no encontrada en la base de Líneas.";
                    var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                    MailApp.sendEmail(recipientE, subjectE, bodyE);
                  }
                } //Le pertenece al correo de sanción de operación. Pasó debajo de la sanción de línea (por si acaso).


                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                  if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                    /*NO PASA NADA*/
                  } else {
                    body = body + "\nFecha de Vencimiento de la Línea: " + nuevaFV;
                  }
                }

                MailApp.sendEmail(recipient, subject, body, options); //HERE DWP FECHA SANCIÓN

                var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com"; //comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com 

                var subjectDWP = CodSol + "##Sanción " + tSanDWP + "##" + AnAsig + "##" + fecha + "##" + CodCentral;
                var bodyDWP = "Información general sobre la operación: " + "\n" + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp + "\nMonto de Sanción (Miles de US$): " + tMontSan + "\nÁmbito de Sanción: " + ambitoDWP;
                MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);

                if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                  if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                    /*NO PASA NADA*/
                  } else {
                    var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com"; //comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com 

                    var subjectDWP = CodCentral + "##" + nuevaFV;
                    var bodyDWP = "";
                    MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                  }
                }

                sheetWF.getRange(fEncontrada, cFS).setValue(fecha); //Escribe la fecha.

                celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                celda.setBackgroundRGB(85, 199, 104);
                sheetWF.getRange(fEncontrada, cRating).setValue(tRating);
                sheetWF.getRange(fEncontrada, cEEFF).setValue(yEEFF);
                sheetWF.getRange(fEncontrada, cHerramienta).setValue(herramienta);
                sheetWF.getRange(fEncontrada, cBuro).setValue(buro);
                sheetWF.getRange(fEncontrada, cMitig).setValue(condicMitig);
                sheetWF.getRange(fEncontrada, cEstratSanc).setValue(estrategia);
                /*if(Cliente === "GRUPO ECONÓMICO"){
                  sheetWF.getRange(fEncontrada,cCodRelacionado).setValue(codRelacionado)
                }
                else{
                  //No pasa nada. sheetWF.getRange(fEncontrada,cCodRelacionado).setValue(CodCentral)
                }*/

                if (tipoSan === "APROBADO CM") {
                  sheetWF.getRange(fEncontrada, cCas).setValue(tCas);
                } else if (tipoSan === "DENEGADO") {
                  sheetWF.getRange(fEncontrada, cCas).setValue(tCasDeneg);
                } else if (tipoSan === "DEVUELTO") {
                  sheetWF.getRange(fEncontrada, cCas).setValue(tMotDev);
                }

                sheetWF.getRange(fEncontrada, cAmb).setValue(ambito); //Escribe el ámbito.

                sheetWF.getRange(fEncontrada, cTS).setValue(tipoSan); //Escribe el tipo de sanción.

                sheetWF.getRange(fEncontrada, cMontSan).setValue(tMontSan);
                Browser.msgBox("Registrado", "Fecha de sanción, ámbito, tipo de sanción y monto sancionado registrados con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                ejecfWorkflowRiesgos(visor);
                return;
              } //}
              //}

            }
          }

          break;
        }
      }
    } //Solo para la estación 2
    //Solo para la estación 3
    else if (valEstEst === "I3") {
        for (var i = finicioEO; i <= lRowEO; i++) {
          var celda = sheetEO.getRange(i, cDatosE3);
          var val = celda.getValue();

          if (val === strRiReg || val === strRiReg2 || val === strRiReg3) {
            //Encuentra la instancia en la que riesgos puede ingresar una fecha.
            var celdavaldev = sheetEO.getRange(rNDEV);
            var valdev = celdavaldev.getValue();
            var celdaEO = sheetEO.getRange(i, cHeadingsE3);
            var headingEO = celdaEO.getValue(); //Obtiene el heading de la hoja de estación.

            if (headingEO === "Asignación RVGL" || headingEO === "Consulta de Riesgos a Cliente" || headingEO === "Consulta de BE a Cliente" || headingEO === "Devolución" || headingEO === "Consulta de Riesgos a Cliente (2)" || headingEO === "Consulta de BE a Cliente (2)" || headingEO === "Devolución (2)" || headingEO === "Fin Evaluación (VB Jefe)" || headingEO === "Fecha Sanción" || headingEO === "Asignación Evaluación II" || headingEO === "Consulta de Riesgos a Cliente (3)" || headingEO === "Consulta de BE a Cliente (3)" || headingEO === "Devolución (3)" || headingEO === "Fin Evaluación (VB Jefe) (2)" || headingEO === "Fecha Sanción (2)") {
              /*Continuar*/
            } else {
              Browser.msgBox("Riesgos", "El registro de fecha de esta operación le pertenece a la Oficina.", Browser.Buttons.OK);
              ss.toast("Cuando pueda registrar una fecha, aparecerá una celda de color amarillo (o celeste).", "Tip", 5);
            }

            if (headingEO === "Fin Evaluación (VB Jefe) (2)") {
              var cval = sheetEO.getRange(finicioEO + 1, cHeadingsE3);
              var valfvb = cval.getValue();

              if (valfvb === "Fin Evaluación (VB Jefe) (2)") {
                var respuesta = Browser.msgBox("Riesgos", "¿La evaluación ha concluido? Presione 'Sí' solo si concluyó. En caso contrario presione 'No' para registrar la fecha de las consultas enviadas a la Oficina.", Browser.Buttons.YES_NO);

                if (respuesta == "no") {
                  var decRi = "no";
                  var respuestamail = "yes"; //Browser.msgBox("Correo de Solicitud", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                  if (respuestamail === "yes") {
                    if (Correo != "SIN CORREO") {
                      var recipient = Correo;
                    } else if (Correo === "SIN CORREO") {
                      var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                    }

                    if (recipient === "cancel") {
                      return;
                    }

                    if (Cliente === "GRUPO ECONÓMICO") {
                      var subject = "SIO OC: " + Grupo + " - Solicitud de Ingreso de Respuestas Para la Operación " + CodSol;
                    } else {
                      var subject = "SIO OC: " + Cliente + " - Solicitud de Ingreso de Respuestas Para la Operación " + CodSol;
                    }

                    var body = "Se solicita su intervención para el ingreso de respuestas en el Workflow para la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                    var options = {
                      cc: correoGOF + ", " + correoAsist + ", siom@bbva.com"
                    };
                    MailApp.sendEmail(recipient, subject, body, options); //HERE DWP Consulta de BE a Cliente (3)

                    var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                    var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                    var subjectDWP = CodSol + "##Consulta de Riesgos a Cliente (3)##" + AnAsig + "##" + fecha + "##" + CodCentral;
                    var bodyDWP = Oper;
                    MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                  }
                } else if (respuesta === "cancel") {
                  return;
                } else {
                  var decRi = "yes";
                }
              } else {
                var decRi = "no";
              }
            } else if (headingEO === "Devolución (3)") {} else if (headingEO === "Reingreso con Respuestas (3)") {} else if (headingEO === "Fecha Sanción (2)") {} else if (headingEO === "Asignación Evaluación II") {
              var rpta = Browser.msgBox("Riesgos", "¿Desea reabrir la operación?", Browser.Buttons.YES_NO);

              if (rpta != "yes") {
                return;
              }
            } else {
              var respuesta = Browser.msgBox("Riesgos", "¿Desea devolver la operación?", Browser.Buttons.YES_NO);

              if (respuesta == "no") {
                var decRi = "no";
              } else if (respuesta === "cancel") {
                return;
              } else {
                var decRi = "yes";
              }
            }

            for (var j = 1; j <= lColumnWF; j++) {
              //Recorre las columnas de la base WF.
              var celdaWF = sheetWF.getRange(fHeadingsWF, j);
              var headingWF = celdaWF.getValue(); //Obtiene el heading de la base WF.

              if (headingWF === headingEO) {
                //Encuentra la instancia en la que se igualan ambos valores de los headings.
                var celdaHEOCRC = sheetEO.getRange(finicioEO + 1, cHeadingsE3);
                var HEOCRC = celdaHEOCRC.getValue();
                var celdanDEV = sheetEO.getRange(rNDEV);
                var valNDEV = celdanDEV.getValue();

                if (headingWF === "Asignación Evaluación II") {
                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                    var valCodCentEnc = false;
                    var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                    for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                      var codCentPF = arrayPF[i][cCodCentPF - 1];
                      var codSolPF = arrayPF[i][cCodSolPF - 1];
                      var nuevaFV = "REABIERTO";
                      var tMontSan = "";

                      if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                        valCodCentEnc = true;
                        sheetPF.getRange(i + fInicioPF, cFechaOG).setValue(nuevaFV);
                        sheetPF.getRange(i + fInicioPF, cTraPF).setValue("SÍ");
                        sheetPF.getRange(i + fInicioPF, cMontoSanc).setValue(tMontSan);
                        break;
                      }
                    }

                    if (valCodCentEnc === false) {
                      Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                      var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                      var subjectE = "Operación no encontrada en la base de Líneas.";
                      var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                      MailApp.sendEmail(recipientE, subjectE, bodyE);
                    }
                  }

                  var respuestamail = "yes"; //Browser.msgBox("Correo de Solicitud", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                  if (respuestamail === "yes") {
                    if (Correo != "SIN CORREO") {
                      var recipient = Correo;
                    } else if (Correo === "SIN CORREO") {
                      var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                    }

                    if (recipient === "cancel") {
                      return;
                    }

                    if (Cliente === "GRUPO ECONÓMICO") {
                      var subject = "SIO OC: " + Grupo + " - La Operación " + CodSol + " Ha Sido Reabierta por Riesgos";
                    } else {
                      var subject = "SIO OC: " + Cliente + " - La Operación " + CodSol + " Ha Sido Reabierta por Riesgos";
                    }

                    var body = "Se reabrió la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                    var options = {
                      cc: correoGOF + ", " + correoAsist + ", siom@bbva.com"
                    };
                    MailApp.sendEmail(recipient, subject, body, options); //HERE DWP Asignación Evaluación II

                    var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                    var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                    var subjectDWP = CodSol + "##Asignación Evaluación II##" + AnAsig + "##" + fecha + "##" + CodCentral;
                    var bodyDWP = Oper;
                    MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                  }

                  sheetWF.getRange(fEncontrada, cAsEv).setValue(fecha);
                  sheetWF.getRange(fEncontrada, cMontSan).setValue("");
                  Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                  ejecfWorkflowRiesgos(visor);
                  return;
                }

                if (headingWF === "Fin Evaluación (VB Jefe) (2)" && decRi === "no" && HEOCRC != "Consulta de BE a Cliente (3)") {
                  sheetWF.getRange(fEncontrada, c3raCBC).setValue(fecha);
                  Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                  ejecfWorkflowRiesgos(visor);
                  return;
                } else if (headingWF === "Fin Evaluación (VB Jefe) (2)" && decRi === "yes" && valNDEV === "") {
                  sheetWF.getRange(fEncontrada, cFev2).setValue(fecha); //Escribe la fecha.

                  celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                  celda.setBackgroundRGB(85, 199, 104);
                  Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                  ejecfWorkflowRiesgos(visor);
                  return;
                } else if (headingWF === "Fin Evaluación (VB Jefe) (2)" && valNDEV === "NDEV") {
                  sheetWF.getRange(fEncontrada, cFev2).setValue(fecha);
                  celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                  celda.setBackgroundRGB(85, 199, 104);
                  Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                  ejecfWorkflowRiesgos(visor);
                  return;
                } /////////Primera ronda de preguntas                


                if (headingWF === "Devolución (3)") {
                  var respuesta = Browser.msgBox("Riesgos", "¿El cliente le ha contestado a la oficina? Si elige no, entonces se procederá a devolver la operación hasta que la oficina reingrese las respuestas.", Browser.Buttons.YES_NO);

                  if (respuesta === "cancel") {
                    return;
                  } else if (respuesta === "yes") {
                    Browser.msgBox("Riesgos", "Espere hasta que la oficina registre la fecha de respuestas.", Browser.Buttons.OK);
                    return;
                  }

                  var valCas = false;

                  while (valCas === false) {
                    var tMotDev = Browser.inputBox("Casuística de Devolución", "Digite la opción de la casuística de la siguiente lista:\\nA. DEVUELTA A SOLICITUD DE LA OFICINA\\nB. DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO\\nC. DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA\\nD. DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING\\nE. DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR\\nF. DEVUELTA POR FALTA DE INFORMACION EN GENERAL\\nG. OTROS\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                    if (tMotDev === "cancel") {
                      return;
                    }

                    tMotDev = tMotDev.toUpperCase();
                    var split_str = tMotDev.split("+");

                    for (var iStr = 0; iStr < split_str.length; iStr++) {
                      if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G") {
                        if (split_str[iStr] === "A") {
                          split_str[iStr] = "DEVUELTA A SOLICITUD DE LA OFICINA";
                        } else if (split_str[iStr] === "B") {
                          split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO";
                        } else if (split_str[iStr] === "C") {
                          valCasImp = true;
                          split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA";
                        } else if (split_str[iStr] === "D") {
                          split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING";
                        } else if (split_str[iStr] === "E") {
                          split_str[iStr] = "DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR";
                        } else if (split_str[iStr] === "F") {
                          split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION EN GENERAL";
                        } else if (split_str[iStr] === "G") {
                          split_str[iStr] = "OTROS";
                        }

                        valCas = true;
                      } else {
                        Browser.msgBox("Error", "Motivo de devolución no válido.", Browser.Buttons.OK);
                        valCas = false;
                      }
                    }
                  }

                  tMotDev = split_str.join("+");
                  tMotDev = tMotDev.toUpperCase();
                  var respuestamail = "yes"; //Browser.msgBox("Correo de Devolución", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                  if (respuestamail === "yes") {
                    if (Correo != "SIN CORREO") {
                      var recipient = Correo;
                    } else if (Correo === "SIN CORREO") {
                      var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                    }

                    if (recipient === "cancel") {
                      return;
                    }

                    if (Cliente === "GRUPO ECONÓMICO") {
                      var subject = "SIO OC: " + Grupo + " - Devolución de la Operación " + CodSol;
                    } else {
                      var subject = "SIO OC: " + Cliente + " - Devolución de la Operación " + CodSol;
                    }

                    var body = "Riesgos ha devuelto la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                    var tCasCorreo = tMotDev.split('+').join(" || ");
                    body = body + "\nCasuística de Devolución: " + tCasCorreo;
                    var options = {
                      cc: correoGOF + ", " + correoAsist + ", siom@bbva.com"
                    };
                    MailApp.sendEmail(recipient, subject, body, options); //HERE Devolución (3)

                    var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                    var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                    var subjectDWP = CodSol + "##Devolución (3)##" + AnAsig + "##" + fecha + "##" + CodCentral;
                    var bodyDWP = Oper;
                    MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                    sheetWF.getRange(fEncontrada, cCas).setValue(tMotDev);
                  }

                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                    var valCodCentEnc = false;
                    var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                    for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                      var codCentPF = arrayPF[i][cCodCentPF - 1];
                      var codSolPF = arrayPF[i][cCodSolPF - 1]; //var nuevaFV = "DEVUELTO"
                      //var tMontSan = "DEVUELTO"

                      if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                        valCodCentEnc = true;

                        if (tMontSan === "DENEGADO" || tMontSan === "DEVUELTO") {
                          nuevaFV === tMontSan;
                        } //sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)


                        sheetPF.getRange(i + fInicioPF, cTraPF).setValue("NO"); //sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)

                        break;
                      }
                    }

                    if (valCodCentEnc === false) {
                      Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                      var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                      var subjectE = "Operación no encontrada en la base de Líneas.";
                      var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                      MailApp.sendEmail(recipientE, subjectE, bodyE);
                    }
                  }

                  sheetWF.getRange(fEncontrada, c3raDev).setValue(fecha);
                  Browser.msgBox("Registrado", "Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                  celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                  celda.setBackgroundRGB(85, 199, 104);
                  ejecfWorkflowRiesgos(visor);
                  return;
                } /////////Primera ronda de preguntas


                if (headingWF === "Fecha Sanción (2)") {
                  /*if(Cliente === "GRUPO ECONÓMICO"){
                    var valCodRelacionado = false
                    while(valCodRelacionado === false){
                      valCodRelacionado = true
                      
                      var codRelacionado = Browser.inputBox("Registrar clientes del grupo y sus respectivos montos sancionados de líneas", "Digite los códigos centrales relacionados al grupo y su respectivo monto sancionado de línas en miles de USD. Si fuera una devolución o denegación, el monto sería 0. \\n\\nDeben estar separados con la siguiente estructura: Código Central 1//Monto Sancionado 1;Código Central 2//Monto Sancionado 2. \\n\\nEjemplo: 00399280//10000;20728211//8000.", Browser.Buttons.OK_CANCEL)
                      codRelacionado = codRelacionado.toString()
                      if(codRelacionado === "" || codRelacionado === "cancel"){
                        return;
                      }
                      
                      
                      var valFaltaSep1 = codRelacionado.indexOf("//");
                      if(valFaltaSep1 < 0){
                        valCodRelacionado = false
                      }
                      
                      var valFaltaSep2 = codRelacionado.indexOf(";");
                      if(valFaltaSep2 < 0){
                        valCodRelacionado = false
                      }
                      
                      var valFaltaSep3 = codRelacionado.indexOf(",");
                      if(valFaltaSep3 > -1){
                        valCodRelacionado = false
                      }
                      
                      if(valCodRelacionado === false){
                        Browser.msgBox("No se ha encontrado el separador // o el separador ;. \\n\\nRecuerde que la estructura es: Código Central 1 // Monto Sancionado 1; Código Central 2 // Monto Sancionado 2")
                      }
                    }
                  }*/
                  var valAmb = false;

                  while (valAmb === false) {
                    valAmb = true;
                    var amb = Browser.inputBox("Registrar Nivel de Aprobación", "Digite la letra correspondiente a un nivel de aprobación de la siguiente lista:\\nA. ANALISTA\\nB. JEFE DE GRUPO\\nC. JEFE DE EQUIPO\\nD. SUBGERENTE\\nE. GERENTE DE UNIDAD\\nF. GERENTE DE ÁREA\\nG. COMITÉ DE CONTRASTE \\nH. CTO \\nI. CEC\\nJ. WCRMC\\nK. GCRMC\\nL. RECONDUCCIÓN\\nM. ECCMWO", Browser.Buttons.OK_CANCEL);
                    var tempAmb = amb;

                    if (amb === "") {
                      return;
                    }

                    if (amb === "cancel") {
                      return;
                    }

                    amb = amb.toUpperCase();

                    if (amb === "A") {
                      amb = "ANALISTA";
                    } else if (amb === "B") {
                      amb = "JEFE DE GRUPO";
                    } else if (amb === "C") {
                      amb = "JEFE DE EQUIPO";
                    } else if (amb === "D") {
                      amb = "SUBGERENTE";
                    } else if (amb === "E") {
                      amb = "GERENTE DE UNIDAD";
                    } else if (amb === "F") {
                      amb = "GERENTE DE ÁREA";
                    } else if (amb === "G") {
                      amb = "COMITÉ DE CONTRASTE";
                    } else if (amb === "H") {
                      amb = "CTO";
                    } else if (amb === "I") {
                      amb = "CEC";
                    } else if (amb === "J") {
                      amb = "WCRMC";
                    } else if (amb === "K") {
                      amb = "GCRMC";
                    } else if (amb === "L") {
                      amb = "RECONDUCCIÓN";
                    } else if (amb === "M") {
                      amb = "ECCMWO";
                    } else {
                      valAmb = false;
                      Browser.msgBox("Error", "No se ingresó una letra válida de la lista de nivel de aprobación.", Browser.Buttons.OK);
                      return;
                    }
                  }

                  amb = amb.toUpperCase();
                  var ambito = amb;
                  var ambitoDWP = ambito;
                  var valAmb = false;

                  while (valAmb === false) {
                    valAmb = true;
                    var amb = Browser.inputBox("Registrar Ámbito", "Digite la letra correspondiente al ámbito de la siguiente lista:\\nA. LOCAL\\nB. GCR ARGENTINA\\nC. GCR BRASIL\\nD. GCR COLOMBIA\\nE. GCR COMPASS\\nF. GCR MEXICO\\nG. GCR NY\\nH. GCR PARAGUAY\\nI. GCR PANAMA\\nJ. GCR URUGUAY\\nK. RPM ALEMANIA\\nL. RPM BELGICA\\nM. RPM COREA\\nN. RPM ESPAÑA\\nO. RPM FRANCIA\\nP. RPM HK\\nQ. RPM ITALIA\\nR. RPM JAPON\\nS. RPM SINGAPUR\\nT. RPM UK", Browser.Buttons.OK_CANCEL);
                    var tempAmb = amb;

                    if (amb === "") {
                      return;
                    }

                    if (amb === "cancel") {
                      return;
                    }

                    amb = amb.toUpperCase();

                    if (amb === "A") {
                      amb = "LOCAL";
                    } else if (amb === "B") {
                      amb = "GCR ARGENTINA";
                    } else if (amb === "C") {
                      amb = "GCR BRASIL";
                    } else if (amb === "D") {
                      amb = "GCR COLOMBIA";
                    } else if (amb === "E") {
                      amb = "GCR COMPASS";
                    } else if (amb === "F") {
                      amb = "GCR MEXICO";
                    } else if (amb === "G") {
                      amb = "GCR NY";
                    } else if (amb === "H") {
                      amb = "GCR PARAGUAY";
                    } else if (amb === "I") {
                      amb = "GCR PANAMA";
                    } else if (amb === "J") {
                      amb = "GCR URUGUAY";
                    } else if (amb === "K") {
                      amb = "RPM ALEMANIA";
                    } else if (amb === "L") {
                      amb = "RPM BELGICA";
                    } else if (amb === "M") {
                      amb = "RPM COREA";
                    } else if (amb === "N") {
                      amb = "RPM ESPAÑA";
                    } else if (amb === "O") {
                      amb = "RPM FRANCIA";
                    } else if (amb === "P") {
                      amb = "RPM HK";
                    } else if (amb === "Q") {
                      amb = "RPM ITALIA";
                    } else if (amb === "R") {
                      amb = "RPM JAPON";
                    } else if (amb === "S") {
                      amb = "RPM SINGAPUR";
                    } else if (amb === "T") {
                      amb = "RPM UK";
                    } else {
                      valAmb = false;
                      Browser.msgBox("Error", "No se ingresó una letra válida de la lista de ámbitos.", Browser.Buttons.OK);
                      return;
                    }
                  }

                  amb = amb.toUpperCase();
                  ambito = ambito + " // " + amb; //var valRating = false

                  var tRating = "(vacío)";
                  /*while(valRating === false){
                    var tRating = Browser.inputBox("Registrar el Rating", "Digite el rating. Ej. AAA, AA+, BBB+1, etc. \\nSi no aplicara, digite SIN RATING. ", Browser.Buttons.OK_CANCEL);
                    if(tRating === "cancel"){
                      return
                    }
                    else if(tRating === ""){
                      return
                    }
                    valRating = true
                    tRating = tRating.toUpperCase()
                    if (tRating === "SIN RATING" || tRating === "AAA" || tRating === "AA+" || tRating === "AA" || tRating === "AA-" || tRating === "A+" || tRating === "A" || tRating === "A-" || tRating === "BBB+1" || tRating === "BBB+2" || tRating === "BBB1" || tRating === "BBB2" || tRating === "BBB-1" || tRating === "BBB-2" || tRating === "BB+1" || tRating === "BB+2" || tRating === "BB1" || tRating === "BB2" || tRating === "BB-1" || tRating === "BB-2" || tRating === "B+1" || tRating === "B+2" || tRating === "B+3" || tRating === "B1" || tRating === "B2" || tRating === "B3" || tRating === "B-1" || tRating === "B-2" || tRating === "B-3" || tRating === "CCC+" || tRating === "CCC" || tRating === "CCC-" || tRating === "CC+" || tRating === "CC" || tRating === "CC-" ){/*SIGUE A LA SIGUIENTE*/

                  /*}
                  else{  
                  valRating = false
                  Browser.msgBox("Error", "Digite un rating válido.", Browser.Buttons.OK)
                  }   
                  }*/

                  var yEEFF = "(vacío)";
                  var herramienta = "(vacío)";
                  /*if(tRating === "SIN RATING"){
                    var yEEFF = "SIN RATING"
                    var herramienta = "SIN RATING"
                    }
                  else{
                    var valYEEFF = false
                    while(valYEEFF === false){
                      var yEEFF = Browser.inputBox("Registrar el año del EEFF", "Digite el año del EEFF. Ej. 2015, 2016, etc.", Browser.Buttons.OK_CANCEL);
                      if(yEEFF === "cancel"){
                        return
                      }
                      else if(yEEFF === ""){
                        return
                      }
                      valYEEFF = true
                      if (isNaN(yEEFF) != true){
                        if (yEEFF >= 2015 && yEEFF <= 2020){/*SIGUE A LA SIGUIENTE*/

                  /*}
                  else{  
                  valYEEFF = false
                  Browser.msgBox("Error", "Digite un año válido.", Browser.Buttons.OK)
                  }   
                  }
                  else{
                  valYEEFF = false
                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  }   
                  }
                  var valHerramienta = false
                  while(valHerramienta === false){
                  valHerramienta = true
                  var herramienta = Browser.inputBox("Registrar Herramienta", "Digite la letra correspondiente a la herramienta de la siguiente lista:\\nA. NACAR\\nB. RA", Browser.Buttons.OK_CANCEL)
                  if(herramienta === ""){return}
                  if(herramienta === "cancel"){return}
                  herramienta = herramienta.toUpperCase()
                  if(herramienta === "A"){herramienta = "NACAR"}
                  else if(herramienta === "B"){herramienta = "RA"}
                  else{
                  valHerramienta = false
                  Browser.msgBox("Error", "No se ingresó una letra válida de la lista de herramientas.", Browser.Buttons.OK)
                  }
                  }
                  }*/

                  var buro = "(vacío)";
                  /*var valBuro = false
                  while(valBuro === false){
                    valBuro = true
                    var buro = Browser.inputBox("Registrar Buró", "Digite la letra correspondiente al buró de la siguiente lista:\\nA. BURO G1\\nB. BURO G2\\nC. BURO G3\\nD. BURO G4\\nE. BURO G5\\nF. BURO G6\\nG. BURO G7 \\nH. BURO G8 \\nI. NO BANCARIZADO", Browser.Buttons.OK_CANCEL);
                    
                    if(buro === ""){return}
                    if(buro === "cancel"){return}
                    
                    buro = buro.toUpperCase()
                    if(buro === "A"){buro = "BURO G1"}
                    else if(buro === "B"){buro = "BURO G2"}
                    else if(buro === "C"){buro = "BURO G3"}
                    else if(buro === "D"){buro = "BURO G4"}
                    else if(buro === "E"){buro = "BURO G5"}
                    else if(buro === "F"){buro = "BURO G6"}
                    else if(buro === "G"){buro = "BURO G7"}
                    else if(buro === "H"){buro = "BURO G8"}
                    else if(buro === "I"){buro = "NO BANCARIZADO"}
                    else{
                      valBuro = false
                      Browser.msgBox("Error", "No se ingresó una letra válida de la lista de burós.", Browser.Buttons.OK)
                      return;
                    }
                  }*/

                  var estrategia = "(vacío)";
                  /*var valEstrategia = false
                  while(valEstrategia === false){
                    valEstrategia = true
                    var estrategia = Browser.inputBox("Registrar Estrategia", "Digite la letra correspondiente a la estrategia de la siguiente lista:\\nA. Liderar\\nB. Crecer\\nC. Vigilar\\nD. Reducir\\nE. Extinguir\\nF. A Potenciar\\nG. No Sugerido \\nH. Sin Estrategia", Browser.Buttons.OK_CANCEL);
                    
                    if(estrategia === ""){return}
                    if(estrategia === "cancel"){return}
                    
                    estrategia = estrategia.toUpperCase()
                    if(estrategia === "A"){estrategia = "Liderar"}
                    else if(estrategia === "B"){estrategia = "Crecer"}
                    else if(estrategia === "C"){estrategia = "Vigilar"}
                    else if(estrategia === "D"){estrategia = "Reducir"}
                    else if(estrategia === "E"){estrategia = "Extinguir"}
                    else if(estrategia === "F"){estrategia = "A potenciar"}
                    else if(estrategia === "G"){estrategia = "No Sugerido"}
                    else if(estrategia === "H"){estrategia = "(en blanco)"}
                    else{
                      valEstrategia = false
                      Browser.msgBox("Error", "No se ingresó una letra válida de la lista de estrategias.", Browser.Buttons.OK)
                      return;
                    }
                  }*/

                  var valTSan = false;

                  while (valTSan === false) {
                    var tSan = Browser.inputBox("Registrar el Tipo de Sanción", "Digite el número correspondiente a un tipo de sanción de la siguiente lista:\\n1. Aprobado Sin Modificación\\n2. Denegado\\n3. Devuelto\\n4. Aprobado Con Modificación", Browser.Buttons.OK_CANCEL);
                    valTSan = true;
                    var temptSan = tSan;

                    if (tSan === "") {
                      return;
                    }

                    if (tSan === "cancel") {
                      return;
                    }

                    if (tSan == 1) {
                      tSan = "Aprobado SM";
                    } else if (tSan == 2) {
                      tSan = "Denegado";
                    } else if (tSan == 3) {
                      tSan = "Devuelto";
                    } else if (tSan == 4) {
                      tSan = "Aprobado CM";
                    } else {
                      Browser.msgBox("Error", "No se digitó un número válido de la lista de tipos de sanción.", Browser.Buttons.OK);
                      valTSan = false;
                    }
                  }

                  var tSanDWP = tSan;
                  tSan = tSan.toUpperCase();

                  switch (tSan) {
                    case "APROBADO SM":
                      var tipoSan = tSan;
                      var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue();
                      break;

                    case "APROBADO CM":
                      var tipoSan = tSan;
                      var valCasImp = false;
                      var valCas = false;

                      while (valCas === false) {
                        var tCas = Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. Plazo\\nB. Garantía\\nC. Importe\\nD. Condicionantes Previas al Desembolso\\nE. Otros\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                        if (tCas === "cancel") {
                          return;
                        }

                        tCas = tCas.toUpperCase();
                        var split_str = tCas.split("+");

                        for (var iStr = 0; iStr < split_str.length; iStr++) {
                          if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E") {
                            if (split_str[iStr] === "A") {
                              split_str[iStr] = "Plazo";
                            } else if (split_str[iStr] === "B") {
                              split_str[iStr] = "Garantía";
                            } else if (split_str[iStr] === "C") {
                              valCasImp = true;
                              split_str[iStr] = "Importe";
                            } else if (split_str[iStr] === "D") {
                              split_str[iStr] = "Condicionantes Previas al Desembolso";
                            } else if (split_str[iStr] === "E") {
                              split_str[iStr] = "Otros";
                            }

                            valCas = true;
                          } else {
                            Browser.msgBox("Error", "Casuística no válida.", Browser.Buttons.OK);
                            valCas = false;
                            break;
                          }
                        }
                      }

                      tCas = split_str.join("+");
                      tCas = tCas.toUpperCase();
                      var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue();

                      if (valCasImp === true) {
                        var valTMont = false;

                        while (valTMont === false) {
                          var tMontSan = Browser.inputBox("Registrar el Monto Sancionado", "Registre el monto sancionado de esta operación en miles de US$.", Browser.Buttons.OK_CANCEL);

                          if (tMontSan === "cancel") {
                            return;
                          } else if (tMontSan === "") {
                            return;
                          }

                          valTMont = true;

                          if (isNaN(tMontSan) != false) {
                            valTMont = false;
                            Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK);
                          }
                        }
                      } else {
                        var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue();
                      } //sheetWF.getRange(fEncontrada,cCas).setValue(tCas)


                      break;

                    case "DEVUELTO":
                      var tipoSan = tSan;
                      var tMontSan = "DEVUELTO";
                      var valCas = false;

                      while (valCas === false) {
                        var tMotDev = Browser.inputBox("Casuística de Devolución", "Digite la opción de la casuística de la siguiente lista:\\nA. DEVUELTA A SOLICITUD DE LA OFICINA\\nB. DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO\\nC. DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA\\nD. DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING\\nE. DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR\\nF. DEVUELTA POR FALTA DE INFORMACION EN GENERAL\\nG. OTROS\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                        if (tMotDev === "cancel") {
                          return;
                        }

                        tMotDev = tMotDev.toUpperCase();
                        var split_str = tMotDev.split("+");

                        for (var iStr = 0; iStr < split_str.length; iStr++) {
                          if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G") {
                            if (split_str[iStr] === "A") {
                              split_str[iStr] = "DEVUELTA A SOLICITUD DE LA OFICINA";
                            } else if (split_str[iStr] === "B") {
                              split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DEL NEGOCIO";
                            } else if (split_str[iStr] === "C") {
                              valCasImp = true;
                              split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION DE LA PROPUESTA";
                            } else if (split_str[iStr] === "D") {
                              split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION PARA VALIDAR EL RATING";
                            } else if (split_str[iStr] === "E") {
                              split_str[iStr] = "DEVUELTA POR INFORMACIÓN ESPECIFICA DEL SECTOR";
                            } else if (split_str[iStr] === "F") {
                              split_str[iStr] = "DEVUELTA POR FALTA DE INFORMACION EN GENERAL";
                            } else if (split_str[iStr] === "G") {
                              split_str[iStr] = "OTROS";
                            }

                            valCas = true;
                          } else {
                            Browser.msgBox("Error", "Motivo de devolución no válido.", Browser.Buttons.OK);
                            valCas = false;
                          }
                        }
                      }

                      tMotDev = split_str.join("+");
                      tMotDev = tMotDev.toUpperCase();
                      break;

                    case "DENEGADO":
                      var tipoSan = tSan;
                      var tMontSan = "DENEGADO";
                      var valCasDeneg = false;

                      while (valCasDeneg === false) {
                        var tCasDeneg = Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. ESTRUCTURA ECONÓMICA FINANCIERA HISTORICA NO FAVORABLE\\nB. ALTO NIVEL DE ENDEUDAMIENTO/NO HAY CAPACIDAD DE PAGO PARA EL RIESGO PROPUESTO\\nC. DETERIORO DE LAS CIFRAS DE SITUACIÓN\\nD. ALTO NIVEL DE ENDEUDAMIENTO Y DEUDA ESTRUCTURAL ACTUAL DESCALZADA\\nE. ALERTAS EN EL COMPORTAMIENTO DE PAGO/ALERTAS EN LOS INDICES DE GESTION\\nF. EMPRESAS VINCULADAS /ACCIONSITAS CON ALERTAS NEGATIVAS.\\nG. INVERSIÓN DE ACTIVO FIJO NO CORRESPONDE AL CORE BUSSINES DEL NEGOCIO\\nH. ALERTAS EN EL SECTOR\\nI. ESTRUCTURA DE LA OPERACIÓN (PRODUCTO / PLAZO / CONDICIONES) SUPERA LA DIMENSIÓN DEL NEGOCIO\\nJ. PFA RECIENTEMENTE SANCIONADO Y/O CLIENTE CON TECHO DE RIESGO\\nK. SOLICITUD DE CP NO CORRESPONDE AL DIMENSIONAMIENTO DE LAS NECESIDADES OPERATIVAS\\nL. NO CUENTA CON EEFF DE CIERRE DE EVIDENCIEN GENERACIÓN DE CAJA SUFICIENTE\\nM. NO CUMPLE CON CONVENANTS/CONDICIONES ESTABLECIDAS EN PF/SANCION ANTERIOR\\nN. OTROS\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);

                        if (tCasDeneg === "cancel") {
                          return;
                        }

                        tCasDeneg = tCasDeneg.toUpperCase();
                        var split_str = tCasDeneg.split("+");

                        for (var iStr = 0; iStr < split_str.length; iStr++) {
                          if (split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L" || split_str[iStr] === "M" || split_str[iStr] === "N") {
                            if (split_str[iStr] === "A") {
                              split_str[iStr] = "ESTRUCTURA ECONÓMICA FINANCIERA HISTORICA NO FAVORABLE";
                            } else if (split_str[iStr] === "B") {
                              split_str[iStr] = "ALTO NIVEL DE ENDEUDAMIENTO/NO HAY CAPACIDAD DE PAGO PARA EL RIESGO PROPUESTO";
                            } else if (split_str[iStr] === "C") {
                              valCasImp = true;
                              split_str[iStr] = "DETERIORO DE LAS CIFRAS DE SITUACIÓN";
                            } else if (split_str[iStr] === "D") {
                              split_str[iStr] = "ALTO NIVEL DE ENDEUDAMIENTO Y DEUDA ESTRUCTURAL ACTUAL DESCALZADA";
                            } else if (split_str[iStr] === "E") {
                              split_str[iStr] = "ALERTAS EN EL COMPORTAMIENTO DE PAGO/ALERTAS EN LOS INDICES DE GESTION";
                            } else if (split_str[iStr] === "F") {
                              split_str[iStr] = "EMPRESAS VINCULADAS /ACCIONSITAS CON ALERTAS NEGATIVAS";
                            } else if (split_str[iStr] === "G") {
                              split_str[iStr] = "INVERSIÓN DE ACTIVO FIJO NO CORRESPONDE AL CORE BUSSINES DEL NEGOCIO";
                            } else if (split_str[iStr] === "H") {
                              split_str[iStr] = "ALERTAS EN EL SECTOR";
                            } else if (split_str[iStr] === "I") {
                              split_str[iStr] = "ESTRUCTURA DE LA OPERACIÓN (PRODUCTO / PLAZO / CONDICIONES) SUPERA LA DIMENSIÓN DEL NEGOCIO";
                            } else if (split_str[iStr] === "J") {
                              split_str[iStr] = "PFA RECIENTEMENTE SANCIONADO Y/O CLIENTE CON TECHO DE RIESGO";
                            } else if (split_str[iStr] === "K") {
                              split_str[iStr] = "SOLICITUD DE CP NO CORRESPONDE AL DIMENSIONAMIENTO DE LAS NECESIDADES OPERATIVAS";
                            } else if (split_str[iStr] === "L") {
                              split_str[iStr] = "NO CUENTA CON EEFF DE CIERRE DE EVIDENCIEN GENERACIÓN DE CAJA SUFICIENTE";
                            } else if (split_str[iStr] === "M") {
                              split_str[iStr] = "NO CUMPLE CON CONVENANTS/CONDICIONES ESTABLECIDAS EN PF/SANCION ANTERIOR";
                            } else if (split_str[iStr] === "N") {
                              split_str[iStr] = "OTROS";
                            }

                            valCasDeneg = true;
                          } else {
                            Browser.msgBox("Error", "Casuística no válida.", Browser.Buttons.OK);
                            valCasDeneg = false;
                            break;
                          }
                        }
                      }

                      tCasDeneg = split_str.join("+");
                      tCasDeneg = tCasDeneg.toUpperCase();
                      break;

                    default:
                      Browser.msgBox("Error", "No se ingresó un número válido de la lista de tipos de sanción.", Browser.Buttons.OK);
                      ss.toast("Digite el número correspondiente para el tipo de sanción; por ejemplo, para Aprobado Sin Modificación coloque un 1.", "Tip", 8);
                      return;
                  } //CONDICMITIG    


                  var condicMitig = "(vacío)"; //              if(tipoSan === "APROBADO CM" || tipoSan === "APROBADO SM"){
                  //                var condic = sheetWF.getRange(fEncontrada,cCondicion).getValue()
                  //                
                  //                if (tRating === "AAA" || tRating === "AA+" || tRating === "AA" || tRating === "AA-" || tRating === "A+" || tRating === "A" || tRating === "A-" || tRating === "BBB+1" || tRating === "BBB+2" || tRating === "BBB1" || tRating === "BBB2" || tRating === "BBB-1" || tRating === "BBB-2" || tRating === "BB+1" || tRating === "BB+2" || tRating === "BB1" || tRating === "BB2" || tRating === "BB-1" || tRating === "BB-2"){
                  //                  var condicMitig = "SIN MITIGANTES"
                  //                  }
                  //                else if (tRating === "B+1" || tRating === "B+2" || tRating === "B+3" || tRating === "B1" || tRating === "B2" || tRating === "B3" || tRating === "B-1" || tRating === "B-2" || tRating === "B-3" || tRating === "CCC+" || tRating === "CCC" || tRating === "CCC-" || tRating === "CC+" || tRating === "CC" || tRating === "CC-" ){
                  //                  
                  //                  var valCodCentralOpc1 = false
                  //                  
                  //                  var valCondicOpc1 = false
                  //                  var valCondicOpc2 = false
                  //                  var valCondicOpc3 = false
                  //                  var valCondicOpc4 = false
                  //                  var valCondicOpc5 = false
                  //                  var valCondicOpc6 = false
                  //                  
                  //                  var valCondicMitig = false
                  //                  while(valCondicMitig === false){
                  //                    var condicMitig = Browser.inputBox("Registrar Mitigantes", "Digite la opción de la mitigante de la siguiente lista:\\nA. <=B+1 (CESION DE FLUJOS COMO MEDIO DE PAGO)\\nB. <=B+1 (CON ARRENDATARIO SUTITUTO)\\nC. <=B+1 (CON POLIZA ENDOSADA AL BANCO)\\nD. <=B+1 (CON FIANZA SOLIDARIA)\\nE. <=B+1 (GAR. DEPÓSITO)\\nF. <=B+1 (GAR. FIDEICOMISO DE ACTIVOS Y/O FLUJOS)\\nG. <=B+1 (GAR. HIPOTECARIA EN TRAMITE)\\nH. <=B+1 (GAR. SBLC)\\nI. <=B+1 (GAR. WARRANT)\\nJ. <=B+1 (GAR.HIPOTECARIA)\\nK. <=B+1 (MEJORA DE GARANTÍAS)\\nL. <=B+1 (REPERFILAMIENTO DE DEUDA)\\nM. RIESGO EN BLANCO\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);
                  //                    if(condicMitig === "cancel"){return}
                  //                    condicMitig = condicMitig.toUpperCase()
                  //                    
                  //                    var split_str = condicMitig.split("+");
                  //                    
                  //                    valCondicMitig = true
                  //                    for(var iStr = 0; iStr < split_str.length; iStr++){
                  //                      if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L"|| split_str[iStr] === "M"){/*No pasa nada*/}
                  //                      else{
                  //                        Browser.msgBox("Error", "Mitigante no válida.", Browser.Buttons.OK)
                  //                        valCondicMitig = false
                  //                        break;
                  //                      }
                  //                    }
                  //                    
                  //                    if(valCondicMitig === true){
                  //                      for(var iStr = 0; iStr < split_str.length; iStr++){
                  //                        if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D" || split_str[iStr] === "E" || split_str[iStr] === "F" || split_str[iStr] === "G" || split_str[iStr] === "H" || split_str[iStr] === "I" || split_str[iStr] === "J" || split_str[iStr] === "K" || split_str[iStr] === "L" || split_str[iStr] === "M" ){
                  //                          if(split_str[iStr] === "A"){
                  //                            split_str[iStr] = "<=B+1 (CESION DE FLUJOS COMO MEDIO DE PAGO)"
                  //                          }
                  //                          else if(split_str[iStr] === "B"){
                  //                            split_str[iStr] = "<=B+1 (CON ARRENDATARIO SUTITUTO)"
                  //                          }
                  //                          else if(split_str[iStr] === "C"){
                  //                            split_str[iStr] = "<=B+1 (CON POLIZA ENDOSADA AL BANCO)"
                  //                          }
                  //                          else if(split_str[iStr] === "D"){
                  //                            valCodCentralOpc1 = true
                  //                            if(valCodCentralOpc1 === true){
                  //                              var valTCodCentralOpc1 = false
                  //                              while(valTCodCentralOpc1 === false){
                  //                                var codCentralOpc1 = Browser.inputBox("Registrar el código central", "Registre el código central relacionado a esta opción de mitigante (incluyendo 0s, por ejemplo: 00399280).", Browser.Buttons.OK_CANCEL);
                  //                                if(codCentralOpc1 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(codCentralOpc1 === ""){
                  //                                  return
                  //                                }
                  //                                valTCodCentralOpc1 = true
                  //                                if (codCentralOpc1.length != 8){
                  //                                  valTCodCentralOpc1 = false
                  //                                  Browser.msgBox("Error", "Digite un código central válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (CON FIANZA SOLIDARIA), " + codCentralOpc1
                  //                          }
                  //                          else if(split_str[iStr] === "E"){
                  //                            valCondicOpc1 = true
                  //                            if(valCondicOpc1 === true){
                  //                              var valTCondicOpc1 = false
                  //                              while(valTCondicOpc1 === false){
                  //                                var condicOpc1 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. DEPÓSITO)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                  //                                if(condicOpc1 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(condicOpc1 === ""){
                  //                                  return
                  //                                }
                  //                                valTCondicOpc1 = true
                  //                                if (isNaN(condicOpc1) != false){
                  //                                  valTCondicOpc1 = false
                  //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (GAR. DEPÓSITO), " + condicOpc1 + "%"
                  //                          }
                  //                          else if(split_str[iStr] === "F"){
                  //                            valCondicOpc2 = true
                  //                            if(valCondicOpc2 === true){
                  //                              var valTCondicOpc2 = false
                  //                              while(valTCondicOpc2 === false){
                  //                                var condicOpc2 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. FIDEICOMISO DE ACTIVOS Y/O FLUJOS)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                  //                                if(condicOpc2 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(condicOpc2 === ""){
                  //                                  return
                  //                                }
                  //                                valTCondicOpc2 = true
                  //                                if (isNaN(condicOpc2) != false){
                  //                                  valTCondicOpc2 = false
                  //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (GAR. FIDEICOMISO DE ACTIVOS Y/O FLUJOS), " + condicOpc2 + "%"
                  //                          }
                  //                          else if(split_str[iStr] === "G"){
                  //                            valCondicOpc3 = true
                  //                            if(valCondicOpc3 === true){
                  //                              var valTCondicOpc3 = false
                  //                              while(valTCondicOpc3 === false){
                  //                                var condicOpc3 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. HIPOTECARIA EN TRAMITE)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                  //                                if(condicOpc3 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(condicOpc3 === ""){
                  //                                  return
                  //                                }
                  //                                valTCondicOpc3 = true
                  //                                if (isNaN(condicOpc3) != false){
                  //                                  valTCondicOpc3 = false
                  //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (GAR. HIPOTECARIA EN TRAMITE), " + condicOpc3 + "%"
                  //                          }
                  //                          else if(split_str[iStr] === "H"){
                  //                            valCondicOpc4 = true
                  //                            if(valCondicOpc4 === true){
                  //                              var valTCondicOpc4 = false
                  //                              while(valTCondicOpc4 === false){
                  //                                var condicOpc4 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. SBLC)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                  //                                if(condicOpc4 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(condicOpc4 === ""){
                  //                                  return
                  //                                }
                  //                                valTCondicOpc4 = true
                  //                                if (isNaN(condicOpc4) != false){
                  //                                  valTCondicOpc4 = false
                  //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (GAR. SBLC), " + condicOpc4 + "%"
                  //                          }
                  //                          else if(split_str[iStr] === "I"){
                  //                            valCondicOpc5 = true
                  //                            if(valCondicOpc5 === true){
                  //                              var valTCondicOpc5 = false
                  //                              while(valTCondicOpc5 === false){
                  //                                var condicOpc5 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR. WARRANT)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                  //                                if(condicOpc5 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(condicOpc5 === ""){
                  //                                  return
                  //                                }
                  //                                valTCondicOpc5 = true
                  //                                if (isNaN(condicOpc5) != false){
                  //                                  valTCondicOpc5 = false
                  //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (GAR. WARRANT), " + condicOpc5 + "%"
                  //                          }
                  //                          else if(split_str[iStr] === "J"){
                  //                            valCondicOpc6 = true
                  //                            if(valCondicOpc6 === true){
                  //                              var valTCondicOpc6 = false
                  //                              while(valTCondicOpc6 === false){
                  //                                var condicOpc6 = Browser.inputBox("Registrar el % de cobertura de <=B+1 (GAR.HIPOTECARIA)", "Registre el % de cobertura de esta mitigación. NO digite el símbolo de %. Ej. 25.", Browser.Buttons.OK_CANCEL);
                  //                                if(condicOpc6 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(condicOpc6 === ""){
                  //                                  return
                  //                                }
                  //                                valTCondicOpc6 = true
                  //                                if (isNaN(condicOpc6) != false){
                  //                                  valTCondicOpc6 = false
                  //                                  Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            
                  //                            split_str[iStr] = "<=B+1 (GAR.HIPOTECARIA), " + condicOpc6 + "%"
                  //                          }
                  //                          else if(split_str[iStr] === "K"){
                  //                            split_str[iStr] = "<=B+1 (MEJORA DE GARANTÍAS)"
                  //                          }
                  //                          else if(split_str[iStr] === "L"){
                  //                            split_str[iStr] = "<=B+1 (REPERFILAMIENTO DE DEUDA)"
                  //                          }
                  //                          else if(split_str[iStr] === "M"){
                  //                            split_str[iStr] = "RIESGO EN BLANCO"
                  //                          }
                  //                          
                  //                          valCondicMitig = true   
                  //                        }
                  //                      }
                  //                    }
                  //                  }
                  //                  
                  //                  condicMitig = split_str.join(" // ")                  
                  //                  condicMitig = condicMitig.toUpperCase()
                  //                }
                  //                else if(tRating === "SIN RATING"){    
                  //                  
                  //                  var valCodCentralOpc1 = false
                  //                  
                  //                  var valCondicMitig = false
                  //                  while(valCondicMitig === false){
                  //                    var condicMitig = Browser.inputBox("Registrar Mitigantes", "Digite la opción de la mitigante de la siguiente lista:\\nA. SIN.RAT 100% GARANTIZADO\\nB. SIN.RAT FIANZA SOLIDARIA\\nC. SIN.RAT PROJECT FINANCE\\nD. SIN.RAT Ventas < 2.8MM\\nE. RIESGO EN BLANCO\\n\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B.", Browser.Buttons.OK_CANCEL);
                  //                    if(condicMitig === "cancel"){return}
                  //                    condicMitig = condicMitig.toUpperCase()
                  //                    
                  //                    var split_str = condicMitig.split("+");
                  //                    
                  //                    valCondicMitig = true
                  //                    for(var iStr = 0; iStr < split_str.length; iStr++){
                  //                      if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D"|| split_str[iStr] === "E"){/*No pasa nada*/}
                  //                      else{
                  //                        Browser.msgBox("Error", "Mitigante no válida.", Browser.Buttons.OK)
                  //                        valCondicMitig = false
                  //                        break;
                  //                      }
                  //                    }
                  //                    
                  //                    if(valCondicMitig === true){
                  //                      for(var iStr = 0; iStr < split_str.length; iStr++){
                  //                        if(split_str[iStr] === "A" || split_str[iStr] === "B" || split_str[iStr] === "C" || split_str[iStr] === "D"|| split_str[iStr] === "E"){
                  //                          if(split_str[iStr] === "A"){
                  //                            split_str[iStr] = "SIN.RAT 100% GARANTIZADO"
                  //                          }
                  //                          else if(split_str[iStr] === "B"){
                  //                            valCodCentralOpc1 = true
                  //                            if(valCodCentralOpc1 === true){
                  //                              var valTCodCentralOpc1 = false
                  //                              while(valTCodCentralOpc1 === false){
                  //                                var codCentralOpc1 = Browser.inputBox("Registrar el código central", "Registre el código central relacionado a esta opción de mitigante (incluyendo 0s, por ejemplo: 00399280).", Browser.Buttons.OK_CANCEL);
                  //                                if(codCentralOpc1 === "cancel"){
                  //                                  return
                  //                                }
                  //                                else if(codCentralOpc1 === ""){
                  //                                  return
                  //                                }
                  //                                valTCodCentralOpc1 = true
                  //                                if (codCentralOpc1.length != 8){
                  //                                  valTCodCentralOpc1 = false
                  //                                  Browser.msgBox("Error", "Digite un código central válido.", Browser.Buttons.OK)
                  //                                }
                  //                              }
                  //                            }
                  //                            split_str[iStr] = "SIN.RAT FIANZA SOLIDARIA, " + codCentralOpc1
                  //                          }
                  //                          else if(split_str[iStr] === "C"){
                  //                            split_str[iStr] = "SIN.RAT PROJECT FINANCE"
                  //                          }
                  //                          else if(split_str[iStr] === "D"){
                  //                            split_str[iStr] = "SIN.RAT Ventas < 2.8MM"
                  //                          }
                  //                          else if(split_str[iStr] === "E"){
                  //                            split_str[iStr] = "RIESGO EN BLANCO"
                  //                          }
                  //                          
                  //                          valCondicMitig = true   
                  //                        }
                  //                      }
                  //                    }
                  //                  }
                  //                  
                  //                  condicMitig = split_str.join(" // ")                  
                  //                  condicMitig = condicMitig.toUpperCase()
                  //                }
                  //                
                  //                var split_str = condic.split("+");
                  //                var fen = false
                  //                for(var iStr = 0; iStr < split_str.length; iStr++){
                  //                  if(split_str[iStr] === "FENOMENO DEL NIÑO"){
                  //                    fen = true
                  //                  }
                  //                }
                  //                
                  //                if(fen === true){
                  //                  if(condicMitig === "SIN MITIGANTES"){
                  //                    condicMitig = "SIN EEFF SUNAT 2016 FEN"
                  //                  }
                  //                  else{
                  //                    condicMitig = condicMitig + " // SIN EEFF SUNAT 2016 FEN"
                  //                  }
                  //                }
                  //              }
                  //CONDICMITIG

                  var respuestamail = "yes"; //Browser.msgBox("Correo de Sanción", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)

                  if (respuestamail === "yes") {
                    if (Correo != "SIN CORREO") {
                      var recipient = Correo;
                    } else if (Correo === "SIN CORREO") {
                      var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL);
                    }

                    if (recipient === "cancel") {
                      return;
                    }

                    if (Cliente === "GRUPO ECONÓMICO") {
                      var subject = "SIO OC: " + Grupo + " - Operación " + CodSol + " Sancionada";
                    } else {
                      var subject = "SIO OC: " + Cliente + " - Operación " + CodSol + " Sancionada";
                    }

                    var body = "Riesgos ha sancionado la operación con código de solicitud: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp + "\nMonto de Sanción (Miles de US$): " + tMontSan + "\nÁmbito de Sanción: " + ambito + "\nTipo de Sanción: " + tipoSan;

                    if (tipoSan === "APROBADO CM") {
                      var tCasCorreo = tCas.split('+').join(" || ");
                      body = body + "\nCasuística de Modificación: " + tCasCorreo;
                    } else if (tipoSan === "DENEGADO") {
                      var tCasCorreo = tCasDeneg.split('+').join(" || ");
                      body = body + "\nCasuística de Denegación: " + tCasCorreo;
                    } else if (tipoSan === "DEVUELTO") {
                      var tCasCorreo = tMotDev.split('+').join(" || ");
                      body = body + "\nCasuística de Devolución: " + tCasCorreo;
                    }

                    if (tipoOp === "LP") {
                      body = body + "\nProducto: " + sheetWF.getRange(fEncontrada, cProducto).getValue();
                      var options = {
                        cc: correoGOF + ", " + correoAsist + ", " + "siom@bbva.com, vcardena@bbva.com, aazcoytia@bbva.com, ucastillo@bbva.com, kgonzalesp@bbva.com, christian.blanch@bbva.com"
                      };
                    } else {
                      var options = {
                        cc: correoGOF + ", " + correoAsist + ", " + "siom@bbva.com"
                      };
                    } //MailApp.sendEmail(recipient,subject,body,options); Pasa después de la sanción de la línea

                  }

                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                    //HERE DIGITAR LOS CÓDIGOS CENTRALES DEL GRUPO
                    if (tipoSan === "DEVUELTO" || tipoSan === "DENEGADO") {
                      var nuevaFV = tipoSan;
                    } else {
                      var valTMontLin = false;

                      while (valTMontLin === false) {
                        var tMontSanLin = tMontSan; //Browser.inputBox("Registrar el Monto Sancionado de la LÍNEA", "Registre el monto sancionado para la línea del grupo " + Grupo + " en miles de US$.", Browser.Buttons.OK_CANCEL);

                        if (tMontSanLin === "cancel") {
                          return;
                        } else if (tMontSanLin === "") {
                          return;
                        }

                        valTMontLin = true;

                        if (isNaN(tMontSanLin) != false) {
                          valTMontLin = false;
                          Browser.msgBox("Error", "Digite un número válido.", Browser.Buttons.OK);
                        }
                      }

                      var valFecha = false;

                      while (valFecha === false) {
                        var nuevaFV = Browser.inputBox("Fecha de Vencimiento", "Ingrese una fecha de vencimiento. Este cambio se reflejará en la base de Líneas. La fecha debe estar en formato 'dd/mm/aaaa'. Por ejemplo, 31/12/2017.", Browser.Buttons.OK_CANCEL);

                        if (nuevaFV === "cancel" || nuevaFV === "") {
                          Browser.msgBox("Ha decidido no registrar una fecha de vencimiento. La operación no será registrada.");
                          return;
                        }

                        var valFecha = isValidDate(nuevaFV);

                        if (valFecha === false) {
                          Browser.msgBox("Fecha no válida. Por favor intente de nuevo.");
                        }
                      }
                    }
                  }

                  if (tipoOp === "CP") {
                    var valCodCentEnc2 = false;
                    var lRowPF = Avals.filter(String).length;
                    var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                    for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                      var codCentPF = arrayPF[i][cCodCentPF - 1];
                      var codSolPF = arrayPF[i][cCodSolPF - 1];

                      if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                        valCodCentEnc2 = true;
                        var fechaVencLin = arrayPF[i][cFechaSanc]; //No se le ha puesto -1 a la columna porque estoy usando el último día

                        if (fechaVencLin === "") {
                          valCodCentEnc2 = false;
                          break;
                        } else {
                          if (fechaVencLin != "Vencidas") {
                            fechaVencLin = new Date(fechaVencLin.getFullYear(), fechaVencLin.getMonth(), fechaVencLin.getDate());
                            fechaVencLin.setMonth(fechaVencLin.getMonth() - 6);
                            var valFechaVencLin = fechaVencLin.valueOf();
                            var valFecha = fecha.valueOf();
                            break;
                          }
                        }
                      }
                    }

                    if (valCodCentEnc2 === true) {
                      if (fechaVencLin === "Vencidas") {
                        sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(2);
                      } else {
                        if (valFechaVencLin < valFecha) {
                          sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(4);
                        } else {
                          sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(3);
                        }
                      }
                    } else {
                      sheetWF.getRange(fEncontrada, cMarcaPuntual).setValue(1);
                    }
                  }

                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA") {
                    var valCodCentEnc2 = false;
                    var lRowPF = Avals.filter(String).length;
                    var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                    for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                      var codCentPF = arrayPF[i][cCodCentPF - 1];
                      var codSolPF = arrayPF[i][cCodSolPF - 1];

                      if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                        valCodCentEnc2 = true;
                        var fechaVencLin = arrayPF[i][cFechaSanc - 1];

                        if (fechaVencLin === "") {
                          valCodCentEnc2 = false;
                          break;
                        } else {
                          if (fechaVencLin != "Vencidas") {
                            valCodCentEnc2 = true;
                            break;
                          }
                        }
                      }
                    }

                    if (valCodCentEnc2 === true) {
                      sheetWF.getRange(fEncontrada, cTipodePF).setValue("RENOVACIÓN");
                    } else {
                      sheetWF.getRange(fEncontrada, cTipodePF).setValue("NUEVO");
                    }
                  }

                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                    if (nuevaFV === "cancel" || nuevaFV === "") {
                      return;
                    }

                    var valCodCentEnc = false;
                    var arrayPF = sheetPF.getRange(fInicioPF, 1, lRowPF - fInicioPF + 1, lColumnPF).getValues();

                    for (var i = 0; i <= lRowPF - fInicioPF; i++) {
                      var codCentPF = arrayPF[i][cCodCentPF - 1];
                      var codSolPF = arrayPF[i][cCodSolPF - 1];

                      if (codCentPF === CodCentral || codCentPF === codGE || CodSol === codSolPF) {
                        valCodCentEnc = true;

                        if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                          nuevaFV = tipoSan;
                        } else {
                          sheetPF.getRange(i + fInicioPF, cFechaSanc).setValue(nuevaFV);
                        } //HERE MOD STAGE 3


                        tipoCliente = sheetWF.getRange(fEncontrada, cTipoCliente).getValue();
                        sheetPF.getRange(i + fInicioPF, cTipoClientePF).setValue(tipoCliente);
                        sheetPF.getRange(i + fInicioPF, cComentariosPF).setValue("");
                        sheetPF.getRange(i + fInicioPF, cUCodSolPF).setValue(CodSol);
                        sheetPF.getRange(i + fInicioPF, cFSPF).setValue(fecha);

                        if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                          /*NO PASA NADA*/
                        } else {
                          sheetPF.getRange(i + fInicioPF, cFechaOG).setValue(nuevaFV);
                        }

                        sheetPF.getRange(i + fInicioPF, cTraPF).setValue("NO");

                        if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                          /*NO PASA NADA*/
                        } else {
                          sheetPF.getRange(i + fInicioPF, cMontoSanc).setValue(tMontSanLin);
                        }

                        break;
                      }
                    }

                    if (valCodCentEnc === false) {
                      Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.");
                      var recipientE = "berly.joaquin@bbva.com, luis.luna.cruz@bbva.com";
                      var subjectE = "Operación no encontrada en la base de Líneas.";
                      var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol + ".\n" + "\nInformación general sobre la operación: " + "\nCódigo Central: " + CodCentral + "\nGestor: " + Ejecutivo + "\nCliente: " + Cliente + "\nGrupo : " + Grupo + "\nOperación: " + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp;
                      MailApp.sendEmail(recipientE, subjectE, bodyE);
                    }
                  } //Le pertenece al correo de sanción de operación. Pasó debajo de la sanción de línea (por si acaso).


                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                    if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                      /*NO PASA NADA*/
                    } else {
                      body = body + "\nFecha de Vencimiento de la Línea: " + nuevaFV;
                    }
                  }

                  MailApp.sendEmail(recipient, subject, body, options); //HERE DWP FECHA SANCIÓN (2)

                  var AnAsig = sheetWF.getRange(fEncontrada, cAnAsig).getValue();
                  var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                  var subjectDWP = CodSol + "##Sanción " + tSanDWP + " (2)" + "##" + AnAsig + "##" + fecha + "##" + CodCentral;
                  var bodyDWP = "Información general sobre la operación: " + "\n" + Oper + "\nMonto Propuesto (Miles de US$): " + MontoProp + "\nMonto de Sanción (Miles de US$): " + tMontSan + "\nÁmbito de Sanción: " + ambitoDWP + "\n" + ambitoDWP;
                  MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);

                  if (tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track") {
                    if (tipoSan === "DENEGADO" || tipoSan === "DEVUELTO") {
                      /*NO PASA NADA*/
                    } else {
                      var recipientDWP = "comunicacion@i-3mzoei034gfd7jvxwzd9fa3v67l4ftdu972yrtkdy5wpk2blw.1i-3okbeuaq.na73.apex.salesforce.com, siom@bbva.com";
                      var subjectDWP = CodCentral + "##" + nuevaFV;
                      var bodyDWP = "";
                      MailApp.sendEmail(recipientDWP, subjectDWP, bodyDWP);
                    }
                  }

                  sheetWF.getRange(fEncontrada, cFS2).setValue(fecha); //Escribe la fecha.

                  celda.setValue(fecha); //Escribe la fecha ingresada en la hoja de estaciones.

                  celda.setBackgroundRGB(85, 199, 104);
                  sheetWF.getRange(fEncontrada, cRating).setValue(tRating);
                  sheetWF.getRange(fEncontrada, cEEFF).setValue(yEEFF);
                  sheetWF.getRange(fEncontrada, cHerramienta).setValue(herramienta);
                  sheetWF.getRange(fEncontrada, cBuro).setValue(buro);
                  sheetWF.getRange(fEncontrada, cMitig).setValue(condicMitig);
                  sheetWF.getRange(fEncontrada, cEstratSanc).setValue(estrategia);
                  /*if(Cliente === "GRUPO ECONÓMICO"){
                    sheetWF.getRange(fEncontrada,cCodRelacionado).setValue(codRelacionado)
                  }
                  else{
                    //sheetWF.getRange(fEncontrada,cCodRelacionado).setValue(CodCentral)
                  }*/

                  if (tipoSan === "APROBADO CM") {
                    sheetWF.getRange(fEncontrada, cCas).setValue(tCas);
                  } else if (tipoSan === "DENEGADO") {
                    sheetWF.getRange(fEncontrada, cCas).setValue(tCasDeneg);
                  } else if (tipoSan === "DEVUELTO") {
                    sheetWF.getRange(fEncontrada, cCas).setValue(tMotDev);
                  }

                  sheetWF.getRange(fEncontrada, cAmb2).setValue(ambito); //Escribe el ámbito.

                  sheetWF.getRange(fEncontrada, cTS2).setValue(tipoSan); //Escribe el tipo de sanción.

                  sheetWF.getRange(fEncontrada, cMontSan).setValue(tMontSan);
                  Browser.msgBox("Registrado", "Fecha de sanción, ámbito, tipo de sanción y monto sancionado registrados con éxito. Revisando nuevamente...", Browser.Buttons.OK);
                  ejecfWorkflowRiesgos(visor);
                  return;
                }
              }
            }

            break;
          }
        }
      } //Solo para la estación 3


  Browser.msgBox("Riesgos", "El registro de fecha de esta operación le pertenece a la Oficina.", Browser.Buttons.OK);
  ss.toast("Cuando pueda registrar una fecha, aparecerá una celda de color amarillo (o celeste).", "Tip", 5);
}