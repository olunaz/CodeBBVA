function fDesOp2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetForm = ss.getSheetByName("Ingreso Metodización")
  var sheetWF = ss.getSheetByName("WF");
  var sheetEO = ss.getSheetByName("HY")
  var finicioEO = 12 //Inicio de la revisión de operaciones
  var finicioWF = 10 //Inicio de la BD del WF
  var lRowWF = sheetWF.getLastRow()
  var lColumnWF = sheetWF.getLastColumn()
  var lRowEO = sheetEO.getLastRow()
  var lColumnEO = sheetEO.getLastColumn()
  var lRowForm = sheetForm.getLastRow()
  var lColumnForm = sheetForm.getLastColumn()
  
  var sh = ss.getSheetByName("Ingreso Metodización");
  var lRowSH = sh.getLastRow()
  
  var cDatosE1 = 2
  var cHeadingsE1 = 1
  var cDatosE2 = 4
  var cHeadingsE2 = 3
  var cDatosE3 = 6
  var cHeadingsE3 = 5  
  var fHeadingsWF = 9
  var valcellcodsol = "H8"
  var rNDEV = "B99"
  var valcellEstEst = "B98"
  var verCodSol = sheetEO.getRange("D1").getValue()
  
  //Columnas de la base de registro del formulario.
  var finicioSH = 2
  var cCorreoResp = 8
  var cCodCentForm = 16
  var cPetForm = 17
  var cEstadoForm = 18
  var cCantidadForm = 19
  var cPesoForm = 20
  var cResultadoForm = 21
  var cAnAsigForm = 22
  //Columnas de la base de registro del formulario.
  
  //Modificar estos valores para que se ajusten a las nuevas coordenadas de los datos.
  var datosGenerales = sheetEO.getRange("A2:D8").getValues()
  var CodCentral = datosGenerales[0][1]
  var Ejecutivo = datosGenerales[2][3]
  var Cliente = datosGenerales[1][1]
  var Grupo   = datosGenerales[2][1]
  var Oper = datosGenerales[3][1]
  var tipoOp = Oper
  var DBG = datosGenerales[4][1]
  var DER = datosGenerales[5][1]
  var ArchAdj = datosGenerales[6][1]
  var Correo = datosGenerales[3][3]
  var AnAsig = datosGenerales[5][3]
  var correoAnAsig = datosGenerales[6][3]
  var correoResp = datosGenerales[4][3]
  
  var strOfReg = "OFICINA: REGISTRE FECHA"
  var strRiReg = "SEGMENTO EMPRESAS: CONFIRMAR O DEVOLVER"
  var strRiReg2 = "SEGMENTO EMPRESAS: REGISTRE FECHA"
  var strRiReg3 = "SEGMENTO EMPRESAS: CONFIRMAR O CONSULTAR"
  
  // Columnas del Drive
  var cCodSol = 1
  var cCodCentral = 2
  var cRUC = 3
  var cTipoDOI = 4
  var cCodOfi = 5
  var cNomOfi = 6
  var cCodGes = 7
  var cEjecutivo = 8
  var cProc = 9
  var cFechaBase = 10
  var cFechaFinal = 11
  var cCodTerr = 12
  var cNomTerr = 13
  var cClienteWF = 14
  var cGrupoWF = 15
  var cArchivoWF = 16
  var cCodDoc = 17
  var cTipoEstadoWF = 18
  var cDABGWF = 19
  var cDAERWF = 20
  var cAnAsig = 21
  var cCheckList = 22
  var c1erCheckList = 23
  var cDevuelveCheckList = 24
  var cArchivoDevuelve = 25
  var cRev = 26
  var cCBC = 27
  var cDev = 28
  var cArchivoResp = 29
  var cRcR1 = 30
  var cFS = 31
  var cAmb = 32
  var cTS = 33
  var cMotivCL = 34
  var cGiroNegocioWF = 35
  var cNumTrabWF = 36
  var cTipoMetodWF = 37
  
  var cNotasAdicionales = 40
  var cDesestimado = 41
  var valtoast = true //Determina cuándo empieza el toast para la fase 2.
  var valtoast2 = true //Determina cuándo empieza el toast para la fase 3.
  var valtoastrec = true //Determina si el recordatorio se dará.
  // Columnas del Drive
  
  var fechaHoy = new Date()
  fechaHoy.setHours(0,0,0,0) 
  var fecha = fechaHoy //Obtiene fecha
  
  var date = new Date();
  var hour = date.getHours();
  if(hour >= 18 && date.getDay() >= 1 && date.getDay() <= 5){
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
  }
  
  var celdaCodSol = sheetEO.getRange("B1")
  var CodSol = celdaCodSol.getValue()
  
  if(CodSol === ""){return}
  
  if(CodSol != verCodSol){
    Browser.msgBox("¡Alerta!", "La fecha NO ha sido registrada debido a que la revisión ha expirado. Actualice las estaciones nuevamente.", Browser.Buttons.OK)
    ss.toast("Antes de registrar una fecha, por favor presione el botón 'Actualizar Estaciones'.", "Recordatorio", 10)
    return;
  }
  
  var cEstEst = sheetEO.getRange(valcellEstEst)
  var valEstEst = cEstEst.getValue()
  
  var arrayWF = sheetWF.getRange(finicioWF-1,1,lRowWF - finicioWF+2,lColumnWF).getValues() //lRowWF - fInicioWF +1
  
  //Loop que encuentra el código de solicitud
  for(var i = 1; i <= lRowWF - finicioWF +1; i++){ //Acá empieza en 1 porque se incluyeron los Headings
    var codSolWF = arrayWF[i][cCodSol-1]
    if(codSolWF === CodSol){           //Compara códigos de solicitud
      var fEncontrada = i + finicioWF -1  //Guarda el valor de la fila del código de solicitud encontrado dentro de la base WF.
      break;
    }
  }

  var fechaDesest = arrayWF[i][cDesestimado-1]
  if(fechaDesest != ""){
    Browser.msgBox("Error", "Esta petición ya fue desestimada anteriormente.", Browser.Buttons.OK)
    return;
  }
  
  var fechaSanc = arrayWF[i][cFS-1]
  if(fechaSanc != ""){
    Browser.msgBox("Error", "Esta petición ya está cerrada.", Browser.Buttons.OK)
    return;
  
  }
  
  var arraySH = sh.getRange(finicioSH,cPetForm,lRowSH - finicioSH +1,1).getValues() //lRowWF - fInicioWF +1
  for(var iSH = 0; iSH <= lRowSH - finicioSH +1 -1; iSH++){        
    var codPetSH = arraySH[iSH][0]
    
    if(Number(CodSol) === Number(codPetSH)){
      sh.getRange(iSH+finicioSH, cCantidadForm).setValue(0)
      break;
    }
  }
  
  

   var notasAd = Browser.inputBox("Registrar Motivo de Desestimación", "Digite la letra correspondiente a un motivo de desestimación de la siguiente lista:\\nA. Validado\\nB. Desestimado por oficina\\nC. RA\\nD. LISTA PRIORIZADA\\nE. Otros", Browser.Buttons.OK_CANCEL);
  var tempNotasAD = notasAd
  
  if(notasAd === ""){return}
  if(notasAd === "cancel"){return}
  
  notasAd = notasAd.toUpperCase()
  if(notasAd === "A"){notasAd = "VALIDADO"}
  else if(notasAd === "B"){notasAd = "DESESTIMADO POR OFICINA"}
  else if(notasAd === "C"){notasAd = "RA"}
  else if(notasAd === "D"){notasAd = "LISTA PRIORIZADA"}
  else if(notasAd === "E"){
    var notasAdComent = Browser.inputBox("¿Cuál es el motivo de la desestimación?", Browser.Buttons.OK_CANCEL)
    if(notasAdComent === "" || notasAdComent === "cancel"){
      return;
    }
  }
  
  notasAd = notasAd.toUpperCase()
  
  switch(notasAd){
      
    case "RA":
      var notasAdicionales = notasAd;
      break;
    case "VALIDADO":
      var notasAdicionales = notasAd;
      break;
    case "DESESTIMADO POR OFICINA":
      var notasAdicionales = notasAd;
      break;
    case "LISTA PRIORIZADA":
      var notasAdicionales = notasAd;
      break;
    case "E":
      var notasAdicionales = notasAdComent
      break;
    default:
      Browser.msgBox("Error", "No se ingresó una letra válida de la lista de motivos de desestimación.", Browser.Buttons.OK)
      ss.toast("Digite la letra correspondiente para el motivo de desestimación; por ejemplo, para 'Rating' digite una 'A'.","Tip",8)
      return
  }

  var fechaCL = arrayWF[i][cCheckList-1]
  var fecha1erController = arrayWF[i][c1erCheckList-1]
  var fechaDevuelveController = arrayWF[i][cDevuelveCheckList-1]
  var fechaRev = arrayWF[i][cRev-1]
  var fechaCBC = arrayWF[i][cCBC-1]
  var fechaDev = arrayWF[i][cDev-1]
  var fechaRcR1 = arrayWF[i][cRcR1-1]
  var fechaFS = arrayWF[i][cFS-1]

  if(fechaRev === "" && fechaDev === "" && fechaCL != ""){
    sheetWF.getRange(fEncontrada,cRev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
    sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
    
    //Correo a BEC
    
    Browser.msgBox("Operación desestimada","Desestimación exitosa. Revisando nuevamente...", Browser.Buttons.OK)
    fWorkflowRiesgos2()
  }
  else if(fechaRev === "" && fecha1erController != "" && fechaDevuelveController != "" && fechaCL === ""){
    sheetWF.getRange(fEncontrada,cCheckList).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDevuelveCheckList).setValue(fecha)
    sheetWF.getRange(fEncontrada,cArchivoDevuelve).setValue("N/A")
    sheetWF.getRange(fEncontrada,cRev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
    
    sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
    sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
    
    //Correo a BEC
    
    Browser.msgBox("Operación desestimada","Desestimación exitosa. Revisando nuevamente...", Browser.Buttons.OK)
    fWorkflowRiesgos2()
  }
  else if(fechaRev === "" && fechaDevuelveController != "" && fechaCL != ""){
    sheetWF.getRange(fEncontrada,cRev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
    sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
    
    //Correo a BEC
    
    Browser.msgBox("Operación desestimada","Desestimación exitosa. Revisando nuevamente...", Browser.Buttons.OK)
    fWorkflowRiesgos2()
  }
  else if(fechaRev != "" && fechaCBC === "" && fechaFS === ""){
    sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
    sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
    
    //Correo a BEC
    
    Browser.msgBox("Operación desestimada","Desestimación exitosa. Revisando nuevamente...", Browser.Buttons.OK)
    fWorkflowRiesgos2()
  }
  else if(fechaRev != "" && fechaCBC != "" && fechaRcR1 === "" && fechaFS === ""){
    sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
    sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
    
    //Correo a BEC
    
    Browser.msgBox("Operación desestimada","Desestimación exitosa. Revisando nuevamente...", Browser.Buttons.OK)
    fWorkflowRiesgos2()
  }
  else if(fechaRev != "" && fechaCBC != "" && fechaRcR1 != "" && fechaFS === ""){
    sheetWF.getRange(fEncontrada,cDev).setValue(fecha)
    sheetWF.getRange(fEncontrada,cDesestimado).setValue(fecha)
    sheetWF.getRange(fEncontrada,cNotasAdicionales).setValue(notasAdicionales) //Escribe la fecha.
    
    //Correo a BEC
    
    Browser.msgBox("Operación desestimada","Desestimación exitosa. Revisando nuevamente...", Browser.Buttons.OK)
    fWorkflowRiesgos2()
  }
}