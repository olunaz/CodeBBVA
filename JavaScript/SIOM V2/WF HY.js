function fWorkflowRiesgos2() {
  //Se necesitan dos de estas macros para asegurar que se esté trabajando en la hoja correcta.
  //Revisa la base de datos en base a un código de solicitud.  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetWF = ss.getSheetByName("WF");
  var sheetEO = ss.getSheetByName("HY") //ESTO ES LO ÚNICO QUE SE TIENE QUE CAMBIAR PARA QUE FUNCIONE EN LA HOJA DE RIESGOS
  var finicioEO = 12
  var finicioWF = 10 //Inicio de la BD del WF
  var lRowWF = sheetWF.getLastRow()
  var lColumnWF = sheetWF.getLastColumn()
  var cDatosE1 = 2
  var cHeadingsE1 = 1
  var cDatosE2 = 4
  var cHeadingsE2 = 3
  var cDatosE3 = 6
  var cHeadingsE3 = 5
  var rNDEV = "B99"
  var rEstEst = "B98"
  var strOfReg = "OFICINA: REGISTRE FECHA"
  var strRiReg = "SEGMENTO EMPRESAS: CONFIRMAR O DEVOLVER"
  var strRiReg2 = "SEGMENTO EMPRESAS: REGISTRE FECHA"
  var strRiReg3 = "SEGMENTO EMPRESAS: CONFIRMAR O CONSULTAR"
  
  var rVerCodSol = sheetEO.getRange("D1")
  
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
  var valtoast = true //Determina cuándo empieza el toast para la fase 2.
  var valtoast2 = true //Determina cuándo empieza el toast para la fase 3.
  var valtoastrec = true //Determina si el recordatorio se dará.
  // Columnas del Drive
  
  //Borrar rango previamente utilizado
  var rngdelete = sheetEO.getRange("A10:I100")
  rngdelete.clear()
  rngdelete.clearNote()
  rngdelete.setVerticalAlignment("middle")
  //Borrar rango previamente utilizado
  
  var val = sheetEO.getRange(1, 3).getValue() //Validación de ENCONTRADO o NO ENCONTRADO
  if(val === "NO ENCONTRADO"){
    ss.toast("Error","Código de petición no encontrado.",3)
    return
  }
  
  ss.toast("Espere por favor...","Inicializando",3) //Se pueden mostrar pop ups de ejecución de macro.
  
  ///////////////////////////////////////////  
  //Formatos de los títulos del WorkFlow
  //Escribir "Estado del WorkFlow"
  
  var rngFormat = sheetEO.getRange("A10:D11")
  
  sheetEO.getRange("A10:D10").merge();
  sheetEO.getRange("A11:B11").merge(); 
  sheetEO.getRange("C11:D11").merge(); 
  
  rngFormat.setBorder(true, true, true, true, true, true)
  rngFormat.setFontWeight("Bold")
  rngFormat.setHorizontalAlignment("center")
  rngFormat.setVerticalAlignment("middle")
  
  var myrngFontSize =[
    [14, "","",""],
    [12,"",12,""]
  ]
  
  var myrngColor =[
    ["#1c4587","","",""],
    ["#a5c2f4","","#a5c2f4",""]
  ]
  
  var myrngHeadings =[
    ["Estado del WorkFlow","","",""],
    ["Ingreso","","Evaluación",""]
  ]
  
  var myrngFontColor =[
    ["white","","",""],
    ["black","","black",""]
  ]
  
  rngFormat.setFontSizes(myrngFontSize)
  rngFormat.setBackgrounds(myrngColor)
  rngFormat.setValues(myrngHeadings)
  rngFormat.setFontColors(myrngFontColor)
  //Formatos de los títulos del WorkFlow
  ///////////////////////////////////////////
  
  
  ss.toast("Recuperando historial de la operación para la fase 'Ingreso'...","Ejecutando",3)
  var codSolEO = sheetEO.getRange(1,2).getValue() //Extrae código de solicitud
  rVerCodSol.setValue(codSolEO)
  
  var arrayWF = sheetWF.getRange(finicioWF-1,1,lRowWF - finicioWF+2,lColumnWF).getValues() //lRowWF - fInicioWF +1
  
  //Loop que encuentra el código de solicitud
  for(var i = 1; i <= lRowWF - finicioWF +1; i++){ //Acá empieza en 1 porque se incluyeron los Headings
    /*for(var i = finicioWF; i <= lRowWF; i++){*/
    var codSolWF = arrayWF[i][cCodSol-1]
    if(codSolWF === codSolEO){           //Compara códigos de solicitud
      var fEncontrada = i + finicioWF -1  //Guarda el valor de la fila del código de solicitud encontrado dentro de la base WF.
      var valSwitch = true //Regula qué caso empieza y cuál no.
      var fEO = finicioEO  //La fila en donde se registrarán los campos del "reporte". 
      //Esta variable cambiará a lo largo del siguiente loop cada vez que se entre a un caso.
      for(var j = cCheckList; j <= lColumnWF; j++){ //Asume que la primera fecha a registrar es el la de la columna del CheckList
        var datoWF = sheetWF.getRange(fEncontrada, j).getValue() //Obtiene el valor de la BD en la fila del código de solicitud encontrada, listo para ser transferido. Si fuera nulo en las validaciones siguientes, entonces termina.
        var celdaEO = sheetEO.getRange(fEO,cDatosE1) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
        var celdaHEO = sheetEO.getRange(fEO, cHeadingsE1) //Fija celda de los headings en la hoja de las estaciones.
        
        switch(j){   //La variable "j" equivale a las columnas de la base WF. Cada caso es una columna que corresponde a un heading de la base.
          case cCheckList:
            var valMoveCL = false
            var valFechaDev = arrayWF[i][cDevuelveCheckList-1]
            if(valFechaDev === ""){
              celdaHEO.setValue("Check List")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)
              fEO = fEO +1 //Aquí cambia la variable fEO y aumenta una fila más. Esto se repite en cada caso.
              if (datoWF === ""){
                if(sheetWF.getRange(fEncontrada, c1erCheckList).getValue() === ""){
                  var valSwitch = false
                  }
                else{
                  var valSwitchOff = true
                  }
                celdaEO.setValue(strOfReg)
                celdaEO.setFontWeight("bold")
                celdaEO.setBackgroundRGB(255,153,0)
                celdaEO.setBorder(true, true, true, true, true, true)    
                celdaEO.setHorizontalAlignment("center")
                sheetEO.getRange(rEstEst).setValue("I1")
                sheetEO.getRange(rEstEst).setFontColor("white")              
              }
              else{
                celdaEO.setValue(datoWF)
                celdaEO.setFontWeight("bold")
                celdaEO.setBackgroundRGB(85,199,104)
                celdaEO.setBorder(true, true, true, true, true, true)
                celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                celdaEO.setHorizontalAlignment("center")
                sheetEO.getRange(rEstEst).setValue("I1")
                sheetEO.getRange(rEstEst).setFontColor("white")
              }
            }
            else if(valFechaDev != ""){
              valMoveCL = true
            }
            break;
            
          case c1erCheckList:
            if (valSwitch === false){break}
            if (datoWF != ""){
              celdaHEO.setValue("1er Ingreso Check List")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)  
              
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I1")
              sheetEO.getRange(rEstEst).setFontColor("white")              
              
              fEO = fEO + 1
            }
            break;
            
          case cDevuelveCheckList:
            if (valSwitch === false){break}
            if (datoWF != ""){
              celdaHEO.setValue("Devuelve")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)               
              
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I1")
              sheetEO.getRange(rEstEst).setFontColor("white")
              
              var offsetmore = 0
              
              var motivoDev = sheetWF.getRange(fEncontrada, cMotivCL).getValue()
              if(motivoDev != ""){
                offsetmore = 1
                celdaHEO.offset(1,0).setValue("Motivo Devolución")
                celdaHEO.offset(1,0).setFontWeight("bold")
                celdaHEO.offset(1,0).setHorizontalAlignment("center")
                celdaHEO.offset(1,0).setBackgroundRGB(233,233,233)
                celdaHEO.offset(1,0).setFontColor("black")
                celdaHEO.offset(1,0).setBorder(true, true, true, true, true, true)               
                
                
                celdaEO.offset(1,0).setFontSize(8)
                celdaEO.offset(1,0).setWrap(true)
                celdaEO.offset(1,0).setFontWeight("bold")
                celdaEO.offset(1,0).setValue(motivoDev)
                celdaEO.offset(1,0).setBackgroundRGB(85,199,104)
                celdaEO.offset(1,0).setBorder(true, true, true, true, true, true)
                celdaEO.offset(1,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                celdaEO.offset(1,0).setHorizontalAlignment("center")
                sheetEO.getRange(rEstEst).setValue("I1")
                sheetEO.getRange(rEstEst).setFontColor("white")
                
                fEO = fEO + 1
              }

              var ArchivoDev = sheetWF.getRange(fEncontrada, cArchivoDevuelve).getValue()
              if(ArchivoDev != ""){
                celdaHEO.offset(1+offsetmore,0).setValue("Archivo Devuelve")
                celdaHEO.offset(1+offsetmore,0).setFontWeight("bold")
                celdaHEO.offset(1+offsetmore,0).setHorizontalAlignment("center")
                celdaHEO.offset(1+offsetmore,0).setBackgroundRGB(233,233,233)
                celdaHEO.offset(1+offsetmore,0).setFontColor("black")
                celdaHEO.offset(1+offsetmore,0).setBorder(true, true, true, true, true, true)               
                
                
                celdaEO.offset(1+offsetmore,0).setFontSize(8)
                celdaEO.offset(1+offsetmore,0).setWrap(true)
                celdaEO.offset(1+offsetmore,0).setFontWeight("bold")
                celdaEO.offset(1+offsetmore,0).setValue(ArchivoDev)
                celdaEO.offset(1+offsetmore,0).setBackgroundRGB(85,199,104)
                celdaEO.offset(1+offsetmore,0).setBorder(true, true, true, true, true, true)
                celdaEO.offset(1+offsetmore,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                celdaEO.offset(1+offsetmore,0).setHorizontalAlignment("center")
                sheetEO.getRange(rEstEst).setValue("I1")
                sheetEO.getRange(rEstEst).setFontColor("white")
                
                fEO = fEO + 1
              }
              

              
              if(valMoveCL === true){
                fEO = fEO + 1
                var datoWF = sheetWF.getRange(fEncontrada, cCheckList).getValue() //Obtiene el valor de la BD en la fila del código de solicitud encontrada, listo para ser transferido. Si fuera nulo en las validaciones siguientes, entonces termina.
                var celdaEO = sheetEO.getRange(fEO,cDatosE1) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
                var celdaHEO = sheetEO.getRange(fEO, cHeadingsE1) //Fija celda de los headings en la hoja de las estaciones.
                celdaHEO.setValue("Check List")
                celdaHEO.setFontWeight("bold")
                celdaHEO.setHorizontalAlignment("center")
                celdaHEO.setBackgroundRGB(233,233,233)
                celdaHEO.setFontColor("black")
                celdaHEO.setBorder(true, true, true, true, true, true)
                fEO = fEO +1 //Aquí cambia la variable fEO y aumenta una fila más. Esto se repite en cada caso.
                if (datoWF === ""){
                  if(sheetWF.getRange(fEncontrada, cRev).getValue() === ""){
                    var valSwitch = false
                    }
                  else{
                    var valSwitchOff = true
                    }
                  celdaEO.setValue(strOfReg)
                  celdaEO.setFontWeight("bold")
                  celdaEO.setBackgroundRGB(255,153,0)
                  celdaEO.setBorder(true, true, true, true, true, true)    
                  celdaEO.setHorizontalAlignment("center")
                  sheetEO.getRange(rEstEst).setValue("I1")
                  sheetEO.getRange(rEstEst).setFontColor("white")              
                }
                else{
                  celdaEO.setValue(datoWF)
                  celdaEO.setFontWeight("bold")
                  celdaEO.setBackgroundRGB(85,199,104)
                  celdaEO.setBorder(true, true, true, true, true, true)
                  celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                  celdaEO.setHorizontalAlignment("center")
                  sheetEO.getRange(rEstEst).setValue("I1")
                  sheetEO.getRange(rEstEst).setFontColor("white")
                }
              }
              
              if(valMoveCL === false){fEO = fEO + 1} //Se usa para que no se muevan las cajas de las estaciones. Si no, se saltaría un fEO.
            }
            break;          
            
          case cRev:
            if (valSwitch === false){break}
            if (valSwitchOff === false){
              ss.toast("Estaciones actualizadas.","Fin",5)
              return;
            }
            fEO = fEO +1
            if (datoWF === ""){
              celdaHEO.setValue("Revisión")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)
              var valSwitch = false
              celdaEO.setValue(strRiReg)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(255,255,0)
              celdaEO.setBorder(true, true, true, true, true, true)    
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I1")
              sheetEO.getRange(rEstEst).setFontColor("white")
              
              if(arrayWF[i][cCheckList-1] != "" && arrayWF[i][c1erCheckList -1] != "" && arrayWF[i][cRev -1] === ""){
                //var celdaHEOhistorial = sheetEO.getRange(fEO, cHeadingsE1).offset(-4,0)
                //var historial = celdaHEOhistorial.getValue()
                //if(historial === "1er Ingreso Check List"){
                celdaEO.setValue(strRiReg2)
                var nombrehoja = sheetEO.getSheetName()
                if (nombrehoja === "HY"){
                  sheetEO.getRange(rNDEV).setValue("NDEV") //Aquí se define si se puede devolver o no una operación.
                  sheetEO.getRange(rNDEV).setFontColor("white")
                }
              }
            }
            else{
              fEO = finicioEO
              var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
              var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.                                       
              
              celdaHEO.setValue("Revisión")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)
              fEO = fEO +1
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
            }
            break;     
            
          case cCBC:
            if (valSwitch === false){break} 
            if (datoWF != ""){
              var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
              var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.              
              celdaHEO.setValue("Consulta de SE a BE")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)            
              fEO = fEO + 1
              
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")    

              var valFS = sheetWF.getRange(fEncontrada, cFS).getValue()
              var valDevRespuestas = sheetWF.getRange(fEncontrada, cDev).getValue()
              if(valFS != "" && valDevRespuestas === ""){
                celdaHEO.offset(1,0).setValue("Fecha Cierre")
                celdaHEO.offset(1,0).setFontWeight("bold")
                celdaHEO.offset(1,0).setHorizontalAlignment("center")
                celdaHEO.offset(1,0).setBackgroundRGB(233,233,233)
                celdaHEO.offset(1,0).setFontColor("black")
                celdaHEO.offset(1,0).setBorder(true, true, true, true, true, true)
                fEO = fEO +1
                
                var valDevFS = arrayWF[i][cCBC-1]
                var valRespFS = arrayWF[i][cRcR1-1]
                var valFS = arrayWF[i][cFS -1]
                var valDecCL = arrayWF[i][cArchivoDevuelve -1] 
                var valTM = arrayWF[i][cTipoMetodWF-1]
                
                if (valDevFS === "" && valRespFS === "" && valFS === ""){
                }
                else if(valRespFS != "" && valFS === "" ){
                }
                else if(valFS != ""){
                  valtoastrec = false
                  celdaEO.offset(1,0).setValue(datoWF)
                  celdaEO.offset(1,0).setFontWeight("bold")
                  celdaEO.offset(1,0).setBackgroundRGB(85,199,104)
                  celdaEO.offset(1,0).setBorder(true, true, true, true, true, true)
                  celdaEO.offset(1,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                  celdaEO.offset(1,0).setHorizontalAlignment("center")
                  
                  sheetEO.getRange(rEstEst).setValue("I3")
                  sheetEO.getRange(rEstEst).setFontColor("white")
                  sheetEO.getRange(rNDEV).setValue("NDEV") //Aquí se define si se puede devolver o no una operación.
                  sheetEO.getRange(rNDEV).setFontColor("white")
                  
                  
                  celdaHEO.offset(2,0).setValue("Tipo de Asignación")
                  celdaHEO.offset(2,0).setFontWeight("bold")
                  celdaHEO.offset(2,0).setHorizontalAlignment("center")
                  celdaHEO.offset(2,0).setBackgroundRGB(233,233,233)
                  celdaHEO.offset(2,0).setFontColor("black")
                  celdaHEO.offset(2,0).setBorder(true, true, true, true, true, true)
                  
                  celdaHEO.offset(3,0).setValue("Tipo de Sanción")
                  celdaHEO.offset(3,0).setFontWeight("bold")
                  celdaHEO.offset(3,0).setHorizontalAlignment("center")
                  celdaHEO.offset(3,0).setBackgroundRGB(233,233,233)
                  celdaHEO.offset(3,0).setFontColor("black")
                  celdaHEO.offset(3,0).setBorder(true, true, true, true, true, true)
                  
                  
                  datoWF = sheetWF.getRange(fEncontrada, cAmb).getValue()
                  
                  celdaEO.offset(2,0).setValue(datoWF)
                  celdaEO.offset(2,0).setFontWeight("bold")
                  celdaEO.offset(2,0).setBackgroundRGB(85,199,104)
                  celdaEO.offset(2,0).setBorder(true, true, true, true, true, true)
                  celdaEO.offset(2,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                  celdaEO.offset(2,0).setHorizontalAlignment("center")
                  
                  datoWF = sheetWF.getRange(fEncontrada, cTS).getValue()
                  
                  celdaEO.offset(3,0).setValue(datoWF)
                  celdaEO.offset(3,0).setFontWeight("bold")
                  celdaEO.offset(3,0).setBackgroundRGB(85,199,104)
                  celdaEO.offset(3,0).setBorder(true, true, true, true, true, true)
                  celdaEO.offset(3,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                  celdaEO.offset(3,0).setHorizontalAlignment("center")              
                }                
                valSwitch = false
              }

            }
            break;
            
          case cRcR1:
            if (valSwitch === false){break} 
            
            var celdavaldatoWF2 = sheetWF.getRange(fEncontrada, cCBC)
            var valdatoWF2 = celdavaldatoWF2.getValue()
            
            if (valdatoWF2 === ""){break}          
            
            //Campo "Devolución"
            
            if (datoWF === ""){
              
              var datoDevWF = sheetWF.getRange(fEncontrada, cDev).getValue()
              var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
              var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.              
              celdaHEO.setValue("Devolución")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)            
              
              fEO = fEO +1
              
              if (datoDevWF === ""){
                var valSwitch = false
                celdaEO.setValue(strRiReg2)
                celdaEO.setFontWeight("bold")
                celdaEO.setBackgroundRGB(255,255,0)
                celdaEO.setBorder(true, true, true, true, true, true)    
                celdaEO.setHorizontalAlignment("center")
                sheetEO.getRange(rEstEst).setValue("I2")
                sheetEO.getRange(rEstEst).setFontColor("white")
              }
              else{
                var valDev = true
                celdaEO.setValue(datoDevWF)
                celdaEO.setFontWeight("bold")
                celdaEO.setBackgroundRGB(85,199,104)
                celdaEO.setBorder(true, true, true, true, true, true)
                celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
                celdaEO.setHorizontalAlignment("center")
                
                sheetEO.getRange(rEstEst).setValue("I2")
                sheetEO.getRange(rEstEst).setFontColor("white")
                
              }                 
            }
            
            var datoDevWF = sheetWF.getRange(fEncontrada, cDev).getValue()            
            if(datoWF != "" && datoDevWF != ""){
              var datoDevWF = sheetWF.getRange(fEncontrada, cDev).getValue()
              var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
              var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.              
              celdaHEO.setValue("Devolución")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)            
              
              fEO = fEO +1
              
              var valDev = true
              celdaEO.setValue(datoDevWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")              
              
            }
            
            //Campo Reingreso con Respuestas (1)            
            var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
            var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.              
            celdaHEO.setValue("Reingreso con Respuestas")
            celdaHEO.setFontWeight("bold")
            celdaHEO.setHorizontalAlignment("center")
            celdaHEO.setBackgroundRGB(233,233,233)
            celdaHEO.setFontColor("black")
            celdaHEO.setBorder(true, true, true, true, true, true)            
            fEO = fEO + 1
            
            if (datoWF === ""){
              var valSwitch = false
              
              var valdatoCBC = sheetWF.getRange(fEncontrada, cCBC).getValue()              
              
              celdaEO.setValue(strOfReg)
              celdaEO.setBackgroundRGB(255,153,0)
              celdaEO.setFontWeight("bold")
              
              celdaEO.setBorder(true, true, true, true, true, true)    
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
              
              if(sheetEO.getName() === "E1" && celdaEO.offset(-1,0).getValue() != strRiReg2){Browser.msgBox("Mensaje para Oficina", "Segmento Empresas ha devuelto la operación por falta de respuestas.", Browser.Buttons.OK)}  
            }
            else{
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
            }   
            
            var ArchivoResp = sheetWF.getRange(fEncontrada, cArchivoResp).getValue()
            if(ArchivoResp != ""){
              var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
              var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.              
              celdaHEO.setValue("Archivo Respuestas")
              celdaHEO.setFontWeight("bold")
              celdaHEO.setHorizontalAlignment("center")
              celdaHEO.setBackgroundRGB(233,233,233)
              celdaHEO.setFontColor("black")
              celdaHEO.setBorder(true, true, true, true, true, true)               
              
              celdaEO.setFontSize(8)
              celdaEO.setWrap(true)
              celdaEO.setFontWeight("bold")
              celdaEO.setValue(ArchivoResp)
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I1")
              sheetEO.getRange(rEstEst).setFontColor("white")
              
              fEO = fEO + 1
            }
            
            break;              
            /////////////////////////////////////////////////////1ra ronda de preguntas
            
            
          /*case cFEv:
            if (valSwitch === false){break}
            
            if (valtoast === true){ss.toast("Recuperando historial de la operación para la fase 'Evaluación I'...","Ejecutando",3)}
            valtoast = false
            
            var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
            var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.                   
            celdaHEO.setValue("Fin Evaluación (VB Jefe)")
            celdaHEO.setFontWeight("bold")
            celdaHEO.setHorizontalAlignment("center")
            celdaHEO.setBackgroundRGB(233,233,233)
            celdaHEO.setFontColor("black")
            celdaHEO.setBorder(true, true, true, true, true, true)
            fEO = fEO +1
            if (datoWF === ""){
              var valSwitch = false
              
              var valnDEV = sheetEO.getRange(rNDEV)
              var vnDEV = valnDEV.getValue()
              
              if (vnDEV === ""){celdaEO.setValue(strRiReg3)}
              else{celdaEO.setValue(strRiReg2)}
              
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(255,255,0)
              celdaEO.setBorder(true, true, true, true, true, true)    
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
            }
            else{
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
            }
            break;*/
       
          case cFS:
            if (valSwitch === false){break}
            var celdaEO = sheetEO.getRange(fEO,cDatosE2) //Fija celda de datos a los que se transferirá. Usa la columna de los datos de estación 1.
            var celdaHEO = sheetEO.getRange(fEO, cHeadingsE2) //Fija celda de los headings en la hoja de las estaciones.                   
            celdaHEO.setValue("Fecha Cierre")
            celdaHEO.setFontWeight("bold")
            celdaHEO.setHorizontalAlignment("center")
            celdaHEO.setBackgroundRGB(233,233,233)
            celdaHEO.setFontColor("black")
            celdaHEO.setBorder(true, true, true, true, true, true)
            fEO = fEO +1
            
            var valDevFS = arrayWF[i][cCBC-1]
            var valRespFS = arrayWF[i][cRcR1-1]
            var valFS = arrayWF[i][cFS -1]
            var valDecCL = arrayWF[i][cArchivoDevuelve -1] 
            var valTM = arrayWF[i][cTipoMetodWF-1]
            
            Logger.clear()
            Logger.log(valDecCL)
            Logger.log(valTM)
            
            if (valDevFS === "" && valRespFS === "" && valFS === ""){
              var valSwitch = false
              
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(255,255,0)
              celdaEO.setBorder(true, true, true, true, true, true)    
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
              if(valDecCL != "N/A" && valTM === "Metodización Avanzada"){
                celdaEO.setValue(strRiReg)
                sheetEO.getRange(rNDEV).setValue("") //Aquí se define si se puede devolver o no una operación.
                sheetEO.getRange(rNDEV).setFontColor("white")              
              }
              else{
                celdaEO.setValue(strRiReg2)
                sheetEO.getRange(rNDEV).setValue("NDEV") //Aquí se define si se puede devolver o no una operación.
                sheetEO.getRange(rNDEV).setFontColor("white")              
              }
            }
            else if(valRespFS != "" && valFS === "" ){
              var valSwitch = false
              celdaEO.setValue(strRiReg2)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(255,255,0)
              celdaEO.setBorder(true, true, true, true, true, true)    
              celdaEO.setHorizontalAlignment("center")
              sheetEO.getRange(rEstEst).setValue("I2")
              sheetEO.getRange(rEstEst).setFontColor("white")
              sheetEO.getRange(rNDEV).setValue("NDEV") //Aquí se define si se puede devolver o no una operación.
              sheetEO.getRange(rNDEV).setFontColor("white")
            }
            else if(valFS != ""){
              valtoastrec = false
              celdaEO.setValue(datoWF)
              celdaEO.setFontWeight("bold")
              celdaEO.setBackgroundRGB(85,199,104)
              celdaEO.setBorder(true, true, true, true, true, true)
              celdaEO.setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.setHorizontalAlignment("center")
              
              sheetEO.getRange(rEstEst).setValue("I3")
              sheetEO.getRange(rEstEst).setFontColor("white")
              sheetEO.getRange(rNDEV).setValue("NDEV") //Aquí se define si se puede devolver o no una operación.
              sheetEO.getRange(rNDEV).setFontColor("white")
              
              
              celdaHEO.offset(1,0).setValue("Tipo de Asignación")
              celdaHEO.offset(1,0).setFontWeight("bold")
              celdaHEO.offset(1,0).setHorizontalAlignment("center")
              celdaHEO.offset(1,0).setBackgroundRGB(233,233,233)
              celdaHEO.offset(1,0).setFontColor("black")
              celdaHEO.offset(1,0).setBorder(true, true, true, true, true, true)
              
              celdaHEO.offset(2,0).setValue("Tipo de Sanción")
              celdaHEO.offset(2,0).setFontWeight("bold")
              celdaHEO.offset(2,0).setHorizontalAlignment("center")
              celdaHEO.offset(2,0).setBackgroundRGB(233,233,233)
              celdaHEO.offset(2,0).setFontColor("black")
              celdaHEO.offset(2,0).setBorder(true, true, true, true, true, true)
              
              
              datoWF = sheetWF.getRange(fEncontrada, j).offset(0,1).getValue()
              
              celdaEO.offset(1,0).setValue(datoWF)
              celdaEO.offset(1,0).setFontWeight("bold")
              celdaEO.offset(1,0).setBackgroundRGB(85,199,104)
              celdaEO.offset(1,0).setBorder(true, true, true, true, true, true)
              celdaEO.offset(1,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.offset(1,0).setHorizontalAlignment("center")
              
              datoWF = sheetWF.getRange(fEncontrada, j).offset(0,2).getValue()
              
              celdaEO.offset(2,0).setValue(datoWF)
              celdaEO.offset(2,0).setFontWeight("bold")
              celdaEO.offset(2,0).setBackgroundRGB(85,199,104)
              celdaEO.offset(2,0).setBorder(true, true, true, true, true, true)
              celdaEO.offset(2,0).setNumberFormat("DD"+"/"+"MM"+"/"+"YYYY")
              celdaEO.offset(2,0).setHorizontalAlignment("center")              
            }
            break;                              
            //Empieza Evaluación II            
        } 
      }   
      break; //Rompe el Loop que busca la fila del código de consulta
    }
  }
  
  SpreadsheetApp.flush()
  
  if(valtoastrec === true){
    ss.toast("Estaciones actualizadas.","Fin",5)
  }
  else{
    ss.toast("La operación que ha consultado ha concluido.","Operación Finalizada",5)
  }
}




function RegistrarFechaRiesgos2() {
  //Registra cambios en fechas
  //Compara los headings de la sheetEO con los de la base WF. Ambos tienen que ser IDÉNTICOS para que funcione. 
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
  var cHerramienta = 44
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
    }
  }
  
  //Solo para la estación 1
  if(valEstEst === "I1"){
    for(var i = finicioEO; i<=lRowEO; i++){
      var celda = sheetEO.getRange(i, cDatosE1)
      var val = celda.getValue()
      if (val === strRiReg || val === strRiReg2 || val === strRiReg3){ //Encuentra la instancia en la que riesgos puede ingresar una fecha.        
        var celdavaldev = sheetEO.getRange(rNDEV)
        var valdev = celdavaldev.getValue()
        
        if(valdev != "NDEV"){
          var respuesta = Browser.msgBox("Segmento Empresas", "¿Desea devolver la operación a la oficina? Si presiona no, entonces se procederá a registrar la fecha de la Revisión. En caso contrario, se mandará un correo a la oficina para que responda a las consultas de devolución.", Browser.Buttons.YES_NO)
          if (respuesta === "no") {var decRi = "no"}
          else if(respuesta === "cancel"){return;}
          else{var decRi = "yes"}
          
        }
        else{var decRi = "no"}

        var celdaEO = sheetEO.getRange(i, cHeadingsE1)
        var headingEO = celdaEO.getValue() //Obtiene el heading de la hoja de estación.
        
        for (var j = 1; j<=lColumnWF; j++){ //Recorre las columnas de la base WF.
          var celdaWF = sheetWF.getRange(fHeadingsWF,j)
          var headingWF = celdaWF.getValue() //Obtiene el heading de la base WF.
          
          if(headingWF === headingEO){ //Encuentra la instancia en la que se igualan ambos valores de los headings.
            
            if(headingWF === "Revisión" && decRi === "yes"){
              var motivo = Browser.inputBox("Segmento Empresas", "Digite la letra correspondiente al motivo de la devolución: \\nA. Errores en el Check List \\nB. Calidad de Documentos", Browser.Buttons.OK_CANCEL)
              
              if (motivo === "cancel"){return}
              
              motivo = motivo.toUpperCase()
              if(motivo === "A"){motivo = "Errores en el Check List"}
              else if(motivo === "B"){motivo = "Calidad de Documentos"}
              else{
                Browser.msgBox("Error", "Digite A o B", Browser.Buttons.OK)
                return;
              }
              
              
              var respuestamail = "yes"
              if (respuestamail === "yes"){
                if(Correo != "SIN CORREO"){
                  var recipient = Correo
                  }
                else if (Correo === "SIN CORREO"){
                  var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
                  }
                if (recipient === "cancel"){return}
                var subject = "SIOM: Devolución de Petición " + CodSol + " (" + Cliente + ") en el Workflow de Metodización"
                var body = "Segmento Empresas ha devuelto la operación con código de petición: " + CodSol
                + ".\n"
                + "\nInformación general sobre la operación: "
                + "\nCódigo Central: " + CodCentral
                + "\nGestor: " + Ejecutivo
                + "\nCliente: " + Cliente 
                + "\nGrupo : " + Grupo 
                + "\nTipo de Estado: " + Oper
                +  "\nAnalista Asignado: " + AnAsig
                + "\n\nDetalle de Balance General: " + DBG
                + "\n\nDetalle de Estado de Resultados: " + DER
                + "\nArchivo Adjunto: " + ArchAdj
                + "\n\nMotivo de Devolución: " + motivo
                + "\n\n\nIngrese al siguiente formulario para registrar las respuestas a esta devolución. Elija la opción 'Respuestas sobre Check List' en la pregunta de '¿Qué desea reingresar?':"
                + "\nhttps://docs.google.com/a/bbva.com/forms/d/e/1FAIpQLScJXlEeqfg1FqWzhvPWGxR3804XNvdDxMHR_oE4uc2tv9HrUw/viewform?usp=sf_link"
                
                var options = {cc: correoResp + ", wendy.amaya.tobon@bbva.com, siom@bbva.com"}
                
                MailApp.sendEmail(recipient,subject,body,options);
              }
              var fechainicial = sheetWF.getRange(fEncontrada,cCheckList).getValue()
              
              var arraySH = sh.getRange(finicioSH,cPetForm,lRowSH - finicioSH +1,1).getValues() //lRowWF - fInicioWF +1
              for(var iSH = 0; iSH <= lRowSH - finicioSH +1 -1; iSH++){        
                var codPetSH = arraySH[iSH][0]
                
                if(Number(CodSol) === Number(codPetSH)){
                  sh.getRange(iSH+finicioSH, cCantidadForm).setValue(0)
                  break;
                }
              }
              
              sheetWF.getRange(fEncontrada,c1erCheckList).setValue(fechainicial)
              sheetWF.getRange(fEncontrada,cCheckList).setValue("")
              sheetWF.getRange(fEncontrada,cDevuelveCheckList).setValue(fecha)
              sheetWF.getRange(fEncontrada,cMotivCL).setValue(motivo) //Escribe la fecha.
              Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
              fWorkflowRiesgos2()
              return;
            }
            else if(headingWF === "Revisión" && decRi === "no"){
              var respuestamail = "yes" //Browser.msgBox("Correo de Asignación de RVGL", "Esta acción notificará al gestor del registro de fecha. ¿Desea continuar?", Browser.Buttons.YES_NO)      
              if (respuestamail === "yes"){
                var recipient = Correo
                if (recipient === "cancel"){return}
                var subject = "SIOM: La Petición de Metodización con Código " + CodSol + " (" + Cliente + ") Ha Sido Aceptada"
                var body = "Segmento Empresas ha aceptado la petición con código: " + CodSol
                + ".\n"
                + "\nInformación general sobre la operación: "
                + "\nCódigo Central: " + CodCentral
                + "\nGestor: " + Ejecutivo
                + "\nCliente: " + Cliente 
                + "\nGrupo : " + Grupo 
                + "\nTipo de Estado: " + Oper
                +  "\nAnalista Asignado: " + AnAsig
                + "\n\nDetalle de Balance General: " + DBG
                + "\n\nDetalle de Estado de Resultados: " + DER
                + "\nArchivo Adjunto: " + ArchAdj
                + "\n\n\nPor el momento, no requiere realizar ninguna acción adicional en cuanto a esta petición."
              }
              else{return;}
              
              var options = {cc: correoResp + ", wendy.amaya.tobon@bbva.com, siom@bbva.com"}
              MailApp.sendEmail(recipient,subject,body,options);
              
              sheetWF.getRange(fEncontrada,cRev).setValue(fecha) //Escribe la fecha.
              celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
              celda.setBackgroundRGB(85,199,104)
              
              Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
              
              fWorkflowRiesgos2()
              
              return;
            }
          }
        }
        break;
      }
    }
  }
  //Solo para la estación 1
  
  //Solo para la estación 2
  else if(valEstEst === "I2"){
    for(var i = finicioEO; i<=lRowEO; i++){
      var celda = sheetEO.getRange(i, cDatosE2)
      var val = celda.getValue()
      if (val === strRiReg || val === strRiReg2 || val === strRiReg3){ //Encuentra la instancia en la que riesgos puede ingresar una fecha.
        var celdavaldev = sheetEO.getRange(rNDEV)
        var valdev = celdavaldev.getValue()        
        var celdaEO = sheetEO.getRange(i, cHeadingsE2)
        var headingEO = celdaEO.getValue() //Obtiene el heading de la hoja de estación.
        
        if(headingEO === "Fecha Cierre"){
          if(valdev != "NDEV"){
            
            var respuesta = Browser.msgBox("Segmento Empresas", "¿Desea proceder al cierre de la petición? Presione 'Sí' solo si concluyó. En caso contrario presione 'No' para registrar la fecha de las consultas enviadas a la Oficina.", Browser.Buttons.YES_NO)
            
            if (respuesta == "no") {
              var decRi = "no"
              }
            else if(respuesta === "cancel"){
              return
            }
            else{
              var decRi = "yes"
              }
          }
          else{
            var decRi = "no"
            }                  
        }
        /*        else if(headingEO === "Consulta de SE a BE"){
        
        var respuesta = Browser.msgBox("Riesgos", "¿El cliente respondió? Presione 'Sí' solo si se cuenta con las respuestas. En caso contrario presione 'No' para solicitar intervención del ejecutivo de cuenta.", Browser.Buttons.YES_NO)
        
        if (respuesta == "no") {
        var decRi = "no"
        
        var respuestamail = Browser.msgBox("Correo de Solicitud", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)
        if (respuestamail === "yes"){
        if(Correo != "SIN CORREO"){
        var recipient = Correo
        }
        else if (Correo === "SIN CORREO"){
        var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
        }
        if (recipient === "cancel"){return}
        var subject = "Solicitud de Reingreso de Respuestas en el Workflow"
        var body = "Se solicita su intervención para el reingreso de respuestas en el Workflow para la operación con código de solicitud: " + CodSol
        
        + ".\n"
        + "\nInformación general sobre la operación: "
        + "\nCódigo Central: " + CodCentral
        + "\nGestor: " + Ejecutivo
        + "\nCliente: " + Cliente 
        + "\nGrupo : " + Grupo 
        + "\nOperación: " + Oper
        //+ "\nMonto Propuesto (Miles de US$): " +MontoProp
        
        //var options = {bcc: correoGOF}
        
        MailApp.sendEmail(recipient,subject,body/*,options*//*);
        }              
        }
        else if(respuesta === "cancel"){
        return
        }
        else{
        var decRi = "yes"
        }
        }
        else if(headingEO === "Consulta de BE a Cliente (2)"){
        var respuesta = Browser.msgBox("Riesgos", "¿El cliente respondió? Presione 'Sí' solo si se cuenta con las respuestas. En caso contrario presione 'No' para solicitar intervención del ejecutivo de cuenta.", Browser.Buttons.YES_NO)
        
        if (respuesta == "no") {
        var decRi = "no"
        
        var respuestamail = Browser.msgBox("Correo de Solicitud", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)
        if (respuestamail === "yes"){
        if(Correo != "SIN CORREO"){
        var recipient = Correo
        }
        else if (Correo === "SIN CORREO"){
        var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
        }
        if (recipient === "cancel"){return}
        var subject = "Solicitud de Reingreso de Respuestas en el Workflow"
        var body = "Se solicita su intervención para el reingreso de respuestas en el Workflow para la operación con código de solicitud: " + CodSol
        + ".\n"
        + "\nInformación general sobre la operación: "
        + "\nCódigo Central: " + CodCentral
        + "\nGestor: " + Ejecutivo
        + "\nCliente: " + Cliente 
        + "\nGrupo : " + Grupo 
        + "\nOperación: " + Oper
        + "\nMonto Propuesto (Miles de US$): " +MontoProp
        
        var options = {bcc: correoGOF}
        
        MailApp.sendEmail(recipient,subject,body,options);
        }
        }
        else if(respuesta === "cancel"){
        return
        }
        else{
        var decRi = "yes"
        }        
        }*/
        
        
        else if(headingEO === "Devolución"){}
        else if(headingEO === "Reingreso con Respuestas"){}
        /*else if(headingEO === "Devolución"){
        var respuesta = Browser.msgBox("Segmento Empresas", "¿Desea devolver la operación?", Browser.Buttons.YES_NO)
        
        if(respuesta  == "no") {
        var decRi = "no"
        }
        else if(respuesta === "cancel"){
        return 
        }
        else{
        var decRi = "yes"
        }
        }*/
        
        
        for (var j = 1; j<=lColumnWF; j++){ //Recorre las columnas de la base WF.
          var celdaWF = sheetWF.getRange(fHeadingsWF,j)
          var headingWF = celdaWF.getValue() //Obtiene el heading de la base WF.
          
          if(headingWF === headingEO){ //Encuentra la instancia en la que se igualan ambos valores de los headings.
            var celdaHEOCRC = sheetEO.getRange(finicioEO+1, cHeadingsE2)
            var HEOCRC = celdaHEOCRC.getValue()     
            var celdanDEV = sheetEO.getRange(rNDEV)
            var valNDEV = celdanDEV.getValue()
            
            if(headingWF === "Fecha Cierre" && decRi === "no"/* && HEOCRC != "Consulta de Riesgos a Cliente" */&& valNDEV === ""){
              
              var recipient = Correo
              var subject = "SIOM: Consultas de Segmento Empresas Acerca de la Petición " + CodSol + " (" + Cliente + ") en el Workflow de Metodización"
              var body = "Segmento Empresas tiene consultas acerca de la operación con código de petición: " + CodSol
              + ".\n"
              + "\nInformación general sobre la operación: "
              + "\nCódigo Central: " + CodCentral
              + "\nGestor: " + Ejecutivo
              + "\nCliente: " + Cliente 
              + "\nGrupo : " + Grupo 
              + "\nTipo de Estado: " + Oper
              +  "\nAnalista Asignado: " + AnAsig
              + "\n\nDetalle de Balance General: " + DBG
              + "\n\nDetalle de Estado de Resultados: " + DER
              + "\nArchivo Adjunto: " + ArchAdj
              + "\n\n\nIngrese al siguiente formulario para registrar las respuestas a las consultas de metodización. Elija la opción 'Respuestas sobre Metodización' en la pregunta de '¿Qué desea reingresar?':"
              + "\nhttps://docs.google.com/a/bbva.com/forms/d/e/1FAIpQLScJXlEeqfg1FqWzhvPWGxR3804XNvdDxMHR_oE4uc2tv9HrUw/viewform?usp=sf_link"
              
              var options = {cc: correoResp + ", wendy.amaya.tobon@bbva.com, siom@bbva.com"}
              
              MailApp.sendEmail(recipient,subject,body,options);
              
              var arraySH = sh.getRange(finicioSH,cPetForm,lRowSH - finicioSH +1,1).getValues() //lRowWF - fInicioWF +1
              for(var iSH = 0; iSH <= lRowSH - finicioSH +1 -1; iSH++){        
                var codPetSH = arraySH[iSH][0]
                
                if(Number(CodSol) === Number(codPetSH)){
                  sh.getRange(iSH+finicioSH, cCantidadForm).setValue(0)
                  break;
                }
              }
              
              sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
              Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
              fWorkflowRiesgos2()
              return;
            }
            /*else if(headingWF === "Fin Evaluación (VB Jefe)" && decRi === "no" && HEOCRC === "Consulta de Riesgos a Cliente" && valNDEV === ""){
            sheetWF.getRange(fEncontrada,c2daCRC).setValue(fecha)
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            fWorkflowRiesgos2()
            return;
            }*/
            else if(headingWF === "Fecha Cierre" && decRi === "yes" && valNDEV === ""){
              
              var decRespCL = sheetWF.getRange(fEncontrada, cArchivoDevuelve).getValue()
              var decRespMetod = sheetWF.getRange(fEncontrada, cArchivoResp).getValue()
              var taFS = "" //El tipo de asignación (Limpia o No Disponible)
              var tsFS = "" //El tipo de sanción.
              
              var respCompl = ""
              if(sheetWF.getRange(fEncontrada, cRcR1).getValue() != ""){
                respCompl = Browser.msgBox("Segmento Empresas", "¿La información que brindó la oficina fue la necesaria para completar el proceso?", Browser.Buttons.YES_NO)
              }
              
              if(decRespCL === "N/A"){taFS = "No Disponible"}
              else{taFS = "Limpia"}
              
              if(decRespMetod === "N/A"){tsFS = "Procesada Sin Respuestas"}
              
              if(respCompl === ""){tsFS = "Procesada"}
              else if(respCompl === "yes"){tsFS = "Procesada con Respuestas Completas"}
              else if(respCompl === "no"){tsFS = "Procesada con Respuestas Parciales"}
              
              var valHerramienta = false
              while(valHerramienta === false){
                valHerramienta = true
                var herramienta = Browser.inputBox("Registrar Herramienta", "Digite la letra correspondiente a la herramienta de metodización usada de la siguiente lista:\\nA. EEFF Web\\nB. NACAR\\nC. Manual", Browser.Buttons.OK_CANCEL)
                
                if(herramienta === ""){return}
                if(herramienta === "cancel"){return}
                
                herramienta = herramienta.toUpperCase()
                if(herramienta === "A"){herramienta = "EEFF Web"}
                else if(herramienta === "B"){herramienta = "NACAR"}
                else if(herramienta === "C"){herramienta = "Manual"}
                else{
                  valHerramienta = false
                  Browser.msgBox("Error", "No se ingresó una letra válida de la lista de herramientas.", Browser.Buttons.OK)
                }
              }
              
              sheetWF.getRange(fEncontrada,cFS).setValue(fecha) //Escribe la fecha.
              sheetWF.getRange(fEncontrada,cAmb).setValue(taFS)
              sheetWF.getRange(fEncontrada, cTS).setValue(tsFS)
              sheetWF.getRange(fEncontrada, cHerramienta).setValue(herramienta)
              
              var arraySH = sh.getRange(finicioSH,cPetForm,lRowSH - finicioSH +1,1).getValues() //lRowWF - fInicioWF +1
              for(var iSH = 0; iSH <= lRowSH - finicioSH +1 -1; iSH++){        
                var codPetSH = arraySH[iSH][0]
                
                if(Number(CodSol) === Number(codPetSH)){
                  sh.getRange(iSH+finicioSH, cCantidadForm).setValue(0)
                  break;
                }
              }
              
              var recipient = Correo
              var subject = "SIOM: Cierre de Petición " + CodSol + " (" + Cliente + ") en el Workflow de Metodización"
              var body = "Segmento Empresas ha terminado de metodizar la operación con código de petición: " + CodSol + "."
              + ".\n"
              + "\nInformación general sobre la operación: "
              + "\nCódigo Central: " + CodCentral
              + "\nGestor: " + Ejecutivo
              + "\nCliente: " + Cliente 
              + "\nGrupo : " + Grupo 
              + "\nTipo de Estado: " + Oper
              +  "\nAnalista Asignado: " + AnAsig
              + "\n\nDetalle de Balance General: " + DBG
              + "\n\nDetalle de Estado de Resultados: " + DER
              + "\nArchivo Adjunto: " + ArchAdj
              + "\nTipo de Asignación: " + taFS
              + "\nTipo de Sanción: " + tsFS
              
              var options = {cc: correoResp + ", wendy.amaya.tobon@bbva.com, siom@bbva.com"}
              
              MailApp.sendEmail(recipient,subject,body,options);
              
              celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
              celda.setBackgroundRGB(85,199,104)
              Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
              
              fWorkflowRiesgos2()
              
              return;
            }
            else if(headingWF === "Fecha Cierre" && valNDEV === "NDEV"){
              var decRespCL = sheetWF.getRange(fEncontrada, cArchivoDevuelve).getValue()
              var decRespMetod = sheetWF.getRange(fEncontrada, cArchivoResp).getValue()
              var taFS = "" //El tipo de asignación (Limpia o No Disponible)
              var tsFS = "" //El tipo de sanción.
              
              var respCompl = ""
              if(sheetWF.getRange(fEncontrada, cArchivoResp).getValue() === "N/A"){/*No hace nada*/}
              else if(sheetWF.getRange(fEncontrada, cRcR1).getValue() != ""){
                respCompl = Browser.msgBox("Segmento Empresas", "¿La información que brindó la oficina fue la necesaria para completar el proceso?", Browser.Buttons.YES_NO)
              }
              
              if(respCompl === "cancel"){return;}
              
              if(decRespCL === "N/A"){taFS = "No Disponible"}
              else{taFS = "Limpia"}
              

              
              if(respCompl === ""){tsFS = "Procesada"}
              else if(respCompl === "yes"){tsFS = "Procesada con Respuestas Completas"}
              else if(respCompl === "no"){tsFS = "Procesada con Respuestas Parciales"}

              if(decRespMetod === "N/A"){tsFS = "Procesada Sin Respuestas"}

              var valHerramienta = false
              while(valHerramienta === false){
                valHerramienta = true
                var herramienta = Browser.inputBox("Registrar Herramienta", "Digite la letra correspondiente a la herramienta de metodización usada de la siguiente lista:\\nA. EEFF Web\\nB. NACAR\\nC. Manual", Browser.Buttons.OK_CANCEL)
                
                if(herramienta === ""){return}
                if(herramienta === "cancel"){return}
                
                herramienta = herramienta.toUpperCase()
                if(herramienta === "A"){herramienta = "EEFF Web"}
                else if(herramienta === "B"){herramienta = "NACAR"}
                else if(herramienta === "C"){herramienta = "Manual"}
                else{
                  valHerramienta = false
                  Browser.msgBox("Error", "No se ingresó una letra válida de la lista de herramientas.", Browser.Buttons.OK)
                }
              }
              
              sheetWF.getRange(fEncontrada,cFS).setValue(fecha) //Escribe la fecha.
              sheetWF.getRange(fEncontrada,cAmb).setValue(taFS)
              sheetWF.getRange(fEncontrada, cTS).setValue(tsFS)
              sheetWF.getRange(fEncontrada, cHerramienta).setValue(herramienta)
              
              var arraySH = sh.getRange(finicioSH,cPetForm,lRowSH - finicioSH +1,1).getValues() //lRowWF - fInicioWF +1
              for(var iSH = 0; iSH <= lRowSH - finicioSH +1 -1; iSH++){        
                var codPetSH = arraySH[iSH][0]
                
                if(Number(CodSol) === Number(codPetSH)){
                  sh.getRange(iSH+finicioSH, cCantidadForm).setValue(0)
                  break;
                }
              }
              
              var recipient = Correo
              var subject = "SIOM: Cierre de Petición " + CodSol + " (" + Cliente + ") en el Workflow de Metodización"
              var body = "Segmento Empresas ha terminado de metodizar la operación con código de petición: " + CodSol
              + ".\n"
              + "\nInformación general sobre la operación: "
              + "\nCódigo Central: " + CodCentral
              + "\nGestor: " + Ejecutivo
              + "\nCliente: " + Cliente 
              + "\nGrupo : " + Grupo 
              + "\nTipo de Estado: " + Oper
              +  "\nAnalista Asignado: " + AnAsig
              + "\n\nDetalle de Balance General: " + DBG
              + "\n\nDetalle de Estado de Resultados: " + DER
              + "\nArchivo Adjunto: " + ArchAdj
              + "\nTipo de Asignación: " + taFS
              + "\nTipo de Sanción: " + tsFS
              
              var options = {cc: correoResp + ", wendy.amaya.tobon@bbva.com, siom@bbva.com"}
              
              MailApp.sendEmail(recipient,subject,body,options);
              
              celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
              celda.setBackgroundRGB(85,199,104)  
              Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
              fWorkflowRiesgos2()
              return;
            }
            /*else if(headingWF === "Fin Evaluación (VB Jefe)" && valNDEV === ""){
            sheetWF.getRange(fEncontrada,c2daCRC).setValue(fecha)
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)                  
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            fWorkflowRiesgos2()
            return;
            } */               
            
            /////////Primera ronda de preguntas                
            /*if(headingWF === "Consulta de BE a Cliente" && decRi === "no"){                 
            sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)
            fWorkflowRiesgos2()
            return;
            }
            else if(headingWF === "Consulta de BE a Cliente" && decRi === "yes"){
            sheetWF.getRange(fEncontrada,cRcR1).setValue(fecha) //Escribe la fecha.
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            
            fWorkflowRiesgos2()
            
            return;
            }*/             
            
            if(headingWF === "Devolución"){
              var respuesta = Browser.msgBox("Segmento Empresas", "¿La oficina le ha contestado? Si elige no, entonces se procederá a cerrar la operación sin las respuestas.", Browser.Buttons.YES_NO)
              
              if(respuesta === "cancel"){return}
              else if(respuesta === "yes"){
                Browser.msgBox("Segmento Empresas", "Espere hasta que la oficina registre la fecha de respuestas.", Browser.Buttons.OK)
                return
              }
              
              var respuestamail = "yes"
              if (respuestamail === "yes"){
                if(Correo != "SIN CORREO"){
                  var recipient = Correo
                  }
                else if (Correo === "SIN CORREO"){
                  var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
                  }
                if (recipient === "cancel"){return}
                var subject = "SIOM: Devolución y Cierre de Petición " + CodSol + " (" + Cliente + ") en el Workflow de Metodización"
                var body = "Segmento Empresas ha cerrado la operación con código de petición " + CodSol + " sin las respuestas solicitadas."
                + ".\n"
                + "\nInformación general sobre la operación: "
                + "\nCódigo Central: " + CodCentral
                + "\nGestor: " + Ejecutivo
                + "\nCliente: " + Cliente 
                + "\nGrupo : " + Grupo 
                + "\nTipo de Estado: " + Oper
                +  "\nAnalista Asignado: " + AnAsig
                + "\n\nDetalle de Balance General: " + DBG
                + "\n\nDetalle de Estado de Resultados: " + DER
                + "\nArchivo Adjunto: " + ArchAdj
                + "\n\n\nIngrese al siguiente formulario para registrar las respuestas a esta devolución. Eligiendo la opción 'Consultas sobre Check List' en la pregunta de '¿Qué desea reingresar?'. Esto reabrirá la petición."
                + "\nhttps://docs.google.com/a/bbva.com/forms/d/e/1FAIpQLScJXlEeqfg1FqWzhvPWGxR3804XNvdDxMHR_oE4uc2tv9HrUw/viewform?usp=sf_link"     
                var options = {cc: correoResp + ", wendy.amaya.tobon@bbva.com, siom@bbva.com"}
                
                MailApp.sendEmail(recipient,subject,body,options);
              }
              
              //Acá se registra como 0 la carga la base de ingreso de formulario.
              var arraySH = sh.getRange(finicioSH,cPetForm,lRowSH - finicioSH +1,1).getValues() //lRowWF - fInicioWF +1
              for(var iSH = 0; iSH <= lRowSH - finicioSH +1 -1; iSH++){        
                var codPetSH = arraySH[iSH][0]
                
                if(Number(CodSol) === Number(codPetSH)){
                  sh.getRange(iSH+finicioSH, cCantidadForm).setValue(0)
                  break;
                }
              }


              var decRespCL = sheetWF.getRange(fEncontrada, cArchivoDevuelve).getValue()
              if(decRespCL === "N/A"){taFS = "No Disponible"}
              else{taFS = "Limpia"}
              sheetWF.getRange(fEncontrada,cFS).setValue(fecha) //Escribe la fecha.
              sheetWF.getRange(fEncontrada,cAmb).setValue(taFS)
              sheetWF.getRange(fEncontrada, cTS).setValue("Procesada Sin Respuestas")
              Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
              celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
              celda.setBackgroundRGB(85,199,104)
              fWorkflowRiesgos2()
              return;
              
            }

            /////////Primera ronda de preguntas
            
            /*if(headingWF === "Fecha Sanción"){
            var amb = Browser.inputBox("Registrar Ámbito", "Digite la letra correspondiente a un ámbito de la siguiente lista:\\nA. Jefe Grupo\\nB. Jefe Equipo\\nC. Sub Gerente\\nD. Gerente de Unidad\\nE. CTO\\nF. CEC", Browser.Buttons.OK_CANCEL);
            var tempAmb = amb
            
            if(amb === ""){return}
            if(amb === "cancel"){return}
            
            amb = amb.toUpperCase()
            if(amb === "A"){amb = "Jefe Grupo"}
            else if(amb === "B"){amb = "Jefe Equipo"}
            else if(amb === "C"){amb = "Sub Gerente"}
            else if(amb === "D"){amb = "Gerente de Unidad"}
            else if(amb === "E"){amb = "CTO"}
            else if(amb === "F"){amb = "CEC"}
            amb = amb.toUpperCase()
            
            switch(amb){
            
            case "JEFE GRUPO":
            var ambito = amb;
            break;
            case "JEFE EQUIPO":
            var ambito = amb;
            break;
            case "SUB GERENTE":
            var ambito = amb;
            break;
            case "GERENTE":
            var ambito = amb;
            break;
            case "CTO":
            var ambito = amb;
            break;
            case "CEC":
            var ambito = amb;
            break;
            case "GERENTE DE UNIDAD":
            var ambito = amb;
            break;
            default:
            Browser.msgBox("Error", "No se ingresó una letra válida de la lista de ámbitos.", Browser.Buttons.OK)
            ss.toast("Digite la letra correspondiente para el ámbito; por ejemplo, para Jefe Grupo coloque una 'A'.","Tip",8)
            return
            }
            
            var tSan = Browser.inputBox("Registrar el Tipo de Sanción", "Digite el número correspondiente a un tipo de sanción de la siguiente lista:\\n1. Aprobado Sin Modificación\\n2. Denegado\\n3. Devuelto\\n4. Aprobado Con Modificación", Browser.Buttons.OK_CANCEL);
            
            var temptSan = tSan
            
            if(tSan === ""){return}
            if(tSan === "cancel"){return}
            
            if(tSan == 1){tSan = "Aprobado SM"}
            else if(tSan == 2){tSan = "Denegado"}
            else if(tSan == 3){tSan = "Devuelto"}
            else if(tSan == 4){tSan = "Aprobado CM"}
            
            tSan = tSan.toUpperCase()
            
            switch(tSan){
            
            case "APROBADO SM":
            var tipoSan = tSan;
            var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue()
            break;
            case "APROBADO CM":
            var tipoSan = tSan;
            var tMontSan = Browser.inputBox("Registrar el Monto Sancionado", "Registre el monto sancionado en miles de US$.", Browser.Buttons.OK_CANCEL);
            if(tMontSan === "cancel"){return}
            if(tMontSan === ""){return}
            var MontSol = sheetWF.getRange(fEncontrada, cMontSol).getValue()
            if (tMontSan > MontSol){
            Browser.msgBox("Error", "El monto sancionado debe ser menor al monto solicitado.", Browser.Buttons.OK)
            return
            }                      
            var tCas = Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. Plazo\\nB. Garantía\\nC. Importe\\nD. Otros\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B. Si hace esto, respetar el orden alfabético.", Browser.Buttons.OK_CANCEL);
            if(tCas === "cancel"){return}
            tCas = tCas.toUpperCase()
            if(tCas === "A" || tCas === "B" || tCas === "C" || tCas === "D" || tCas === "A+B" || tCas === "A+C" || tCas === "A+D" || tCas === "B+C" || tCas === "B+D" || tCas === "C+D" || tCas === "A+B+C" || tCas === "A+C+D" || tCas === "A+B+C+D" || tCas === "B+C+D"){/*Pasa*//*}
            else{
            Browser.msgBox("Error", "Casuística no válida.", Browser.Buttons.OK)
            return
            }
            sheetWF.getRange(fEncontrada,cCas).setValue(tCas)
            break;
            case "DEVUELTO":
            var tipoSan = tSan;
            var tMontSan = "DEVUELTO"
            break;
            case "DENEGADO":
            var tipoSan = tSan;
            var tMontSan = "DENEGADO"
            break;
            default:
            Browser.msgBox("Error", "No se ingresó un número válido de la lista de tipos de sanción.", Browser.Buttons.OK)
            ss.toast("Digite el número correspondiente para el tipo de sanción; por ejemplo, para Aprobado Sin Modificación coloque un 1.","Tip",8)
            return
            }
            
            var respuestamail = Browser.msgBox("Correo de Sanción", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)
            if (respuestamail === "yes"){
            if(Correo != "SIN CORREO"){
            var recipient = Correo
            }
            else if (Correo === "SIN CORREO"){
            var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
            }
            if (recipient === "cancel"){return}
            var subject = "Operación Sancionada"
            
            var body = "Riesgos ha sancionado la operación con código de solicitud: " + CodSol
            + ".\n"
            + "\nInformación general sobre la operación: "
            + "\nCódigo Central: " + CodCentral
            + "\nGestor: " + Ejecutivo
            + "\nCliente: " + Cliente 
            + "\nGrupo : " + Grupo 
            + "\nOperación: " + Oper
            //+ "\nMonto Propuesto (Miles de US$): " +MontoProp
            + "\nÁmbito de Sanción: " +ambito
            + "\nTipo de Sanción: " +tipoSan
            //+ "\nMonto de Sanción (Miles de US$): " +tMontSan
            
            //var options = {bcc: correoGOF}
            
            MailApp.sendEmail(recipient,subject,body/*,options*//*);
            }
            
            if(tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track"){
            if(tipoSan === "DEVUELTO" || tipoSan === "DENEGADO"){var nuevaFV = tipoSan}
            else{
            var nuevaFV = Browser.inputBox("Fecha de Vencimiento", "Ingrese una fecha de vencimiento. Este cambio se reflejará en la base de Líneas.", Browser.Buttons.OK_CANCEL)
            if (nuevaFV === "cancel" || nuevaFV === ""){return}
            }
            }
            if(tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track"){
            if (nuevaFV === "cancel" || nuevaFV === ""){return}
            var ssL = SpreadsheetApp.openById("1Eyjey8wJlCDxzQcShgLxemZAcWA_JkE8KFlEEf7P-IM")
            var sheetPF = ssL.getSheetByName("Líneas")
            var Avals = sheetPF.getRange("A1:A").getValues()
            var lRowPF = Avals.filter(String).length
            var lColumnPF = sheetPF.getLastColumn()
            var fInicioPF = 2
            var cCodCentPF = 1
            var cGrupoPF = 2
            var cEjecPF = 3
            var cFechaOG = 4
            var cGrupoL = 6
            var cEstRiesgos = 9
            var cTraPF = 10
            var cCodSolPF = 11
            var cTipoOpPF = 12
            var cMontoPF = 14
            var cMontoSanc = 15
            var cFechaSanc = 18
            var codGE = sheetEO.getRange("C2").getValue()
            
            var valCodCentEnc = false
            
            var arrayPF = sheetPF.getRange(fInicioPF,1,lRowPF - fInicioPF+1,lColumnPF).getValues()
            for(var i = 0; i <= lRowPF - fInicioPF; i++){
            var codCentPF = arrayPF[i][cCodCentPF-1]
            var codSolPF = arrayPF[i][cCodSolPF-1]
            if(codCentPF === CodCentral  || codCentPF === codGE || CodSol === codSolPF){
            valCodCentEnc = true
            if(tipoSan === "DENEGADO" || tipoSan === "DEVUELTO"){
            nuevaFV = tipoSan
            }
            else{
            sheetPF.getRange(i+fInicioPF,cFechaSanc).setValue(nuevaFV)
            }
            sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)
            sheetPF.getRange(i+fInicioPF,cTraPF).setValue("NO")
            sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)
            
            break;
            }
            }  
            if(valCodCentEnc === false){
            Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.")     
            var recipientE = "dennis.delgado@bbva.com"
            var subjectE = "Operación no encontrada en la base de Líneas."
            var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol
            + ".\n"
            + "\nInformación general sobre la operación: "
            + "\nCódigo Central: " + CodCentral
            + "\nGestor: " + Ejecutivo
            + "\nCliente: " + Cliente 
            + "\nGrupo : " + Grupo 
            + "\nOperación: " + Oper
            //+ "\nMonto Propuesto (Miles de US$): " +MontoProp
            
            MailApp.sendEmail(recipientE,subjectE,bodyE);
            }
            }
            
            sheetWF.getRange(fEncontrada,cFS).setValue(fecha) //Escribe la fecha.
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)
            
            sheetWF.getRange(fEncontrada,cAmb).setValue(ambito) //Escribe el ámbito.
            sheetWF.getRange(fEncontrada,cTS).setValue(tipoSan) //Escribe el tipo de sanción.
            
            //sheetWF.getRange(fEncontrada,cMontSan).setValue(tMontSan)                  
            
            Browser.msgBox("Registrado","Fecha de sanción, ámbito, tipo de sanción y monto sancionado registrados con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            
            fWorkflowRiesgos2()
            
            return;
            }*/
            //}
            //}
          }
        }
        break;
      }
    }
  }
  //Solo para la estación 2
  
  //Solo para la estación 3
  /*else if(valEstEst === "I3"){
  for(var i = finicioEO; i<=lRowEO; i++){
  var celda = sheetEO.getRange(i, cDatosE3)
  var val = celda.getValue()
  if (val === strRiReg || val === strRiReg2 || val === strRiReg3){ //Encuentra la instancia en la que riesgos puede ingresar una fecha.
  
  var celdavaldev = sheetEO.getRange(rNDEV)
  var valdev = celdavaldev.getValue()        
  var celdaEO = sheetEO.getRange(i, cHeadingsE3)
  var headingEO = celdaEO.getValue() //Obtiene el heading de la hoja de estación.
  
  
  if(headingEO === "Fin Evaluación (VB Jefe) (2)"){
  
  var cval = sheetEO.getRange(finicioEO+1,cHeadingsE3)
  var valfvb = cval.getValue()
  if (valfvb === "Fin Evaluación (VB Jefe) (2)"){
  var respuesta = Browser.msgBox("Riesgos", "¿La evaluación ha concluido? Presione 'Sí' solo si la evaluación ha concluido. En caso contrario presione 'No' para registrar la fecha de las consultas enviadas al cliente.", Browser.Buttons.YES_NO)
  
  if (respuesta == "no") {
  var decRi = "no"
  }
  else if(respuesta === "cancel"){
  return
  }
  else{
  var decRi = "yes"
  }
  }
  else{
  var decRi = "no"
  }                  
  }
  else if(headingEO === "Consulta de BE a Cliente (3)"){
  
  var respuesta = Browser.msgBox("Riesgos", "¿El cliente respondió? Presione 'Sí' solo si se cuenta con las respuestas. En caso contrario presione 'No' para solicitar intervención del ejecutivo de cuenta.", Browser.Buttons.YES_NO)
  
  if (respuesta == "no") {
  var decRi = "no"
  
  var respuestamail = Browser.msgBox("Correo de Solicitud", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)
  if (respuestamail === "yes"){
  if(Correo != "SIN CORREO"){
  var recipient = Correo
  }
  else if (Correo === "SIN CORREO"){
  var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
  }
  if (recipient === "cancel"){return}
  var subject = "Solicitud de Reingreso de Respuestas"
  var body = "Se solicita su intervención para el reingreso de respuestas en el Workflow para la operación con código de solicitud: " + CodSol
  + ".\n"
  + "\nInformación general sobre la operación: "
  + "\nCódigo Central: " + CodCentral
  + "\nGestor: " + Ejecutivo
  + "\nCliente: " + Cliente 
  + "\nGrupo : " + Grupo 
  + "\nOperación: " + Oper
  + "\nMonto Propuesto (Miles de US$): " +MontoProp
  
  var options = {bcc: correoGOF}
  
  MailApp.sendEmail(recipient,subject,body,options);
  }
  
  }
  else if(respuesta === "cancel"){
  return
  }
  else{
  var decRi = "yes"
  }     
  }
  else if(headingEO === "Devolución (3)"){}
  else if(headingEO === "Reingreso con Respuestas (3)"){}
  else if(headingEO === "Fecha Sanción (2)"){}
  else if(headingEO === "Asignación Evaluación II"){
  var rpta = Browser.msgBox("Riesgos","¿Desea reabrir la operación?", Browser.Buttons.YES_NO)
  if(rpta != "yes"){return}
  } 
  else{
  var respuesta = Browser.msgBox("Riesgos", "¿Desea devolver la operación?", Browser.Buttons.YES_NO)
  if(respuesta  == "no") {
  var decRi = "no"
  }
  else if(respuesta === "cancel"){
  return 
  }
  else{
  var decRi = "yes"
  } 
  }
  
  
  for (var j = 1; j<=lColumnWF; j++){ //Recorre las columnas de la base WF.
  var celdaWF = sheetWF.getRange(fHeadingsWF,j)
  var headingWF = celdaWF.getValue() //Obtiene el heading de la base WF.
  
  if(headingWF === headingEO){ //Encuentra la instancia en la que se igualan ambos valores de los headings.
  var celdaHEOCRC = sheetEO.getRange(finicioEO+1, cHeadingsE3)
  var HEOCRC = celdaHEOCRC.getValue()
  
  var celdanDEV = sheetEO.getRange(rNDEV)
  var valNDEV = celdanDEV.getValue()
  
  
  
  if(headingWF === "Asignación Evaluación II"){
  if(tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track"){
  var ssL = SpreadsheetApp.openById("1Eyjey8wJlCDxzQcShgLxemZAcWA_JkE8KFlEEf7P-IM")
  var sheetPF = ssL.getSheetByName("Líneas")
  var Avals = sheetPF.getRange("A1:A").getValues()
  var lRowPF = Avals.filter(String).length
  var lColumnPF = sheetPF.getLastColumn()
  var fInicioPF = 2
  var cCodCentPF = 1
  var cGrupoPF = 2
  var cEjecPF = 3
  var cFechaOG = 4
  var cGrupoL = 6
  var cEstRiesgos = 9
  var cTraPF = 10
  var cCodSolPF = 11
  var cTipoOpPF = 12
  var cMontoPF = 14
  var cMontoSanc = 15
  var codGE = sheetEO.getRange("C2").getValue()
  
  var valCodCentEnc = false
  
  var arrayPF = sheetPF.getRange(fInicioPF,1,lRowPF - fInicioPF+1,lColumnPF).getValues()
  for(var i = 0; i <= lRowPF - fInicioPF; i++){
  var codCentPF = arrayPF[i][cCodCentPF-1]
  var codSolPF = arrayPF[i][cCodSolPF-1]
  var nuevaFV = "REABIERTO"
  var tMontSan = ""
  if(codCentPF === CodCentral  || codCentPF === codGE || CodSol === codSolPF){
  valCodCentEnc = true
  sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)
  sheetPF.getRange(i+fInicioPF,cTraPF).setValue("SÍ")
  sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)
  break;
  }
  }  
  if(valCodCentEnc === false){
  Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.")     
  var recipientE = "dennis.delgado@bbva.com"
  var subjectE = "Operación no encontrada en la base de Líneas."
  var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol
  + ".\n"
  + "\nInformación general sobre la operación: "
  + "\nCódigo Central: " + CodCentral
  + "\nGestor: " + Ejecutivo
  + "\nCliente: " + Cliente 
  + "\nGrupo : " + Grupo 
  + "\nOperación: " + Oper
  + "\nMonto Propuesto (Miles de US$): " +MontoProp
  
  MailApp.sendEmail(recipientE,subjectE,bodyE);
  }
  }
  
  sheetWF.getRange(fEncontrada,cAsEv).setValue(fecha)
  sheetWF.getRange(fEncontrada,cMontSan).setValue("")
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  fWorkflowRiesgos2()
  return;
  }
  if(headingWF === "Fin Evaluación (VB Jefe) (2)" && decRi === "no" && HEOCRC != "Consulta de Riesgos a Cliente (3)"){
  sheetWF.getRange(fEncontrada,c3raCRC).setValue(fecha)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  fWorkflowRiesgos2()
  return;
  }
  else if(headingWF === "Fin Evaluación (VB Jefe) (2)" && decRi === "yes" && valNDEV === ""){
  sheetWF.getRange(fEncontrada,cFev2).setValue(fecha) //Escribe la fecha.
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  
  fWorkflowRiesgos2()
  
  return;
  }
  else if(headingWF === "Fin Evaluación (VB Jefe) (2)" && valNDEV === "NDEV"){
  sheetWF.getRange(fEncontrada,cFev2).setValue(fecha)
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)  
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  fWorkflowRiesgos2()
  return;                        
  }        
  
  /////////Primera ronda de preguntas                
  if(headingWF === "Consulta de BE a Cliente (3)" && decRi === "no"){ 
  sheetWF.getRange(fEncontrada,c3raCBC).setValue(fecha)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  fWorkflowRiesgos2()
  return;
  }
  else if(headingWF === "Consulta de BE a Cliente (3)" && decRi === "yes"){
  sheetWF.getRange(fEncontrada,cRcR3).setValue(fecha) //Escribe la fecha.
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  
  fWorkflowRiesgos2()
  
  return;
  }                
  
  if(headingWF === "Devolución (3)"){
  var respuesta = Browser.msgBox("Riesgos", "¿El cliente le ha contestado a la oficina? Si elige no, entonces se procederá a devolver la operación hasta que la oficina reingrese las respuestas.", Browser.Buttons.YES_NO)
  
  if(respuesta === "cancel"){return}
  else if(respuesta === "yes"){
  Browser.msgBox("Riesgos", "Espere hasta que la oficina registre la fecha de respuestas.", Browser.Buttons.OK)
  return
  }
  
  var respuestamail = Browser.msgBox("Correo de Devolución", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)
  if (respuestamail === "yes"){
  if(Correo != "SIN CORREO"){
  var recipient = Correo
  }
  else if (Correo === "SIN CORREO"){
  var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
  }
  if (recipient === "cancel"){return}
  var subject = "Devolución de Operación en el Workflow"
  var body = "Riesgos ha devuelto la operación con código de solicitud: " + CodSol
  + ".\n"
  + "\nInformación general sobre la operación: "
  + "\nCódigo Central: " + CodCentral
  + "\nGestor: " + Ejecutivo
  + "\nCliente: " + Cliente 
  + "\nGrupo : " + Grupo 
  + "\nOperación: " + Oper
  + "\nMonto Propuesto (Miles de US$): " +MontoProp
  
  var options = {bcc: correoGOF}
  
  MailApp.sendEmail(recipient,subject,body,options);
  }             
  
  if(tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track"){
  var ssL = SpreadsheetApp.openById("1Eyjey8wJlCDxzQcShgLxemZAcWA_JkE8KFlEEf7P-IM")
  var sheetPF = ssL.getSheetByName("Líneas")
  var Avals = sheetPF.getRange("A1:A").getValues()
  var lRowPF = Avals.filter(String).length
  var lColumnPF = sheetPF.getLastColumn()
  var fInicioPF = 2
  var cCodCentPF = 1
  var cGrupoPF = 2
  var cEjecPF = 3
  var cFechaOG = 4
  var cGrupoL = 6
  var cEstRiesgos = 9
  var cTraPF = 10
  var cCodSolPF = 11
  var cTipoOpPF = 12
  var cMontoPF = 14
  var cMontoSanc = 15
  var cFechaSanc = 18
  var codGE = sheetEO.getRange("C2").getValue()
  
  var valCodCentEnc = false
  
  var arrayPF = sheetPF.getRange(fInicioPF,1,lRowPF - fInicioPF+1,lColumnPF).getValues()
  for(var i = 0; i <= lRowPF - fInicioPF; i++){
  var codCentPF = arrayPF[i][cCodCentPF-1]
  var codSolPF = arrayPF[i][cCodSolPF-1]
  var nuevaFV = "DEVUELTO"
  var tMontSan = "DEVUELTO"
  if(codCentPF === CodCentral  || codCentPF === codGE || CodSol === codSolPF){
  valCodCentEnc = true
  if(tMontSan === "DENEGADO" || tMontSan === "DEVUELTO"){nuevaFV === tMontSan}
  sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)
  sheetPF.getRange(i+fInicioPF,cTraPF).setValue("NO")
  sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)
  break;
  }
  }  
  if(valCodCentEnc === false){
  Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.")     
  var recipientE = "dennis.delgado@bbva.com"
  var subjectE = "Operación no encontrada en la base de Líneas."
  var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol
  + ".\n"
  + "\nInformación general sobre la operación: "
  + "\nCódigo Central: " + CodCentral
  + "\nGestor: " + Ejecutivo
  + "\nCliente: " + Cliente 
  + "\nGrupo : " + Grupo 
  + "\nOperación: " + Oper
  + "\nMonto Propuesto (Miles de US$): " +MontoProp
  
  MailApp.sendEmail(recipientE,subjectE,bodyE);
  }
  }
  
  
  sheetWF.getRange(fEncontrada,c3raDev).setValue(fecha)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  fWorkflowRiesgos2()
  return;
  
  }
  
  /////////Primera ronda de preguntas
  if(headingWF === "Fecha Sanción (2)"){
  var amb = Browser.inputBox("Registrar Ámbito", "Digite la letra correspondiente a un ámbito de la siguiente lista:\\nA. Jefe Grupo\\nB. Jefe Equipo\\nC. Sub Gerente\\nD. Gerente de Unidad\\nE. CTO\\nF. CEC", Browser.Buttons.OK_CANCEL);
  var tempAmb = amb
  
  if(amb === ""){return}
  if(amb === "cancel"){return}
  
  amb = amb.toUpperCase()
  if(amb === "A"){amb = "Jefe Grupo"}
  else if(amb === "B"){amb = "Jefe Equipo"}
  else if(amb === "C"){amb = "Sub Gerente"}
  else if(amb === "D"){amb = "Gerente de Unidad"}
  else if(amb === "E"){amb = "CTO"}
  else if(amb === "F"){amb = "CEC"}
  amb = amb.toUpperCase()
  
  switch(amb){
  
  case "JEFE GRUPO":
  var ambito = amb;
  break;
  case "JEFE EQUIPO":
  var ambito = amb;
  break;
  case "SUB GERENTE":
  var ambito = amb;
  break;
  case "GERENTE":
  var ambito = amb;
  break;
  case "CTO":
  var ambito = amb;
  break;
  case "CEC":
  var ambito = amb;
  break;
  case "GERENTE DE UNIDAD":
  var ambito = amb;
  break;
  default:
  Browser.msgBox("Error", "No se ingresó una letra válida de la lista de ámbitos.", Browser.Buttons.OK)
  ss.toast("Digite la letra correspondiente para el ámbito; por ejemplo, para Jefe Grupo coloque una 'A'.","Tip",8)
  return
  }
  
  var tSan = Browser.inputBox("Registrar el Tipo de Sanción", "Digite el número correspondiente a un tipo de sanción de la siguiente lista:\\n1. Aprobado Sin Modificación\\n2. Denegado\\n3. Devuelto\\n4. Aprobado Con Modificación", Browser.Buttons.OK_CANCEL);
  
  var temptSan = tSan
  
  if(tSan === ""){return}
  if(tSan === "cancel"){return}
  
  if(tSan == 1){tSan = "Aprobado SM"}
  else if(tSan == 2){tSan = "Denegado"}
  else if(tSan == 3){tSan = "Devuelto"}
  else if(tSan == 4){tSan = "Aprobado CM"}
  
  tSan = tSan.toUpperCase()
  
  switch(tSan){
  
  case "APROBADO SM":
  var tipoSan = tSan;
  var tMontSan = sheetWF.getRange(fEncontrada, cMontSol).getValue()
  break;
  case "APROBADO CM":
  var tipoSan = tSan;
  var tMontSan = Browser.inputBox("Registrar el Monto Sancionado", "Registre el monto sancionado en miles de US$.", Browser.Buttons.OK_CANCEL);
  if(tMontSan === "cancel"){return}
  if(tMontSan === ""){return}
  var MontSol = sheetWF.getRange(fEncontrada, cMontSol).getValue()
  if (tMontSan > MontSol){
  Browser.msgBox("Error", "El monto sancionado debe ser menor al monto solicitado.", Browser.Buttons.OK)
  return
  }                      
  var tCas = Browser.inputBox("Casuística", "Digite la opción de la casuística de la siguiente lista:\\nA. Plazo\\nB. Garantía\\nC. Importe\\nD. Otros\\nSe aceptan combinaciones con un símbolo de '+', por ejemplo: A+B. Si hace esto, respetar el orden alfabético.", Browser.Buttons.OK_CANCEL);
  if(tCas === "cancel"){return}
  tCas = tCas.toUpperCase()
  if(tCas === "A" || tCas === "B" || tCas === "C" || tCas === "D" || tCas === "A+B" || tCas === "A+C" || tCas === "A+D" || tCas === "B+C" || tCas === "B+D" || tCas === "C+D" || tCas === "A+B+C" || tCas === "A+C+D" || tCas === "A+B+C+D" || tCas === "B+C+D"){/*Pasa*//*}
  else{
  Browser.msgBox("Error", "Casuística no válida.", Browser.Buttons.OK)
  return
  }
  sheetWF.getRange(fEncontrada,cCas).setValue(tCas)
  break;
  case "DEVUELTO":
  var tipoSan = tSan;
  var tMontSan = "DEVUELTO"
  break;
  case "DENEGADO":
  var tipoSan = tSan;
  var tMontSan = "DENEGADO"
  break;
  default:
  Browser.msgBox("Error", "No se ingresó un número válido de la lista de tipos de sanción.", Browser.Buttons.OK)
  ss.toast("Digite el número correspondiente para el tipo de sanción; por ejemplo, para Aprobado Sin Modificación coloque un 1.","Tip",8)
  return
  }
  
  var respuestamail = Browser.msgBox("Correo de Sanción", "¿Desea enviar un correo a la oficina en este momento?", Browser.Buttons.YES_NO)
  if (respuestamail === "yes"){
  if(Correo != "SIN CORREO"){
  var recipient = Correo
  }
  else if (Correo === "SIN CORREO"){
  var recipient = Browser.inputBox("Ingrese el correo al que quiere enviar la solicitud:", Browser.Buttons.OK_CANCEL)
  }
  if (recipient === "cancel"){return}
  var subject = "Operación Sancionada"
  
  var body = "Riesgos ha sancionado la operación con código de solicitud: " + CodSol
  + ".\n"
  + "\nInformación general sobre la operación: "
  + "\nCódigo Central: " + CodCentral
  + "\nGestor: " + Ejecutivo
  + "\nCliente: " + Cliente 
  + "\nGrupo : " + Grupo 
  + "\nOperación: " + Oper
  + "\nMonto Propuesto (Miles de US$): " +MontoProp
  + "\nÁmbito de Sanción: " +ambito
  + "\nTipo de Sanción: " +tipoSan
  + "\nMonto de Sanción (Miles de US$): " +tMontSan
  
  var options = {bcc: correoGOF}
  
  MailApp.sendEmail(recipient,subject,body,options);
  }
  
  if(tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track"){
  if(tipoSan === "DEVUELTO" || tipoSan === "DENEGADO"){var nuevaFV = tipoSan}
  else{
  var nuevaFV = Browser.inputBox("Fecha de Vencimiento", "Ingrese una fecha de vencimiento. Este cambio se reflejará en la base de Líneas.", Browser.Buttons.OK_CANCEL)
  if (nuevaFV === "cancel" || nuevaFV === ""){return}
  }
  }
  if(tipoOp === "PFA" || tipoOp === "Prórroga PFA" || tipoOp === "Fast Track"){
  if (nuevaFV === "cancel" || nuevaFV === ""){return}
  var ssL = SpreadsheetApp.openById("1Eyjey8wJlCDxzQcShgLxemZAcWA_JkE8KFlEEf7P-IM")
  var sheetPF = ssL.getSheetByName("Líneas")
  var Avals = sheetPF.getRange("A1:A").getValues()
  var lRowPF = Avals.filter(String).length
  var lColumnPF = sheetPF.getLastColumn()
  var fInicioPF = 2
  var cCodCentPF = 1
  var cGrupoPF = 2
  var cEjecPF = 3
  var cFechaOG = 4
  var cGrupoL = 6
  var cEstRiesgos = 9
  var cTraPF = 10
  var cCodSolPF = 11
  var cTipoOpPF = 12
  var cMontoPF = 14
  var cMontoSanc = 15
  var cFechaSanc = 18
  var codGE = sheetEO.getRange("C2").getValue()
  
  var valCodCentEnc = false
  
  var arrayPF = sheetPF.getRange(fInicioPF,1,lRowPF - fInicioPF+1,lColumnPF).getValues()
  for(var i = 0; i <= lRowPF - fInicioPF; i++){
  var codCentPF = arrayPF[i][cCodCentPF-1]
  var codSolPF = arrayPF[i][cCodSolPF-1]
  if(codCentPF === CodCentral  || codCentPF === codGE || CodSol === codSolPF){
  valCodCentEnc = true
  if(tipoSan === "DENEGADO" || tipoSan === "DEVUELTO"){
  nuevaFV = tipoSan
  }
  else{
  sheetPF.getRange(i+fInicioPF,cFechaSanc).setValue(nuevaFV)
  }
  sheetPF.getRange(i+fInicioPF,cFechaOG).setValue(nuevaFV)
  sheetPF.getRange(i+fInicioPF,cTraPF).setValue("NO")
  sheetPF.getRange(i+fInicioPF, cMontoSanc).setValue(tMontSan)
  break;
  }
  }  
  if(valCodCentEnc === false){
  Browser.msgBox("No se encontró esta operación en la base de Líneas; sin embargo, la fecha sí se registró con éxito en el SIO. Se procedió a enviar un correo al administrador de la herramienta para revisar esto.")     
  var recipientE = "dennis.delgado@bbva.com"
  var subjectE = "Operación no encontrada en la base de Líneas."
  var bodyE = "No se encontró la siguiente operación en la base de líneas. El código de solicitud es: " + CodSol
  + ".\n"
  + "\nInformación general sobre la operación: "
  + "\nCódigo Central: " + CodCentral
  + "\nGestor: " + Ejecutivo
  + "\nCliente: " + Cliente 
  + "\nGrupo : " + Grupo 
  + "\nOperación: " + Oper
  + "\nMonto Propuesto (Miles de US$): " +MontoProp
  
  MailApp.sendEmail(recipientE,subjectE,bodyE);
  }
  }
  
  sheetWF.getRange(fEncontrada,cFS2).setValue(fecha) //Escribe la fecha.
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  
  sheetWF.getRange(fEncontrada,cAmb2).setValue(ambito) //Escribe el ámbito.
  sheetWF.getRange(fEncontrada,cTS2).setValue(tipoSan) //Escribe el tipo de sanción.
  
  
  
  sheetWF.getRange(fEncontrada,cMontSan).setValue(tMontSan)                  
  
  Browser.msgBox("Registrado","Fecha de sanción, ámbito, tipo de sanción y monto sancionado registrados con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  
  fWorkflowRiesgos2()
  
  return;
  }
  }
  }
  break;
  }
  }
  }*/
  //Solo para la estación 3
  
  Browser.msgBox("Segmento Empresas","El registro de fecha de esta operación le pertenece a la Oficina.", Browser.Buttons.OK)
  ss.toast("Cuando pueda registrar una fecha, aparecerá una celda de color amarillo.","Tip",5) 
}

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

function fModEst2() {
  //Registra cambios en fechas
  //Compara los headings de la sheetEO con los de la base WF. Ambos tienen que ser IDÉNTICOS para que funcione. 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetForm = ss.getSheetByName("Ingreso Metodización")
  var sheetWF = ss.getSheetByName("WF");
  var sheetEO = ss.getSheetByName("HY")
  
  var modEstado = sheetEO.getRange("I2").getValue()
  
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
  
  var arrayWF = sheetWF.getRange(finicioWF-1,1,lRowWF - finicioWF+2,lColumnWF).getValues() //lRowWF - fInicioWF +1
  
  //Loop que encuentra el código de solicitud
  for(var i = 1; i <= lRowWF - finicioWF +1; i++){ //Acá empieza en 1 porque se incluyeron los Headings
    var codSolWF = arrayWF[i][cCodSol-1]
    if(codSolWF === CodSol){           //Compara códigos de solicitud
      var fEncontrada = i + finicioWF -1  //Guarda el valor de la fila del código de solicitud encontrado dentro de la base WF.
    }
  }
  
  if(modEstado === ""){
    return;
  }
  else if(modEstado === "Balance Situacional"){
    sheetWF.getRange(fEncontrada,cFechaFinal).setValue(6)
    sheetWF.getRange(fEncontrada,cCodDoc).setValue(161)
    sheetWF.getRange(fEncontrada,cTipoEstadoWF).setValue("Balance Situacional")
  }
  else if(modEstado === "Declaración Jurada Anual SUNAT"){
    sheetWF.getRange(fEncontrada,cFechaFinal).setValue(18)
    sheetWF.getRange(fEncontrada,cCodDoc).setValue(217)
    sheetWF.getRange(fEncontrada,cTipoEstadoWF).setValue("Declaración Jurada Anual SUNAT")
  }
  else if(modEstado === "Estado Financiero Anual Auditado"){
    sheetWF.getRange(fEncontrada,cFechaFinal).setValue("N/A")
    sheetWF.getRange(fEncontrada,cCodDoc).setValue(160)
    sheetWF.getRange(fEncontrada,cTipoEstadoWF).setValue("Estado Financiero Anual Auditado")    
  }
  
  var decPeriodo = Browser.msgBox("Tipo de Estado Modificado", "Se ha modificado el tipo de estado. ¿Desea cambiar también el periodo?", Browser.Buttons.YES_NO)
  
  if(decPeriodo === "yes"){
    var valFecha = false
    while(valFecha === false){
      var nuevaFV = Browser.inputBox("Periodo", "Ingrese una fecha para el periodo. La fecha debe estar en formato 'dd/mm/aaaa'. Por ejemplo, 31/12/2017.", Browser.Buttons.OK_CANCEL)
      if (nuevaFV === "cancel" || nuevaFV === ""){
        Browser.msgBox("Periodo no modificado.")
        return;
      }
      var valFecha = isValidDate(nuevaFV)
      if(valFecha === false){
        Browser.msgBox("Fecha no válida. Por favor intente de nuevo.",Browser.Buttons.OK)
      }
    }
  }
  else{
    return;
  }
  
  
  sheetWF.getRange(fEncontrada,cFechaBase).setValue(nuevaFV) 
  Browser.msgBox("Periodo Modificado", "El periodo fue modificado con éxito.",Browser.Buttons.OK)
  
}