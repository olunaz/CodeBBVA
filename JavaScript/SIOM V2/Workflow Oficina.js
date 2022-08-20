function fWorkflow() {
    //Se necesitan dos de estas macros para asegurar que se esté trabajando en la hoja correcta.
    //Revisa la base de datos en base a un código de solicitud.  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetWF = ss.getSheetByName("WF");
    var sheetEO = ss.getSheetByName("E1") //ESTO ES LO ÚNICO QUE SE TIENE QUE CAMBIAR PARA QUE FUNCIONE EN LA HOJA DE RIESGOS
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
                  if (nombrehoja === "ER1"){
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