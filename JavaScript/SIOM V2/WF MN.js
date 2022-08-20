function RegistrarFechaRiesgos() {
  //Registra cambios en fechas
  //Compara los headings de la sheetEO con los de la base WF. Ambos tienen que ser IDÉNTICOS para que funcione. 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetForm = ss.getSheetByName("Ingreso Metodización")
  var sheetWF = ss.getSheetByName("WF");
  var sheetEO = ss.getSheetByName("MN")
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
                var subject = "SIOM: Devolución de Petición " + CodSol + " ("+ Cliente + ") en el Workflow de Metodización"
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
              fWorkflowRiesgos()
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
              
              fWorkflowRiesgos()
              
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
              fWorkflowRiesgos()
              return;
            }
            /*else if(headingWF === "Fin Evaluación (VB Jefe)" && decRi === "no" && HEOCRC === "Consulta de Riesgos a Cliente" && valNDEV === ""){
            sheetWF.getRange(fEncontrada,c2daCRC).setValue(fecha)
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            fWorkflowRiesgos()
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
              
              fWorkflowRiesgos()
              
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
              fWorkflowRiesgos()
              return;
            }
            /*else if(headingWF === "Fin Evaluación (VB Jefe)" && valNDEV === ""){
            sheetWF.getRange(fEncontrada,c2daCRC).setValue(fecha)
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)                  
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            fWorkflowRiesgos()
            return;
            } */               
            
            /////////Primera ronda de preguntas                
            /*if(headingWF === "Consulta de BE a Cliente" && decRi === "no"){                 
            sheetWF.getRange(fEncontrada,cCBC).setValue(fecha)
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)
            fWorkflowRiesgos()
            return;
            }
            else if(headingWF === "Consulta de BE a Cliente" && decRi === "yes"){
            sheetWF.getRange(fEncontrada,cRcR1).setValue(fecha) //Escribe la fecha.
            celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
            celda.setBackgroundRGB(85,199,104)
            Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
            
            fWorkflowRiesgos()
            
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
              fWorkflowRiesgos()
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
            
            fWorkflowRiesgos()
            
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
  fWorkflowRiesgos()
  return;
  }
  if(headingWF === "Fin Evaluación (VB Jefe) (2)" && decRi === "no" && HEOCRC != "Consulta de Riesgos a Cliente (3)"){
  sheetWF.getRange(fEncontrada,c3raCRC).setValue(fecha)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  fWorkflowRiesgos()
  return;
  }
  else if(headingWF === "Fin Evaluación (VB Jefe) (2)" && decRi === "yes" && valNDEV === ""){
  sheetWF.getRange(fEncontrada,cFev2).setValue(fecha) //Escribe la fecha.
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  
  fWorkflowRiesgos()
  
  return;
  }
  else if(headingWF === "Fin Evaluación (VB Jefe) (2)" && valNDEV === "NDEV"){
  sheetWF.getRange(fEncontrada,cFev2).setValue(fecha)
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)  
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  fWorkflowRiesgos()
  return;                        
  }        
  
  /////////Primera ronda de preguntas                
  if(headingWF === "Consulta de BE a Cliente (3)" && decRi === "no"){ 
  sheetWF.getRange(fEncontrada,c3raCBC).setValue(fecha)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  fWorkflowRiesgos()
  return;
  }
  else if(headingWF === "Consulta de BE a Cliente (3)" && decRi === "yes"){
  sheetWF.getRange(fEncontrada,cRcR3).setValue(fecha) //Escribe la fecha.
  celda.setValue(fecha) //Escribe la fecha ingresada en la hoja de estaciones.
  celda.setBackgroundRGB(85,199,104)
  Browser.msgBox("Registrado","Fecha registrada con éxito. Revisando nuevamente...", Browser.Buttons.OK)
  
  fWorkflowRiesgos()
  
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
  fWorkflowRiesgos()
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
  
  fWorkflowRiesgos()
  
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