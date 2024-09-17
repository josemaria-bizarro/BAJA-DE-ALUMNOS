function doGet()
{
    const html=HtmlService.createTemplateFromFile('Caratula');
    const salida=html.evaluate();
    return salida;
}

function include(nombArch)
{
    return HtmlService.createHtmlOutputFromFile(nombArch).getContent();
}

function DataAlumnos()
{
    const dataAlumnos=bdikasleDatuak.filter(ilara=>ilara[13]=="ACTIVO")
    return dataAlumnos;
}

function cambiaStatusAlumno(alumnoSelec,inputStatus,inputSolicita,inputMotivo)
{
    //cambia el status en la BD alumnos y desactiva la cuenta institucional
    var lsr=SheetDATA.getLastRow()+1;
    let wsalida=[];
    let wfecha=Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");
   
    var wikasle=bdikasleDatuak.filter(ilara=>ilara[0]==alumnoSelec)
    if(wikasle.length>0)
    {
        inputCorreo=wikasle[0][1];
        inputOpEd=wikasle[0][9];
    }
    else
    {
        wmsg="No encontó alumno con estos datos :"+alumnoSelec
        werror.push[1,wmsg];
        return
    }

    var wserror =enviar(inputSolicita,wfecha,inputMotivo,alumnoSelec,inputCorreo,inputStatus,inputOpEd,lsr)
    if(wserror[0]==0)
    {
        wenvia=wserror[2];
        wactualiza=wserror[1];

        var wsalidaP=[];
        wsalidaP.push(inputSolicita,wfecha,"",inputMotivo,alumnoSelec,inputOpEd,inputCorreo,"","","","","ACTIVO",inputStatus,"","",wenvia,wactualiza)
        wsalida=[wsalidaP];


        try     //ESCRIBE REGISTRO EN TABLA DE MOVIMIENTOS DE ALUMNOS
        {
            
            SheetDATA.getRange(lsr,1,1,17).setValues(wsalida);
            wmsg="Cargó bien datos informativos"
            werror.push(0,wmsg,inputCorreo,inputStatus);
        }
        catch(err)
        {
            wmsg="Error al grabar datos :"+err
            werror.push(1,wmsg);
        }
    }
    else
    {
        wmsg=wserror[1]
        werror.push(1,wmsg);  
    }

    return werror;
}

function enviar(inputSolicita,wfecha,inputMotivo,alumnoSelec,inputCorreo,inputStatus,inputOpEd,lsr)
                    //(SOLICITANTE,DIA,MOTIVO,ALUMNO,OPCION,CORREO,PLANTEL,FILA,STATUS)
{
    let wserror =[];
    let statusPar="";
  //var template = SheetDB.getRange(2,1).getValue();

  var mensaje = wmensaje.replace("{SOLICITANTE}",inputSolicita)
                            .replace("{DIA}",wfecha)
                            .replace("{MOTIVO}",inputMotivo)
                            .replace("{ALUMNO}",alumnoSelec)
                            .replace("{OPCION}",inputOpEd)
                            .replace("{STATUS}",inputStatus)

    var bajak=0

    var fila=lsr
    
      switch (inputStatus)
        {
          
          default:
            break;
          case "BAJA TEMPORAL":
            try
            {
              var newstatus = {suspended: true,orgUnitPath:"/Suspendidos"};
              //AdminDirectory.Users.update(newstatus,inputCorreo);
              we=0;
              wm="ACTUALIZADO";
              bajak=1;inputCorreo
            }
            catch(err)
            {
                we=1;
                wm="ERROR BAJA TEMPORAL :"+err;
            }
            break;

          case "BAJA DEFINITIVA":
           try
            {
              var newstatus = {suspended: true,orgUnitPath:"/Suspendidos"};
              //AdminDirectory.Users.update(newstatus,inputCorreo);
              we=0;
              wm="ACTUALIZADO";
              bajak=1;
            }
            catch(err)
            {
                we=1;
                wm="ERROR BAJA DEFINITIVA :"+err;
            }
          break;

          case "BAJA ADMINISTRATIVA":
            try
            {
              var newstatus = {suspended: true,orgUnitPath:"/Suspendidos"};
              //AdminDirectory.Users.update(newstatus,inputCorreo);
              we=0;
              wm="ACTUALIZADO";
              bajak=1;
            }
            catch(err)
            {
                we=1;
                wm="ERROR BAJA ADMINISTRATIVA :"+err;
            }
          break;
       };


    if (we==0)
    {
        
        //GmailApp.sendEmail("ceec.cafetales@ceeccafetales.edu.mx,cobranza@ceeccafetales.edu.mx,servicios.escolares@ceeccafetales.edu.mx","CAMBIO DE STATUS ACADÉMICO DEL ALUMNO "+ALUMNO, mensaje,
        //                        {name: 'STATUS ACADÉMICO | CEEC',noReply: true});
        GmailApp.sendEmail("ceec.cafetales@ceeccafetales.edu.mx","CAMBIO DE STATUS ACADÉMICO DEL ALUMNO "+alumnoSelec, mensaje,
                                    {name: 'STATUS ACADÉMICO | CEEC',noReply: true});

        var ws="ENVIADO";

        //OBTIENE EL REGISTRO DEL ALUMNO EN EL CATALOGO
      let ikasleDatuak =bdikasleOrria.getDataRange().getDisplayValues();
      let ikasleFilter = ikasleDatuak.filter(ilara=> ilara[1]==inputCorreo);
      
      if (ikasleFilter.length>0)   //ENCONTRÓ AL ALUMNO EN EL CATALOGO
      {

          var wilara =ikasleFilter[0][44]   //DATO DE NUMERO DE FILA
          //MODIFICA DATOS DE ALUMNOS CON BAJA
                  //obtiene el alumno a modificar en el catalogo 
            statusPar="BAJA" 
            bdikasleOrria.getRange(wilara,13).setValue(inputStatus);//COLOCA EL STATUS EN EL CATALOGO DE ALUMNOS
            bdikasleOrria.getRange(wilara,14).setValue(statusPar);//COLOCA EL STATUS PARCIAL EN EL CATALOGO DE ALUMNOS
            bdikasleOrria.getRange(wilara,40).setValue(wfecha);//COLOCA LA FECHA DE ULTIMO MOVMTO EN EL CATALOGO DE ALUMNOS
      }
      else
      {
        ws="";
          
      }
    }
   
    wserror.push(we,wm,ws);
    return wserror;
}

function eliminaAulas(correoAlum,inputStatus)
{
  //BUSCA EN QUE AULAS ESTA EL ALUMNO
  const aulasBD=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ZnOn67CEQpaPSlYA0SKl_C_vSL5ziKALM2H-UhleuIM/')
  const aulasBDOrria=aulasBD.getSheetByName('BAJAS');
  var nptk=null;
  var optionalArgs= {pageToken:nptk, studentId:correoAlum};
  var busca= {userId:correoAlum}; //,courseId:idCurso};
  var aulaNom="";
  var aulaId="";
  var werror=0;
  let arrayBaja=[];
  
  try
  {
    var activo=Classroom.Courses.list(optionalArgs).courses //aulas donde el alumno ya entró
    var invita=Classroom.Invitations.list(busca).invitations    //obtiene el id de STUDENT
    
  }
  catch(err)
  {
    console.log(err)
    return
  }

    //PROCESA LAS AULAS ACTIVAS DEL ALUMNO
  try{
        if (activo.length>0)
        {
          var cursos=activo.filter(ilara=>ilara.courseState=="ACTIVE")
          if(cursos.length>0)
          {
            cursos.forEach(ilara=>      //busca en los cursos donde aparece el alumno
            {
              // identifica el ide del aula para dar de baja, el nombre del aula para mostrar al final
              var cursoId=ilara.id;
              var cursoIzan=ilara.name;
              try
              {
              Classroom.Courses.Students.remove(cursoId,correoAlum);

              console.log("Alumno "+correoAlum+" eliminado de aula :"+cursoIzan);
              werror=0;
              aulaNom=cursoIzan;  //nombre de curso para tabla de bajas
              aulaId=cursoId;      //id de aula para Tabla de bajas
              try
                {
                  llenaTabla(arrayBaja,bdikasleDatuak,correoAlum,aulaNom,aulaId,inputStatus,aulasBDOrria)
                }
              catch(err)
                {
                  console.log("error al cargar tabla de aulas, alumno :"+correoAlum+"  error :"+err)
                }
              }
              catch(err)
              {
                console.log("error al retirar alumno "+correoAlum+" de aula "+cursoIzan+"  error:"+err);
                werror=1;
              }

              
            })
          }
          else
          {
            aulaNom="SIN AULAS";
            aulaId="";
             llenaTabla(arrayBaja,bdikasleDatuak,correoAlum,aulaNom,aulaId,inputStatus,aulasBDOrria)
          }
        }
      }
      catch(err)
      {
        console.log("AULA SIN EL ALUMNO ACTIVO, ERROR :"+err)
      }

  //PROCESA LAS AULAS INVITADO
  try{
        if(invita .length>0)
        {
          var aulaInvita=invita.filter(ilara=>ilara.role=="STUDENT")
          if(aulaInvita.length>0)
          {
            aulaInvita.forEach(ilara=>
            {
              var cursoInvitaId=ilara.courseId;
              var invitaId=ilara.id;
              
              var cursoInvita=Classroom.Courses.get(cursoInvitaId)
              var cursoInvitaIzan=cursoInvita.name

              try
              {
                Classroom.Invitations.remove(invitaId);
                console.log("Alumno "+correoAlum+", invitación eliminada de aula :"+cursoInvitaIzan)
                werror=0;
                aulaId=cursoInvitaId;  //nombre de curso para tabla de bajas
                aulaNom=cursoInvitaIzan;   //id de aula para Tabla de bajas

                try
                  {
                    llenaTabla(arrayBaja,bdikasleDatuak,correoAlum,aulaNom,aulaId,inputStatus,aulasBDOrria)
                  }
                catch(err)
                  {
                    console.log("error al cargar tabla de aulas, alumno :"+correoAlum+"  error :"+err)
                  }
              }
              catch(err)
              {
                console.log("error al retirar invitación alumno "+correoAlum+" de aula "+cursoInvitaIzan+"  error:"+err);
                werror=1;
              }

              
            })
          }
          else
            {
              aulaNom="SIN INVITACION AULAS";
            aulaId="";
             llenaTabla(arrayBaja,bdikasleDatuak,correoAlum,aulaNom,aulaId,inputStatus,aulasBDOrria)
            }
        }
        else
        {
          console.log("alumno sin invitaciones :"+correoAlum);
          aulaNom="SIN INVITACION AULAS";
            aulaId="";
             llenaTabla(arrayBaja,bdikasleDatuak,correoAlum,aulaNom,aulaId,inputStatus,aulasBDOrria)
        }
      }
      catch(err)
      {
        console.log("AULA SIN EL ALUMNO INVITADO, ERROR :"+err)
      }

  if(werror>1)  // ALUMNO SI AULAS Y SIN INVITACIONES
  {
    werror=1;
  }
  arrayBaja.push(werror)
  return arrayBaja;
}

function llenaTabla(arrayBaja,bdikasleDatuak,correoAlum,aulaNom,aulaId,inputStatus,aulasBDOrria)
{
//escribir información en tabla de registro de bajas
var arrayPrint=[];
  if(werror==0)
  {
    var datuakIkasle=bdikasleDatuak.filter(ilara=>ilara[1]==correoAlum); //OBTIENE EL NOMBRE DEL ALUMNO
    if(datuakIkasle.length>0)
    {
      var ikasleIzan=datuakIkasle[0][0];  //NOMBRE
      var ikasleTaldea=datuakIkasle[0][5]+datuakIkasle[0][6]; //GRUPO Y SUBGRUPO
      // aulaNom      nombre de curso para tabla de bajas
      // aulaId       id de aula para Tabla de bajas
      // inputStatus  tipo de movimiento
      var wfecha=Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");

      //arma registro de salida
      arrayBaja.push([ikasleIzan,correoAlum,ikasleTaldea,aulaNom,aulaId,inputStatus,wfecha])

      // **GRABA LO QUE SE DA DE BAJA  *****
      arrayPrint.push([ikasleIzan,correoAlum,ikasleTaldea,aulaNom,aulaId,inputStatus,wfecha])
      var lsr=aulasBDOrria.getLastRow()+1;
      aulasBDOrria.getRange(lsr,1,1,7).setValues(arrayPrint);
      
      return arrayBaja;
    }
    else
    {
      console.log("error al accesar catalogo de alumnos para el alumno:"+correoAlum)
    }
  }
}

function eliminaListas(arrayBaja)
{
  //OBTIENE LA OPCION EDUCATIVA
  let wcorreo=arrayBaja[0][1];
  let dautaI=bdikasleOrria.getDataRange().getDisplayValues();
  let datuaIF=dautaI.filter(ilara=>ilara[1]==wcorreo);
  try{
    inputOpEd=datuaIF[0][9]
  }
  catch
  {
    console.log("error al obtener opc edu del alumno :"+wcorreo)
    return
  }
  //BUSCAR LISTAS DONDE SE ENCUENTRA EL ALUMNO LAS CORRESPONDIENTES A LAS AULAS
  let datuaOpED=bdOpEdu.getDataRange().getDisplayValues();
  let datuaOpEDF=datuaOpED.filter(ilara=>ilara[0]==inputOpEd);
  try{
    var idContenedor=datuaOpEDF[0][17];
  }
  catch(err)
  {
    console.log("error al obtener id Contenedor para la opc edu :"+inputOpEd+" error :"+err)
  }
  // IR UNA POR UNA BUSCANDO AL ALUMNO O PUEDE BUSCAR LAS DE LAS AULAS
  //      IDENTIFICAR OPC EDU
  //      
  let jatorri=DriveApp.getFolderById(idContenedor);
  let filesInFolder=jatorri.getFiles();
  let arrayLista=[];
  let arrayListaBaja=[];
  let fechaBaja=(new Date());
  while (filesInFolder.hasNext())
  {
    let wFile=filesInFolder.next()
    let fileType=wFile.getMimeType();
    // VERIFICA QUE USE SOLO ARCHIVOS TIPO HOJA DE CALCULO
    if(fileType==="application/vnd.google-apps.spreadsheet")
    {
      var fileW=wFile.getId();
      var fileWP="";
      var fileIzan="SIN LISTAS";
      
      var filePaso=SpreadsheetApp.openById(fileW);
      var hojaW=filePaso.getSheetByName('ORIGEN');
      
    //BUSCAR EL ALUMNO EN LA LISTA
      var datuaHoja=hojaW.getDataRange().getDisplayValues();
      var lsrow=hojaW.getLastRow();
      var lscol=hojaW.getLastColumn();
      var numRows=lsrow-10;

      var datuaHojaF=datuaHoja.filter(ilara=> ilara[2]==arrayBaja[0][0])
      //      si se encuentra el alumno:
      //          seleccionar desde la C10 hasta el final de hoja
      if (datuaHojaF.length>0)
      {
        //RESPALDAR LA LISTA DUPLICANDO LA HOJA POR SI HAY PROBLEMAS
        hojaW.copyTo(filePaso);

        fileIzan=wFile.getName();
        fileWP=wFile.getUrl();
        var rangoW=hojaW.getRange(10,3,numRows,lscol).getDisplayValues();
        var arrasyIkasle=[];

        //        RECORRE EL RANGO UNO A UNO ESCRIBIENDO EN UN ARRAY LOS ALUMNOS DIFERENTES AL SELECCIONADO
        rangoW.forEach(ilara=>
        {
          //  si posición 0 ==""descarta
          //  si la posicion 0 == nombre alumno
          //      No la considera
          if (ilara[0]!="")
          {
            if(ilara[0]!=arrayBaja[0][0])
            {
              //     asigna un numero consecutivo EN la posicion 0 push
              //     guarda toda la fila push
              //arrasyIkasle.push([windex,ilara])
              arrasyIkasle.push(ilara)
            }
          }
           
        })
         //            BORRA INFORMACIÓN Y REESCRIBE SIN EL ALUMNO DADO DE BAJA
            hojaW.getRange(10,3,numRows,lscol).clearContent()            
            hojaW.getRange(10,3,arrasyIkasle.length,arrasyIkasle[0].length).setValues(arrasyIkasle);
            
        //REGISTRO DE LISTAS PARA MOSTRAR EN CARATULA
            arrayLista.push(fileIzan);
      }
      else
      {
        console.log("LISTA "+fileIzan+" SIN ALUMNO :"+arrayBaja[0][0])
      }
    }
  }

  //REGISTRAR EN LOG DE BAJAS
  const bajaListas=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1mGd1qkCBfErQV-Z8eJqIGhQ8002XuP6SpEgqrbaMtl8/')
  const bajaListasOrria=bajaListas.getSheetByName('BAJAS');

  var lsr=bajaListasOrria.getLastRow()+1;
  arrayListaBaja.push(arrayBaja[0][0],fileIzan,fechaBaja,fileWP)
  try{
    var arraySale=[arrayListaBaja]
       bajaListasOrria.getRange(lsr,1,1,4).setValues(arraySale);
  }
  catch(err)
  {
      console.log("error al grabar en tabla BAJA LISTAS error :"+err+"  "+arrayListaBaja)
  }
     

  return arrayLista;



  
}
