const bdikasleId='1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY';  //CATALOGO ALUMNOS
const bdikasle=SpreadsheetApp.openById(bdikasleId);
const bdikasleOrria=bdikasle.getSheetByName('ACTIVOS_FORMATEADO');
const bdikasleDatuak=bdikasleOrria.getDataRange().getDisplayValues();
                                                                //CATALOGO DE MOVIMIENTOS DE STATUS
const ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1btm8_Qxyqq8mTD-M8lzWsCoUVm_q5wHSSBVtWxFqGzU/")
const aulasBD=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ZnOn67CEQpaPSlYA0SKl_C_vSL5ziKALM2H-UhleuIM/')
const aulasBDOrria=aulasBD.getSheetByName('BAJAS');
let SheetDATA =ss.getSheetByName("DATA");
let SheetDB=ss.getSheetByName('DB');
const bdOpEdu=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1T5KQCvXiYsGewSoc_A8qugUwSs-ft1NNds4oaNgQxkw/').getSheetByName("OPCIONES EDUCATIVAS")
let werror=[];
let wmsg="";
let statusAlumno="";
let wenvia="";
let wactualiza="";
let inputOpEd="vacio";
const wmensaje="A los encargados de las áreas se informa que el alumno {ALUMNO} perteneciente a la opción educativa de {OPCION} ha cambiado su estatus academico a: {STATUS}, a partir del día: {DIA}.    A solicitud de: {SOLICITANTE}, por el siguiente motivo: {MOTIVO}>>FAVOR DE REALIZAR LAS ACCIONES CORRESPONDIENTES EN CADA ÁREA<<."