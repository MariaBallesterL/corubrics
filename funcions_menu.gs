function enviaResultsC() {
  enviaResults(0);
};
function enviaResultsA() {
  enviaResults(1);
};
function enviaResultsP() {
  enviaResults(2);
};

/**
 * Canvia el formulari enllaçat amb CoRubrics
 * 
 */
function nouformulari(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";   
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nouformulari = Browser.inputBox('Nou formulari','Escriu la URL del nou formulari amb que vols enllaçar CoRubrics; ha de ser de la forma https://docs.google.com/a/domini.cat/forms/d/1beCTQ2QfB5weyZa6XK5Yay4wQ7kRqScwFsll4M/edit', Browser.Buttons.OK_CANCEL);
      break;
    case "es":
      var nouformulari = Browser.inputBox('Nuevo formulario','Escribe la URL del nuevo formulario con que quieres enlazar CoRubrics; tiene que ser de la forma https://docs.google.com/a/dominio.es/forms/d/1beCTFYy8HQ2QfB5we6XK5Yay4wQ7kRqScwFsll4M/edit', Browser.Buttons.OK_CANCEL);
      break;
    case "eu":
      var nouformulari = Browser.inputBox('Inprimaki berria','CoRubricsera lotu nahi zenukeen inprimakiaren URLa idatz ezazu; formatu honetako izan behar du https://docs.google.com/a/dominio.cat/forms/d/1beCTFYy8HQ2QfB5weyZa6wQ7kRqScwFsll4M/edit', Browser.Buttons.OK_CANCEL);
      break;
    case "fr":
      var nouformulari = Browser.inputBox('Nouveau formulaire','URL du nouveau formulaire (p.ex. https://docs.google.com/a/domaini.edu/forms/d/1beCTFYy8HQ2QfB5weyZa6XK5Yay4wQ7kRqScwFsll4M/edit', Browser.Buttons.YES_NO);
      break;
    default: 
      var nouformulari = Browser.inputBox('New form','URL from the new form linked (ex: https://docs.google.com/a/domaini.edu/forms/d/1beCTFYy8HQ2QfB5weyZa6XK5Yay4wQ7kRqScwFsll4M/edit', Browser.Buttons.YES_NO);
  };
  
  var form = FormApp.openByUrl(nouformulari);
  var idform = form.getId();
  var formurlbona = nouformulari.slice(0,-4)+"viewform"; 
  esborradB();
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Formid', idform);
  documentProperties.setProperty('Formurl', formurlbona);
  documentProperties.setProperty('Formnom', "Form");
  documentProperties.setProperty('Formulari', "1");
}


/**
 * Mostra una barra lateral per triar què cal enviar als alumnes
 *
 */

function enviament() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var html = HtmlService.createHtmlOutputFromFile('index_cat')
      .setTitle('Opcions d\'enviament')
      .setWidth(400);
      break;
    case "es":
      var html = HtmlService.createHtmlOutputFromFile('index_es')
      .setTitle('Opcions de envío')
      .setWidth(400);
      break;
    case "eu":
      var html = HtmlService.createHtmlOutputFromFile('index_eu')
      .setTitle('Bidalketa auker')
      .setWidth(400);
      break;
    case "fr":
      var html = HtmlService.createHtmlOutputFromFile('index_fr')
      .setTitle('Options d\'envoi')
      .setWidth(400);
      break;          
    default:
      var html = HtmlService.createHtmlOutputFromFile('index_en')
      .setTitle('Shipping Options')
      .setWidth(400);
  } 

  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

/**
 * Recull les respostes de la barra lateral
 * Envia els resultats de cada alumne
 * per mail
 */
function envianotes(formObject) {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  /*Recollim respostes de la barra*/
  
  var nf = formObject.nf;
  if (nf==="0"){
    var notafinal="no";
  }else{
    var notafinal="yes";
    if (nf==="1"){
      var tipusnotafinal="yes";
    }else{
      var tipusnotafinal="no";
    };
  };
  var ng = formObject.ng;
  if (ng==="0"){
    var notaglobal="no";
  }else{
    var notaglobal="yes";
  };
  var av = formObject.av;
  if (av==="on") {
    av = true;
  }else{
    av = false;
  };
  var co = formObject.co;
  if (co==="on") {
    co = true;
  }else{
    co = false;
  };
  var pf = formObject.pf;
  if (pf==="on") {
    pf = true;
  }else{
    pf = false
  };
  var cco = formObject.cco;
  if (cco==="on") {
    var coment_alu="yes";
  }else{
   var coment_alu="no";
  };
  var cpf = formObject.cpf;
  if (cpf==="on") {
    var coment_prof="yes";
  }else{
    var coment_prof="no";
  };
  var cav = formObject.cav;
  if (cav==="on") {
    var coment_auto="yes";
  }else{
    var coment_auto="no";
  };
  var dest = formObject.dest;
  if (dest==="all") {
    var dest=true;
  }else{
    var dest=false;
  };
  if (dest===false){
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var nalum = Browser.inputBox('Alumne','A quin alumne/grup vols enviar els resultats? Indica el número de la columna A del full de resultats.', Browser.Buttons.OK);
        break;
      case "es":
        var nalum = Browser.inputBox('Alumno','¿A que alumno/grupo quieres mandar los resultados? Indica el número de la columna A de la hoja de resultados.', Browser.Buttons.OK);
        break;
      case "eu":
        var nalum = Browser.inputBox('Ikasle','Zein ikasle ala talderi bidali nahi zenioke? Jasotzeaileak emaitzen A zutabean duen zenbakia adieraz ezazu.', Browser.Buttons.OK);
        break;
      case "fr":
        var nalum = Browser.inputBox('Élève','Pour quel élève/groupe désirez-vous envoyer les résultats?  Indiquez le numéro de la colonne A de l\'élève/du groupe dans la feuille des résultats.', Browser.Buttons.OK);
        break;
      default:
        var nalum = Browser.inputBox('Student','Which student/group do you want to send the results to? Indicate the number in column A in the result sheet.', Browser.Buttons.OK);
    }  
  };
  /*Tanquem la barra*/
  
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var html = HtmlService.createHtmlOutputFromFile('enviat_ca')
      .setTitle('Opcions d\'enviament')
      .setWidth(400);
      break;
    case "es":
      var html = HtmlService.createHtmlOutputFromFile('enviat_es')
      .setTitle('Opcions de envío')
      .setWidth(400);
      break;
    case "eu":
      var html = HtmlService.createHtmlOutputFromFile('enviat_eu')
      .setTitle('Bidalketa auker')
      .setWidth(400);
      break;
    case "fr":
      var html = HtmlService.createHtmlOutputFromFile('enviat_fr')
      .setTitle('Options d\'envoi')
      .setWidth(400);
      break; 
    default:
      var html = HtmlService.createHtmlOutputFromFile('enviat_en')
      .setTitle('Shipping Options')
      .setWidth(400);
  } 
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
  
  
  /*Eviem les notes*/
  
  var fullmitjanes = llibreActual.getSheets()[llibreActual.getNumSheets()-1]; //S'agafen els resultat del darrer full
  var rangmitjanes = fullmitjanes.getDataRange();
  var valormitjanes = rangmitjanes.getValues();
  
 
  //Omplo la matriu rúbrica amb les dades de la rúbrica
  var mat_rubrica=rangrubrica.getValues();
 
  //Preparem els destinataris i el cos del missatge
  var documentProperties = PropertiesService.getDocumentProperties();
  var formnom = documentProperties.getProperty('Formnom');
  var cosmissatge="";
  var alumnes = "";
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var titol= "Resultat de la rúbrica: " + formnom;
      var pc = "Puntuació dels companys";
      var pp = "Puntuació pròpia";
      var ppr = "Puntuació del professor";
      var ng = "Nota final"
      var nfg= "Nota global";
      var cpf ="Comentaris del professor: ";
      var cco ="Comentaris dels compays: ";
      var cav ="Comentaris del propi alumne: ";
      break;
    case "es":
      var titol= "Resultado de la rúbrica: " + formnom;
      var pc = "Puntuación de los compañeros";
      var pp = "Puntuación propia";
      var ppr = "Puntuación del profesor";
      var ng = "Nota final"
      var nfg= "Nota global";
      var cpf ="Comentarios del profesor: ";
      var cco ="Comentarios de los compañeros: ";
      var cav ="Comentarios del propio alumno: ";
      break;
    case "eu":
      var titol= "Errubrikaren emaitza: " + formnom;
      var pc = "Kideen kalifikazioa";
      var pp = "Norberaren kalifikazioa";
      var ppr = "Irakaslearen kalifikazioa";
      var ng = "Nota finala"
      var nfg= "Nota globala";
      var cpf ="Irakaslearen iruzkina: ";
      var cco ="Ikasleen iruzkinak: ";
      var cav ="Comentaris del propi alumne: ";
      break;
    case "fr":
      var titol= "Résultat de la grille: " + formnom;
      var pc = "Note d'un collègue";
      var pp = "Note de l'élève-même";
      var ppr = "Note de l'enseignant";
      var ng = "Note finale";
      var nfg= "Note globale";
      var cpf ="Commentaires de l'enseignant: ";
      var cco ="Commentaires des collègues: ";
      var cav ="Commentaires de l'élève-même: ";
      break;
    default:
      var titol= "Rubric result: " + formnom;
      var pc = "Colleagues grade";
      var pp = "Student own grade";
      var ppr = "Teacher grade";
      var ng = "Final grade";
      var nfg= "Overall grade";
      var cpf ="Teacher's comments: ";
      var cco ="Comments from colleagues: ";
      var cav ="Student's own comments: ";
  } 

  var r_al=rangalumnes.getValues();//agafem tots els alumnes
  for (var i=0; i<rangalumnes.getNumRows()-1;i++){
    for (var j=0; j<rangalumnes.getNumColumns()-1;j++){   
      if (dest===false){
        i=nalum-1;
      };
      alumnes = r_al[i+1][j+1];
      if (alumnes!=""){
      
        //Definim el cos del missatge amb la rúbrica original i els resultats
        cosmissatge='<table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000"><tr align="center" valign="middle" bgcolor="#C0C0C0">';
        for (k=1;k<rangrubrica.getNumColumns();k++){
          cosmissatge= cosmissatge+'<td  colspan="1"><p align="center"><strong>'+mat_rubrica[0][k-1]+'</strong></p></td>';
        };
        if (co){
          cosmissatge= cosmissatge+'<td  colspan="1"><p align="center"><strong>'+pc+'</strong></p></td>';
        };
        if (av){
          cosmissatge= cosmissatge+'<td  colspan="1"><p align="center"><strong>'+pp+'</strong></p></td>';
        };
        if (pf){
          cosmissatge= cosmissatge+'<td  colspan="1"><p align="center"><strong>'+ppr+'</strong></p></td>';
        };
        
        cosmissatge = cosmissatge+'</tr><tr>';
        cosmissatge=cosmissatge+'<td><p align="center"><strong><div align="center"></strong></p></div></td>';
        for (k=2;k<rangrubrica.getNumColumns();k++){
          cosmissatge=cosmissatge+'<td><p align="center"><strong><div align="center">'+mat_rubrica[1][k-1]+'</strong></p></div></td>';
        };
        if (co){
          cosmissatge=cosmissatge+'<td><p align="center"><strong><div align="center"></strong></p></div></td>';
        };
        if (av){
          cosmissatge=cosmissatge+'<td><p align="center"><strong><div align="center"></strong></p></div></td>';
        };
        if (pf){
          cosmissatge=cosmissatge+'<td><p align="center"><strong><div align="center"></strong></p></div></td>';
        };        
        cosmissatge = cosmissatge+'</tr><tr>';
        for (z=2;z<rangrubrica.getNumRows();z++){
          cosmissatge=cosmissatge+'<td><p align="center"><strong><div align="center">'+mat_rubrica[z][0]+'</strong></p></div></td>';
          for (k=2;k<rangrubrica.getNumColumns();k++){
            cosmissatge=cosmissatge+'<td><p align="center"><div align="center">'+mat_rubrica[z][k-1]+'</p></div></td>';
          };
          if (co){
            cosmissatge=cosmissatge+'<td bgcolor="#66FF99"><p align="center"><strong><div align="center">'+valormitjanes[i+3][3*z-1]+'</strong></p></div></td>';
          };
          if (av){
            cosmissatge=cosmissatge+'<td bgcolor="#66FF99"><p align="center"><strong><div align="center">'+valormitjanes[i+3][3*z]+'</strong></p></div></td>';
          };
          if (pf){
            cosmissatge=cosmissatge+'<td bgcolor="#66FF99"><p align="center"><strong><div align="center">'+valormitjanes[i+3][3*z+1]+'</strong></p></div></td>';
          };
          cosmissatge = cosmissatge+'</tr><tr>';
        };
        cosmissatge=cosmissatge + '</tr></table><p></p>';
        
        //Si hem d'enviar la nota final, l'afegim
        if (notafinal==="yes") {
          cosmissatge= cosmissatge + '<table border="1" cellpadding="0" cellspacing="0"><tr align="center" valign="middle">';
          cosmissatge= cosmissatge + '<td  bgcolor="#C0C0C0" width="100" colspan="1"><p align="center"><strong>'+ng+'</strong></p></td>';
          if (tipusnotafinal==="yes"){
            if (co){
              cosmissatge= cosmissatge + '<td  bgcolor="#66FF99" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-9]+'</strong></p></td>';
            };
            if (av){
              cosmissatge= cosmissatge + '<td  bgcolor="#66FF99" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-8]+'</strong></p></td>';
            };
            if (pf){
              cosmissatge= cosmissatge + '<td  bgcolor="#66FF99" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-7]+'</strong></p></td>';
            };            
          }else{
            if (co){
              cosmissatge= cosmissatge + '<td  bgcolor="#66FF99" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-12]+'</strong></p></td>';
            };
            if (av){
              cosmissatge= cosmissatge + '<td  bgcolor="#66FF99" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-11]+'</strong></p></td>';
            };
            if (pf){
              cosmissatge= cosmissatge + '<td  bgcolor="#66FF99" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-10]+'</strong></p></td>';
            };            
          };
          cosmissatge= cosmissatge + '</tr></table><p></p>';
        };
          
        //Si hem d'enviar la nota global, l'afegim
        if (notaglobal==="yes") {
          cosmissatge= cosmissatge + '<table border="1" cellpadding="0" cellspacing="0"><tr align="center" valign="middle">';
          cosmissatge= cosmissatge + '<td  bgcolor="#C0C0C0" width="100" colspan="1"><p align="center"><strong>'+nfg+'</strong></p></td>';
          cosmissatge= cosmissatge + '<td  bgcolor="#7ab6ff" width="100" colspan="1"><p align="center"><strong>'+valormitjanes[i+3][rangmitjanes.getNumColumns()-6]+'</strong></p></td>';
          cosmissatge= cosmissatge + '</tr></table><p></p>';
        };            
            
        //Afegim els comentaris del professor
        if (coment_prof==="yes"){
          cosmissatge= cosmissatge + '<p><b>'+cpf+'</b>'+rangmitjanes.getCell(i+4,rangmitjanes.getNumColumns()-2).getValue()+'</p>';
        };

        //Afegim els comentaris dels alumnes
        if (coment_alu==="yes"){
          cosmissatge= cosmissatge + '<p><b>'+cco+'</b>'+rangmitjanes.getCell(i+4,rangmitjanes.getNumColumns()-1).getValue()+'</p>';
        };

        //Afegim els comentaris propis
        if (coment_auto==="yes"){
          cosmissatge= cosmissatge + '<p><b>'+cav+'</b>'+rangmitjanes.getCell(i+4,rangmitjanes.getNumColumns()).getValue()+'</p>';
        };
        
        //Enviem els missatges 
          GmailApp.sendEmail(alumnes, titol, '', {
            htmlBody: cosmissatge
          });
          alumnes="";
          cosmissatge="";

      };
    };
    if (dest===false){
      i=rangalumnes.getNumRows();
    };
    
  };
  
  //Deso al ScripDb que l'he enviat (si no estic reenviant)
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Mail2', "1");
  
  
  //Canviar el menú, treient Enviar formulari i posant el que correspongui
  onOpen();  
};



/**
 * Reiniciar el procés. Esborra la base de
 * dades i torna a crear el menú.
 */
function reinici(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();  
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
   case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  esborradB();
  
  //Canviar el menú
  onOpen();  
};

/**
 * Tornar a crear el formualri
*/
function creanouFormulari(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
     case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;   
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  esborradB();
  creaFormulari();
};

/**
 * Mostra l'enllaç del formulari
 * per pantalla
 */
function enllaFormulari() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();  
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  //Recupero el ID del fomrulari, del ScriptDB
  var documentProperties = PropertiesService.getDocumentProperties();
  var formurl= documentProperties.getProperty('Formurl');
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var enl = '<p>L\'enllaç del formulari és: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    case "es":
      var enl = '<p>El enlace del formulario és: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    case "eu":
      var enl = '<p>Hau da Inprimakiaren helbidea: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    case "fr":
      var enl = '<p>Le lien au formulaire est:: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    default:
      var enl = '<p>The link to the form is: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
  }  
  
  var htmlApp = HtmlService.createHtmlOutput();
  htmlApp.setContent(enl);
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      htmlApp.setTitle('Enllaç del formulari');
      break;
    case "es":
      htmlApp.setTitle('Enlace del formulario');
      break;
    case "eu":
        htmlApp.setTitle('Inprimakiaren helbidea');
      break;
    case "fr":
        htmlApp.setTitle('Lien au formulaire');
      break;
    default:
      htmlApp.setTitle('Form link');
  } 
  htmlApp.setWidth(400);
  htmlApp.setHeight(150);

  SpreadsheetApp.getActive().show(htmlApp);
  
  //Deso al ScripDb que l'he enviat (si no estic reenviant)
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Mail', "1");
    
  
  //Canviar el menú, treient Enviar formulari i posant el que correspongui

  onOpen();  
};


/**
 * Envia l'enllaç del formulari per mail
 * a tots les alumnes.
 */
function enviaFormulari() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  //Recupero el ID del formulari, del ScriptDB
  var documentProperties = PropertiesService.getDocumentProperties();
  var formurl = documentProperties.getProperty('Formurl');
  var formnom = documentProperties.getProperty('Formnom');

  
  //Defineixo el títol (nom del formulari) i el cos del missatge
  var titolform = formnom;
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
        var titol = "Formulari d'avaluació: " + titolform;
        var cosmissatge = "Aquí teniu l'enllaç per puntuar la tasca: " + formurl;
      break;
    case "es":
        var titol = "Formulario de evaluación: " + titolform;
        var cosmissatge = "Aquí teneis el enlace para puntuar la actividad: " + formurl;
      break;
    case "eu":
        var titol = "Ebaluazio  Inprimakia: " + titolform;
        var cosmissatge = "Zuen kideen jarduerak puntuatzeko helbidea hau da:" + formurl;
      break;
    case "fr":
        var titol = "Formulaire d'évaluation: " + titolform;
        var cosmissatge = "Voici le lien pour évaluer l'activité: " + formurl;
      break;
    default:
        var titol = "Evaluation form: " + titolform;
        var cosmissatge = "Here it is the link to rate the activity: " + formurl;
  } 
  
  
  //Envio el formulari a cada un dels alumnes
  var alumnes = "";
  for (i=0; i<rangalumnes.getNumRows()-1;i++){
    for (j=0; j<rangalumnes.getNumColumns()-1;j++){   
      alumnes = rangalumnes.getCell(i+2,j+2).getValue();
      if (alumnes!=""){
        GmailApp.sendEmail(alumnes, titol, cosmissatge);
        alumnes="";
      };
    };
  };
  
  //Deso al ScripDb que l'he enviat (si no estic reenviant)
  documentProperties.setProperty('Mail',"1");
  
  //Canviar el menú, treient Enviar formulari i posant el que correspongui;
  onOpen();  
  
};

/**
 * Crea un formulari a partir de la rúbrica del full actual, amb un pregunta Llista amb el nom
 * de tots els alumnes que hi ha al full Alumnes i amb tantes preguntes Graella com aspectes
 * valora la rúbrica.
 */
function creaFormulari() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
        var nomform = Browser.inputBox('Nom del formulari','Quin nom vols que tingui el formulari?', Browser.Buttons.OK_CANCEL);
      break;
    case "es":
        var nomform = Browser.inputBox('Nombre del formulario','¿Qué nombre quieres que tenga el formulario?', Browser.Buttons.OK_CANCEL);
      break;
    case "eu":
        var nomform = Browser.inputBox('Inprimakiaren izena','Zer izen eman nahi diozu  Inprimakiari?', Browser.Buttons.OK_CANCEL);
      break;
    case "fr":
        var nomform = Browser.inputBox('Nom du formulaire','Quel est le nom du formulaire?', Browser.Buttons.OK_CANCEL);
      break;
    default:
        var nomform = Browser.inputBox('Form name','What is the name of the form?', Browser.Buttons.OK_CANCEL);
  } 
  if (nomform!='cancel'){
    var form = FormApp.create(nomform); //Crea el formulari
    var formid = form.getId();
    var formurl = form.getPublishedUrl();
    var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
    //Movem el formualri creat a la carpeta del full de CoRubrics
    var folders = DriveApp.getFileById(llibreActual.getId()).getParents(); //Agafem les carpetes on està el full de CoRubrics
    var folder = folders.next();  //Agafem la primer carpeta.  
    if (folder.getName()!=DriveApp.getRootFolder().getName()){  //Si es troba a La Meva Unitat, no movem el formulari
      var fitxer = DriveApp.getFileById(formid);
      folder.addFile(fitxer);
      DriveApp.removeFile(fitxer);    
    };
    //Deso l'ID en un ScriptdB   
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('Formid',formid);
    documentProperties.setProperty('Formurl',formurl);
    documentProperties.setProperty('Formnom',nomform);
    documentProperties.setProperty('Formulari',"1");
 
    //Afegim descripció i que no sigui anònim
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        form.setDescription("Aquest formulari servirà per avaluar la feina dels teus companys. Primer selecciona a quin company avalues i després, per cada aspecte, tria la descripció que més s'acosti a la seva tasca");
        break;
      case "es":
        form.setDescription("Este formulario servirá para evaluar las actividades. Primero selecciona a qué compañero evaluas y después, para cada aspecto, elige la descripción que más coincida con su actividad");
        break;
      case "eu":
        form.setDescription("Jarduerak ebaluatzeko inprimakia da hau. Lehenik ebaluatuko duzun kidea aukeratu eta, jarraian, burutu den jarduerarekin gehien hurbiltzen den adierazpena hauta ezazu");
        break;
      case "fr":
        form.setDescription("Ce formulaire est utilisé pour évaluer l'activité.  Premièrement, choisissez l'élève à évaluer.  Ensuite, choisissez la meilleure description pour chaque aspect.");
        break;
      default:
        form.setDescription("This form is used to evaluate the activity. First, choose the student to rate. Then, choose the best description in each aspect.");
    }
    
    try { 
      form.setCollectEmail(true); //Només si es GAFE
    
      // Afegir llista alumnes per seleccionar 
      var preguntaList = form.addListItem();
      var alumnes = [];
      var nombrealumnes=0;
      if (rangalumnes.getNumRows()-1===0){
        var properties = PropertiesService.getDocumentProperties();   
        var idioma = properties.getProperty('Idioma');   
        switch(idioma){
          case "ca":
            Browser.msgBox('Alumnes','No has indicat cap alumne per avaluar!', Browser.Buttons.OK);
            break;
          case "es":
            Browser.msgBox('Alumnos','¡No has indicado ningún alumno a evaluar!', Browser.Buttons.OK);
            break;
          case "eu":
            Browser.msgBox('Ikasleak','Ebaluatzeko ikaslerik ez duzu adierazi!', Browser.Buttons.OK);
            break;
          case "fr":
            Browser.msgBox('Élèves','La liste d\'élèves est vide', Browser.Buttons.OK);
            break;
          default:
            Browser.msgBox('Students','The list of students is empty.', Browser.Buttons.OK);
        }
        
        nombrealumnes=1;
      }
      for (var i=0; i<rangalumnes.getNumRows()-1;i++){
        alumnes[i] = rangalumnes.getCell(i+2,1).getValue();
        alumnes[i]=alumnes[i].toString();
        var fora_espais_finals = alumnes[i].trim();
        if (alumnes[i] != fora_espais_finals){
          rangalumnes.getCell(i+2,1).setValue(fora_espais_finals);
        };
        alumnes[i]=fora_espais_finals;
      };
      var properties = PropertiesService.getDocumentProperties();   
      var idioma = properties.getProperty('Idioma');   
      switch(idioma){
        case "ca":
          preguntaList.setTitle("Alumne a avaluar");
          break;
        case "es":
          preguntaList.setTitle("Alumno a evaluar");
          break;
        case "eu":
          preguntaList.setTitle("Ebaluatuko den ikaslea");
          break;
        case "fr":
          preguntaList.setTitle("Élève à évaluer");
          break;
        default:
          preguntaList.setTitle("Student to rate");
      }
      
      preguntaList.setChoiceValues(alumnes);
      preguntaList.setRequired(true);
      
      //Afegit preguntes grid (rúbrica)
      var columnes=[];
      for (i=1;i<rangrubrica.getNumRows()-1;i++){
        var titol = rangrubrica.getCell(i+2,1).getValue();
        for (j=0; j<rangrubrica.getNumColumns()-2;j++){
          var contingut= rangrubrica.getCell(i+2,j+2).getValue();
          contingut = contingut.replace("\n", "");
          columnes[j]=rangrubrica.getCell(1,j+2).getValue() +": " + contingut;
        };
        var preguntaGrid= form.addGridItem();
        preguntaGrid.setTitle(titol);
        preguntaGrid.setRows([titol]);
        preguntaGrid.setColumns(columnes);
        preguntaGrid.setRequired(true);
        
        columnes=preguntaGrid.getColumns();  //Recuperem el que hem posat al formulari
        for (var j=0;j<columnes.length;j++){  
          columnes[j]
          var desc = columnes[j].split(': ');  //Eliminem el nivell de la descripció
          columnes[j] = desc[1];
          for (var k=2; k<desc.length;k++){
            columnes[j] = columnes[j] + ": " +desc[k];
          }          
        };   
        columnes=[columnes];
        var rangpreguntes=rubricaActual.getRange(i+2, 2, 1,rangrubrica.getNumColumns()-2).setValues(columnes);  //Posem a la rúbrica el mateix que al formulari
      };
      var preguntatext = form.addParagraphTextItem();
      var properties = PropertiesService.getDocumentProperties();   
      var idioma = properties.getProperty('Idioma');   
      switch(idioma){
        case "ca":
          preguntatext.setTitle('Comentaris');
          break;
        case "es":
          preguntatext.setTitle('Comentarios');
          break;
        case "eu":
          preguntatext.setTitle('Iruzkinak');
          break;
        case "fr":
          preguntatext.setTitle('Commentaires');
          break;
        default:
          preguntatext.setTitle('Comments');
      }      
      
      //Canviar el menú, treient Crear formulari i posant el que correspongui
      onOpen();
    }
    catch(err) {

    };
  };
};


/*
 * A partir de les respostes dels alumnes,
 * crea un full de càlcul, hi posa totes les respostes dels alumnes,
 * calcula notes i mitjanes,
 * i les posa en un nou full del llibre on s'executa l'script
 */
function procesFormulari() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();  
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var llistaprofes = llibreActual.getSheetByName(nom_full_prof);
  var rangalumnes = llistaalumnes.getDataRange();
  var rangprofes = llistaprofes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var nombreprofes = rangprofes.getNumRows()-1;
  var mat_rubrica = [];
  
  //Recupero el ID del fomrulari, del ScriptDB
  var documentProperties = PropertiesService.getDocumentProperties();
  var formurl = documentProperties.getProperty('Formurl');  
  var formnom = documentProperties.getProperty('Formnom');
  
  
  //Obro el Formulari, creo un full de càlcul per desar les respostes i
  //fixo aquest full com a destí del formulari
  var formid = documentProperties.getProperty('Formid');
  
  var form = FormApp.openById(formid);
  var record_p = documentProperties.getProperty('Proces');  //Mirem que no estigui reprocessant 
  var reprocess=1;
  if (record_p!="1"){
    
    //Deso al full d'estadística que s'ha processat
    var per_auto = "10%";
    var per_co= "40%";
    var per_prof= "50%";
    reprocess=0;
    var cv="https://docs.google.com/spreadsheets/d/1eNg5xQ1nq_Psm0JgPw0RWKBPatp4-us890tCDTT4Vrg/";
    var fullOrigen = SpreadsheetApp.openByUrl(cv).getSheetByName("Analytics");
    var filesple = fullOrigen.getDataRange().getNumRows()+1;
    var range = fullOrigen.getRange("A" + filesple + ":B" + filesple);
    var avui = new Date();
    var data_actual11 = avui.getDate(); //Trobo dia d'avui
    var data_actual = new Date();
    data_actual.setDate(data_actual11);
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        range.setValues([["CoRubrics ca",data_actual]]);
        break;
      case "es":
        range.setValues([["CoRubrics es",data_actual]]);
        break;
      case "eu":
        range.setValues([["CoRubrics eu",data_actual]]);
        break;
      case "fr":
        range.setValues([["CoRubrics fr",data_actual]]);
        break;
      default:
        range.setValues([["CoRubrics en",data_actual]]);
    }
    
    //Creo un full al llibre de la rúbrica per posar els resultats. Per nom té
    //el dia i l'hora que es fa el processament
    var fullactiu = llibreActual.getActiveSheet();
    var nombrefulls = llibreActual.getNumSheets();
    llibreActual.setActiveSheet(llibreActual.getSheets()[nombrefulls-1]); //Activa el darrer full per insertar el nou al final
    var fullmitjanesc = llibreActual.insertSheet(); 
    var avui = Dataactual(); //busco la data i hora actual amb una funció que defineixo
    documentProperties.setProperty('nom_full_proces',avui);
    fullmitjanesc.setName(avui);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    for (i=1;i<rangrubrica.getNumRows()-6;i++){
      fullmitjanesc.insertColumnAfter(1);
      fullmitjanesc.insertColumnAfter(1);      
      fullmitjanesc.insertColumnAfter(1);
    };
    
    //Deso al ScripDb que he processat
    documentProperties.setProperty('Proces',"1");
    //Deso al ScripDb que l'he obtingt l'enllaç (per si ha processat sense fer-ho)
    documentProperties.setProperty('Mail',"1");
    
    //Canviar el menú, treient Crear formulari i posant el que correspongui
    onOpen();
  };

  var fullmitjanes = llibreActual.getSheets()[llibreActual.getNumSheets()-1]; //Es posaran les mitjanes a l'últim full
  var avui = Dataactual(); //busco la data i hora actual amb una funció que defineixo
  fullmitjanes.setName(avui);
  var confloc = llibreActual.getSpreadsheetLocale();
  var conteEN = confloc.search("en_");
  var conteEN2 = confloc.search("_GB");
  var conteEN3 = confloc.search("_US");
  var canviloc=0;
  if (conteEN2>0 || conteEN3>0 || conteEN==0){   
    var canviloc = 1;
  };
    
  //Poso la capçalera i les mides de les columnes
  var properties = PropertiesService.getDocumentProperties();  
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var num = 'Num';
      var al = 'Alumne avaluat/Grup';
      var np = 'Nombre de puntuacions';
      var coav = 'Coav';
      var auto = 'Auto';
      var prof = 'Prof';
      var nc = 'Nota quantitativa (comptant només l\'ítem més baix)';
      var npon = 'Nota quantitativa (fent mitjana ponderada de tots els ítems)';
      var nf = 'Nota global';
      var cp = 'Comentaris del professor';
      var cc = 'Comentaris dels alumnes (coavaluació)';
      var ca = 'Comentaris del propi alumne (autoavaluació)';
      var na = 'No avaluat';
      var max = "Màx. punt.";
      var mismaxpunt = "Quina serà la nota màxima? (introduïu un nombre enter)";
      break;
    case "es":
      var num = 'Num';
      var al = 'Alumno evaluado/Grupo';
      var np = 'Número de puntuaciones';
      var coav = 'Coev';
      var auto = 'Auto';
      var prof = 'Prof';
      var nc = 'Nota cuantitativa (contando solo el ítem más bajo)';
      var npon = 'Nota cuantitativa (usando la media ponderada de los ítems)';
      var nf = 'Nota global';
      var cp = 'Comentarios del profesor';
      var cc = 'Comentarios de los alumnos (coevaluación)';
      var ca = 'Comentarios del propio alumno (autoevaluación)';
      var ma = 'No evaluado';
      var max = "Máx. punt.";
      var mismaxpunt = "¿Cuál será la puntuación máxima? (introducir un número entero)";
      break;
    case "eu":
      var num = 'Zenb';
      var al = 'Ebaluatutako ikaslea/Taldea';
      var np = 'Puntuazio kopurua';
      var coav = 'Koe';
      var auto = 'Auto';
      var prof = 'Irak';
      var nc = 'Nota kuantitatiboa (item baxuena kontutan hartuz bakarrik)';
      var npon = 'Nota kuantitatiboa (item guztien batezbesteko ponderatua kontutan hartuz bakarrik)';
      var nf = 'Nota globala';
      var cp = 'Irakaslearen iruzkina';
      var cc = 'Ikasle taldearen iruzkinak (koebaluazioa)';
      var ca = 'Ikaslearen beraren iruzkinak (autoebaluazioa)';
      var na = 'No evaluado';
      var max = "Max";
      var mismaxpunt = "Zein izango da puntuazio maximoa? (Zenbakiak osoa izan behar du)";      
      break;
    case "fr":
      var num = 'Élève #';
      var al = 'Élève/Groupe';
      var np = 'Nombre d\'évaluations';
      var coav = 'Pairs';
      var auto = 'Auto';
      var prof = 'Ens';
      var nc = 'Note quantitative (incluant seulement la plus basse note)';
      var npon = 'Note quantitative (utilisant la moyenne pondérée de chaque aspect)';
      var nf = 'Note globale';
      var cp = 'Commentaires de l\'enseignant';
      var cc = 'Commentaires des élèves (évaluation par les pairs)';
      var ca = 'Commentaires des élèves-mêmes (autoévaluation)';
      var na = 'Non évalué';
      var max = "Note max.";
      var mismaxpunt = "Quel sera la plus haute note (indiquez un nombre entier)?";
      break;
    default:
      var num = 'Num';
      var al = 'Student/Group';
      var np = 'Number of ratings';
      var coav = 'Coev';
      var auto = 'Self';
      var prof = 'Teach';
      var nc = 'Quantitative score (counting only the lowest item)';
      var npon = 'Quantitative score (using the weighted average of the items)';
      var nf = 'Overall Grade';
      var cp = 'Teacher comments';
      var cc = 'Students comments (coevaluation)';
      var ca = 'Comments from students themselves (autoevaluation)';
      var na = 'Not assessed';
      var max = "Max grade";
      var mismaxpunt = "What will be the maximum grade?? (enter a integer)";
  }
  if (record_p!="1"){ //si no reporcessem, preguntem la nota màxima
    var maxpunt = Browser.inputBox(max,mismaxpunt, Browser.Buttons.OK_CANCEL);
    if (maxpunt % 1 != 0){
      maxpunt=100;
    }
    properties.setProperty('pmax', maxpunt)
  }else{
      var maxpunt = properties.getProperty('pmax');   
  };
  fullmitjanes.getRange("A1").setValue(num);
  fullmitjanes.getRange("A1:A3").merge();
  fullmitjanes.setColumnWidth(1, 45);
  fullmitjanes.getRange("B1").setValue(al);
  fullmitjanes.getRange("B1:B3").merge();
  fullmitjanes.setColumnWidth(2, 200);
  fullmitjanes.getRange("C1").setValue(np);
  fullmitjanes.getRange("C1:E2").merge();
  fullmitjanes.getRange("C1").setWrap(true);
  fullmitjanes.getRange("C3").setValue(coav);
  fullmitjanes.getRange("D3").setValue(auto);
  fullmitjanes.getRange("E3").setValue(prof);
  fullmitjanes.getRange("C3:E3").setBorder(true,true,true,true,false,false);
  fullmitjanes.getRange("C3:E3").setBackground("#fff2cc");
  fullmitjanes.setColumnWidth(3, 37);
  fullmitjanes.setColumnWidth(4, 37);
  fullmitjanes.setColumnWidth(5, 37);
    
  fullmitjanes.getRange("A1:C2").setBackground("#DDDDDD");
  fullmitjanes.getRange("A1:E2").setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange("A1:C2").setFontWeight("bold");
  fullmitjanes.getRange("A:Z").setVerticalAlignment("middle");
  fullmitjanes.getRange("A:Z").setHorizontalAlignment("center");
  fullmitjanes.getRange("A:Z").setWrap(true);
  fullmitjanes.getRange("C3:E3").setWrap(false);
  
  //Poso la capçalera amb els aspectes de la rúbrica
  var numcolumnes = rangrubrica.getNumColumns();
  var columnafinal=1;
  for (i=1;i<rangrubrica.getNumRows()-1;i++){
      var aspecte = rangrubrica.getCell(i+2,1).getValue();
      var pes = rangrubrica.getCell(i+2,numcolumnes).getValue();
      fullmitjanes.setColumnWidth(3*i+3,37);
      fullmitjanes.setColumnWidth(3*i+4,37);
      fullmitjanes.setColumnWidth(3*i+5,37);
      fullmitjanes.getRange(3,3*i+3,1,1).setValue(coav);
      fullmitjanes.getRange(3,3*i+3+1,1,1).setValue(auto);
      fullmitjanes.getRange(3,3*i+3+2,1,1).setValue(prof);    
      fullmitjanes.getRange(3,3*i+3,1,3).setBackground("#fff2cc");
      fullmitjanes.getRange(1,3*i+3,1,3).merge();
      fullmitjanes.getRange(1,3*i+3,1,3).setWrap(true);
      fullmitjanes.getRange(3,3*i+3,1,3).setWrap(false); 
      fullmitjanes.getRange(2,3*i+3,1,3).merge();
      fullmitjanes.getRange(1,3*i+3,1,1).setValue(aspecte);
      fullmitjanes.getRange(2,3*i+3,1,1).setNumberFormat("0");
      fullmitjanes.getRange(2,3*i+3,1,1).setValue(pes);
      fullmitjanes.getRange(1,3*i+3,3,3).setBorder(true,true,true,true,false,false);
      fullmitjanes.getRange(1,3*i+3,1,1).setBackground("#DDDDDD");
      fullmitjanes.getRange(1,3*i+3,2,1).setFontWeight("bold");
      fullmitjanes.getRange(1,3*i+3,2,1).setVerticalAlignment("middle");
      fullmitjanes.getRange(1,3*i+3,2,1).setHorizontalAlignment("center");    
      fullmitjanes.getRange(2,3*i+3,1,1).setBackground("#cc4125");
      fullmitjanes.getRange(2,3*i+3,1,1).setFontColor("white");
      fullmitjanes.getRange(2,3*i+3,1,1).setNumberFormat("0%");
      fullmitjanes.getRange(2,3*i+3,1,1).setBorder(true,true,true,true,true,true);
      columnafinal = 3*i+6;
  };

  fullmitjanes.setColumnWidth(columnafinal, 37);
  fullmitjanes.setColumnWidth(columnafinal+1, 37);
  fullmitjanes.setColumnWidth(columnafinal+2, 37);
  fullmitjanes.getRange(1,columnafinal,2,3).merge();
  fullmitjanes.getRange(1,columnafinal,2,3).clearFormat();
  fullmitjanes.getRange(1,columnafinal,1,3).merge();
  fullmitjanes.getRange(2,columnafinal,1,2).merge();
  
  fullmitjanes.getRange(1,columnafinal,1,3).setWrap(true);
  fullmitjanes.getRange(1,columnafinal,1,1).setValue(nc);
  fullmitjanes.getRange(1,columnafinal,1,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(1,columnafinal,1,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(1,columnafinal,1,1).setFontWeight("bold");
  fullmitjanes.getRange(1,columnafinal,2,3).setVerticalAlignment("middle");
  fullmitjanes.getRange(1,columnafinal,2,3).setHorizontalAlignment("center");    
  fullmitjanes.getRange(2,columnafinal,1,3).setBackground("#cc4125");
  fullmitjanes.getRange(2,columnafinal,1,3).setFontColor("white");
  fullmitjanes.getRange(2,columnafinal,1,3).setFontWeight("bold");
  fullmitjanes.getRange(2,columnafinal,1,1).setValue(max);
  fullmitjanes.getRange(2,columnafinal+2,1,1).setValue(maxpunt);
  var celmaxpunt = fullmitjanes.getRange(2,columnafinal+2,1,1).getA1Notation();

  
  fullmitjanes.getRange(3,columnafinal,1,1).setValue(coav);
  fullmitjanes.getRange(3,columnafinal+1,1,1).setValue(auto);
  fullmitjanes.getRange(3,columnafinal+2,1,1).setValue(prof);
  fullmitjanes.getRange(3,columnafinal,1,3).setWrap(false);
  fullmitjanes.getRange(3,columnafinal,1,3).setBackground("#fff2cc");
  fullmitjanes.getRange(3,columnafinal,1,3).setBorder(true,true,true,true,false,false);

  fullmitjanes.getRange(1,columnafinal+3,1,3).merge();
  fullmitjanes.setColumnWidth(columnafinal+3, 37);
  fullmitjanes.setColumnWidth(columnafinal+4, 37);
  fullmitjanes.setColumnWidth(columnafinal+5, 37);
  fullmitjanes.getRange(1,columnafinal+3,1,1).setValue(npon);
  fullmitjanes.getRange(1,columnafinal+3,1,1).setWrap(true);
  fullmitjanes.getRange(1,columnafinal+3,1,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(1,columnafinal+3,1,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(1,columnafinal+3,1,1).setFontWeight("bold");
  fullmitjanes.getRange(3,columnafinal+3,1,1).setValue(coav);
  fullmitjanes.getRange(3,columnafinal+4,1,1).setValue(auto);
  fullmitjanes.getRange(3,columnafinal+5,1,1).setValue(prof);
  fullmitjanes.getRange(3,columnafinal+3,1,3).setWrap(false);
  fullmitjanes.getRange(3,columnafinal+3,1,3).setBackground("#fff2cc");
  fullmitjanes.getRange(3,columnafinal+3,1,3).setBorder(true,true,true,true,false,false);
  
  var rangsuma = fullmitjanes.getRange(2,6,1,columnafinal-8).getA1Notation();
  fullmitjanes.getRange(2,columnafinal+3,1,3).merge();
  fullmitjanes.getRange(2,columnafinal+3,1,1).setFormula("=SUM("+rangsuma+")");
  fullmitjanes.getRange(2,columnafinal+3,1,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(2,columnafinal+3,1,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(2,columnafinal+3,1,1).setFontWeight("bold");
  fullmitjanes.getRange(2,columnafinal+3,1,1).setNumberFormat("0%");
  if (fullmitjanes.getRange(2,columnafinal+3,1,1).getCell(1,1).getValue()!=1){
    fullmitjanes.getRange(2,columnafinal+3,1,1).setBackground("#e06666");
  }else{
    fullmitjanes.getRange(2,columnafinal+3,1,1).setBackground("#93c47d");
  };
  
  //Afegim columna per mitjana ponderada de les 3 notes (co, auto i profe)
  fullmitjanes.setColumnWidth(columnafinal+6, 37);
  fullmitjanes.setColumnWidth(columnafinal+7, 37);
  fullmitjanes.setColumnWidth(columnafinal+8, 37);
  fullmitjanes.getRange(1,columnafinal+6,1,3).merge();
  fullmitjanes.getRange(1,columnafinal+6,1,1).setWrap(true);
  fullmitjanes.getRange(1,columnafinal+6,1,1).setValue(nf);
  fullmitjanes.getRange(1,columnafinal+6,1,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(1,columnafinal+6,1,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(1,columnafinal+6,1,1).setFontWeight("bold");
  fullmitjanes.getRange(1,columnafinal+6,1,1).setHorizontalAlignment("center");
  fullmitjanes.getRange(1,columnafinal+6,1,1).setVerticalAlignment("middle");
  fullmitjanes.getRange(3,columnafinal+6,1,3).setFontWeight("bold");
  fullmitjanes.getRange(2,columnafinal+6,1,1).setValue(coav);
  fullmitjanes.getRange(2,columnafinal+7,1,1).setValue(auto);
  fullmitjanes.getRange(2,columnafinal+8,1,1).setValue(prof);
  if (reprocess!="1"){ //si reprocessem, no modifiquem els % que ha indicat
    fullmitjanes.getRange(3,columnafinal+6,1,1).setValue(per_co);
    fullmitjanes.getRange(3,columnafinal+7,1,1).setValue(per_auto);
    fullmitjanes.getRange(3,columnafinal+8,1,1).setValue(per_prof);
  };
  fullmitjanes.getRange(2,columnafinal+6,2,3).setVerticalAlignment("middle");
  fullmitjanes.getRange(2,columnafinal+6,2,3).setHorizontalAlignment("center"); 
  fullmitjanes.getRange(2,columnafinal+6,1,3).setBackground("#f9cb9c");
  fullmitjanes.getRange(3,columnafinal+6,1,3).setBackground("#cc4125");
  fullmitjanes.getRange(3,columnafinal+6,1,3).setFontColor("white");
  fullmitjanes.getRange(3,columnafinal+6,1,3).setNumberFormat("0%");
  
  //Afegim columna per comentaris del profe
  fullmitjanes.setColumnWidth(columnafinal+9, 245);
  fullmitjanes.getRange(1,columnafinal+9,3,1).merge();
  fullmitjanes.getRange(1,columnafinal+9,2,1).setWrap(true);
  fullmitjanes.getRange(1,columnafinal+9,2,1).setValue(cp);
  fullmitjanes.getRange(1,columnafinal+9,2,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(1,columnafinal+9,2,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(1,columnafinal+9,2,1).setFontWeight("bold");
  fullmitjanes.getRange(1,columnafinal+9,3,1).setHorizontalAlignment("center");
  fullmitjanes.getRange(1,columnafinal+9,3,1).setVerticalAlignment("middle");
 
  //Afegim columna per  comentaris dels companys
  fullmitjanes.setColumnWidth(columnafinal+10, 245);
  fullmitjanes.getRange(1,columnafinal+10,3,1).merge();
  fullmitjanes.getRange(1,columnafinal+10,2,1).setWrap(true);
  fullmitjanes.getRange(1,columnafinal+10,2,1).setValue(cc);
  fullmitjanes.getRange(1,columnafinal+10,2,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(1,columnafinal+10,2,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(1,columnafinal+10,2,1).setFontWeight("bold");
  fullmitjanes.getRange(1,columnafinal+10,3,1).setHorizontalAlignment("center");
  fullmitjanes.getRange(1,columnafinal+10,3,1).setVerticalAlignment("middle");  
 
  //Afegim columna per comentaris del propi alumne
  fullmitjanes.setColumnWidth(columnafinal+11, 245);
  fullmitjanes.getRange(1,columnafinal+11,3,1).merge();
  fullmitjanes.getRange(1,columnafinal+11,2,1).setWrap(true);
  fullmitjanes.getRange(1,columnafinal+11,2,1).setValue(ca);
  fullmitjanes.getRange(1,columnafinal+11,2,1).setBorder(true,true,true,true,true,true);
  fullmitjanes.getRange(1,columnafinal+11,2,1).setBackground("#f9cb9c");
  fullmitjanes.getRange(1,columnafinal+11,2,1).setFontWeight("bold");
  fullmitjanes.getRange(1,columnafinal+11,3,1).setHorizontalAlignment("center");
  fullmitjanes.getRange(1,columnafinal+11,3,1).setVerticalAlignment("middle");  
  
  //Agafem les respostes del formulari i les posem en una matriu
  var matriu_res=[];
  var matriu_res_auto=[];
  var matriu_res_profe=[];
  var re = form.getResponses();
  var num_resp = re.length;
  var num_col = numcolumnes;
  var nom_respon=[]; //GAFE Afegim variable per desar qui respon 
  for (var i=0;i<num_resp;i++) {  //Defineixo la matriu com a bidimensional i la poso tot en blanc
     matriu_res[i] = [];
     nom_respon[i]=[]; //GAFE Afegim variable per desar qui respon 
     nom_respon[i][0]=0; //GAFE Afegim variable per desar qui respon 
     nom_respon[i][1]=0; //GAFE Afegim variable per desar qui respon 
    for (var j=0; j<num_col;j++){
       matriu_res[i][j]=0;   
    };
  }; 
 
  var nom_respon_auto=[]; //GAFE Afegim variable per desar qui respon 
  for (var i=0;i<num_resp;i++) {  //Defineixo la matriu com a bidimensional i la poso tot en blanc
     matriu_res_auto[i] = [];
     nom_respon_auto[i]=[]; //GAFE Afegim variable per desar qui respon 
     nom_respon_auto[i][0]=0; //GAFE Afegim variable per desar qui respon 
     nom_respon_auto[i][1]=0; //GAFE Afegim variable per desar qui respon 
     for (var j=0; j<num_col;j++){
       matriu_res_auto[i][j]=0;   
    };
  }; 
  
  var nom_respon_profe=[]; //GAFE Afegim variable per desar qui respon 
  for (var i=0;i<num_resp;i++) {  //Defineixo la matriu com a bidimensional i la poso tot en blanc
     matriu_res_profe[i] = [];
     nom_respon_profe[i]=[]; //GAFE Afegim variable per desar qui respon 
     nom_respon_profe[i][0]=0; //GAFE Afegim variable per desar qui respon 
     nom_respon_profe[i][1]=0; //GAFE Afegim variable per desar qui respon 
     for (var j=0; j<num_col;j++){
       matriu_res_profe[i][j]=0;   
    };
  }; 

  
  var asp = rangalumnes.getNumColumns()-1;  
  //Creo una matriu amb els noms dels alumne i els seus mails
  var alumnes = [];
  for (var i=0;i<nombrealumnes;i++) {  //Defineixo la matriu com a bidimensional i l'omplo amb zeros
    alumnes[i] = [];
    for (j=0; j<asp+1;j++){
       alumnes[i][j]=0;
    };
  }; 
  for (i=0;i<alumnes.length;i++){
    alumnes[i][0]=rangalumnes.getCell(i+2,1).getValue(); //Omplim el primer camp de alumnes amb tots els noms del full Alumnes
    for (var z=1;z<asp+1;z++){
      alumnes[i][z]=rangalumnes.getCell(i+2,z+1).getValue()
      var fora_espais_finals = alumnes[i][z].trim();
        if (alumnes[i][z] != fora_espais_finals){
          rangalumnes.getCell(i+2,z+1).setValue(fora_espais_finals);
        };
        alumnes[i][z]=fora_espais_finals;
      if (alumnes[i][z]==="") {
        alumnes[i].splice(z,1); //Eliminem si algun grup té menys mails que altres
      };
    };
  }; 

  //Creo una matriu amb els noms dels profes i els seus mails
  var profes = [];
  for (var i=0;i<nombreprofes;i++) {  //Defineixo la matriu com a bidimensional i l'omplo amb zeros
    profes[i] = [];
    for (j=0; j<2;j++){
       profes[i][j]=0;
    };
  }; 
  for (i=0;i<profes.length;i++){
    profes[i][0]=rangprofes.getCell(i+2,1).getValue(); //Omplim el primer camp de profes amb tots els noms del full Profes
    profes[i][1]=rangprofes.getCell(i+2,2).getValue(); //Omplim el primer camp de profes amb tots els mails del full Profes
    var fora_espais_finals = profes[i][1].trim();
    if (profes[i][1] != fora_espais_finals){
      rangalumnes.getCell(i+2,2).setValue(fora_espais_finals);
    };
    profes[i][1]=fora_espais_finals;
  };   


  //Creo tres matriu pels comentaris,una per profes, una per alumnes i una per autoavaluació. El primer camp és el nom de l'alumne i el segon seran tots els comentaris
  var comentaris_alu = [];
  var comentaris_prof = [];
  var comentaris_auto = [];
  for (var i=0;i<nombrealumnes;i++) {  //Defineixo la matriu com a bidimensional i l'omplo amb zeros
    comentaris_alu[i] = [];
    comentaris_prof[i] = [];
    comentaris_auto[i] = [];
    for (j=0; j<asp+1;j++){
       comentaris_alu[i][j]="";
       comentaris_prof[i][j]="";
       comentaris_auto[i][j]="";
    };
  }; 
  for (i=0;i<alumnes.length;i++){
    comentaris_alu[i][0]=rangalumnes.getCell(i+2,1).getValue(); //Omplim el primer camp de alumnes amb tots els noms del full Alumnes
    comentaris_prof[i][0]=rangalumnes.getCell(i+2,1).getValue(); //Omplim el primer camp de alumnes amb tots els noms del full Alumnes
    comentaris_auto[i][0]=rangalumnes.getCell(i+2,1).getValue(); //Omplim el primer camp de alumnes amb tots els noms del full Alumnes
  }; 


  //Agafem les respostes del formulari
  var x=0;
  var s=0;
  var y=0;
  for (var i=0; i < num_resp; i++) {
   var respota_f = re[i];
   nom_respon[s][0]=respota_f.getRespondentEmail(); //GAFE Desem qui ha respost 
   var itemResponses = respota_f.getItemResponses();
        
    // Buscar el nom a qui correspon l'adreça de qui respon, dins la matriu alumnes 
    // Un cop trobat el nom, comprobar si és nom_respon[i][1]. Si és així, no contar la resposta (i=i-1)
    // i desar la resposta en una altra matriu d'autoevaluació
    var nom_buscat="";
    for (var k=0;k<alumnes.length;k++){
      for (var r=1;r<alumnes[k].length;r++){
        if (nom_respon[s][0]===alumnes[k][r]){
          nom_buscat=alumnes[k][0];
        };
      };
    };
    
    var profe_trobat=0;
    for (var k=0;k<profes.length;k++){
      for (var r=1;r<profes[k].length;r++){
        if (nom_respon[s][0]===profes[k][r]){
          profe_trobat=1;
        };
      };
    };

    for (var j=0;j<itemResponses.length;j++) {
     var itemResponse = itemResponses[j]; 
     if (j===0) {
       matriu_res[s][j]=itemResponse.getResponse();
       nom_respon[s][1]=itemResponse.getResponse(); //GAFE Desem a qui ha avaluat    
       var avaluat="";
       avaluat=nom_respon[s][1];
       if (nom_respon[s][1]===nom_buscat){ //Miro si es AUTO
         matriu_res_auto[x][j]=matriu_res[s][j];
         nom_respon_auto[x][1]=nom_respon[s][1]; //GAFE Desem a qui ha avaluat 
         nom_respon_auto[x][0]=nom_respon[s][0]; //GAFE Desem qui ha avaluat 
       }else{
         if (profe_trobat===1){ //Miro si es PROFE
           matriu_res_profe[y][j]=matriu_res[s][j];
           nom_respon_profe[y][1]=nom_respon[s][1]; //GAFE Desem a qui ha avaluat 
           nom_respon_profe[y][0]=nom_respon[s][0]; //GAFE Desem qui ha avaluat 
         };
       };           
     }else{
       if (j===itemResponses.length-1){
         var comentaris = itemResponse.getResponse();
       }else{
         matriu_res[s][j]=itemResponse.getResponse()[0]; 
         if (nom_respon[s][1]===nom_buscat){  //Miro si es AUTO
           matriu_res_auto[x][j]=matriu_res[s][j];
         }else{
           if (profe_trobat===1){  //Miro si es PROFE
             matriu_res_profe[y][j]=matriu_res[s][j];
           };
         };
       };
     };
    };
       
   if (nom_respon[s][1]===nom_buscat){ //Si és autoavaluació ho elimino de la matri_res
     matriu_res.splice(s,1);
     nom_respon.splice(s,1);
     s=s-1;
     x=x+1;
     //cerco l'avaluat dins la matriu comentaris_prof
     for (var h=0;h<comentaris_auto.length;h++){
       if (comentaris_auto[h][0]===avaluat) {
         if (comentaris_auto[h][1]===""){
           comentaris_auto[h][1]=comentaris;
         }else{
           comentaris_auto[h][1]=comentaris_auto[h][1] + '; ' + comentaris;           
         };
       };
     };
   }else{
     if (profe_trobat===1){  //Si és avaluació del profe ho elimino de la matriu res
       matriu_res.splice(s,1);
       nom_respon.splice(s,1);
       s=s-1;
       y=y+1;  
       //cerco l'avaluat dins la matriu comentaris_prof
       for (var h=0;h<comentaris_prof.length;h++){
         if (comentaris_prof[h][0]===avaluat) {
           if (comentaris_prof[h][1]===""){
             comentaris_prof[h][1]=comentaris;
           }else{
             comentaris_prof[h][1]=comentaris_prof[h][1] + '; ' + comentaris;
           };
         };
       };
     }else{
       //cerco l'avaluat dins la matriu comentaris_alu
       for (var h=0;h<comentaris_alu.length;h++){
         if (comentaris_alu[h][0]===avaluat) {
           if (comentaris_alu[h][1]===""){
             comentaris_alu[h][1]=comentaris;             
           }else{
             comentaris_alu[h][1]=comentaris_alu[h][1] + '; ' + comentaris;
           };
         };
       };
     };
   };
   s++;
  };  

    
  for (var x=0; x<matriu_res_auto.length;x++) { //Eliminem els registres buits de la matriu de AUTOAVALUACIÓ
    if (matriu_res_auto[x][0]===0) {
      matriu_res_auto.splice(x,1);
      x=x-1;
    };
  };
  for (var x=0; x<nom_respon_auto.length;x++) { //Eliminem els registres buits de la matriu de AUTOAVALUACIÓ
    if (nom_respon_auto[x][1]===0) {
      nom_respon_auto.splice(x,1);
      x=x-1;
    };
  };
  
  for (var y=0; y<matriu_res_profe.length;y++) { //Eliminem els registres buits de la matriu de PROFE
    if (matriu_res_profe[y][1]===0) {
      matriu_res_profe.splice(y,1);
      y=y-1;
    };
  };

   for (var y=0; y<nom_respon_profe.length;y++) { //Eliminem els registres buits de la matriu de PROFE
    if (nom_respon_profe[y][1]===0) {
      nom_respon_profe.splice(y,1);
      y=y-1;
    };
  };
  
    // GAFE Trobem les persones que han avaluat més d'un cop una persona

  var respdup = Trobar_duplicat(nom_respon); //La funció retorna un objecte amb dues propietats.
  var respdup_auto = Trobar_duplicat(nom_respon_auto); //La funció retorna un objecte amb dues propietats.
  var respdup_profe = Trobar_duplicat(nom_respon_profe); //La funció retorna un objecte amb dues propietats.
  var resp_duplicat = respdup.nom_duplicat; //Una és una array amb els usuaris que han repetit
  var resp_duplicat_auto = respdup_auto.nom_duplicat; //Una és una array amb els usuaris que han repetit
  var resp_duplicat_profe = respdup_profe.nom_duplicat; //Una és una array amb els usuaris que han repetit
  var resp_elim = respdup.resp_eliminar; //L'altra és una array amb les respostes que caldrà eliminar si no es volen duplicats
  var resp_elim_auto = respdup_auto.resp_eliminar; //L'altra és una array amb les respostes que caldrà eliminar si no es volen duplicats
  var resp_elim_profe = respdup_profe.resp_eliminar; //L'altra és una array amb les respostes que caldrà eliminar si no es volen duplicats
  var usuaris_dup = ""; 
  var usuaris_dup_auto = ""; 
  var usuaris_dup_profe = ""; 
  if (resp_duplicat.length>0){ 
    for (i=0;i<resp_duplicat.length;i++){ 
      if (i===0) { 
        usuaris_dup = resp_duplicat[i];
      }else{
        usuaris_dup = usuaris_dup + ', ' + resp_duplicat[i];
      };
    };
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var duplicats = Browser.msgBox('Puntuacions duplicades Coavaluació','Els següents usuaris han puntuat més d\'un cop en fer la COAVALUACiÓ: '+ usuaris_dup +'. Vols eliminar les respostes duplicades? (només es comptarà la última puntuació)', Browser.Buttons.YES_NO);
        break;
      case "es":
        var duplicats = Browser.msgBox('Puntuaciones duplicadas Coevaluación','Los siguientes usuarios han puntuado más de una vez a algun alumno en la COEVALUACIÓN: '+ usuaris_dup +'. ¿Quieres eliminar las respuestas duplicadas? (solo se contará la última puntuación)', Browser.Buttons.YES_NO);
        break;
      case "eu":
        var duplicats = Browser.msgBox('Koebaluazioa - Bikoiztutako kalifikazioak','Ondorengo ikasleek behin baino gehiagotan kalifikatu dute hurrengo ikasleak: '+ usuaris_dup +'. Bikoiztutako erantzunak ezabatu nahi al dituzu? (azken kalifikazioa soilik hartuko da kontutan)', Browser.Buttons.YES_NO);
        break;
      case "fr":
        var duplicats = Browser.msgBox('Notes des évaluations des pairs se doublent','Les utilisateurs suivants ont noté un pair plus d\'une fois dans ÉVALUATION DES PAIRS:'+ usuaris_dup +'. Voulez-vous supprimer les doublets?  (Seulement la dernière note comptera.)', Browser.Buttons.YES_NO);
        break;
      default:
        var duplicats = Browser.msgBox('Co-evaluation scores duplicate','The following users have rated a student more than once in COEVALUATION:'+ usuaris_dup +'. Do you want to delete duplicate answers? (only the last score will be counted)', Browser.Buttons.YES_NO);
    }
  };
  
  if (resp_duplicat_auto.length>0){ 
    for (i=0;i<resp_duplicat_auto.length;i++){ 
      if (i===0) { 
        usuaris_dup_auto = resp_duplicat_auto[i];
      }else{
        usuaris_dup_auto = usuaris_dup_auto + ', ' + resp_duplicat_auto[i];
      };
    };
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var duplicats_auto = Browser.msgBox('Puntuacions duplicades Autoavaluació','Els següents usuaris han puntuat més d\'un cop en fer la AUTOAVALUACiÓ: '+ usuaris_dup_auto +'. Vols eliminar les respostes duplicades? (només es comptarà la última puntuació)', Browser.Buttons.YES_NO);
        break;
      case "es":
        var duplicats_auto = Browser.msgBox('Puntuaciones duplicadas Autooevaluación','Los siguientes usuarios se han puntuado más de una vez en la AUTOEVALUACIÓN '+ usuaris_dup_auto +'. ¿Quieres eliminar las respuestas duplicadas? (solo se contará la última puntuación)', Browser.Buttons.YES_NO);
        break;
      case "eu":
        var duplicats_auto = Browser.msgBox('Autoebaluzioa - Bikoiztutako kalifikazioak','Ondorengo ikasleek behin baino gehiagotan kalifikatu dute hurrengo ikasleak: '+ usuaris_dup_auto +'. Bikoiztutako erantzunak ezabatu nahi al dituzu? (azken kalifikazioa soilik hartuko da kontutan)', Browser.Buttons.YES_NO);
        break;
      case "fr":
        var duplicats_auto = Browser.msgBox('Notes de l\'autoévaluation se doublent','Les utilisateurs suivants se sont notés plus d\'une fois dans AUTOÉVALUATION.: '+ usuaris_dup_auto +'. Voulez-vous supprimer les doublets?  (Seulement la dernière note comptera.)', Browser.Buttons.YES_NO);
        break;
      default:
        var duplicats_auto = Browser.msgBox('Self-evaluation scores duplicate','The following users have scored more than once in SELF-EVALUATION: '+ usuaris_dup_auto +'. Do you want to delete duplicate answers? (only the last score will be counted)', Browser.Buttons.YES_NO);
    }    
  };
  
  if (resp_duplicat_profe.length>0){ 
    for (i=0;i<resp_duplicat_profe.length;i++){ 
      if (i===0) { 
        usuaris_dup_profe = resp_duplicat_profe[i];
      }else{
        usuaris_dup_profe = usuaris_dup_profe + ', ' + resp_duplicat_profe[i];
      };
    };
    var properties = PropertiesService.getDocumentProperties();  
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var duplicats_profe = Browser.msgBox('Puntuacions duplicades Professor','Els següents PROFESSORS han puntuat més d\'un cop en fer l\'avaluació: '+ usuaris_dup_profe +'. Vols eliminar les respostes duplicades? (només es comptarà la última puntuació)', Browser.Buttons.YES_NO);
        break;
      case "es":
        var duplicats_profe = Browser.msgBox('Puntuaciones duplicadas Profesor','Los siguientes PROFESORES se han puntuado más de una vez a algun alumno '+ usuaris_dup_profe +'. ¿Quieres eliminar las respuestas duplicadas? (solo se contará la última puntuación)', Browser.Buttons.YES_NO);
        break;
      case "eu":
        var duplicats_profe = Browser.msgBox('Irakaslea - Bikoiztutako kalifikazioak','Ondorengo ikasleek behin baino gehiagotan kalifikatu dute hurrengo ikasleak: '+ usuaris_dup_profe +'. Bikoiztutako erantzunak ezabatu nahi al dituzu? (azken kalifikazioa soilik hartuko da kontutan)', Browser.Buttons.YES_NO);
        break;
      case "fr":
        var duplicats_profe = Browser.msgBox('Notes de l\'enseignant se doublent','Les ENSEIGNANTS suivants ont noté un élève plus d\'une fois.: '+ usuaris_dup_profe +'. Voulez-vous supprimer les doublets?  (Seulement la dernière note comptera.)', Browser.Buttons.YES_NO);
        break;
      default:
        var duplicats_profe = Browser.msgBox('Teacher-evaluation scores duplicate','The following TEACHERS have rated a student more than once: '+ usuaris_dup_profe +'. Do you want to delete duplicate answers? (only the last score will be counted)', Browser.Buttons.YES_NO);
    }
  };    
    
  var num_fil=rangrubrica.getNumRows()-1;
  var matriu_respostes = Omplir_matriu_respostes(matriu_res,matriu_res.length,num_fil); //Creem una matriu respostes, però amb puntuacions enlloc de descripcions
  var matriu_respostes_auto = Omplir_matriu_respostes(matriu_res_auto,matriu_res_auto.length,num_fil); //Creem una matriu respostes, però amb puntuacions enlloc de descripcions
  var matriu_respostes_profe = Omplir_matriu_respostes(matriu_res_profe,matriu_res_profe.length,num_fil); //Creem una matriu respostes, però amb puntuacions enlloc de descripcions
   
  
  if (duplicats==='yes'){  //GAFE Eliminem les files duplicades 
    for (i=0;i<resp_elim.length;i++){
      //eliminar les files de matriu_respostes que calgui
      matriu_respostes.splice(resp_elim[i]-i, 1);
    };
  };
  if (duplicats_auto==='yes'){  //GAFE Eliminem les files duplicades 
    for (i=0;i<resp_elim_auto.length;i++){
      //eliminar les files de matriu_respostes que calgui
      matriu_respostes_auto.splice(resp_elim_auto[i]-i, 1);
    };
  };
  
  if (duplicats_profe==='yes'){  //GAFE Eliminem les files duplicades 
    for (i=0;i<resp_elim_profe.length;i++){
      //eliminar les files de matriu_respostes que calgui
      matriu_respostes_profe.splice(resp_elim_profe[i]-i, 1);
    };
  };
     
  var alumnes = Trobar_alumnes(matriu_respostes);  //Torna un matriu amb una fila per alumne i les puntuacions sumades
  var alumnes_auto = Trobar_alumnes(matriu_respostes_auto);  //Torna un matriu amb una fila per alumne i les puntuacions sumades
  var alumnes_profe = Trobar_alumnes(matriu_respostes_profe);  //Torna un matriu amb una fila per alumne i les puntuacions sumades    
  var rang_resultats = fullmitjanes.getRange(4,1,alumnes.length,3*alumnes[0].length+11);
  var num_alum=1;

  var valormaxim=0; //Busco el grau que val més punts
  for (j=2;j<rangrubrica.getNumColumns();j++){
    if (valormaxim < rangrubrica.getCell(2,j).getValues()){
      valormaxim = rangrubrica.getCell(2,j).getValue();
    };      
  };
  
  for (i=0;i<alumnes.length;i++){
    rang_resultats.getCell(i+1,1).setValue(num_alum);
    num_alum = num_alum+1;
    fullmitjanes.getRange(i+4,1,1,1).setBorder(true,true,true,true,true,true);
    var resp=1;
    var resp_auto=1;
    var resp_profe=1;
     for (j=0;j<alumnes[0].length;j++){
      var l=3*j;
      if (j==0){
        l=2;       
      };
       if (j==1){
         if (alumnes[i][j]==0){
           resp="-";
         };
         if (alumnes_auto[i][j]==0) {
             resp_auto="-";
         };
         if (alumnes_profe[i][j]==0) {
             resp_profe="-";
         };
       };
       if (resp==1){
         rang_resultats.getCell(i+1,l).setValue(alumnes[i][j]);
       }else{
         rang_resultats.getCell(i+1,l).setValue(resp);
       };
       if (resp_auto==1){
         rang_resultats.getCell(i+1,l+1).setValue(alumnes_auto[i][j]);
       }else{
         rang_resultats.getCell(i+1,l+1).setValue(resp_auto);
       };
       if (resp_profe==1){
         rang_resultats.getCell(i+1,l+2).setValue(alumnes_profe[i][j]);
       }else{
         rang_resultats.getCell(i+1,l+2).setValue(resp_profe);
       };       
              
       fullmitjanes.getRange(i+4,l,1,3).setBorder(true,true,true,true,false,false);
    };

    //Busco comentaris del profe i els poso
    var tr=-1;
    for (var f=0;f<comentaris_prof.length;f++){
      if (comentaris_prof[f][0]===alumnes[i][0]) {
        tr= f;
      };
    };
    if (tr!=-1){
      rang_resultats.getCell(i+1,3*alumnes_profe[i].length+9).setValue(comentaris_prof[tr][1]);
    };
      
    //Busco comentaris dels alumnes i els poso
    var tr=-1;
    for (var f=0;f<comentaris_alu.length;f++){
      if (comentaris_alu[f][0]===alumnes[i][0]) {
        tr= f;
      };
    };
    if (tr!=-1){
      rang_resultats.getCell(i+1,3*alumnes_profe[i].length+10).setValue(comentaris_alu[tr][1]);    
    };
        
    //Busco comentaris dels propis alumnes i els poso
    var tr=-1;
    for (var f=0;f<comentaris_auto.length;f++){
      if (comentaris_auto[f][0]===alumnes[i][0]) {
        tr= f;
      };
    };
    if (tr!=-1){
      rang_resultats.getCell(i+1,3*alumnes_profe[i].length+11).setValue(comentaris_auto[tr][1]);
    };
    var rangminim="";
    var rangminim_auto="";
    var rangminim_profe="";
    var pesos="";
    for (h=1;h<alumnes[0].length-1;h++){
      rangminim = rangminim + fullmitjanes.getRange(i+4,3+3*h,1,1).getA1Notation() + ";" ;
      pesos = pesos + fullmitjanes.getRange(2,3+3*h,1,1).getA1Notation() + ";" ;
    };
    rangminim = rangminim.substring(0, rangminim.length-1);
    for (h=1;h<alumnes_auto[0].length-1;h++){
      rangminim_auto = rangminim_auto + fullmitjanes.getRange(i+4,3+3*h+1,1,1).getA1Notation() + ";" ;
    };
    rangminim_auto = rangminim_auto.substring(0, rangminim_auto.length-1);
    for (h=1;h<alumnes_profe[0].length-1;h++){
      rangminim_profe = rangminim_profe + fullmitjanes.getRange(i+4,3+3*h+2,1,1).getA1Notation() + ";" ;
    };
    rangminim_profe = rangminim_profe.substring(0, rangminim_profe.length-1);
    pesos = pesos.substring(0, pesos.length-1);
    if (fullmitjanes.getRange(i+4,3,1,1).getCell(1,1).getValue()===0){
      rang_resultats.getCell(i+1,3*alumnes[i].length).setValue("-");
      rang_resultats.getCell(i+1,3*alumnes[i].length+3).setValue("-");
    }else{
      var vm=10/valormaxim;
      vm = vm + "";
      if (resp==1){
        if (canviloc===1){
          rang_resultats.getCell(i+1,3*alumnes[i].length).setFormula("round(min("+rangminim+")*"+vm+",2)*"+celmaxpunt+"/10");
          rang_resultats.getCell(i+1,3*alumnes[i].length+3).setFormula("round(SUMPRODUCT({"+ rangminim +"};{"+ pesos+"})*"+vm+",2)*"+celmaxpunt+"/10");
        }else{
          vm = vm.replace(".", ",");
          rang_resultats.getCell(i+1,3*alumnes[i].length).setFormula("round(min("+rangminim+")*"+vm+";2)*"+celmaxpunt+"/10");
          rang_resultats.getCell(i+1,3*alumnes[i].length+3).setFormula("round(SUMPRODUCT({"+ rangminim +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
        };
      }else{
        rang_resultats.getCell(i+1,3*alumnes[i].length).setValue("-");
        rang_resultats.getCell(i+1,3*alumnes[i].length+3).setValue("-");
      };
      fullmitjanes.getRange(i+4,3*alumnes[i].length,1,3).setBorder(true,true,true,true,false,false);
      fullmitjanes.getRange(i+4,3*alumnes[i].length+3,1,3).setBorder(true,true,true,true,false,false);
      fullmitjanes.getRange(i+4,3*alumnes[i].length+6,1,3).setBorder(true,true,true,true,true,true);
      fullmitjanes.getRange(i+4,3*alumnes[i].length+9,1,3).setBorder(true,true,true,true,true,true);
    };
    
   
    if (fullmitjanes.getRange(i+4,4,1,1).getCell(1,1).getValue()===0){
      rang_resultats.getCell(i+1,3*alumnes_auto[i].length+1).setValue("-");
      rang_resultats.getCell(i+1,3*alumnes_auto[i].length+4).setValue("-");
    }else{
      if (resp_auto==1){
        if (canviloc===1){
          rang_resultats.getCell(i+1,3*alumnes_auto[i].length+1).setFormula("round(min("+rangminim_auto+")*"+vm+";2)*"+celmaxpunt+"/10");
          rang_resultats.getCell(i+1,3*alumnes_auto[i].length+4).setFormula("round(SUMPRODUCT({"+ rangminim_auto +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
        }else{
          vm = vm.replace(".", ",");
          rang_resultats.getCell(i+1,3*alumnes_auto[i].length+1).setFormula("round(min("+rangminim_auto+")*"+vm+";2)*"+celmaxpunt+"/10");
          rang_resultats.getCell(i+1,3*alumnes_auto[i].length+4).setFormula("round(SUMPRODUCT({"+ rangminim_auto +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
        };          
      }else{
        rang_resultats.getCell(i+1,3*alumnes_auto[i].length+1).setValue("-");
        rang_resultats.getCell(i+1,3*alumnes_auto[i].length+4).setValue("-");
      };
    };
    if (fullmitjanes.getRange(i+4,4,1,1).getCell(1,1).getValue()===0){
      rang_resultats.getCell(i+1,3*alumnes_profe[i].length+2).setValue("-");
      rang_resultats.getCell(i+1,3*alumnes_profe[i].length+5).setValue("-");
    }else{
      if (resp_profe==1){
        if (canviloc===1){
          rang_resultats.getCell(i+1,3*alumnes_profe[i].length+2).setFormula("round(min("+rangminim_profe+")*"+vm+";2)*"+celmaxpunt+"/10");
          rang_resultats.getCell(i+1,3*alumnes_profe[i].length+5).setFormula("round(SUMPRODUCT({"+ rangminim_profe +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
        }else{
          vm = vm.replace(".", ",");
          rang_resultats.getCell(i+1,3*alumnes_profe[i].length+2).setFormula("round(min("+rangminim_profe+")*"+vm+";2)*"+celmaxpunt+"/10");
          rang_resultats.getCell(i+1,3*alumnes_profe[i].length+5).setFormula("round(SUMPRODUCT({"+ rangminim_profe +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
        };          
      }else{
        rang_resultats.getCell(i+1,3*alumnes_profe[i].length+2).setValue("-");
        rang_resultats.getCell(i+1,3*alumnes_profe[i].length+5).setValue("-");
      };
    };
    //Posem la fórmula de la nota final (ponderant les 3)
    fullmitjanes.getRange(i+4,columnafinal+6,1,3).merge();
    var rangpesosfinal = fullmitjanes.getRange(3,columnafinal+6,1,3).getA1Notation();
    var notesfinals = fullmitjanes.getRange(i+4,columnafinal+3,1,3).getA1Notation();
    fullmitjanes.getRange(i+4,columnafinal+6,1,3).setFormula("round(SUMPRODUCT("+ rangpesosfinal +";"+ notesfinals+");2)");  
    fullmitjanes.getRange(i+4,columnafinal+6,1,3).setHorizontalAlignment("center");
    fullmitjanes.getRange(i+4,columnafinal+6,1,3).setVerticalAlignment("middle"); 
  };

};

    
function creaCoRubrics(){
  var properties = PropertiesService.getDocumentProperties(); 
  properties.setProperty('tasca_av_id', "") 
  properties.setProperty('tasca_co_id', "")
  properties.setProperty('tasca_pf_id', "")
  properties.setProperty('tasca_ng_id', "")
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var avis = Browser.msgBox('Crear CoRubrics','Aquest procés eliminarà tots els fulls de càlcul existents i crearà la plantilla CoRubrics. Voleu continuar?', Browser.Buttons.YES_NO);
      break;
    case "es":
      var avis = Browser.msgBox('Crear CoRubrics','Este proceso eliminará todas las hojas de cálculo existentes y creará la plantilla CoRubrics. ¿Continuar?', Browser.Buttons.YES_NO);
      break;
    case "eu":
      var avis = Browser.msgBox('CoRubrics sortu','Pozesu honek kalkulu-orri honetako orrialde guztiak ezabatu eta CoRubrics txantilloia sortuko du. Jarraitu?', Browser.Buttons.YES_NO);
      break;
    case "fr":
      var avis = Browser.msgBox('Créez CoRubrics','Ce processus supprimera toutes les feuilles de la feuille de calcul et créera un gabarit CoRubrics.  Vous voulez continuer?', Browser.Buttons.YES_NO);
      break;
    default:
      var avis = Browser.msgBox('Create CoRubrics','This process deletes all sheets in this spreadsheet and will create a CoRubrics template. Do you wish to continue?', Browser.Buttons.YES_NO);
  }  
  if (avis==='yes'){
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var nom_full_rubrica = "Rúbrica";
        var nom_full_alumnes= "Alumnes";
        var nom_full_prof= "Profes";
        var source = SpreadsheetApp.openById("1TO2r293VzuOv2hajxp6BmCFMtumPCRgtQI9Y_xg-A10");
        break;
      case "es":
        var nom_full_rubrica = "Rúbrica";
        var nom_full_alumnes= "Alumnos";
        var nom_full_prof= "Profes";
        var source = SpreadsheetApp.openById("1tK22TMB-ZFbQKse_wfmpUKp65vxF-pVIide1wfRVRvw");
        break;
      case "eu":
        var nom_full_rubrica = "Errubrika";
        var nom_full_alumnes= "Ikasleak";
        var nom_full_prof= "Irakasleak";
        var source = SpreadsheetApp.openById("1x5Q1rlhwjDcf3DgIcbyoM-Rb6CJ6jVbejtY0-xpXx64");
        break;
      case "fr":
        var nom_full_rubrica = "Grille";
        var nom_full_alumnes= "Élèves";
        var nom_full_prof= "Enseignants";
        var source = SpreadsheetApp.openById("14cI12heSn50_kyp_g7uzi_dBjW6K6odD7YQDA4Te_XU");
        break;
      default:
        var nom_full_rubrica = "Rubric";
        var nom_full_alumnes= "Students";
        var nom_full_prof= "Teachers";
        var source = SpreadsheetApp.openById("1qaZu9J6FzHyxZxNjTJfBWhN7dUWNJ8rVcMDnFzoBTDo");
    }  
    var sheet = source.getSheetByName(nom_full_rubrica);
    var sheet1 = source.getSheetByName(nom_full_alumnes);
    var sheet2 = source.getSheetByName(nom_full_prof);
    var destination = SpreadsheetApp.getActiveSpreadsheet();
    for (var i=0; i<destination.getNumSheets();i++){
      var full1 = destination.getSheets()[i].setName("Full" + i);
    };
    var full_rub=sheet.copyTo(destination); 
    full_rub.setName(nom_full_rubrica); 
    var full_al = sheet1.copyTo(destination); 
    full_al.setName(nom_full_alumnes);
    var full_profs = sheet2.copyTo(destination); 
    full_profs.setName(nom_full_prof);
    sleep (2000);
    var destination = SpreadsheetApp.getActiveSpreadsheet();
    var fulls= destination.getNumSheets();
    for (var i=0; i<fulls-3;i++){
      var full1 = destination.getSheetByName("Full" + i)
      destination.deleteSheet(full1);
    };
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('Importacio', "1");
    
    //Canviar el menú, treient Crear CoRubrics i posant el que correspongui
    esborradB();
    onOpen();
  };
};


function catala(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "ca"); 
  onOpen();
};

function espanol(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "es"); 
  onOpen();
};

function euskara(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "eu"); 
  onOpen();
};

function english(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "en"); 
  onOpen();
};

function français(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "fr"); 
  onOpen();
};

function activaCoRubrics(){
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nouformulari = Browser.msgBox('CoRubrics','CoRubrics s\'ha activat correctament', Browser.Buttons.OK);
      break;
    case "es":
      var nouformulari = Browser.msgBox('CoRubrics','CoRubrics se ha activado correctamente', Browser.Buttons.OK);
      break;
    case "eu":
      var nouformulari = Browser.msgBox('CoRubrics','CoRubrics ondo aktibatu da', Browser.Buttons.OK);
      break;
    case "fr":
      var nouformulari = Browser.msgBox('CoRubrics','CoRubrics est activé', Browser.Buttons.OK);
      break;
    default: 
      var nouformulari = Browser.msgBox('CoRubrics','CoRubrics is enabled', Browser.Buttons.OK);
  };
  onOpen();
};

function impalClasroom(){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  switch(idioma){
    case "ca":
      var nom_html='importal_ca';
      break;
    case "es":
      var nom_html='importal_es';
      break;
    case "eu":
      var nom_html='importal_eu';
      break;
    case "fr":
      var nom_html='importal_fr';
      break;
    default:
      var nom_html='importal_en';
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'CoRubrics');
};

function importacio_al(formObject){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma'); 
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'CoRubrics');
  var cursid = formObject.combo_curs;
  //var cursid = '5110160217';
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      var nom_html='Cal triar un curs de Classroom';
      var curs_m='Curs';
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      var nom_html='Es necesario elegir un curso de Classroom';
      var curs_m='Curso';
      break;
    case "eu":
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      var nom_html='Classroomeko ikasgela hautatzea beharrezkoa da';
      var curs_m='Curso';
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      var nom_html='Il est nécessaire de choisir un cours de Classroom';
      var curs_m='Cours';
      break;
    default:
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
      var nom_html='It is necessary to choose a Classroom course';
      var curs_m='Course';
  }; 
  if (cursid == 0){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    impalClasroom();
  }else{
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('cursid', cursid);
    //importem alumnes
    var alumnes=Classroom.Courses.Students.list(cursid); 
    var matriu=new Array(alumnes.students.length);
    for (var i=0;i<alumnes.students.length;i++){
      var cognom_al=alumnes.students[i].profile.name.familyName;
      var nom_al=alumnes.students[i].profile.name.givenName;
      var mail_al=alumnes.students[i].profile.emailAddress;
      matriu[i]=new Array(2);
      matriu[i][0]=cognom_al+", "+nom_al;
      matriu[i][1]=mail_al;
    };
    matriu.sort();
    var rang_full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_full_alumnes).getRange(2,1,alumnes.students.length,2);
    rang_full.setValues(matriu);
    //importem professors
    var profes=Classroom.Courses.Teachers.list(cursid);
    var matriu=new Array(profes.teachers.length);
    for (var i=0;i<profes.teachers.length;i++){
      var cognom_pr=profes.teachers[i].profile.name.familyName;
      var nom_pr=profes.teachers[i].profile.name.givenName;
      var mail_pr=profes.teachers[i].profile.emailAddress;
      matriu[i]=new Array(2);
      matriu[i][0]=cognom_pr+", "+nom_pr;
      matriu[i][1]=mail_pr;
    };
    matriu.sort();
    var rang_full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_full_prof).getRange(2,1,profes.teachers.length,2);
    rang_full.setValues(matriu);
    
    switch(idioma){
      case "ca":
        var frase = properties.setProperty('frase', 'Els alumnes i els professors s\'han importat correctament');
        var boto = properties.setProperty('boto', 'Tancar finestra');
        break;
      case "es":
        var frase = properties.setProperty('frase', 'Los alumnos y los profesores se han importado correctamente');
        var boto = properties.setProperty('boto', 'Cerrar ventana');
        break;    
      case "eu":
        var frase = properties.setProperty('frase', 'Ikasle zein irakasleak zuzen inportatu dira');
        var boto = properties.setProperty('boto', 'Itxi');
        break;
      case "fr":
        var frase = properties.setProperty('frase', 'Élèves et enseignants ont été importés correctement');
        var boto = properties.setProperty('boto', 'Fermez');
        break;
      default:
        var frase = properties.setProperty('frase', 'Students and teachers have been properly imported');
        var boto = properties.setProperty('boto', 'Close');
    };  
    var nom_html='confirma';
    var html = HtmlService
    .createTemplateFromFile(nom_html)
    .evaluate();
    
    SpreadsheetApp.getUi().showModelessDialog(html, 'CoRubrics'); 
  };   
};

function classFormulari(){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  switch(idioma){
    case "ca":
      var nom_html='clasF_ca';
      break;
    case "es":
      var nom_html='clasF_es';
      break;
    case "eu":
      var nom_html='clasF_eu';
      break;
    case "fr":
      var nom_html='clasF_fr';
      break;
    default:
      var nom_html='clasF_en';
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();  
  SpreadsheetApp.getUi().showModelessDialog(html, 'CoRubrics');  
};

function classform(formObject){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma'); 
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'CoRubrics');
  var cursid = formObject.combo_curs;
  var titol = formObject.titol;
  var descripcio = formObject.descripcio;
  switch(idioma){
    case "ca":
      var nom_html='Cal triar un curs de Classroom';
      var curs_m='Curs';
      break;
    case "es":
      var nom_html='Es necesario elegir un curso de Classroom';
      var curs_m='Curso';
      break;
    case "eu":
      var nom_html='Il est nécessaire de choisir un cours de Classroom';
      var curs_m='Cours';
      break;
    case "fr":
      var nom_html='Classroomeko ikasgela hautatzea beharrezkoa da';
      var curs_m='Curso';
      break;
    default:
      var nom_html='It is necessary to choose a Classroom course';
      var curs_m='Course';
  }; 
  if (cursid == 0){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    classFormulari();
  }else{  
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('cursid', cursid);
    var fomrid = documentProperties.getProperty('Formid');
    var formurl = documentProperties.getProperty('Formurl');
    //CREAR L'ANUNCI AMB EL FORMULARI ADJUNT
    var creo_anunci = {
      "courseId": cursid,
      "text": titol,
      'materials': [  
        {'link': { 'url': formurl }} 
      ],  
      "state": "PUBLISHED"
    }
    var anunci_creat=Classroom.Courses.Announcements.create(creo_anunci, cursid)
    //var tasca_creada=Classroom.Courses.CourseWork.create(creo_tasca, cursid); //Només es poden canviar notes de tasques creades per la API
    var anunci_id=anunci_creat.id;
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('anunci_id', anunci_id);  
    switch(idioma){
      case "ca":
        var frase = properties.setProperty('frase', 'El formulari s\'ha publicat correctament al Classroom');
        var boto = properties.setProperty('boto', 'Tancar finestra');
        break;
      case "es":
        var frase = properties.setProperty('frase', 'El formulario se ha publicado correctamente en Classroom');
        var boto = properties.setProperty('boto', 'Cerrar ventana');
        break;    
      case "eu":
        var frase = properties.setProperty('frase', 'El formulario se ha publicado correctamente en Classroom');
        var boto = properties.setProperty('boto', 'Itxi');
        break;
      case "fr":
        var frase = properties.setProperty('frase', 'Le formulaire a été publié dans Classroom avec succès');
        var boto = properties.setProperty('boto', 'Fermez');
        break;
      default:
        var frase = properties.setProperty('frase', 'The form has been successfully published in Classroom');
        var boto = properties.setProperty('boto', 'Close');
    };  
    var nom_html='confirma';
    var html = HtmlService
    .createTemplateFromFile(nom_html)
    .evaluate();
    
    SpreadsheetApp.getUi().showModelessDialog(html, 'CoRubrics'); 
    
    //Deso al ScripDb que he publicat
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('Mail', "1");
    //Canviar el menú, treient Enviar formulari i posant el que correspongui  
    onOpen();  
  };
};

function nota_classroom(){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  switch(idioma){
    case "ca":
      var nom_html='p_clas_ca';
      break;
    case "es":
      var nom_html='p_clas_es';
      break;
    case "eu":
      var nom_html='p_clas_eu';
      break;
     case "fr":
      var nom_html='p_clas_fr';
      break;
    default:
      var nom_html='p_clas_en';
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)      
  .evaluate();
  html.setTitle('CoRubrics');
  SpreadsheetApp.getUi().showSidebar(html);  
}

function publicanotes(formObject){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var sbtitolav=' - Autoavaluació';
      var sbtitolco=' - Coavaluació';
      var sbtitolpf=' - Professors';
      var sbtitolng=' - Nota global';
      var nom_full_alumnes= "Alumnes";
      break;
    case "es":
      var sbtitolav=' - Autoevaluación';
      var sbtitolco=' - Coevaluación';
      var sbtitolpf=' - Profesor';
      var sbtitolng=' - Nota global';
      var nom_full_alumnes= "Alumnos";
      break;
    case "eu":
      var sbtitolav=' - Autoebaluzioa';
      var sbtitolco=' - Koebaluazioa';
      var sbtitolpf=' - Irakaslea(k)';
      var sbtitolng=' - Nota globala';
      var nom_full_alumnes= "Ikasleak";
      break;
    case "fr":
      var sbtitolav=' - Autoévaluation';
      var sbtitolco=' - Évaluation des pairs';
      var sbtitolpf=' - Enseignants';
      var sbtitolng=' - Note globale';
      var nom_full_alumnes= "Élèves";
      break;
    default:
      var sbtitolav=' - Self-evaluation';
      var sbtitolco=' - Co-evaluation';
      var sbtitolpf=' - Teachers';
      var sbtitolng=' - Overall grade';
      var nom_full_alumnes= "Students";
  };
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  html.setTitle('CoRubrics');
  SpreadsheetApp.getUi().showSidebar(html); 
  var tasca_av_id = properties.getProperty('tasca_av_id'); 
  var tasca_co_id = properties.getProperty('tasca_co_id');
  var tasca_pf_id = properties.getProperty('tasca_pf_id');
  var tasca_ng_id = properties.getProperty('tasca_ng_id');
  if (tasca_av_id == "" && tasca_co_id == "" && tasca_pf_id == "" && tasca_ng_id == ""){
    var cursid = formObject.combo_curs;
    var titol = formObject.titol;
    var descripcio = formObject.descripcio;
    properties.setProperty('titol_tasca',titol); 
    properties.setProperty('desc_tasca',descripcio); 
    properties.setProperty('cursid',cursid); 
  }else{
    var cursid = properties.getProperty('cursid'); 
    var titol = properties.getProperty('titol_tasca');
    var descripcio = properties.getProperty('desc_tasca');
  };
  var nf = formObject.nf;
  var av = formObject.av;
  var co = formObject.co;
  var pf = formObject.pf;
  var ng= formObject.ng;
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var rang_full = llibreActual.getSheetByName(nom_full_alumnes).getDataRange();
  var mails_alumnes= rang_full.getValues();
  var fil_mails = rang_full.getNumRows();
  var col_mails = rang_full.getNumColumns();
  var rangmitjanes = llibreActual.getSheets()[llibreActual.getNumSheets()-1].getDataRange(); //S'agafen els resultat del darrer full
  var fil_mitjanes = rangmitjanes.getNumRows();
  var col_mitjanes = rangmitjanes.getNumColumns();
  var valormitjanes = rangmitjanes.getValues();
  //Agafem la puntuació màxima
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      break;
    default:
      var nom_full_rubrica = "Rubric";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var columnafinal = 3*(rangrubrica.getNumRows()-2)+6;
  var fullmitjanes = llibreActual.getSheets()[llibreActual.getNumSheets()-1];
  var pmax = properties.getProperty('pmax');
  
  if (av==='on'){
    if (tasca_av_id == null || tasca_av_id==""){ //Només creem si no Re-publica
      //CREO TASCA AUTOAVALUACIÓ I AFEGEIXO NOTES  
      var titolav = titol + sbtitolav;
      var creo_tasca = {
        "courseId": cursid,
        "title": titolav,
        "description": descripcio,
        "maxPoints": pmax,
        "workType":"ASSIGNMENT",
        "state": "PUBLISHED"
      }
      var tasca_creada=Classroom.Courses.CourseWork.create(creo_tasca, cursid); //Només es poden canviar notes de tasques creades per la API
      var tasca_av_id=tasca_creada.id;
      //Deso el id de la tasca a la propitat tasca_av_id
      properties.setProperty('tasca_av_id',tasca_av_id); 
    };
    //Assignem notes i comentaris
    var tasques_env=Classroom.Courses.CourseWork.StudentSubmissions.list(cursid, tasca_av_id); //recuperem totes les tasques de tots els alumnes
    for (var i=3; i<fil_mitjanes;i++){ //agafo alumne per alumne. A partir del nom trobem el mailal full alumnes, d'aquí trobarem l'userid. A partir de l'userid,el submissionid
      var nom = valormitjanes[i][1];
      for (var l=1; l<fil_mails;l++){
        if (nom==mails_alumnes[l][0]){  //busquem el nom al full alumnes
          var m = 1;
          while(mails_alumnes[l][m] != "" && m<col_mails){  //posarem les notes a tots els alumnes si és un grup
            var mail_st = mails_alumnes[l][m];
            var nota_st = valormitjanes[i][col_mitjanes-8-3*nf];
            if (!(isNaN(nota_st))){
              var llista_st = Classroom.Courses.Students.list(cursid); //Agafo la llista d'alumnes
              properties.setProperty('totalum', llista_st.students.length);
              for (var j=0;j<llista_st.students.length;j++){
                var mail1=llista_st.students[j].profile.emailAddress;  //agafem alumne per alumne i comparem amb el de la cel·la
                if (mail1==mail_st){
                  var userid=llista_st.students[j].userId; //Trobem userid de l'usuari de la cel·la
                  for (var k=0; k<tasques_env.studentSubmissions.length;k++){
                    var env_us=tasques_env.studentSubmissions[k].userId; //busco la tasca de l'usuari de la cel·la
                    var env_id=tasques_env.studentSubmissions[k].id;
                    if (env_us==userid){
                      var nalum= j;
                      properties.setProperty('nalum', nalum);
                      var html = HtmlService
                      .createTemplateFromFile('updating2')
                      .evaluate();      
                      html.setTitle('CoRubrics');
                      SpreadsheetApp.getUi().showSidebar(html);
                      var reso = {'draftGrade':nota_st};
                      var extra={'updateMask':'draftGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_av_id, env_id,extra); // Actualitzem la nota esborrany
                      var reso = {'assignedGrade':nota_st};
                      var extra={'updateMask':'assignedGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_av_id, env_id,extra);// Actualitzaem la nota que veu l'alumne
                      var reso = {'return':1};
                    }
                  }        
                }
              }
            }
            m++;
          }
        }
      }
    }      
  };
  if (co==='on'){
    if (tasca_co_id==null || tasca_co_id==""){ //Només creem si no Republica
      //CREO TASCA AUTOAVALUACIÓ I AFEGEIXO NOTES I COMENTARIS SI CAL  
      var titolco = titol + sbtitolco;
      var creo_tasca = {
        "courseId": cursid,
        "title": titolco,
        "description": descripcio,
        "maxPoints": pmax,
        "workType":"ASSIGNMENT",
        "state": "PUBLISHED"
      }
      var tasca_creada=Classroom.Courses.CourseWork.create(creo_tasca, cursid); //Només es poden canviar notes de tasques creades per la API
      var tasca_co_id=tasca_creada.id;
      //Deso el id de la tasca a la propitat tasca_co_id
      properties.setProperty('tasca_co_id',tasca_co_id);  
    };
    //Assignem notes i comentaris
    var tasques_env=Classroom.Courses.CourseWork.StudentSubmissions.list(cursid, tasca_co_id); //recuperem totes les tasques de tots els alumnes
    for (var i=3; i<fil_mitjanes;i++){ //agafo alumne per alumne. A partir del nom trobem el mailal full alumnes, d'aquí trobarem l'userid. A partir de l'userid,el submissionid
      var nom = valormitjanes[i][1];
      for (var l=1; l<fil_mails;l++){
        if (nom==mails_alumnes[l][0]){  //busquem el nom al full alumnes
          var m = 1;
          while(mails_alumnes[l][m] != "" && m<col_mails){  //posarem les notes a tots els alumnes si és un grup
            var mail_st = mails_alumnes[l][1];
            var nota_st = valormitjanes[i][col_mitjanes-9-3*nf];
            if (!(isNaN(nota_st))){
              var llista_st = Classroom.Courses.Students.list(cursid); //Agafo la llista d'alumnes
              for (var j=0;j<llista_st.students.length;j++){
                var mail1=llista_st.students[j].profile.emailAddress;  //agafem alumne per alumne i comparem amb el de la cel·la
                if (mail1==mail_st){
                  var userid=llista_st.students[j].userId; //Trobem userid de l'usuari de la cel·la
                  for (var k=0; k<tasques_env.studentSubmissions.length;k++){
                    var env_us=tasques_env.studentSubmissions[k].userId; //busco la tasca de l'usuari de la cel·la
                    var env_id=tasques_env.studentSubmissions[k].id;
                    if (env_us==userid){
                      var reso = {'draftGrade':nota_st};
                      var extra={'updateMask':'draftGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_co_id, env_id,extra); // Actualitzem la nota esborrany
                      var reso = {'assignedGrade':nota_st};
                      var extra={'updateMask':'assignedGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_co_id, env_id,extra);// Actualitzaem la nota que veu l'alumne
                      var reso = {'return':1};
                    }
                  }        
                }
              }
            }
            m++;
          }
        }
      }
    } 
 };
  if (pf==='on'){
    if (tasca_pf_id==null || tasca_pf_id==""){ //Només creem si no Republica
      //CREO TASCA AUTOAVALUACIÓ I AFEGEIXO NOTES I COMENTARIS SI CAL  
      var titolpf = titol + sbtitolpf;
      var creo_tasca = {
        "courseId": cursid,
        "title": titolpf,
        "description": descripcio,
        "maxPoints": pmax,
        "workType":"ASSIGNMENT",
        "state": "PUBLISHED"
      }
      var tasca_creada=Classroom.Courses.CourseWork.create(creo_tasca, cursid); //Només es poden canviar notes de tasques creades per la API
      var tasca_pf_id=tasca_creada.id;
      //Deso el id de la tasca a la propitat tasca_pf_id
      properties.setProperty('tasca_pf_id',tasca_pf_id);
    };
    //Assignem notes
    var tasques_env=Classroom.Courses.CourseWork.StudentSubmissions.list(cursid, tasca_pf_id); //recuperem totes les tasques de tots els alumnes
    for (var i=3; i<fil_mitjanes;i++){ //agafo alumne per alumne. A partir del nom trobem el mailal full alumnes, d'aquí trobarem l'userid. A partir de l'userid,el submissionid
      var nom = valormitjanes[i][1];
      for (var l=1; l<fil_mails;l++){
        if (nom==mails_alumnes[l][0]){  //busquem el nom al full alumnes
          var m = 1;
          while(mails_alumnes[l][m] != "" && m<col_mails){  //posarem les notes a tots els alumnes si és un grup
            var mail_st = mails_alumnes[l][1];
            var nota_st = valormitjanes[i][col_mitjanes-7-3*nf];
            if (!(isNaN(nota_st))){
              var llista_st = Classroom.Courses.Students.list(cursid); //Agafo la llista d'alumnes
              for (var j=0;j<llista_st.students.length;j++){
                var mail1=llista_st.students[j].profile.emailAddress;  //agafem alumne per alumne i comparem amb el de la cel·la
                if (mail1==mail_st){
                  var userid=llista_st.students[j].userId; //Trobem userid de l'usuari de la cel·la
                  for (var k=0; k<tasques_env.studentSubmissions.length;k++){
                    var env_us=tasques_env.studentSubmissions[k].userId; //busco la tasca de l'usuari de la cel·la
                    var env_id=tasques_env.studentSubmissions[k].id;
                    if (env_us==userid){
                      var reso = {'draftGrade':nota_st};
                      var extra={'updateMask':'draftGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_pf_id, env_id,extra); // Actualitzem la nota esborrany
                      var reso = {'assignedGrade':nota_st};
                      var extra={'updateMask':'assignedGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_pf_id, env_id,extra);// Actualitzaem la nota que veu l'alumne
                      var reso = {'return':1};
                    }
                  }        
                }
              }
            }
            m++;
          }
        }
      }
    }       
 };
  if (ng==='on'){
    if (tasca_ng_id==null || tasca_ng_id==""){ //Només creem si no Republica
      //CREO TASCA AUTOAVALUACIÓ I AFEGEIXO NOTES I COMENTARIS SI CAL  
      var titolng = titol + sbtitolng;
      var creo_tasca = {
        "courseId": cursid,
        "title": titolng,
        "description": descripcio,
        "maxPoints": pmax,
        "workType":"ASSIGNMENT",
        "state": "PUBLISHED"
      }
      var tasca_creada=Classroom.Courses.CourseWork.create(creo_tasca, cursid); //Només es poden canviar notes de tasques creades per la API
      var tasca_ng_id=tasca_creada.id;
      //Deso el id de la tasca a la propitat tasca_pf_id
      properties.setProperty('tasca_ng_id',tasca_ng_id);
    };
    //Assignem notes
    var tasques_env=Classroom.Courses.CourseWork.StudentSubmissions.list(cursid, tasca_ng_id); //recuperem totes les tasques de tots els alumnes
    for (var i=3; i<fil_mitjanes;i++){ //agafo alumne per alumne. A partir del nom trobem el mailal full alumnes, d'aquí trobarem l'userid. A partir de l'userid,el submissionid
      var nom = valormitjanes[i][1];
      for (var l=1; l<fil_mails;l++){
        if (nom==mails_alumnes[l][0]){  //busquem el nom al full alumnes
          var m = 1;
          while(mails_alumnes[l][m] != "" && m<col_mails){  //posarem les notes a tots els alumnes si és un grup
            var mail_st = mails_alumnes[l][1];
            var nota_st = valormitjanes[i][col_mitjanes-6];
            if (!(isNaN(nota_st))){
              var llista_st = Classroom.Courses.Students.list(cursid); //Agafo la llista d'alumnes
              for (var j=0;j<llista_st.students.length;j++){
                var mail1=llista_st.students[j].profile.emailAddress;  //agafem alumne per alumne i comparem amb el de la cel·la
                if (mail1==mail_st){
                  var userid=llista_st.students[j].userId; //Trobem userid de l'usuari de la cel·la
                  for (var k=0; k<tasques_env.studentSubmissions.length;k++){
                    var env_us=tasques_env.studentSubmissions[k].userId; //busco la tasca de l'usuari de la cel·la
                    var env_id=tasques_env.studentSubmissions[k].id;
                    if (env_us==userid){
                      var reso = {'draftGrade':nota_st};
                      var extra={'updateMask':'draftGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_ng_id, env_id,extra); // Actualitzem la nota esborrany
                      var reso = {'assignedGrade':nota_st};
                      var extra={'updateMask':'assignedGrade'};
                      var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tasca_ng_id, env_id,extra);// Actualitzaem la nota que veu l'alumne
                      var reso = {'return':1};
                    }
                  }        
                }
              }
            }
            m++;
          }
        }
      }
    }       
  };
  
  
  switch(idioma){
    case "ca":
      var frase = properties.setProperty('frase', 'Les qualificacions s\'han publicat correctament a Classroom');
      var boto = properties.setProperty('boto', 'Tancar finestra');
      break;
    case "es":
      var frase = properties.setProperty('frase', 'La calificaicones se han publicado correctamente en Classroom');
      var boto = properties.setProperty('boto', 'Cerrar ventana');
      break;    
    case "eu":
      var frase = properties.setProperty('frase', 'Classroomen kalifikazioak ondo argitaratu dira');
      var boto = properties.setProperty('boto', 'Itxi');
      break;
    case "fr":
      var frase = properties.setProperty('frase', 'Les notes ont été publiées dans Classroom avec succès');
      var boto = properties.setProperty('boto', 'Fermez');
      break;
    default:
      var frase = properties.setProperty('frase', 'Grades have been successfully published in Classroom');
      var boto = properties.setProperty('boto', 'Close');
  };  
  var nom_html='confirma';
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();
  html.setTitle('CoRubrics');
  SpreadsheetApp.getUi().showSidebar(html);   
  
  //Deso al ScripDb que s'ha publicat
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Mail2', "1");
  
  //Canviar el menú, treient Enviar formulari i posant el que correspongui
  onOpen();  
};



/*
 * Crea un full de càlcul de respostes en blanc, per si un profe vol fer avaluació d'un grup ell sol (i no vol omplir tant formularis)
*/
function fullrespostes() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();  
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnes";
      var nom_full_prof= "Profes";
      break;
    case "es":
      var nom_full_rubrica = "Rúbrica";
      var nom_full_alumnes= "Alumnos";
      var nom_full_prof= "Profes";
      break;
    case "eu":
      var nom_full_rubrica = "Errubrika";
      var nom_full_alumnes= "Ikasleak";
      var nom_full_prof= "Irakasleak";
      break;
    case "fr":
      var nom_full_rubrica = "Grille";
      var nom_full_alumnes= "Élèves";
      var nom_full_prof= "Enseignants";
      break;
    default:
      var nom_full_rubrica = "Rubric";
      var nom_full_alumnes= "Students";
      var nom_full_prof= "Teachers";
  }  
  var rubricaActual = llibreActual.getSheetByName(nom_full_rubrica);
  var rangrubrica = rubricaActual.getDataRange();
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var rangalumnes = llistaalumnes.getDataRange();
  var r_al=rangalumnes.getValues();//agafem tots els alumnes
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var aspectes = rangrubrica.getNumRows();
  var mat_rubrica = [];
  if (rangalumnes.getNumRows()-1===0){
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        Browser.msgBox('Alumnes','No has indicat cap alumne per avaluar!', Browser.Buttons.OK);
        break;
      case "es":
        Browser.msgBox('Alumnos','¡No has indicado ningún alumno a evaluar!', Browser.Buttons.OK);
        break;
      case "eu":
        Browser.msgBox('Ikasleak','Ebaluatzeko ikaslerik ez duzu adierazi!', Browser.Buttons.OK);
        break;
      case "fr":
        Browser.msgBox('Élèves','La liste d\'élèves est vide', Browser.Buttons.OK);
        break;
      default:
        Browser.msgBox('Students','The list of students is empty.', Browser.Buttons.OK);
    }
  }else{
  
    //Deso al full d'estadística que s'ha processat
    var per_auto = "10%";
    var per_co= "40%";
    var per_prof= "50%";
    var cv="https://docs.google.com/spreadsheets/d/1eNg5xQ1nq_Psm0JgPw0RWKBPatp4-us890tCDTT4Vrg/";
    var fullOrigen = SpreadsheetApp.openByUrl(cv).getSheetByName("Analytics");
    var filesple = fullOrigen.getDataRange().getNumRows()+1;
    var range = fullOrigen.getRange("A" + filesple + ":B" + filesple);
    var avui = new Date();
    var data_actual11 = avui.getDate(); //Trobo dia d'avui
    var data_actual = new Date();
    data_actual.setDate(data_actual11);
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        range.setValues([["CoRubrics ca",data_actual]]);
        break;
      case "es":
        range.setValues([["CoRubrics es",data_actual]]);
        break;
      case "eu":
        range.setValues([["CoRubrics eu",data_actual]]);
        break;
      case "fr":
        range.setValues([["CoRubrics fr",data_actual]]);
        break;
      default:
        range.setValues([["CoRubrics en",data_actual]]);
    }
    
    //Creo un full al llibre de la rúbrica per posar els resultats. Per nom té
    //el dia i l'hora que es fa el processament
    var fullactiu = llibreActual.getActiveSheet();
    var nombrefulls = llibreActual.getNumSheets();
    llibreActual.setActiveSheet(llibreActual.getSheets()[nombrefulls-1]); //Activa el darrer full per insertar el nou al final
    var fullmitjanesc = llibreActual.insertSheet(); 
    var avui = Dataactual(); //busco la data i hora actual amb una funció que defineixo
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('nom_full_proces',avui);
    fullmitjanesc.setName(avui);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    fullmitjanesc.insertColumnAfter(1);
    for (i=1;i<rangrubrica.getNumRows()-6;i++){
      fullmitjanesc.insertColumnAfter(1);
      fullmitjanesc.insertColumnAfter(1);      
      fullmitjanesc.insertColumnAfter(1);
    };
    
    //Deso al ScripDb que l'he obtingt l'enllaç (per si ha processat sense fer-ho)
    documentProperties.setProperty('Mail',"1");
    
    //Canviar el menú, treient Crear formulari i posant el que correspongui
    onOpen();
    
    
    var fullmitjanes = llibreActual.getSheets()[llibreActual.getNumSheets()-1]; //Es posaran les mitjanes a l'últim full
    var avui = Dataactual(); //busco la data i hora actual amb una funció que defineixo
    fullmitjanes.setName(avui);
    var confloc = llibreActual.getSpreadsheetLocale();
    var conteEN = confloc.search("en_");
    var conteEN2 = confloc.search("_GB");
    var conteEN3 = confloc.search("_US");
    var canviloc=0;
    if (conteEN2>0 || conteEN3>0 || conteEN==0){   
      var canviloc = 1;
    };
    
    //Poso la capçalera i les mides de les columnes
    var properties = PropertiesService.getDocumentProperties();  
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var num = 'Num';
        var al = 'Alumne avaluat/Grup';
        var np = 'Nombre de puntuacions';
        var coav = 'Coav';
        var auto = 'Auto';
        var prof = 'Prof';
        var nc = 'Nota quantitativa (comptant només l\'ítem més baix)';
        var npon = 'Nota quantitativa (fent mitjana ponderada de tots els ítems)';
        var nf = 'Nota global';
        var cp = 'Comentaris del professor';
        var cc = 'Comentaris dels alumnes (coavaluació)';
        var ca = 'Comentaris del propi alumne (autoavaluació)';
        var na = 'No avaluat';
        var max = "Màx. punt.";
        var maxpunt = 10;
        var documentProperties = PropertiesService.getDocumentProperties();
        documentProperties.setProperty('pmax', maxpunt);
        break;
      case "es":
        var num = 'Num';
        var al = 'Alumno evaluado/Grupo';
        var np = 'Número de puntuaciones';
        var coav = 'Coev';
        var auto = 'Auto';
        var prof = 'Prof';
        var nc = 'Nota cuantitativa (contando solo el ítem más bajo)';
        var npon = 'Nota cuantitativa (usando la media ponderada de los ítems)';
        var nf = 'Nota global';
        var cp = 'Comentarios del profesor';
        var cc = 'Comentarios de los alumnos (coevaluación)';
        var ca = 'Comentarios del propio alumno (autoevaluación)';
        var ma = 'No evaluado';
        var max = "Máx. punt.";
        var maxpunt = 10;
        var documentProperties = PropertiesService.getDocumentProperties();
        documentProperties.setProperty('pmax', maxpunt);
        break;
      case "eu":
        var num = 'Zenb';
        var al = 'Ebaluatutako ikaslea/Taldea';
        var np = 'Puntuazio kopurua';
        var coav = 'Koe';
        var auto = 'Auto';
        var prof = 'Irak';
        var nc = 'Nota kuantitatiboa (item baxuena kontutan hartuz bakarrik)';
        var npon = 'Nota kuantitatiboa (item guztien batezbesteko ponderatua kontutan hartuz bakarrik)';
        var nf = 'Nota globala';
        var cp = 'Irakaslearen iruzkina';
        var cc = 'Ikasle taldearen iruzkinak (koebaluazioa)';
        var ca = 'Ikaslearen beraren iruzkinak (autoebaluazioa)';
        var na = 'No evaluado';
        var max = "Max";
        var maxpunt = 10;
        var documentProperties = PropertiesService.getDocumentProperties();
        documentProperties.setProperty('pmax', maxpunt);
        break;
      case "fr":
        var num = 'Élève #';
        var al = 'Élève/Groupe';
        var np = 'Nombre d\'évaluations';
        var coav = 'Pairs';
        var auto = 'Auto';
        var prof = 'Ens';
        var nc = 'Note quantitative (incluant seulement la plus basse note)';
        var npon = 'Note quantitative (utilisant la moyenne pondérée de chaque aspect)';
        var nf = 'Note globale';
        var cp = 'Commentaires de l\'enseignant';
        var cc = 'Commentaires des élèves (évaluation par les pairs)';
        var ca = 'Commentaires des élèves-mêmes (autoévaluation)';
        var na = 'Non évalué';
        var max = "Note max.";
        var maxpunt = 100;
        var documentProperties = PropertiesService.getDocumentProperties();
        documentProperties.setProperty('pmax', maxpunt);
        break;
      default:
        var num = 'Num';
        var al = 'Student/Group';
        var np = 'Number of ratings';
        var coav = 'Coev';
        var auto = 'Self';
        var prof = 'Teach';
        var nc = 'Quantitative score (counting only the lowest item)';
        var npon = 'Quantitative score (using the weighted average of the items)';
        var nf = 'Overall Grade';
        var cp = 'Teacher comments';
        var cc = 'Students comments (coevaluation)';
        var ca = 'Comments from students themselves (autoevaluation)';
        var na = 'Not assessed';
        var max = "Max grade";
        var maxpunt = 100;
        var documentProperties = PropertiesService.getDocumentProperties();
        documentProperties.setProperty('pmax', maxpunt);
    }
    
    fullmitjanes.getRange("A1").setValue(num);
    fullmitjanes.getRange("A1:A3").merge();
    fullmitjanes.setColumnWidth(1, 45);
    fullmitjanes.getRange("B1").setValue(al);
    fullmitjanes.getRange("B1:B3").merge();
    fullmitjanes.setColumnWidth(2, 200);
    fullmitjanes.getRange("C1").setValue(np);
    fullmitjanes.getRange("C1:E2").merge();
    fullmitjanes.getRange("C1").setWrap(true);
    fullmitjanes.getRange("C3").setValue(coav);
    fullmitjanes.getRange("D3").setValue(auto);
    fullmitjanes.getRange("E3").setValue(prof);
    fullmitjanes.getRange("C3:E3").setBorder(true,true,true,true,false,false);
    fullmitjanes.getRange("C3:E3").setBackground("#fff2cc");
    fullmitjanes.setColumnWidth(3, 37);
    fullmitjanes.setColumnWidth(4, 37);
    fullmitjanes.setColumnWidth(5, 37);
    
    fullmitjanes.getRange("A1:C2").setBackground("#DDDDDD");
    fullmitjanes.getRange("A1:E2").setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange("A1:C2").setFontWeight("bold");
    fullmitjanes.getRange("A:Z").setVerticalAlignment("middle");
    fullmitjanes.getRange("A:Z").setHorizontalAlignment("center");
    fullmitjanes.getRange("A:Z").setWrap(true);
    fullmitjanes.getRange("C3:E3").setWrap(false);
    
    //Poso la capçalera amb els aspectes de la rúbrica
    var numcolumnes = rangrubrica.getNumColumns();
    var columnafinal=1;
    for (var i=1;i<rangrubrica.getNumRows()-1;i++){
      var aspecte = rangrubrica.getCell(i+2,1).getValue();
      var pes = rangrubrica.getCell(i+2,numcolumnes).getValue();
      fullmitjanes.setColumnWidth(3*i+3,37);
      fullmitjanes.setColumnWidth(3*i+4,37);
      fullmitjanes.setColumnWidth(3*i+5,37);
      fullmitjanes.getRange(3,3*i+3,1,1).setValue(coav);
      fullmitjanes.getRange(3,3*i+3+1,1,1).setValue(auto);
      fullmitjanes.getRange(3,3*i+3+2,1,1).setValue(prof);    
      fullmitjanes.getRange(3,3*i+3,1,3).setBackground("#fff2cc");
      fullmitjanes.getRange(1,3*i+3,1,3).merge();
      fullmitjanes.getRange(1,3*i+3,1,3).setWrap(true);
      fullmitjanes.getRange(3,3*i+3,1,3).setWrap(false); 
      fullmitjanes.getRange(2,3*i+3,1,3).merge();
      fullmitjanes.getRange(1,3*i+3,1,1).setValue(aspecte);
      fullmitjanes.getRange(2,3*i+3,1,1).setNumberFormat("0");
      fullmitjanes.getRange(2,3*i+3,1,1).setValue(pes);
      fullmitjanes.getRange(1,3*i+3,3,3).setBorder(true,true,true,true,false,false);
      fullmitjanes.getRange(1,3*i+3,1,1).setBackground("#DDDDDD");
      fullmitjanes.getRange(1,3*i+3,2,1).setFontWeight("bold");
      fullmitjanes.getRange(1,3*i+3,2,1).setVerticalAlignment("middle");
      fullmitjanes.getRange(1,3*i+3,2,1).setHorizontalAlignment("center");    
      fullmitjanes.getRange(2,3*i+3,1,1).setBackground("#cc4125");
      fullmitjanes.getRange(2,3*i+3,1,1).setFontColor("white");
      fullmitjanes.getRange(2,3*i+3,1,1).setNumberFormat("0%");
      fullmitjanes.getRange(2,3*i+3,1,1).setBorder(true,true,true,true,true,true);
      columnafinal = 3*i+6;
    };
    
    fullmitjanes.setColumnWidth(columnafinal, 37);
    fullmitjanes.setColumnWidth(columnafinal+1, 37);
    fullmitjanes.setColumnWidth(columnafinal+2, 37);
    fullmitjanes.getRange(1,columnafinal,2,3).merge();
    fullmitjanes.getRange(1,columnafinal,2,3).clearFormat();
    fullmitjanes.getRange(1,columnafinal,1,3).merge();
    fullmitjanes.getRange(2,columnafinal,1,2).merge();
    
    fullmitjanes.getRange(1,columnafinal,1,3).setWrap(true);
    fullmitjanes.getRange(1,columnafinal,1,1).setValue(nc);
    fullmitjanes.getRange(1,columnafinal,1,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(1,columnafinal,1,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(1,columnafinal,1,1).setFontWeight("bold");
    fullmitjanes.getRange(1,columnafinal,2,3).setVerticalAlignment("middle");
    fullmitjanes.getRange(1,columnafinal,2,3).setHorizontalAlignment("center");    
    fullmitjanes.getRange(2,columnafinal,1,3).setBackground("#cc4125");
    fullmitjanes.getRange(2,columnafinal,1,3).setFontColor("white");
    fullmitjanes.getRange(2,columnafinal,1,3).setFontWeight("bold");
    fullmitjanes.getRange(2,columnafinal,1,1).setValue(max);
    fullmitjanes.getRange(2,columnafinal+2,1,1).setValue(maxpunt);
    var celmaxpunt = fullmitjanes.getRange(2,columnafinal+2,1,1).getA1Notation();
    
    fullmitjanes.getRange(3,columnafinal,1,1).setValue(coav);
    fullmitjanes.getRange(3,columnafinal+1,1,1).setValue(auto);
    fullmitjanes.getRange(3,columnafinal+2,1,1).setValue(prof);
    fullmitjanes.getRange(3,columnafinal,1,3).setWrap(false);
    fullmitjanes.getRange(3,columnafinal,1,3).setBackground("#fff2cc");
    fullmitjanes.getRange(3,columnafinal,1,3).setBorder(true,true,true,true,false,false);
    
    fullmitjanes.getRange(1,columnafinal+3,1,3).merge();
    fullmitjanes.setColumnWidth(columnafinal+3, 37);
    fullmitjanes.setColumnWidth(columnafinal+4, 37);
    fullmitjanes.setColumnWidth(columnafinal+5, 37);
    fullmitjanes.getRange(1,columnafinal+3,1,1).setValue(npon);
    fullmitjanes.getRange(1,columnafinal+3,1,1).setWrap(true);
    fullmitjanes.getRange(1,columnafinal+3,1,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(1,columnafinal+3,1,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(1,columnafinal+3,1,1).setFontWeight("bold");
    fullmitjanes.getRange(3,columnafinal+3,1,1).setValue(coav);
    fullmitjanes.getRange(3,columnafinal+4,1,1).setValue(auto);
    fullmitjanes.getRange(3,columnafinal+5,1,1).setValue(prof);
    fullmitjanes.getRange(3,columnafinal+3,1,3).setWrap(false);
    fullmitjanes.getRange(3,columnafinal+3,1,3).setBackground("#fff2cc");
    fullmitjanes.getRange(3,columnafinal+3,1,3).setBorder(true,true,true,true,false,false);
    
    var rangsuma = fullmitjanes.getRange(2,6,1,columnafinal-8).getA1Notation();
    fullmitjanes.getRange(2,columnafinal+3,1,3).merge();
    fullmitjanes.getRange(2,columnafinal+3,1,1).setFormula("=SUM("+rangsuma+")");
    fullmitjanes.getRange(2,columnafinal+3,1,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(2,columnafinal+3,1,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(2,columnafinal+3,1,1).setFontWeight("bold");
    fullmitjanes.getRange(2,columnafinal+3,1,1).setNumberFormat("0%");
    if (fullmitjanes.getRange(2,columnafinal+3,1,1).getCell(1,1).getValue()!=1){
      fullmitjanes.getRange(2,columnafinal+3,1,1).setBackground("#e06666");
    }else{
      fullmitjanes.getRange(2,columnafinal+3,1,1).setBackground("#93c47d");
    };
    
    //Afegim columna per mitjana ponderada de les 3 notes (co, auto i profe)
    fullmitjanes.setColumnWidth(columnafinal+6, 37);
    fullmitjanes.setColumnWidth(columnafinal+7, 37);
    fullmitjanes.setColumnWidth(columnafinal+8, 37);
    fullmitjanes.getRange(1,columnafinal+6,1,3).merge();
    fullmitjanes.getRange(1,columnafinal+6,1,1).setWrap(true);
    fullmitjanes.getRange(1,columnafinal+6,1,1).setValue(nf);
    fullmitjanes.getRange(1,columnafinal+6,1,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(1,columnafinal+6,1,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(1,columnafinal+6,1,1).setFontWeight("bold");
    fullmitjanes.getRange(1,columnafinal+6,1,1).setHorizontalAlignment("center");
    fullmitjanes.getRange(1,columnafinal+6,1,1).setVerticalAlignment("middle");
    fullmitjanes.getRange(3,columnafinal+6,1,3).setFontWeight("bold");
    fullmitjanes.getRange(2,columnafinal+6,1,1).setValue(coav);
    fullmitjanes.getRange(2,columnafinal+7,1,1).setValue(auto);
    fullmitjanes.getRange(2,columnafinal+8,1,1).setValue(prof);
    fullmitjanes.getRange(3,columnafinal+6,1,1).setValue(per_co);
    fullmitjanes.getRange(3,columnafinal+7,1,1).setValue(per_auto);
    fullmitjanes.getRange(3,columnafinal+8,1,1).setValue(per_prof);
    fullmitjanes.getRange(2,columnafinal+6,2,3).setVerticalAlignment("middle");
    fullmitjanes.getRange(2,columnafinal+6,2,3).setHorizontalAlignment("center"); 
    fullmitjanes.getRange(2,columnafinal+6,1,3).setBackground("#f9cb9c");
    fullmitjanes.getRange(3,columnafinal+6,1,3).setBackground("#cc4125");
    fullmitjanes.getRange(3,columnafinal+6,1,3).setFontColor("white");
    fullmitjanes.getRange(3,columnafinal+6,1,3).setNumberFormat("0%");
    
    //Afegim columna per comentaris del profe
    fullmitjanes.setColumnWidth(columnafinal+9, 245);
    fullmitjanes.getRange(1,columnafinal+9,3,1).merge();
    fullmitjanes.getRange(1,columnafinal+9,2,1).setWrap(true);
    fullmitjanes.getRange(1,columnafinal+9,2,1).setValue(cp);
    fullmitjanes.getRange(1,columnafinal+9,2,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(1,columnafinal+9,2,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(1,columnafinal+9,2,1).setFontWeight("bold");
    fullmitjanes.getRange(1,columnafinal+9,3,1).setHorizontalAlignment("center");
    fullmitjanes.getRange(1,columnafinal+9,3,1).setVerticalAlignment("middle");
    
    //Afegim columna per  comentaris dels companys
    fullmitjanes.setColumnWidth(columnafinal+10, 245);
    fullmitjanes.getRange(1,columnafinal+10,3,1).merge();
    fullmitjanes.getRange(1,columnafinal+10,2,1).setWrap(true);
    fullmitjanes.getRange(1,columnafinal+10,2,1).setValue(cc);
    fullmitjanes.getRange(1,columnafinal+10,2,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(1,columnafinal+10,2,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(1,columnafinal+10,2,1).setFontWeight("bold");
    fullmitjanes.getRange(1,columnafinal+10,3,1).setHorizontalAlignment("center");
    fullmitjanes.getRange(1,columnafinal+10,3,1).setVerticalAlignment("middle");  
    
    //Afegim columna per comentaris del propi alumne
    fullmitjanes.setColumnWidth(columnafinal+11, 245);
    fullmitjanes.getRange(1,columnafinal+11,3,1).merge();
    fullmitjanes.getRange(1,columnafinal+11,2,1).setWrap(true);
    fullmitjanes.getRange(1,columnafinal+11,2,1).setValue(ca);
    fullmitjanes.getRange(1,columnafinal+11,2,1).setBorder(true,true,true,true,true,true);
    fullmitjanes.getRange(1,columnafinal+11,2,1).setBackground("#f9cb9c");
    fullmitjanes.getRange(1,columnafinal+11,2,1).setFontWeight("bold");
    fullmitjanes.getRange(1,columnafinal+11,3,1).setHorizontalAlignment("center");
    fullmitjanes.getRange(1,columnafinal+11,3,1).setVerticalAlignment("middle");  
    
    var rang_resultats = fullmitjanes.getRange(4,1,nombrealumnes,3*(aspectes)+11);
    var num_alum=1;
    var valormaxim=0; //Busco el grau que val més punts
    for (var j=2;j<rangrubrica.getNumColumns();j++){
      if (valormaxim < rangrubrica.getCell(2,j).getValues()){
        valormaxim = rangrubrica.getCell(2,j).getValue();
      };      
    };
    
    for (var i=0;i<nombrealumnes;i++){
      rang_resultats.getCell(i+1,1).setValue(num_alum);
      rang_resultats.getCell(i+1,2).setValue(r_al[i+1][0]);
      num_alum = num_alum+1;
      fullmitjanes.getRange(i+4,1,1,1).setBorder(true,true,true,true,true,true);
      var resp=1;
      var resp_auto=1;
      var resp_profe=1;
      for (j=0;j<aspectes;j++){
        var l=3*j;
        if (j==0){
          l=2;       
        };
        fullmitjanes.getRange(i+4,l,1,3).setBorder(true,true,true,true,false,false);
      };
      
      var rangminim="";
      var rangminim_auto="";
      var rangminim_profe="";
      var pesos="";
      for (var h=1;h<aspectes-1;h++){
        rangminim = rangminim + fullmitjanes.getRange(i+4,3+3*h,1,1).getA1Notation() + ";" ;
        pesos = pesos + fullmitjanes.getRange(2,3+3*h,1,1).getA1Notation() + ";" ;
      };
      rangminim = rangminim.substring(0, rangminim.length-1);
      for (h=1;h<aspectes-1;h++){
        rangminim_auto = rangminim_auto + fullmitjanes.getRange(i+4,3+3*h+1,1,1).getA1Notation() + ";" ;
      };
      rangminim_auto = rangminim_auto.substring(0, rangminim_auto.length-1);
      for (h=1;h<aspectes-1;h++){
        rangminim_profe = rangminim_profe + fullmitjanes.getRange(i+4,3+3*h+2,1,1).getA1Notation() + ";" ;
      };
      rangminim_profe = rangminim_profe.substring(0, rangminim_profe.length-1);
      pesos = pesos.substring(0, pesos.length-1);
      
      var vm=10/valormaxim;
      vm = vm + "";
      
      if (canviloc===1){
        rang_resultats.getCell(i+1,3*aspectes).setFormula("round(min("+rangminim+")*"+vm+",2)*"+celmaxpunt+"/10");
        rang_resultats.getCell(i+1,3*aspectes+3).setFormula("round(SUMPRODUCT({"+ rangminim +"};{"+ pesos+"})*"+vm+",2)*"+celmaxpunt+"/10");
      }else{
        vm = vm.replace(".", ",");
        rang_resultats.getCell(i+1,3*aspectes).setFormula("round(min("+rangminim+")*"+vm+";2)*"+celmaxpunt+"/10");
        rang_resultats.getCell(i+1,3*aspectes+3).setFormula("round(SUMPRODUCT({"+ rangminim +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
      };
      
      fullmitjanes.getRange(i+4,3*aspectes,1,3).setBorder(true,true,true,true,false,false);
      fullmitjanes.getRange(i+4,3*aspectes+3,1,3).setBorder(true,true,true,true,false,false);
      fullmitjanes.getRange(i+4,3*aspectes+6,1,3).setBorder(true,true,true,true,true,true);
      fullmitjanes.getRange(i+4,3*aspectes+9,1,3).setBorder(true,true,true,true,true,true);
      
      if (canviloc===1){
        rang_resultats.getCell(i+1,3*aspectes+1).setFormula("round(min("+rangminim_auto+")*"+vm+";2)*"+celmaxpunt+"/10");
        rang_resultats.getCell(i+1,3*aspectes+4).setFormula("round(SUMPRODUCT({"+ rangminim_auto +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
      }else{
        vm = vm.replace(".", ",");
        rang_resultats.getCell(i+1,3*aspectes+1).setFormula("round(min("+rangminim_auto+")*"+vm+";2)*"+celmaxpunt+"/10");
        rang_resultats.getCell(i+1,3*aspectes+4).setFormula("round(SUMPRODUCT({"+ rangminim_auto +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
      };          
      
      if (canviloc===1){
        rang_resultats.getCell(i+1,3*aspectes+2).setFormula("round(min("+rangminim_profe+")*"+vm+";2)*"+celmaxpunt+"/10");
        rang_resultats.getCell(i+1,3*aspectes+5).setFormula("round(SUMPRODUCT({"+ rangminim_profe +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
      }else{
        vm = vm.replace(".", ",");
        rang_resultats.getCell(i+1,3*aspectes+2).setFormula("round(min("+rangminim_profe+")*"+vm+";2)*"+celmaxpunt+"/10");
        rang_resultats.getCell(i+1,3*aspectes+5).setFormula("round(SUMPRODUCT({"+ rangminim_profe +"};{"+ pesos+"})*"+vm+";2)*"+celmaxpunt+"/10");
      };          
      
      //Posem la fórmula de la nota final (ponderant les 3)
      fullmitjanes.getRange(i+4,columnafinal+6,1,3).merge();
      var rangpesosfinal = fullmitjanes.getRange(3,columnafinal+6,1,3).getA1Notation();
      var notesfinals = fullmitjanes.getRange(i+4,columnafinal+3,1,3).getA1Notation();
      fullmitjanes.getRange(i+4,columnafinal+6,1,3).setFormula("round(SUMPRODUCT("+ rangpesosfinal +";"+ notesfinals+");2)");  
      fullmitjanes.getRange(i+4,columnafinal+6,1,3).setHorizontalAlignment("center");
      fullmitjanes.getRange(i+4,columnafinal+6,1,3).setVerticalAlignment("middle"); 
    };
    //Deso al ScripDb que hm creat el full
    documentProperties.setProperty('Formulari',"1");
    documentProperties.setProperty('Proces',"1");
    documentProperties.setProperty('Mail',"1");
    
    //Canviar el menú, treient Crear formulari i posant el que correspongui
    onOpen();
  };
};
