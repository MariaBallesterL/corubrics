/*
 * Esborra ScriptDb
 */

function esborradB (){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Formulari',"");
  documentProperties.setProperty('Mail',"");
  documentProperties.setProperty('Proces',"");
  documentProperties.setProperty('Mail2',"");
  documentProperties.setProperty('Formid',"");
  documentProperties.setProperty('Formul',"");
  documentProperties.setProperty('Formnom',"");
  documentProperties.setProperty('cursid',"");
  documentProperties.setProperty('tasca_av_id',"");
  documentProperties.setProperty('tasca_co_id',"");
  documentProperties.setProperty('tasca_pf_id',"");
  documentProperties.setProperty('titol_tasca',"");  
  documentProperties.setProperty('desc_tasca',"");
  
};

/*
* Obtinc la data  l'hora actual
*/
function Dataactual(){
  var avui = new Date();
  var dd = avui.getDate();
  var mm = avui.getMonth()+1; //Gener és el mes 0!
  var yyyy = avui.getFullYear();
  if(dd<10){
    dd='0'+dd
  };
  if(mm<10){
    mm='0'+mm
  }; 
  var hh = avui.getHours();
  var min = avui.getMinutes();
  var ss = avui.getSeconds();
  if(hh<10){
    hh='0'+hh
  };
  if(min<10){
    min='0'+min
  }; 
  if(ss<10){
    ss='0'+ss
  };
  var avui = dd+'/'+mm+'/'+yyyy+' '+hh+":"+min+':'+ss;
  return avui;
};

//Omple una matriu amb totes les respostes dels alumnes, però amb puntuacións enlloc de descripcions
function Omplir_matriu_respostes(matriu_res,res,col){
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
  
  var mat_resp=[];
  for (var i=0;i<res;i++) {  //Defineixo la matriu com a bidimensional i la poso tot en blanc
     mat_resp[i] = [];
     for (var j=0; j<col;j++){
       mat_resp[i][j]=0;
    };
  }; 
  
  
  
  //Omplo la matriu rúbrica amb les dades de la rúbrica
    for (var i=0;i<rangrubrica.getNumRows();i++) {  //Defineixo la matriu com a bidimensional i la poso tot en blanc
      mat_rubrica[i] = [];
      for (j=0; j<rangrubrica.getNumColumns();j++){
        mat_rubrica[i][j]="";
      };
    }; 
    
    for (i=0;i<mat_rubrica.length;i++){ //Ompla la matriu  rubrica amb la rúbrica
      for (j=0;j<mat_rubrica[0].length;j++){
        mat_rubrica[i][j]=rangrubrica.getCell(i+1,j+1).getValue()
      };
    };
  
  for (i=1;i<res+1;i++){ //Omplo la matriu resposta amb les respostes dels alumnes
    for (j=1; j<col+1;j++){
      var valorcela = matriu_res[i-1][j-1];
      if (j===1){
        mat_resp[i-1][j-1]=valorcela; //Alumne a qui es valora
      }else{
        for (k=0;k<rangrubrica.getNumColumns()-2;k++){
          var aspecterubrica = mat_rubrica[0][k+1]+": "+mat_rubrica[j][k+1]; //Ageixo el títol com al formulari
          if (valorcela===aspecterubrica) {
            mat_resp[i-1][j-1] = mat_rubrica[1][k+1];
          };
        };
      };
    };
   };
  return mat_resp;
};


//A partir de la matriu de respostes, crea una matriu alumnes amb
// la informació de [alumne avaluat, nombre de respostes, suma puntuació aspecte 1, suma puntuació aspecte 2,...]
function Trobar_alumnes (mat_resp){
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
  
  //Evitem errors si no hi ha respostes en el formulari
  if (mat_resp.length===0){
    var asp = rangrubrica.getNumRows()-1;
  }else{
    var asp = mat_resp[0].length;
  };

  var alumnes = [];
  for (var i=0;i<nombrealumnes;i++) {  //Defineixo la matriu com a bidimensional i l'omplo amb zeros
    alumnes[i] = [];
    for (j=0; j<asp+1;j++){
       alumnes[i][j]=0;
    };
  }; 
  for (i=0;i<alumnes.length;i++){
    alumnes[i][0]=rangalumnes.getCell(i+2,1).getValue().toString(); //Omplim el primer camp de alumnes amb tots els noms del full Alumnes
  };
  for (i=0;i<mat_resp.length;i++){
    for (k=0;k<alumnes.length;k++){
      if (mat_resp[i][0]===alumnes[k][0]){
        alumnes[k][1]=alumnes[k][1]+1;
        for (z=2;z<alumnes[k].length;z++){
          alumnes[k][z]=alumnes[k][z]+parseInt(mat_resp[i][z-1]);
        };
      };
    };
  };
  for (i=0;i<alumnes.length;i++){
    for (j=2;j<alumnes[0].length;j++){
      if (alumnes[i][1]!=0){      
        alumnes[i][j]=alumnes[i][j]/alumnes[i][1];
        alumnes[i][j] = Math.round(alumnes[i][j] * 100) / 100;
      };
    };
  };
  return alumnes;                  
};


//GAFE A partir d'una matriu on a cada fila hi ha qui ha respost i a qui ha evaluat,
//torna una llista de persones que han duplicat respostes i una amb el número de resposte duplicada.
function Trobar_duplicat (nom_resp){
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
  
  var trdup = new Object;
  trdup.nom_duplicat=[];
  trdup.resp_eliminar=[];
  var index=0;
  var cont=0;
  var trobat=0;
  var mesduplicats=0;
  var anterior=0;
  for (i=0;i<nom_resp.length;i++){
    trobat=0;
    mesduplicats=0;
    for (k=i+1; k<nom_resp.length;k++){
      if (nom_resp[i][0]===nom_resp[k][0] && nom_resp[i][1]===nom_resp[k][1]){
        if (mesduplicats===0){
          trdup.resp_eliminar[index]=i;
        }else{
          trdup.resp_eliminar[index]=anterior;
          nom_resp[k][0]=index; //Esborro el contingut per evitar que entrades triplicades es tornin a comptar per eliminar
        };
        anterior=k;
        mesduplicats=1;
        index=index+1;
        for (j=0; j<cont; j++){
          if (trdup.nom_duplicat[j]===nom_resp[i][0]){
            trobat=1;           
          };
        };
        if (trobat===0){
          trdup.nom_duplicat[cont]=nom_resp[i][0];
          cont=cont+1;
        };
      };
    };
  };
  return trdup;
};
  
  
function variables(){
  /**
  * Declaració de variables globals
  */
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
};

function sleep(milliseconds) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}
  
