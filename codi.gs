function onInstall(e) {
  onOpen(e)
};

/**
 * Afegeix el menú al full
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  if (e && e.authMode === ScriptApp.AuthMode.NONE){
    switch(Session.getActiveUserLocale()){
        case "ca":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Activar el complement CoRubrics','activaCoRubrics')
        .addToUi()
        break;
        
        case "es":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Activar el complemento CoRubrics','activaCoRubrics')
        .addToUi()
         break;

        case "eu":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('CoRubrics gehigarria aktibatu','activaCoRubrics')
        .addToUi()
         break;
        
      case "fr":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Activez CoRubrics','activaCoRubrics')
        .addToUi()
        break;

      default:
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Enable CoRubrics','activaCoRubrics')
        .addToUi()
    };
  } else {
    var properties = PropertiesService.getDocumentProperties();
    var importacio = properties.getProperty('Importacio');
    if (importacio===null){
      importacio= "0";
      properties.setProperty('Importacio', "0");
    };
    var idioma = properties.getProperty('Idioma');
    if (idioma===null){
      idioma= Session.getActiveUserLocale();
      properties.setProperty('Idioma', idioma);
    }  
    switch(idioma){
      case "ca":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addItem('Crear la plantilla CoRubrics','creaCoRubrics')
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Canviar idioma')
                      .addItem('Español', 'espanol')
                      .addItem('Euskara', 'euskara')
                      .addItem('English', 'english')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          var record = properties.getProperty('Formulari');
          if (record != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Crear el formulari','creaFormulari')
            .addItem('Importar alumnes i professors de Google Classroom','impalClasroom')
            .addSeparator()
            .addItem('Crea el full de resultats en blanc','fullrespostes')
            .addToUi()
          }else{
            var record1 = properties.getProperty('Mail');
            if (record1 != "1"){
              SpreadsheetApp.getUi()
              .createAddonMenu()
              .addItem('Enviar el formulari als alumnes','enviaFormulari')
              .addItem('Obtenir l\'enllaç del formulari','enllaFormulari')
              .addItem('Publicar l\'enllaç del formulari a Classroom com un anunci','classFormulari')
              .addSeparator()
              .addItem('Processar les respostes del formulari','procesFormulari')
              .addItem('Tornar a crear el formulari','creanouFormulari')
              .addItem('Reiniciar el procés', 'reinici')
              .addToUi()
            }else{
              var record2 = properties.getProperty('Proces');
              if (record2 != "1"){
                SpreadsheetApp.getUi()
                .createAddonMenu()
                .addItem('Processar les respostes del formulari','procesFormulari')
                .addSeparator()
                .addItem('Enviar el formulari als alumnes','enviaFormulari')
                .addItem('Obtenir l\'enllaç del formulari','enllaFormulari')
                .addItem('Publicar l\'enllaç del formulari a Classroom com un anunci','classFormulari')
                .addItem('Canvia el formulari enllaçat','nouformulari')
                .addItem('Tornar a crear el formulari','creanouFormulari')
                .addItem('Reiniciar el procés','reinici')
                .addToUi()
              }else{
                var record3 = properties.getProperty('Mail2');
                if (record3 != "1"){
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Enviar resultats als alumnes', 'enviament')
                  .addItem('Publica la nota final a Classroom', 'nota_classroom')
                  .addSeparator()
                  .addItem('Reprocessar les respostes del formulari','procesFormulari')
                  .addItem('Enviar el formulari als alumnes','enviaFormulari')
                  .addItem('Obtenir l\'enllaç del formulari','enllaFormulari')
                  .addItem('Canvia el formulari enllaçat','nouformulari')
                  .addItem('Reiniciar el procés','reinici')
                  .addToUi()
                }else{
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Reiniciar el procés','reinici')
                  .addSeparator()
                  .addItem('Enviar resultats als alumnes', 'enviament')
                  .addItem('Publicar/Actualitzar la nota final a Classroom', 'nota_classroom')
                  .addItem('Obtenir l\'enllaç del formulari','enllaFormulari')
                  .addItem('Canvia el formulari enllaçat','nouformulari')
                  .addItem('Reprocessar les respostes del formulari','procesFormulari')   
                  .addToUi()
                };
              }; 
            };
          };
        };
        break;
      case "es":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addItem('Crear la plantilla CoRubrics','creaCoRubrics')
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Cambiar idioma')
                      .addItem('Català', 'catala')
                      .addItem('Euskara', 'euskara')
                      .addItem('English', 'english')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          var record = properties.getProperty('Formulari');
          if (record != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Crear el formulario','creaFormulari')
            .addItem('Importar alumnos y profesores de Google Classroom','impalClasroom') 
            .addSeparator()
            .addItem('Crea la hoja de respuesta vacía','fullrespostes')
            .addToUi()
          }else{
            var record1 = properties.getProperty('Mail');
            if (record1 != "1"){
              SpreadsheetApp.getUi()
              .createAddonMenu()
              .addItem('Enviar el formulario a los alumnos','enviaFormulari')
              .addItem('Obtener el enlace del formulario','enllaFormulari')
              .addItem('Publicar el enlace del formulario en Classroom como un anuncio','classFormulari')
              .addSeparator()
              .addItem('Procesar las respuestas del formulario','procesFormulari')
              .addItem('Volver a crear el formulario','creanouFormulari')
              .addItem('Reiniciar el proceso', 'reinici')
              .addToUi()
            }else{
              var record2 = properties.getProperty('Proces');
              if (record2 != "1"){
                SpreadsheetApp.getUi()
                .createAddonMenu()
                .addItem('Procesar las respuestas del formulario','procesFormulari')
                .addSeparator()
                .addItem('Enviar el formulario a los alumnos','enviaFormulari')
                .addItem('Obtener el enlace del formulario','enllaFormulari')
                .addItem('Publicar el enlace del formulario en Classroom como un anuncio','classFormulari')
                .addItem('Cambiar el formulario enlazado','nouformulari')
                .addItem('Volver a crear el formulario','creanouFormulari')
                .addItem('Reiniciar el proceso','reinici')
                .addToUi()
              }else{
                var record3 = properties.getProperty('Mail2');
                if (record3 != "1"){
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Enviar los resultados a los alumnos', 'enviament')
                  .addItem('Publicar la nota final en Classroom', 'nota_classroom')
                  .addSeparator()
                  .addItem('Reprocesar las respuestas del formulario','procesFormulari')
                  .addItem('Enviar el formulario a los alumnos','enviaFormulari')
                  .addItem('Obtener el enlace del formulario','enllaFormulari')
                  .addItem('Cambiar el formulario enlazado','nouformulari')
                  .addItem('Reiniciar el proceso','reinici')
                  .addToUi()
                }else{
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Reiniciar el proceso','reinici')
                  .addSeparator()
                  .addItem('Enviar los resultados a los alumnos', 'enviament')
                  .addItem('Publicar/Actualizar la nota final en Classroom', 'nota_classroom')
                  .addItem('Obtener el enlace del formulario','enllaFormulari')
                  .addItem('Cambiar el formulario enlazado','nouformulari')
                  .addItem('Reprocesar las respuestas del formulario','procesFormulari')  
                  .addToUi()
                };
              }; 
            };
          };
        };
        break;
      case "eu":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addItem('CoRubrics txantiloia sortu','creaCoRubrics')
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Hizkuntza aldatu')
                      .addItem('Català', 'catala')
                      .addItem('Español', 'espanol')
                      .addItem('English', 'english')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          var record = properties.getProperty('Formulari');
          if (record != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Errubrikarekin ebaluatzeko  Inprimakia sortu','creaFormulari')
            .addItem('Classroometik Ikasle eta irakasleak inportatu','impalClasroom')
            .addSeparator()
            .addItem('Crea la hoja de respuesta vacía','fullrespostes')
            .addToUi()
          }else{
            var record1 = properties.getProperty('Mail');
            if (record1 != "1"){
              SpreadsheetApp.getUi()
              .createAddonMenu()
              .addItem('Inprimakia ikasleei bidali','enviaFormulari')
              .addItem('Inprimakiaren helbidea lortu','enllaFormulari')
              .addItem('Inprimakiaren lotura iragarpen gisa Classromen argitaratu','classFormulari')
              .addSeparator()
              .addItem('Inprimakiaren erantzunak prozesatu','procesFormulari')
              .addItem('Inprimakia berriro sortu','creanouFormulari')
              .addItem('Prozesua berriro hasi', 'reinici')
              .addToUi()
            }else{
              var record2 = properties.getProperty('Proces');
              if (record2 != "1"){
                SpreadsheetApp.getUi()
                .createAddonMenu()
                .addItem('Inprimakiaren erantzunak prozesatu','procesFormulari')
                .addSeparator()
                .addItem('Inprimakia ikasleei bidali','enviaFormulari')
                .addItem('Inprimakiaren helbidea lortu','enllaFormulari')
                .addItem('Inprimakiaren lotura iragarpen gisa Classromen argitaratu','classFormulari')
                .addItem('Beste inprimaki batera lotu','nouformulari')
                .addItem('Inprimakia berriro sortu','creanouFormulari')
                .addItem('Prozesua berriro hasi','reinici')
                .addToUi()
              }else{
                var record3 = properties.getProperty('Mail2');
                if (record3 != "1"){
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Emaitzak ikasleei bidali', 'enviament')
                  .addItem('Azken emaitza Classroomen argitaratu', 'nota_classroom')
                  .addSeparator()
                  .addItem('Inprimakiaren erantzunak prozesatu','procesFormulari')
                  .addItem('Inprimakia ikasleei bidali','enviaFormulari')
                  .addItem('Inprimakiaren helbidea lortu','enllaFormulari')
                  .addItem('Beste inprimaki batera lotu','nouformulari')
                  .addItem('Prozesua berriro hasi','reinici')
                  .addToUi()
                }else{
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Prozesua berriro hasi','reinici')
                  .addSeparator()
                  .addItem('Emaitzak ikasleei bidali', 'enviament')
                  .addItem('Eguneratu azken emaitza Classroomen', 'nota_classroom')
                  .addItem('Inprimakiaren helbidea lortu','enllaFormulari')
                  .addItem('Beste inprimaki batera lotu','nouformulari')
                  .addItem('Inprimakiaren erantzunak prozesatu','procesFormulari')   
                  .addToUi()
                };
              }; 
            };
          };
        };
        break;
      case "fr":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addItem('Créez un gabarit de CoRubrics','creaCoRubrics')
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Changez de langue')
                      .addItem('Català', 'catala')
                      .addItem('Español', 'espanol')
                      .addItem('Euskara', 'euskara')
                      .addItem('English', 'english'))
          .addToUi()
        }else{ 
          var record = properties.getProperty('Formulari');
          if (record != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Créez le formulaire afin d\'évaluer avec la grille','creaFormulari')
            .addItem('Importez élèves et enseignants de Google Classroom','impalClasroom')
            .addSeparator()
            .addItem('Crée la feuille de réponse vide','fullrespostes')
            .addToUi()
          }else{
            var record1 = properties.getProperty('Mail');
            if (record1 != "1"){
              SpreadsheetApp.getUi()
              .createAddonMenu()
              .addItem('Envoyez le formulaire aux élèves','enviaFormulari')
              .addItem('Obtenez le lien au formulairee','enllaFormulari')
              .addItem('Publiez le lien au formulaire dans Classroom comme annonce','classFormulari')
              .addSeparator()
              .addItem('Traitez les réponses aux formulaire','procesFormulari')
              .addItem('Recréez le formulaire','creanouFormulari')
              .addItem('Recommencez du début ', 'reinici')
              .addToUi()
            }else{
              var record2 = properties.getProperty('Proces');
              if (record2 != "1"){
                SpreadsheetApp.getUi()
                .createAddonMenu()
                .addItem('Traitez les réponses aux formulaire','procesFormulari')
                .addSeparator()
                .addItem('Envoyez le formulaire aux élèves','enviaFormulari')
                .addItem('Obtenez le lien au formulaire','enllaFormulari')
                .addItem('Publiez le lien au formulaire dans Classroom comme annonce','classFormulari')
                .addItem('Changez le formulaire','nouformulari')
                .addItem('Recréez le formulaire','creanouFormulari')
                .addItem('Recommencez du début','reinici')
                .addToUi()
              }else{
                var record3 = properties.getProperty('Mail2');
                if (record3 != "1"){
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Envoyez les résultats aux élèves', 'enviament')
                  .addItem('Publiez les notes dans Classroom', 'nota_classroom')
                  .addSeparator()
                  .addItem('Retraitez les réponses aux formulaires','procesFormulari')
                  .addItem('Envoyez le formulaire aux élèves','enviaFormulari')
                  .addItem('Obtenez le lien au formulaire','enllaFormulari')
                  .addItem('Changez le formulaire','nouformulari')
                  .addItem('Recommencez du début','reinici')
                  .addToUi()
                }else{
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Recommencez du début ','reinici')
                  .addSeparator()
                  .addItem('Renvoyez les résultats aux élèves', 'enviament')
                  .addItem('Publiez/Mettre à jour les notes finales dans Classroom', 'nota_classroom')
                  .addItem('Obtenez le lien au formulaire','enllaFormulari')
                  .addItem('Changez le formulaire','nouformulari')
                  .addItem('Retraitez les réponses aux formulaires','procesFormulari')   
                  .addToUi()
                };
              }; 
            };
          };
        };
        break;
      default:
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addItem('Create template of CoRubrics','creaCoRubrics')
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Change language')
                      .addItem('Català', 'catala')
                      .addItem('Español', 'espanol')
                      .addItem('Euskara', 'euskara')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          var record = properties.getProperty('Formulari');
          if (record != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Create the form to evaluate with the rubric','creaFormulari')
            .addItem('Import students and teachers from Google Classroom','impalClasroom')
            .addSeparator()
            .addItem('Creates the empty answer sheet','fullrespostes')
            .addToUi()
          }else{
            var record1 = properties.getProperty('Mail');
            if (record1 != "1"){
              SpreadsheetApp.getUi()
              .createAddonMenu()
              .addItem('Send form to students','enviaFormulari')
              .addItem('Get the form link','enllaFormulari')
              .addItem('Publish the form link in Classroom like an annoucement','classFormulari')
              .addSeparator()
              .addItem('Process form responses','procesFormulari')
              .addItem('Recreate the form','creanouFormulari')
              .addItem('Restart the process', 'reinici')
              .addToUi()
            }else{
              var record2 = properties.getProperty('Proces');
              if (record2 != "1"){
                SpreadsheetApp.getUi()
                .createAddonMenu()
                .addItem('Process form responses','procesFormulari')
                .addSeparator()
                .addItem('Resend form to students','enviaFormulari')
                .addItem('Get the form link','enllaFormulari')
                .addItem('Publish the form link in Classroom like an annoucement','classFormulari')
                .addItem('Change linked form','nouformulari')
                .addItem('Recreate the form','creanouFormulari')                
                .addItem('Restart the process','reinici')
                .addToUi()
              }else{
                var record3 = properties.getProperty('Mail2');
                if (record3 != "1"){
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Send results to students', 'enviament')
                  .addItem('Publish grades in Classroom', 'nota_classroom')
                  .addSeparator()
                  .addItem('Reprocess form responses','procesFormulari')
                  .addItem('Send form to students','enviaFormulari')
                  .addItem('Get the form link','enllaFormulari')
                  .addItem('Change linked form','nouformulari')
                  .addItem('Restart the process','reinici')
                  .addToUi()
                }else{
                  SpreadsheetApp.getUi()
                  .createAddonMenu()
                  .addItem('Restart the process','reinici')
                  .addSeparator()
                  .addItem('Send results to students', 'enviament')
                  .addItem('Publish/Update final grade in Classroom', 'nota_classroom')
                  .addItem('Get the form link','enllaFormulari')
                  .addItem('Change linked form','nouformulari')
                  .addItem('Reprocess form responses','procesFormulari')      
                  .addSeparator()
                  .addToUi()
                };
              }; 
            };
          };
        };            
    };
  };
};
