let incidentTimer = {};

Office.onReady(info => {
    if(info.host === Office.HostType.Excel) {
        $("#initialisation").on("click", () => tryCatch(initialisation));
        $("#AddIncident").on("click", () => tryCatch(addIncident));
        
    }
})

/** the function for the initialisation of all timers */
function initialisation() {
    return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Save");
        var usedRange = sheet.getUsedRange(true);
        usedRange.load("rowCount");
        return context.sync().then(function () {
            var lastRow = usedRange.rowCount;
            var range = sheet.getRange("A2:A" + lastRow);
            range.load("values"); // Charger les valeurs
            return context.sync().then(function () {
                var values = range.values;
                for (var i = 0; i < values.length; i++) {
                    let key = values[i][0];
                    incidentTimer[key] = new Date();
                }
                console.log("Tout est bien initialisé.");
            });
        });
    });
}

function addIncident() {
    return Excel.run(function (context) {
        var creaSuivi = context.workbook.worksheets.getItem("Création suivi");
        var suivi = context.workbook.worksheets.getItem("Suivi");

        // Chargement du nombre de lignes utilisées dans 'Suivi' pour déterminer où ajouter la nouvelle valeur
        var usedRange = suivi.getUsedRange();
        usedRange.load("rowCount");

        // Chargement de la valeur à ajouter depuis 'Création Suivi'
        var cellCreaSuivi = creaSuivi.getCell(2, 0); // Cela récupère la cellule en A4 (l'indexation commence à 0)
        cellCreaSuivi.load("values");

        return context.sync().then(function () {
            var lastRow = usedRange.rowCount; // Dernière ligne utilisée dans 'Suivi'
            var cellValue = cellCreaSuivi.values; // Valeur à ajouter

            // Vérification si la valeur est déjà présente dans 'Suivi'
            var rangeSuivi = suivi.getRange("A2:A" + lastRow); // Plage actuelle des valeurs
            let foundRange = rangeSuivi.findOrNullObject(cellValue[0][0], {
                completeMatch: true,
                matchCase: false,
                searchDirection: Excel.SearchDirection.forward
            });

            return context.sync().then(function () {
                if (foundRange.isNullObject) {
                    // Si la valeur n'est pas trouvée, elle est ajoutée à la fin de la liste des IDs dans 'Suivi'
                    var firstEmptyCell = suivi.getCell(lastRow, 0); // La première cellule vide après la dernière ligne utilisée
                    console.log(cellValue[0][0]);
                    firstEmptyCell.values = [[cellValue[0][0]]];
                    return context.sync();
                }
                else {
                    console.log("l'incident entrée existe déjà");
                }
                // Si la valeur est trouvée, aucune action n'est nécessaire. Vous pouvez ajouter un else ici si nécessaire.
            });
        });
    });
}

/** Default helper for invoking an action and handling errors. */
function tryCatch(callback) {
    Promise.resolve()
        .then(callback)
        .catch(function (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            console.error(error);
        });
}