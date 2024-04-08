Office.onReady((info) => {
  console.log("Office.js is now ready in ${info.host} host.");
  $("#initialisation").on("click", () => tryCatch(initialisation));
  $("#AddIncident").on("click", () => tryCatch(addIncident));
  $("#AllRetake").on("click", () => tryCatch(allRetake));
  $("#Retake").on("click", () => tryCatch(retake));
  $("#Sollicitation").on("click", () => tryCatch(sollicitation()));
  $("#RetourGA").on("click", () => tryCatch(retourGA()));
  $("#DemandeValCom").on("click", () => tryCatch(demandeValCom()));
  $("#RetourValCom").on("click", () => tryCatch(retourValCom()));
  $("#FinPre").on("click", () => tryCatch(finPre()));
});

let incidentTimer = {};
let timerForSave = {};

// This function is for refresh the select when the number of incident is modify.
function refreshList(ids) {
  var select = document.getElementById("IdIncident");

  select.innerHTML = "";
  ids.forEach(function(option) {
    var el = document.createElement("option");
    el.textContent = option;
    el.value = option;
    select.appendChild(el);
  });
}

function addCellSave(cellpos, values, key) {
  return Excel.run(function(context) {
    var save = context.workbook.worksheets.getItem("Save");
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");
    save.load("values");
    return context.sync().then(function() {
      var lastRow = usedRangeSave.rowCount;
      var cell = save.getCell(cellpos[0] + 1, cellpos[1]);
      var cellNNI = save.getCell(cellpos[0] + 1, 2);
      cell.load("values");
      return context.sync().then(function() {
        cell.values = [[values + (cell.values[0][0] - timerForSave[key])]];
        cellNNI.values = [[""]];
        timerForSave[key] = values;
        return context.sync();
      });
    });
  });
}

/** the function for the initialisation of all timers */
function initialisation() {
  // We get the NNI of the users
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;

  if (nniValue) {
    return Excel.run(function(context) {
      // We get the sheet Save for doing operation on it.
      var save = context.workbook.worksheets.getItem("Save");
      var usedRange = save.getUsedRange(true);
      usedRange.load("rowCount");

      return context.sync().then(function() {
        var lastRow = usedRange.rowCount;
        var range = save.getRange("A2:A" + lastRow);
        range.load("values"); // Charger les valeurs
        var rangeNNI = save.getRange("C2:C" + lastRow);
        rangeNNI.load("values");

        return context.sync().then(function() {
          var values = range.values;
          var valuesNNI = rangeNNI.values;
          for (var i = 0; i < values.length; i++) {
            if (!valuesNNI[i][0] && values[i][0]) {
              let key = values[i][0];
              timerForSave[key] = 0;
              incidentTimer[key] = new Date();
              save.getCell(i + 1, 2).values = [[nniValue]];
            }
          }
          refreshList(Object.keys(incidentTimer));
          console.log("Tout est bien initialisé.");
        });
      });
    });
  } else {
    console.log("Entrer un NNI");
  }
}

/**
 * SaveTimer need to by modify, is here for save the timer and for all other agents when he's take in charge the incident
 */
function allRetake() {
  return Excel.run(function(context) {
    var save = context.workbook.worksheets.getItem("Save");
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");
    save.load("values");
    return context.sync().then(function() {
      var lastRow = usedRangeSave.rowCount;
      var rangeSave = save.getRange("A2:A" + lastRow);
      // Ici vous devez charger les valeurs pour pouvoir les utiliser après context.sync()
      rangeSave.load("values");

      return context.sync().then(function() {
        var values = rangeSave.values;
        let promises = [];
        for (let key in incidentTimer) {
          for (let i = 0; i < values.length; i++) {
            if (values[i][0] === key) {
              // Ici vous pouvez accéder à la cellule et mettre à jour les valeurs
              var actualTime = new Date();

              promises.push(
                promises,
                addCellSave([i, 1], (actualTime.getTime() - incidentTimer[key].getTime()) / 1000 / 60, key)
              );
            }
          }
        }
        return Promise.all(promises).then(() => {
          for (let key in incidentTimer) delete incidentTimer[key];
          refreshList(Object.keys(incidentTimer));
          console.log("Liste d'incident vide");
          return context.sync();
        });
      });
    });
  });
}

/**
 * SaveTimer need to by modify, is here for save the timer and for all other agents when he's take in charge the incident
 */
function retake() {
  return Excel.run(function(context) {
    var save = context.workbook.worksheets.getItem("Save");
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");
    save.load("values");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    return context.sync().then(function() {
      var lastRow = usedRangeSave.rowCount;
      var rangeSave = save.getRange("A2:A" + lastRow);
      // Ici vous devez charger les valeurs pour pouvoir les utiliser après context.sync()
      rangeSave.load("values");

      return context.sync().then(function() {
        var values = rangeSave.values;
        let promises = [];
        for (let i = 0; i < values.length; i++) {
          if (values[i][0] === id) {
            // Ici vous pouvez accéder à la cellule et mettre à jour les valeurs
            var cell = save.getCell(i + 1, 1); // i + 1 car les index dans Excel commencent à 1, et non 0
            var actualTime = new Date();
            promises.push(
              promises,
              addCellSave([i, 1], (actualTime.getTime() - incidentTimer[id].getTime()) / 1000 / 60, id)
            );
          }
        }
        refreshList(Object.keys(incidentTimer));
        return Promise.all(promises).then(() => {
          delete incidentTimer[id];
          refreshList(Object.keys(incidentTimer));
          console.log("Incident retiré");
          return context.sync();
        });
      });
    });
  });
}

// For adding an incident
function addIncident() {
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;

  const app = document.getElementById("Application");
  const appValue = app.value;

  const type_inc = document.getElementById("type_inc");
  const type_incValue = type_inc.value;

  if (nniValue && appValue) {
    return Excel.run(function(context) {
      var suivi = context.workbook.worksheets.getItem("Suivi");
      var save = context.workbook.worksheets.getItem("Save");

      // Load all row for check where we are going to put the value
      var usedRangeSuivi = suivi.getUsedRange();
      usedRangeSuivi.load("rowCount");

      // Load all row for check where we are going to put the value
      var usedRangeSave = save.getUsedRange();
      usedRangeSave.load("rowCount");

      return context.sync().then(function() {
        var lastRowSuivi = usedRangeSuivi.rowCount; // The last row used in 'Suivi'
        var lastRowSave = usedRangeSave.rowCount; // The last row used in 'Save'

        return context.sync().then(function() {
          var domaineCell = suivi.getCell(lastRowSuivi, 1);
          var idCellSuivi = suivi.getCell(lastRowSuivi, 0);
          var idCellSave = save.getCell(lastRowSave, 0);
          var cellSaveTimer = save.getCell(lastRowSave, 1);
          var idCellNNI = save.getCell(lastRowSave, 2);
          if (type_incValue == "MA") {
            var id = "MA-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
          } else if (type_incValue == "MCP") {
            var id = "MCP-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
          } else {
            var id = "MA-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
            domaineCell.values = [[appValue]];
            idCellSuivi.values = [[id]];
            idCellSave.values = [[id]];
            idCellNNI.values = [[nniValue]];
            cellSaveTimer.values = [[0]];

            // Adding the new timer in the variable reserved to it
            incidentTimer[id] = new Date();
            timerForSave[id] = 0;

            domaineCell = suivi.getCell(lastRowSuivi + 1, 1);
            idCellSuivi = suivi.getCell(lastRowSuivi + 1, 0);
            idCellSave = save.getCell(lastRowSave + 1, 0);
            cellSaveTimer = save.getCell(lastRowSave + 1, 1);
            idCellNNI = save.getCell(lastRowSave + 1, 2);

            var id = "MCP-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
          }

          domaineCell.values = [[appValue]];
          idCellSuivi.values = [[id]];
          idCellSave.values = [[id]];
          idCellNNI.values = [[nniValue]];
          cellSaveTimer.values = [[0]];

          // Adding the new timer in the variable reserved to it
          incidentTimer[id] = new Date();
          timerForSave[id] = 0;
          refreshList(Object.keys(incidentTimer));
          return context.sync();
        });
      });
    });
  } else {
    console.log("Entrer le NNI et l'application concerné avant de vouloir ajouter un incident !");
  }
}

/** Default helper for invoking an action and handling errors. */
function tryCatch(callback) {
  Promise.resolve()
    .then(callback)
    .catch(function(error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    });
}

// Set the past time between "prise en charge" and "sollicitation" in the correct cell
function sollicitation() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");
      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");
      return context.sync().then(function() {
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        sollicitationCell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (sollicitationCell.values[0][0] === null || sollicitationCell.values[0][0] === "") {
            var actualTime = new Date();
            sollicitationCell.values = [
              [(actualTime.getTime() - incidentTimer[id]) / 1000 / 60 + saveTimer.values[0][0]]
            ];
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// Set the past time between "sollicitation" and "retour du GA" in the correct cell
function retourGA() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        sollicitationCell.load("values");
        retourGACell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (retourGACell.values[0][0] === null || retourGACell.values[0][0] === "") {
            var actualTime = new Date();
            retourGACell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] -
                  sollicitationCell.values[0][0]
              ]
            ];
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// Set the past time between "retour du GA" and "demande de validation de comm." in the correct cell
function demandeValCom() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var demandeValComCell = suivi.getCell(searchCell.rowIndex, 4);
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        demandeValComCell.load("values");
        sollicitationCell.load("values");
        retourGACell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (demandeValComCell.values[0][0] === null || demandeValComCell.values[0][0] === "") {
            var actualTime = new Date();
            demandeValComCell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] -
                  (sollicitationCell.values[0][0] + retourGACell.values[0][0])
              ]
            ];
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// Set the past time between the "demande de validation" and "retour sur la validation"
function retourValCom() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var dureeTtlCell = suivi.getCell(searchCell.rowIndex, 6);
        var validationCell = suivi.getCell(searchCell.rowIndex, 5);
        var demandeValComCell = suivi.getCell(searchCell.rowIndex, 4);
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        demandeValComCell.load("values");
        retourGACell.load("values");
        sollicitationCell.load("values");
        validationCell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (validationCell.values[0][0] === null || validationCell.values[0][0] === "") {
            var actualTime = new Date();
            console.log(saveTimer.values[0][0]);
            validationCell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] -
                  (sollicitationCell.values[0][0] + retourGACell.values[0][0] + demandeValComCell.values[0][0])
              ]
            ];
            dureeTtlCell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] +
                  (sollicitationCell.values[0][0] + retourGACell.values[0][0] + demandeValComCell.values[0][0])
              ]
            ];
            save.getRange("A" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("B" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("C" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            delete incidentTimer[id];
            refreshList(Object.keys(incidentTimer));
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// This for the case where the incident doesn't need more investigation on it.
function finPre() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var dureeTtlCell = suivi.getCell(searchCell.rowIndex, 6);
        var validationCell = suivi.getCell(searchCell.rowIndex, 5);
        var demandeValComCell = suivi.getCell(searchCell.rowIndex, 4);
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        demandeValComCell.load("values");
        retourGACell.load("values");
        sollicitationCell.load("values");
        validationCell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (demandeValComCell.values[0][0] === null || demandeValComCell.values[0][0] === "") {
            var actualTime = new Date();

            demandeValComCell.values[0][0] = 0;
            validationCell.values[0][0] = 0;

            dureeTtlCell.values = [
              [
                (sollicitationCell.values[0][0] +
                  retourGACell.values[0][0] +
                  demandeValComCell.values[0][0] +
                  actualTime.getTime() -
                  incidentTimer[id]) /
                  1000 /
                  60
              ]
            ];

            save.getRange("A" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("B" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("C" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            delete incidentTimer[id];
            refreshList(Object.keys(incidentTimer));
          } else {
            console.log(
              "Cette option n'est plus disponible pour cette incident, ou vous n'avez pas rempli les condition pour confirmer cette action."
            );
          }

          return context.sync();
        });
      });
    });
  });
}
