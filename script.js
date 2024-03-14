Office.onReady(info => {
    // Vérifie que l'add-in est utilisé dans Excel avant d'ajouter des événements
    if (info.host === Office.HostType.Excel) {
        // Initialise votre Add-In avec ses fonctionnalités ici.
        document.getElementById("writeDataBtn").onclick = writeDataToExcel;
    }
});


function writeDataToExcel() {
    Excel.run(function (context) {
        // Récupère la plage de cellules actuellement sélectionnée et la charge.
        const range = context.workbook.getSelectedRange();
        
        // Écrit dans la plage de cellules sélectionnée.
        range.values = [["Hello, Excel!"]];
        
        // Exécute toutes les commandes en attente dans le contexte de votre Add-In.
        return context.sync();
    }).catch(function (error) {
        console.error("Error: " + error);
        if (error.debugInfo) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
