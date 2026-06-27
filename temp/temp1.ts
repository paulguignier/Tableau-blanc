Log.startTimer(`${this.name}.load()`);

// Vérifie si la table à charger existe déjà.
if (this.size > 0) {
    if (!erase) return;
    this.clear();
}

// Charge les parcours s'ils ne sont pas encore chargés.
Paths.load(); 

// Récupère les lignes de la base de données.
const rows = WorkbookServices.getRows({
    sheetName: this.SHEET,
    tableName: this.TABLE
});
if (!rows) {
    Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
    return;
}

// Parcourt les lignes (hors en-tête).
let excelRow: number = 0;
try {

    for (const [rowIndex, row] of rows) {

        // Vérifie si la ligne est vide.
        if (row.length === 0) continue;

        // Calcule le numéro de ligne Excel.
        excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

        // Récupère les champs.
        const params = TableSerializer.loadRow(
            row,
            this.DATABASE_COLUMNS,
            this.COLUMN_DEFINITIONS,
            [{ column: "key" }]
        ) as TrainParams;
        
        // Crée l'objet et l'insère dans la base de données.
        const train = this.create(params);
    } 

} catch (e) {
    throw new Error(`${this.name}.load (ligne ${excelRow}) : ${e}`);
} 

Log.timer(`${this.name}.load()`);
}