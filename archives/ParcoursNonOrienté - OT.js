"use strict";
function main(workbook) {
    let sheet = workbook.getActiveWorksheet();
    // 🔹 Vider la colonne F avant d'écrire
    sheet.getRange("F2:F100").clear(); // Efface jusqu'à 100 lignes pour éviter les anciens résultats
    // 🔹 Récupération des gares de départ, arrivée et via
    let start = sheet.getRange("B1").getValue();
    let via = sheet.getRange("B3").getValue();
    let end = sheet.getRange("B2").getValue();
    // 🔹 Récupération des connexions de A5:D55
    let data = sheet.getRange("H1:K92").getValues();
    // 🔹 Création du graphe
    let graph = {};
    for (let row of data) {
        let station = row[0]; // Colonne A = gare principale
        let connections = row.slice(1).filter(g => g); // Colonnes B-D = connexions non vides
        graph[station] = connections;
    }
    // 🔹 Recherche du trajet avec passage obligatoire par "Via"
    let path = null;
    if (via) {
        let firstLeg = findStops(graph, start, via);
        let secondLeg = findStops(graph, via, end);
        if (firstLeg && secondLeg) {
            path = [...firstLeg, ...secondLeg.slice(1)]; // Fusionner les chemins sans dupliquer "Via"
        }
    }
    else {
        path = findStops(graph, start, end);
    }
    // 🔹 Affichage du résultat en colonne F (à partir de F2)
    if (path) {
        let resultRange = sheet.getRange(`F2:F${path.length + 1}`);
        resultRange.setValues(path.map(station => [station]));
    }
    else {
        sheet.getRange("F2").setValue("Aucun chemin trouvé");
    }
}
function findStops(graph, start, end) {
    let queue = [[start, [start]]];
    while (queue.length > 0) {
        let [current, path] = queue.shift();
        if (current === end) {
            return path;
        }
        for (let neighbor of graph[current] || []) {
            if (!path.includes(neighbor)) {
                queue.push([neighbor, [...path, neighbor]]);
            }
        }
    }
    return null;
}
