"use strict";
function main(workbook) {
    let sheet = workbook.getActiveWorksheet();
    // 🔹 Vider la colonne F avant d'écrire
    sheet.getRange("F2:F100").clear();
    // 🔹 Récupération des gares de départ, via et arrivée
    let start = sheet.getRange("B1").getValue();
    let viaRaw = sheet.getRange("B3").getValue();
    let end = sheet.getRange("B2").getValue();
    // 🔹 Transformation des gares "Via" en tableau (séparateur ";")
    let vias = viaRaw ? viaRaw.split(";").map(s => s.trim()).filter(s => s) : [];
    // 🔹 Récupération des connexions de A5:D55
    let data = sheet.getRange("H1:M108").getValues();
    // 🔹 Création du graphe
    let graph = {};
    for (let row of data) {
        let station = row[0];
        let connections = row.slice(1).filter(g => g);
        graph[station] = connections;
    }
    // 🔹 Trouver le trajet optimal en testant toutes les permutations des "Via"
    let bestPath = findOptimalPath(graph, start, vias, end);
    // 🔹 Affichage du résultat en colonne F
    if (bestPath) {
        let resultRange = sheet.getRange(`F2:F${bestPath.length + 1}`);
        resultRange.setValues(bestPath.map(station => [station]));
    }
    else {
        sheet.getRange("F2").setValue("Aucun chemin trouvé");
    }
}
// 🔹 Trouve l'ordre optimal des "Via" et le chemin le plus court
function findOptimalPath(graph, start, vias, end) {
    let bestPath = null;
    let bestLength = Infinity;
    let permutations = generatePermutations(vias);
    for (let permutedVias of permutations) {
        let fullPath = [start, ...permutedVias, end];
        let path = findCompletePath(graph, fullPath);
        if (path && path.length < bestLength) {
            bestPath = path;
            bestLength = path.length;
        }
    }
    return bestPath;
}
// 🔹 Génère toutes les permutations possibles des gares "Via"
function generatePermutations(arr) {
    if (arr.length === 0)
        return [[]];
    let result = [];
    for (let i = 0; i < arr.length; i++) {
        let rest = arr.slice(0, i).concat(arr.slice(i + 1));
        for (let perm of generatePermutations(rest)) {
            result.push([arr[i], ...perm]);
        }
    }
    return result;
}
// 🔹 Trouve le chemin complet en suivant un ordre précis
function findCompletePath(graph, stations) {
    let path = [];
    for (let i = 0; i < stations.length - 1; i++) {
        let segment = findShortestPath(graph, stations[i], stations[i + 1]);
        if (!segment)
            return null;
        path = [...path, ...segment.slice(i > 0 ? 1 : 0)];
    }
    return path;
}
// 🔹 Algorithme de Dijkstra pour trouver le chemin le plus court entre 2 gares
function findShortestPath(graph, start, end) {
    let queue = [{ station: start, path: [start] }];
    let visited = new Set();
    while (queue.length > 0) {
        let { station, path } = queue.shift();
        if (station === end)
            return path;
        if (!visited.has(station)) {
            visited.add(station);
            for (let neighbor of graph[station] || []) {
                queue.push({ station: neighbor, path: [...path, neighbor] });
            }
        }
    }
    return null;
}
