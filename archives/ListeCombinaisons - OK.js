"use strict";
function main(workbook) {
    let sheet = workbook.getActiveWorksheet();
    // Nettoyer les anciennes données 
    sheet.getRange("F:Z").clear();
    // Récupérer les valeurs de départ, arrivée et via
    let start = sheet.getRange("B1").getValue();
    let end = sheet.getRange("B2").getValue();
    let via = sheet.getRange("B3").getValue().split(";");
    // Créer les connexions et les variantes
    let { connections, variants } = createConnectionsAndVariants(sheet);
    // Générer toutes les combinaisons possibles
    let allCombinations = generateCombinations(start, end, via, variants);
    // Afficher les combinaisons dans la colonne E
    for (let i = 0; i < allCombinations.length; i++) {
        let combination = allCombinations[i].join(' -> ');
        sheet.getRange("F" + (i + 2)).setValue(combination); // Afficher à partir de la ligne 2
    }
}
function createConnectionsAndVariants(sheet) {
    let rawConnections = sheet.getRange("A5:C340").getValues();
    let connections = new Map();
    let variants = new Map();
    for (let row of rawConnections) {
        let from = row[0];
        let to = row[1];
        let time = row[2];
        if (!connections.has(from)) {
            connections.set(from, new Map());
        }
        connections.get(from).set(to, time);
        // Créer les variantes pour la gare 'from'
        let baseFrom = from.split('_')[0];
        if (!variants.has(baseFrom)) {
            variants.set(baseFrom, []);
        }
        if (!variants.get(baseFrom).includes(from)) {
            variants.get(baseFrom).push(from);
        }
        // Créer les variantes pour la gare 'to'
        let baseTo = to.split('_')[0];
        if (!variants.has(baseTo)) {
            variants.set(baseTo, []);
        }
        if (!variants.get(baseTo).includes(to)) {
            variants.get(baseTo).push(to);
        }
    }
    return { connections, variants };
}
function generateCombinations(start, end, via, variants) {
    // Filtrer les gares intermédiaires pour éliminer les chaînes vides
    let filteredVia = via.filter(v => v.trim() !== "");
    // Générer les permutations des gares intermédiaires
    let viaPermutations = permute(filteredVia);
    // Ajouter start au début et end à la fin de chaque permutation
    let routes = viaPermutations.map(permutation => [start, ...permutation, end]);
    // Étendre chaque route pour inclure toutes les variantes possibles
    let allCombinations = routes.flatMap(route => expandPermutations(route, variants));
    return allCombinations;
}
// Fonction pour obtenir toutes les variantes d'une gare
function getAllVariants(gare, variants) {
    // Si la gare a un suffixe (_), renvoyer uniquement [gare]
    if (gare.includes('_')) {
        return [gare];
    }
    // Sinon, renvoyer toutes les variantes associées
    return variants.get(gare) || [];
}
// Fonction pour générer toutes les permutations possibles d'un tableau de chaînes
function permute(arr) {
    if (arr.length === 0)
        return [[]];
    if (arr.length === 1)
        return [[arr[0]]];
    let result = [];
    for (let i = 0; i < arr.length; i++) {
        let rest = [...arr.slice(0, i), ...arr.slice(i + 1)];
        let restPermutations = permute(rest);
        for (let perm of restPermutations) {
            result.push([arr[i], ...perm]);
        }
    }
    return result;
}
function expandPermutations(permutation, variants) {
    if (permutation.length === 0)
        return [[]];
    let first = getAllVariants(permutation[0], variants);
    let restExpanded = expandPermutations(permutation.slice(1), variants);
    let result = [];
    for (let f of first) {
        for (let r of restExpanded) {
            result.push([f, ...r]);
        }
    }
    return result;
}
