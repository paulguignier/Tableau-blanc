"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.loadParams = loadParams;
exports.loadConnections = loadConnections;
exports.findShortestPath = findShortestPath;
exports.generateCombinations = generateCombinations;
/* Classeur principal. */
var WORKBOOK;
/* Liste des paramètres. */
const PARAM = {
    loaded: false,
    evenLetter: "",
    oddLetter: "",
    evenFigure: 0,
    oddFigure: 0,
    maxConnectionNumber: 0,
    turnaroundTime: 0,
    wTrainsRegex: new RegExp(""),
    trains4FiguresRegex: new RegExp(""),
    days: new Map(),
};
/* Liste des éléments sauvegardés en cache. */
const CACHE = {
    daysCombinations: new Map(),
};
/* Liste des trains plannifiés sur un ou plusieurs jours, avec les mêmes horaires. Ils sont référencés pour chaque jour de circulation */
const TRAINS = new Map();
/* Liste des trains associés à un jour donné et leurs réutilisations */
const REUSES = new Map();
/* Liste gares et leurs coordonnées. */
const STATIONS = new Map();
/* Liste des connexions entre les gares, incluant le temps de trajet et l'information sur le besoin de changement de sens. */
const CONNECTIONS = new Map();
function main(workbook) {
    WORKBOOK = workbook;
    const sheet = WORKBOOK.getActiveWorksheet();
    // Lire les paramètres
    loadParams();
    loadConnections();
    loadTrains("", "141243");
    loadStops();
    printTrains("Test", "Trains1");
    TRAINS.get("141243_2").findPath();
    printStops("Test", "Stops1", "A10");
    // findPath2(TRAINS.get("147490_2"));
    console.log(TRAINS.get("141243_2"));
    // const allCombinations = generateCombinations("MPU", "ETP", "".split(";"));
    // console.log(allCombinations);
    // const shortestPath = findShortestPath(allCombinations);
    //     console.log(shortestPath);
    return;
}
function tests() {
    loadParams();
    console.log(isWTrain("146490"));
    console.log(isWTrain("569907"));
    console.log(isWTrain("147490"));
    console.log(renameWith4Figures("146490"));
    console.log(renameWith4Figures("569907"));
    console.log(renameWith4Figures("147490"));
    console.log(daysToNumbers("146"));
    loadTrains("", "147490");
    loadStops();
    printTrains("Test", "Test");
    console.log(TRAINS.get("147490_2"));
    console.log(TRAINS.get("147490_2").number4Figures());
}
/**
 * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
 * @param allCombinations - La liste de toutes les combinaisons de parcours à évaluer.
 * @returns Un objet contenant le chemin le plus court et sa distance totale, ou null si aucun chemin n'est trouvé.
 */
function findShortestPath(allCombinations) {
    let shortestPath = null;
    for (let combination of allCombinations) {
        // Calculer le chemin complet et la distance totale pour la combinaison actuelle
        let { path, totalDistance } = calculateCompletePath(combination);
        if (path.length > 0) {
            if (shortestPath === null || totalDistance < shortestPath.totalDistance) {
                shortestPath = { path, totalDistance };
            }
        }
    }
    return shortestPath;
}
/**
 * Calcule le chemin complet et la distance totale pour une combinaison de gares.
 * @param combination - La liste ordonnée des gares à parcourir.
 * @returns Un objet contenant le chemin complet et la distance totale.
 */
function calculateCompletePath(combination) {
    let completePath = [];
    let totalDistance = 0;
    for (let i = 0; i < combination.length - 1; i++) {
        let segmentStart = combination[i];
        let segmentEnd = combination[i + 1];
        // Trouver le chemin le plus court pour le tronçon actuel
        let segmentPath = dijkstra(segmentStart, segmentEnd);
        if (segmentPath.length === 0) {
            // Si aucun chemin n'est trouvé pour ce tronçon, retourner un chemin vide
            return { path: [], totalDistance: 0 };
        }
        // Calculer la distance pour ce tronçon
        let segmentDistance = calculatePathTime(segmentPath);
        // Ajouter la distance du tronçon à la distance totale
        totalDistance += segmentDistance;
        // Ajouter le chemin du tronçon au chemin complet
        // Éviter de dupliquer les gares intermédiaires
        if (completePath.length > 0) {
            segmentPath.shift(); // Retirer la première gare pour éviter la duplication
        }
        completePath.push(...segmentPath);
    }
    return { path: completePath, totalDistance };
}
/**
 * Calcule le temps total pour un chemin donné en tenant compte des temps de trajet
 * et des éventuels temps de changement de sens.
 * @param path - La liste ordonnée des gares constituant le chemin.
 * @returns Le temps total du chemin, incluant les temps de trajet et de changement de sens.
 */
function calculatePathTime(path) {
    var _a;
    let totalTime = 0;
    for (let i = 0; i < path.length - 1; i++) {
        let from = path[i];
        let to = path[i + 1];
        let connection = (_a = CONNECTIONS.get(from)) === null || _a === void 0 ? void 0 : _a.get(to);
        if (connection) {
            totalTime += connection.time;
            // Ajouter le temps de changement de sens sauf pour le premier segment
            if (i > 0 && connection.withTurnaround) {
                totalTime += PARAM.turnaroundTime;
            }
        }
    }
    return totalTime;
}
/**
 * Cherche le chemin le plus court entre le départ et l'arrivée
 * en appliquant Dijkstra.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @returns Le chemin le plus court.
 */
function dijkstra(start, end) {
    const distances = new Map();
    const previousNodes = new Map();
    const unvisited = new Set(CONNECTIONS.keys());
    let path = [];
    // Initialisation des distances
    for (let node of unvisited) {
        distances.set(node, Infinity);
        previousNodes.set(node, null);
    }
    distances.set(start, 0);
    while (unvisited.size > 0) {
        let currentNode = Array.from(unvisited).reduce((minNode, node) => distances.get(node) < distances.get(minNode) ? node : minNode);
        if (distances.get(currentNode) === Infinity)
            break; // Aucun chemin
        unvisited.delete(currentNode);
        // Examiner les voisins avec les nouveaux attributs
        for (let [neighbor, connexion] of CONNECTIONS.get(currentNode) || []) {
            let additionalTime = connexion.time;
            if (connexion.withTurnaround && currentNode !== start) { // Si un changement de sens est nécessaire, ajouter du temps
                additionalTime += PARAM.turnaroundTime;
            }
            let newDist = distances.get(currentNode) + additionalTime;
            if (newDist < distances.get(neighbor)) {
                distances.set(neighbor, newDist);
                previousNodes.set(neighbor, currentNode);
            }
        }
    }
    // Retracer le chemin
    let step = end;
    while (step) {
        path.unshift(step);
        step = previousNodes.get(step);
    }
    // Si le chemin est valide
    return path[0] === start ? path : [];
}
/**
 * Génère toutes les combinaisons de routes possibles pour aller de start à end en passant par les gares intermédiaires via.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @param via - Les gares intermédiaires à passer par.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une combinaison de route possible.
 */
function generateCombinations(start, end, via) {
    // Filtrer les gares intermédiaires pour éliminer les chaînes vides
    const filteredVia = via.filter(v => v.trim() !== "");
    // Générer les permutations des gares intermédiaires
    const viaPermutations = permute(filteredVia);
    // Ajouter start au début et end à la fin de chaque permutation
    const routes = viaPermutations.map(permutation => [start, ...permutation, end]);
    // Étendre chaque route pour inclure toutes les variantes possibles
    const allCombinations = routes.flatMap((route) => expandPermutations(route));
    return allCombinations;
}
/**
 * Renvoie toutes les variantes possibles pour une gare.
 * Une variante correspond au sens de passage dans la gare : GARE_0 en pair, GARE_1 en impair.
 * Seules les gares de retournement permettent de passer d'une gare à l'autre
 * Si la gare a déjà un suffixe imposé (_), renvoie uniquement cette gare avec suffixe [gare].
 * Sinon, renvoie toutes les variantes associées.
 * @param gare - La gare dont on cherche les variantes.
 * @returns Un tableau contenant toutes les variantes possibles pour la gare.
 */
function getAllVariants(gare) {
    var _a;
    // Si la gare a un suffixe (_), renvoyer uniquement [gare]
    if (gare.includes('_')) {
        return [gare];
    }
    // Sinon, renvoyer toutes les variantes associées
    return ((_a = STATIONS.get(gare)) === null || _a === void 0 ? void 0 : _a.variants) || [];
}
/**
 * Génère toutes les permutations possibles d'un tableau de chaînes.
 * @param array - Le tableau de chaînes à permuter.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une permutation possible.
 */
function permute(array) {
    if (array.length === 0)
        return [[]];
    if (array.length === 1)
        return [[array[0]]];
    let result = [];
    for (let i = 0; i < array.length; i++) {
        let rest = [...array.slice(0, i), ...array.slice(i + 1)];
        let restPermutations = permute(rest);
        for (let perm of restPermutations) {
            result.push([array[i], ...perm]);
        }
    }
    return result;
}
/**
 * Étend une permutation de gares pour inclure toutes les variantes possibles.
 * @param permutation - La permutation de gares à étendre.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une permutation possible avec toutes les variantes.
 */
function expandPermutations(permutation) {
    if (permutation.length === 0)
        return [[]];
    const first = getAllVariants(permutation[0]);
    const restExpanded = expandPermutations(permutation.slice(1));
    let result = [];
    for (let f of first) {
        for (let r of restExpanded) {
            result.push([f, ...r]);
        }
    }
    return result;
}
/**
 * Renvoie la feuille de calcul Excel correspondant au nom donné.
 * Si la feuille n'existe pas, renvoie null si failOnError est à false, sinon lance une exception.
 * @param sheetName - Le nom de la feuille de calcul à chercher.
 * @param failOnError - Si true (par défaut), lance une exception si la feuille n'existe pas. Si false, renvoie null.
 * @returns La feuille de calcul Excel correspondant au nom donné, ou null si elle n'existe pas.
 */
function getSheetOrFail(sheetName, failOnError = true) {
    const sheet = WORKBOOK.getWorksheet(sheetName);
    if (!sheet) {
        const msg = `La feuille "${sheetName}" n'existe pas.`;
        if (failOnError)
            throw new Error(msg);
        console.log(msg);
        return null;
    }
    return sheet;
}
/**
 * Renvoie le tableau Excel correspondant au nom donné dans la feuille de calcul donnée.
 * Si le tableau n'existe pas, renvoie null si failOnError est à false, sinon lance une exception.
 * @param sheetName - Le nom de la feuille de calcul où chercher le tableau.
 * @param tableName - Le nom du tableau à chercher.
 * @param failOnError - Si true (par défaut), lance une exception si le tableau n'existe pas. Si false, renvoie null.
 * @returns Le tableau Excel correspondant au nom donné, ou null si il n'existe pas.
 */
function getTableOrFail(sheetName, tableName, failOnError = true) {
    const sheet = getSheetOrFail(sheetName, failOnError);
    const table = sheet.getTable(tableName);
    if (!table) {
        const msg = `Le tableau "${tableName}" n'existe pas dans la feuille "${sheetName}".`;
        if (failOnError)
            throw new Error(msg);
        console.log(msg);
        return null;
    }
    return table;
}
/**
 * Renvoie les données du tableau Excel correspondant au nom donné dans la feuille de calcul donnée.
 * Si le tableau n'existe pas, renvoie null si failOnError est à false, sinon lance une exception.
 * @param sheetName - Le nom de la feuille de calcul où chercher le tableau.
 * @param tableName - Le nom du tableau à chercher.
 * @param failOnError - Si true (par défaut), lance une exception si le tableau n'existe pas. Si false, renvoie null.
 * @returns Les données du tableau Excel correspondant au nom donné, ou null si il n'existe pas.
 */
function getDataFromTable(sheetName, tableName, failOnError = true) {
    const table = getTableOrFail(sheetName, tableName, failOnError);
    return table.getRange().getValues();
}
/**
 * Vérifie si l'adresse de cellule donnée est valide.
 * Si elle est valide, la renvoie telle quelle.
 * Si elle est invalide, lance une exception si failOnError est à true, sinon renvoie une chaîne vide.
 * @param cellName - L'adresse de cellule à vérifier.
 * @param failOnError - Si true (par défaut), lance une exception si l'adresse est invalide. Si false, renvoie une chaîne vide.
 * @returns L'adresse de cellule si elle est valide, une chaîne vide sinon.
 */
function checkCellName(cellName, failOnError = true) {
    // Convertir startCell en majuscules pour éviter les problèmes de casse
    cellName = cellName.toUpperCase();
    // Vérifier si cellName est une adresse de cellule valide
    if (!/^([A-Z]+)(\d+)$/.test(cellName)) {
        const msg = `L'adresse de départ ${cellName} n'est pas valide.`;
        if (failOnError)
            throw new Error(msg);
        console.log(msg);
        return "";
    }
    return cellName;
}
/**
 * Affiche un tableau avec en-têtes et données dans une feuille de calcul Excel.
 * Combine les en-têtes et les données fournies, puis les insère à partir
 * de la cellule de départ spécifiée. Efface le contenu existant de la plage
 * de cellules ciblée et supprime tout tableau existant avec le même nom avant
 * d'ajouter un nouveau tableau avec les données fournies.
 * @param headers - Les en-têtes du tableau.
 * @param data - Les données du tableau.
 * @param sheetName - Le nom de la feuille de calcul où afficher le tableau.
 * @param tableName - Le nom du tableau à afficher.
 * @param startCell - La cellule où commencer à afficher le tableau. Si non fourni, commence à "A1".
 * @param failOnError - Si true (par défaut), lance une exception si des erreurs surviennent. Si false, renvoie null.
 * @returns Le tableau Excel créé, ou null si une erreur survient.
 */
function printTable(headers, data, sheetName, tableName, startCell = "A1", failOnError = true) {
    // Combiner les en-têtes et les données
    const tableData = headers.concat(data);
    // Vérifier si les données sont non vides
    if (tableData.length === 0 || tableData[0].length === 0) {
        const msg = `Aucune donnée à insérer dans la table "${tableName}".`;
        if (failOnError)
            throw new Error(msg);
        console.log(msg);
        return;
    }
    // Vérifier si un tableau avec le même nom existe déjà et le supprimer si nécessaire
    const sheet = getSheetOrFail(sheetName, failOnError);
    const existingTable = sheet.getTables().find(table => table.getName() === tableName);
    if (existingTable) {
        existingTable.delete();
    }
    // Déterminer la plage où écrire les données
    const startRange = sheet.getRange(checkCellName(startCell));
    const writeRange = startRange.getResizedRange(tableData.length - 1, tableData[0].length - 1);
    // Effacer le contenu de la plage
    writeRange.clear(ExcelScript.ClearApplyTo.contents);
    // Écrire les données dans la plage
    writeRange.setValues(tableData);
    // Ajouter un nouveau tableau
    const table = sheet.addTable(writeRange.getAddress(), true);
    table.setName(tableName);
    console.log(`Le tableau "${tableName}" a été créé avec succès dans la feuille "${sheetName}".`);
    return table;
}
const PARAM_SHEET = "Param";
const PARAM_TABLE = "Paramètres";
const PARAM_LINE_EVEN_LETTER = 1;
const PARAM_LINE_ODD_LETTER = 2;
const PARAM_LINE_EVEN_DIRECTION = 3;
const PARAM_LINE_ODD_DIRECTION = 4;
const PARAM_LINE_MAX_CONNEXIONS_NUMBER = 5;
const PARAM_LINE_TURNAROUND_TIME = 6;
/**
 * Charge les paramètres du tableau "Paramètres" de la feuille "Param".
 * Si PARAM.loaded est true et que erase est false, ne fait rien.
 * Charge les paramètres dans l'objet PARAM et met à jour son champ "loaded".
 * Appelle les fonctions loadWTrainsRegex(), load4FiguresTrainsRegex() et loadDays().
 * @param erase - Si true, force le rechargement des paramètres. Si false (par défaut), ne recharge pas si déjà chargé.
 */
function loadParams(erase = false) {
    if (PARAM.loaded && !erase)
        return;
    const data = getDataFromTable(PARAM_SHEET, PARAM_TABLE);
    // Extraction des valeurs
    PARAM.evenLetter = data[PARAM_LINE_EVEN_LETTER][1];
    PARAM.oddLetter = data[PARAM_LINE_ODD_LETTER][1];
    PARAM.evenFigure = data[PARAM_LINE_EVEN_DIRECTION][1];
    PARAM.oddFigure = data[PARAM_LINE_ODD_DIRECTION][1];
    PARAM.maxConnectionNumber = data[PARAM_LINE_MAX_CONNEXIONS_NUMBER][1];
    PARAM.turnaroundTime = data[PARAM_LINE_TURNAROUND_TIME][1];
    loadWTrainsRegex();
    load4FiguresTrainsRegex();
    loadDays();
    PARAM.loaded = true;
}
const W_SHEET = "Param";
const W_TABLE = "W";
// Fonction pour charger les motifs W depuis la feuille
function loadWTrainsRegex() {
    const data = getDataFromTable(W_SHEET, W_TABLE);
    // Transformer chaque motif en regex partielle
    const regexParts = data
        .flat()
        .filter(v => typeof v === "string" && v.trim() !== "")
        .map(pattern => {
        return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
    });
    // Créer une regex globale combinée
    PARAM.wTrainsRegex = new RegExp(regexParts.join('|'));
}
// Fonction pour tester si un train est W (vide voyageur)
function isWTrain(trainNumber) {
    if (!PARAM.wTrainsRegex)
        loadWTrainsRegex(); // Charge si non encore fait
    return PARAM.wTrainsRegex.test(trainNumber);
}
const TRAINS_4FIGURES_SHEET = "Param";
const TRAINS_4FIGURES_TABLE = "LigneC4chiffres";
// Fonction pour charger les motifs des trains commerciaux que l'on nomme à 4 chiffres sur la ligne C
function load4FiguresTrainsRegex() {
    const data = getDataFromTable(TRAINS_4FIGURES_SHEET, TRAINS_4FIGURES_TABLE);
    // Transformer chaque motif en regex partielle
    const regexParts = data
        .flat()
        .filter(v => typeof v === "string" && v.trim() !== "")
        .map(pattern => {
        return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
    });
    // Créer une regex globale combinée
    // const fullRegex = new RegExp(regexParts.join('|'));
    PARAM.trains4FiguresRegex = new RegExp(regexParts.join('|'));
}
// Fonction transformer un numéro de train de 6 à 4 chiffres pour les trains commerciaux de la ligne C
function renameWith4Figures(trainNumber) {
    return PARAM.trains4FiguresRegex.test(trainNumber.substring(0, 6)) ? trainNumber.substring(2) : trainNumber;
}
/**
 * Classe jour qui défini les jours de la semaine individuellement.
 * ou les groupes de jours (JOB du lundi au vendredi, WE pour samedi et dimanche...).
 */
class Day {
    constructor(numbersString, fullName, abreviation) {
        this.number = isNaN(Number(numbersString)) ? 0 : Number(numbersString);
        this.numbersString = numbersString;
        this.fullName = fullName;
        this.abreviation = abreviation;
    }
}
const DAYS_SHEET = "Param";
const DAYS_TABLE = "Jours";
const DAYS_COL_NUMBERS = 2;
const DAYS_COL_FULL_NAME = 0;
const DAYS_COL_ABBREVIATION = 1;
/**
 * Charge les jours de la semaine à partir du tableau "Jours" de la feuille "Param".
 * Les jours sont stockés dans la structure PARAM.days, sous forme de map, avec
 * le nom complet et l'abréviation du jour comme clés, et leur numéro correspondant
 * comme valeur.
 */
function loadDays() {
    const data = getDataFromTable(DAYS_SHEET, DAYS_TABLE);
    for (let row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        // Extraction des valeurs
        let numbersString = row[DAYS_COL_NUMBERS];
        let fullName = row[DAYS_COL_FULL_NAME];
        let abreviation = row[DAYS_COL_ABBREVIATION];
        // Création du jour
        let day = new Day(numbersString, fullName, abreviation);
        PARAM.days.set(day.number ? day.number : day.numbersString, day);
    }
}
/**
 * Prend une chaîne de caractères en entrée et la convertit en un tableau de
 * nombres de jours de la semaine.
 *
 * Les noms de jours de la semaine peuvent être fournis en plein ou en abrégé.
 * Les noms en abrégé peuvent être des noms de jours en majuscule (par exemple,
 * "Lundi" ou "LUN") ou des noms de jours en minuscule (par exemple, "lundi" ou
 * "lun").
 *
 * Les noms de jours de la semaine sont remplacés par leurs numéros
 * correspondants. Les numéros sont extraits, les doublons sont supprimés et
 * les numéros sont triés.
 *
 * Pour améliorer les performances, les résultats sont stockés dans un cache.
 *
 * @param {string} input - La chaîne de caractères contenant les noms de jours
 *     de la semaine.
 * @returns {number[]} - Un tableau de nombres de jours de la semaine.
 */
function daysToNumbers(input) {
    input = input.toString().toLowerCase();
    // Vérifier si le résultat est déjà dans le cache
    if (CACHE.daysCombinations.has(input)) {
        return CACHE.daysCombinations.get(input);
    }
    // Remplacer chaque motif par ses numéros correspondants
    PARAM.days.forEach((day) => {
        // Crée une expression régulière combinée avec les trois motifs
        const regex = new RegExp(`${day.numbersString}|${day.abreviation.toLowerCase()}|${day.fullName.toLowerCase()}`, 'g');
        // Remplace toutes les occurrences des motifs par 'day.numbersString;'
        input = input.replace(regex, `${day.numbersString};`);
    });
    // Extraire les numéros, liminer les doublons et trier
    const numbers = new Set(input.split(';')
        .filter(num => num !== '')
        .map(num => parseInt(num)));
    const result = Array.from(numbers).sort((a, b) => a - b);
    // Stocker le résultat dans le cache
    CACHE.daysCombinations.set(input, result);
    return result;
}
/* Classe Parity qui permet de manipuler la parité */
class Parity {
    constructor(value) {
        // Si numéro de train avec double parité explicite (contient un '/')
        if (value.toString().includes('/')) {
            this.value = Parity.double;
            return;
        }
        switch (value) {
            case Parity.odd:
            case PARAM.oddLetter:
            case PARAM.oddFigure:
            case PARAM.oddFigure.toString():
                this.value = Parity.odd;
                break;
            case Parity.even:
            case PARAM.evenLetter:
            case PARAM.evenFigure:
            case PARAM.evenFigure.toString():
                this.value = Parity.even;
                break;
            case Parity.double:
                this.value = Parity.double;
                break;
            case Parity.undefined:
                this.value = Parity.undefined;
                break;
            default:
                switch (isNaN(parseInt(value)) || parseInt(value) <= 0 ? null : parseInt(value) % 2) {
                    case 0:
                        this.value = Parity.even;
                        break;
                    case 1:
                        this.value = Parity.odd;
                        break;
                    default:
                        this.value = Parity.undefined;
                        break;
                }
                break;
        }
    }
    /**
     * Inverse la parité actuelle.
     * Si la parité actuelle est paire, elle devient impaire, et inversement.
     * Si la parité actuelle est indéfinie, elle reste inchangée.
     */
    invert() {
        switch (this.value) {
            case Parity.even:
                this.value = Parity.odd;
                break;
            case Parity.odd:
                this.value = Parity.even;
                break;
            default:
                this.value = Parity.undefined;
                break;
        }
    }
    /**
     * Adapte le numéro du train en fonction de la parité demandée.
     * Si le numéro du train est pair, il est inchangé si la parité demandée est paire,
     * et incrémenté de 1 si la parité demandée est impaire.
     * Si le numéro du train est impair, il est décrémenté de 1 si la parité demandée est paire,
     * et inchangé si la parité demandée est impaire.
     * Si la parité demandée est indéfinie, le numéro du train est inchangé.
     * @param trainNumber Numéro du train, qui peut être un nombre ou une chaine de caractères
     * @returns Numéro du train adapté
     */
    adaptTrainNumber(trainNumber) {
        if (isNaN(parseInt(trainNumber))) {
            return 0;
        }
        let evenTrainNumber = parseInt(trainNumber);
        evenTrainNumber = evenTrainNumber - (evenTrainNumber % 2);
        switch (this.value) {
            case Parity.even:
                return evenTrainNumber;
            case Parity.odd:
                return evenTrainNumber + 1;
            case Parity.double:
            default:
                return trainNumber;
        }
    }
}
Parity.even = 0;
Parity.odd = 1;
Parity.double = -2;
Parity.undefined = -1;
/**
 * Classe Train qui définit un train, plannifié sur un ou plusieurs jours de la semaine,
 * ou sur plusieurs dates précises, avec les mêmes horaires.
 * Plusieurs trains associés à un jour donné et leurs réutilisations y font référence
 */
class Train {
    constructor(number, lineParity, days, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations) {
        this.number = parseInt(number);
        this.lineParity = new Parity(lineParity);
        this.days = days;
        this.trainsByDay = new Map();
        this.missionCode = missionCode;
        this.departureTime = departureTime;
        this.departureStation = departureStation;
        this.arrivalTime = arrivalTime;
        this.arrivalStation = arrivalStation;
        this.viaStations = viaStations ? viaStations.split(';') : [];
        this.stops = new Map();
        this.firstStop = departureStation;
        this.lastStop = arrivalStation;
        // Calcule Détermine si la gare d'arrivée change de parité, auquel cas le train a une double parité
        const departureStationObj = STATIONS.get(departureStation);
        const arrivalStationObj = STATIONS.get(arrivalStation);
        this.doubleParity = departureStationObj && arrivalStationObj && departureStationObj.reverseLineParity !== arrivalStationObj.reverseLineParity;
        // Détermine le sens principal pour la ligne C
        this.lineParity = new Parity(number);
        if (departureStationObj || departureStationObj.reverseLineParity) {
            this.lineParity.invert();
        }
    }
    /**
     * Retourne la clé du train qui est composée du numéro du train
     * suivi de la liste des jours de circulation ou de la première date de circulation.
     * @returns {string} Clé du train plannifié
     */
    get key() {
        return `${this.number}_${this.days.toString().split(';')[0]}`;
    }
    /**
     * Retourne le numéro du train avec changement de parité.
     * Si le train change de parité, le numéro est sous la forme "XX/Y".
     * Sinon, le numéro est sous la forme "XX".
     * @param {boolean} [withNumber4Figures=false] - Si true, le numéro est renommé en 4 chiffres pour les trains commerciaux de la ligne C
     * @returns {string} Le numéro du train avec changement de parité.
     */
    doubleParityNumber(withNumber4Figures = false) {
        const evenNumber = this.number - (this.number % 2);
        const doubleParityNumber = this.doubleParity ? evenNumber + "/" + ((evenNumber + 1) % 10) : this.number.toString();
        return withNumber4Figures ? renameWith4Figures(doubleParityNumber) : doubleParityNumber;
    }
    /**
     * Retourne le numéro du train utilisé par les opérateurs de la ligne C,
     * avec un numéro de train à 4 chiffres si le train est commercial
     * Si withDoubleParity est true, le numéro est renommé en prenant en compte le changement de parité.
     * @param {boolean} [withDoubleParity=false] - Si true, le numéro est renommé en prenant en compte le changement de parité.
     * @returns {string} Le numéro du train à 4 ou 6 chiffres utilisé par les opérateurs de la ligne C.
     */
    number4Figures(withDoubleParity = false) {
        return withDoubleParity ? this.doubleParityNumber(true) : renameWith4Figures(this.number.toString());
    }
    /**
     * Ajoute un arrêt au train.
     * Si le train est déjà passé par l'arrêt et que erase est faux, lance une erreur.
     * @param {Stop} stop - Arrêt à ajouter
     * @param {boolean} [erase=false] - Si true, remplace l'arrêt si il existe déjà
     * @throws {Error} Si le train est déjà passé par l'arrêt et que erase est faux
     */
    addStop(stop, erase = false) {
        if (this.stops.has(stop.key) && !erase) {
            const msg = `L'arrêt "${stop.key}" est déjà associé au train ${this.key}. Un même train ne peut pas revenir dans la même gare et avec le même sens.`;
            throw new Error(msg);
        }
        this.stops.set(stop.key, stop);
    }
    /**
     * Renvoie le numéro du train au départ de l'arrêt.
     * Si le train est terminus, renvoie le numéro de train de réutlisation.
     * Si l'arrêt est un rebroussement, renvoie la parité modifiée.
     * Sinon, renvoie le numéro du train en cours.
     * @returns {number} Numéro du train au départ
     */
    getStop(station, updateParity = false) {
        if (this.stops.has(station)) {
            return this.stops.get(station);
        }
        let [stationName, parity] = station.split("_");
        parity = parity !== null && parity !== void 0 ? parity : this.number % 2;
        const stop = this.stops.get(stationName + "_" + parity)
            || this.stops.get(stationName + "_" + (1 - parity))
            || this.stops.get(stationName)
            || null;
        if (updateParity && stop) {
            stop.parity = parity;
        }
        return stop;
    }
    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée du train,
     * puis génère la liste des arrêts calculés
     * @returns {void}
     */
    findPath() {
        var _a, _b;
        // Cherche toutes les combinaisons possibles de départ, d'arrivée et de passages via
        const allCombinations = generateCombinations(this.departureStation, this.arrivalStation, this.viaStations);
        // Trouve le chemin le plus court parmi toutes les combinaisons
        const shortestPath = findShortestPath(allCombinations);
        // Quitte la fonction si aucun chemin n'est trouvé
        if (shortestPath.path.length === 0) {
            return;
        }
        // Crée la nouvelle liste d'arrêts
        const newStops = new Map();
        let lastTimedStopName = null;
        let lastTimedTime = 0;
        let segmentPath = [];
        // Remplit la liste des arrêts en reprenant les arrêts déjà renseignés et en ajoutant les gares de passage     
        for (const stopName of shortestPath.path) {
            const currentStop = this.getStop(stopName, true) || Stop.newStopIncludingParity(stopName);
            newStops.set(stopName, currentStop);
            // Recherche de la gare suivante et de la connexion entre les deux gares
            const nextStopName = shortestPath.path[shortestPath.path.indexOf(stopName) + 1];
            if (nextStopName) {
                const nextStop = newStops.get(nextStopName);
                if (nextStop) {
                    const connection = (_a = CONNECTIONS.get(stopName)) === null || _a === void 0 ? void 0 : _a.get(nextStopName);
                    if (connection) {
                        currentStop.nextStop = nextStop;
                        currentStop.connection = connection;
                    }
                }
            }
            // Calcul du temps de parcours depuis la dernière gare avec une heure de départ
            const hasTime = currentStop.arrivalTime > 0 || currentStop.passageTime > 0;
            if (hasTime) {
                const currentTime = currentStop.arrivalTime || currentStop.passageTime;
                if (lastTimedStopName !== null && segmentPath.length > 0) {
                    // Calcule le temps total entre les deux points connus
                    let totalTime = 0;
                    const segmentTimes = [];
                    let from = lastTimedStopName;
                    for (const to of segmentPath) {
                        const connection = (_b = CONNECTIONS.get(from)) === null || _b === void 0 ? void 0 : _b.get(to);
                        if (!connection) {
                            console.warn(`Pas de connexion entre ${from} et ${to}`);
                            segmentTimes.push(0);
                            continue;
                        }
                        segmentTimes.push(connection.time);
                        totalTime += connection.time;
                        from = to;
                    }
                    // Répartir le temps aux arrêts intermédiaires
                    let cumulativeTime = 0;
                    for (let i = 0; i < segmentPath.length; i++) {
                        cumulativeTime += segmentTimes[i];
                        const stop = newStops.get(segmentPath[i]);
                        if (stop && stop.passageTime === 0 && stop.arrivalTime === 0) {
                            const interpolatedTime = Math.round(lastTimedTime + (cumulativeTime * (currentTime - lastTimedTime)) / totalTime);
                            stop.passageTime = interpolatedTime;
                        }
                    }
                }
                lastTimedStopName = stopName;
                lastTimedTime = currentTime;
                segmentPath = [];
            }
            else {
                segmentPath.push(stopName);
            }
        }
        this.stops = newStops;
    }
}
const TRAINS_SHEET = "Trains";
const TRAINS_TABLE = "Trains";
const TRAINS_HEADERS = [[
        "Id",
        "Numéro du train",
        "Parité de ligne",
        "Jours",
        "Code mission",
        "Heure de départ",
        "Gare de départ",
        "Heure d'arrivée",
        "Gare d'arrivée",
        "Gares intermédiaires"
    ]];
const TRAINS_COL_KEY = 0; // Non lue car calculée
const TRAINS_COL_NUMBER = 1;
const TRAINS_COL_LINE_PARITY = 2;
const TRAINS_COL_DAYS = 3;
const TRAINS_COL_MISSION_CODE = 4;
const TRAINS_COL_DEPARTURE_TIME = 5;
const TRAINS_COL_DEPARTURE_STATION = 6;
const TRAINS_COL_ARRIVAL_TIME = 7;
const TRAINS_COL_ARRIVAL_STATION = 8;
const TRAINS_COL_VIA_STATIONS = 9;
/**
 * Charge les trains à partir du tableau "Trains" de la feuille "Trains".
 * Les trains sont stockés dans un objet avec comme clés le numéro de train
 * suivi du jour et comme valeur l'objet Train.
 */
function loadTrains(days = "JW", trains = "") {
    loadStations();
    const data = getDataFromTable(TRAINS_SHEET, TRAINS_TABLE);
    const daysToLoadTable = daysToNumbers(days || "JW");
    const trainDaysCache = new Map();
    const trainsToLoadTable = new Set(String(trains).split(';'));
    for (let row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        const number = row[TRAINS_COL_NUMBER];
        // Analyse si une sélection par numéro de train est demandée
        if (trains !== "" && !trainsToLoadTable.has(number.toString())) {
            continue;
        }
        const days = row[TRAINS_COL_DAYS];
        let commonDays;
        // Lecture du cache pour vérifier si la liste des jours du train a déjà été rencontrée
        if (trainDaysCache.has(days)) {
            commonDays = trainDaysCache.get(days);
        }
        else {
            // Dans ce cas, conversion des jours du train en table et enregistrement dans le cache
            const trainDaysTable = daysToNumbers(days);
            commonDays = trainDaysTable.filter(day => daysToLoadTable.includes(day));
            trainDaysCache.set(days, commonDays);
        }
        // Si les jours du train ne correspondent pas aux jours demandés, passer au train suivant
        if (commonDays.length === 0) {
            continue;
        }
        // Extraction des valeurs
        const lineParity = row[TRAINS_COL_LINE_PARITY];
        const missionCode = row[TRAINS_COL_MISSION_CODE];
        const departureTime = row[TRAINS_COL_DEPARTURE_TIME];
        const departureStation = row[TRAINS_COL_DEPARTURE_STATION];
        const arrivalTime = row[TRAINS_COL_ARRIVAL_TIME];
        const arrivalStation = row[TRAINS_COL_ARRIVAL_STATION];
        const viaStations = row[TRAINS_COL_VIA_STATIONS];
        // Création de l'objet Train
        const train = new Train(number, lineParity, days, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations);
        // Insertion du train dans la table avec plusieurs clés d'accès
        // - une référence pour la clé unique du train
        TRAINS.set(train.key, train);
        // - une référence pour chacun des jours demandés
        commonDays.forEach((day) => {
            const key = number + "_" + day;
            TRAINS.set(key, train);
        });
    }
}
/**
 * Affiche les trains dans un tableau.
 * Les données sont celles stockées dans l'objet TRAINS.
 * @param {string} sheetName - Nom de la feuille de calcul.
 * @param {string} tableName - Nom du tableau.
 * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
 */
function printTrains(sheetName, tableName, startCell = "A1") {
    // Filtrer l'objet TRAINS en ne prennant qu'une seule fois les trains ayant la même clé   
    const seenKeys = new Set();
    const uniqueTrains = Array.from(TRAINS.entries())
        .filter(([mapKey, train]) => mapKey === train.key)
        .map(([_, train]) => train);
    // Convertir l'objet TRAINS filtré en un tableau de données
    const data = uniqueTrains.map(train => [
        train.key,
        train.number,
        train.lineParity,
        train.days,
        train.missionCode,
        train.departureTime,
        train.departureStation,
        train.arrivalTime,
        train.arrivalStation,
        train.viaStations.join(';'),
    ]);
    // Imprimer le tableau
    const table = printTable(TRAINS_HEADERS, data, sheetName, tableName, startCell);
    // Mettre les heures au format "hh:mm:ss"
    const timeColumns = [
        TRAINS_COL_DEPARTURE_TIME,
        TRAINS_COL_ARRIVAL_TIME,
    ];
    for (const col of timeColumns) {
        table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
    }
}
/**
 * Classe Reuse qui définit un train pour un unique jour, étant la réutilisation
 * d'un ou deux trains précédents, et ayant une ou deux réutilisations,
 * en faisant référence à un Train avec horaires pouvant circuler plusieurs jours par semaine
 */
class Reuse {
    constructor(number, train, day, unit1 = "", unit2 = "", previous1 = "", previous2 = "", reuse1 = "", reuse2 = "") {
        this.number = number;
        this.train = train;
        this.day = day;
        this.unit1 = unit1;
        this.unit2 = unit2;
        this.previous1 = previous1;
        this.previous2 = previous2;
        this.reuse1 = reuse1;
        this.reuse2 = reuse2;
    }
    /**
     * Retourne la clé du train qui est composée du numéro du train
     * suivi de la liste des jours de circulation.
     * @returns {string} Clé du train
     */
    get key() {
        return plannedTrain ? `${this.number}_${this.day}` : "";
    }
    /**
     * Retourne le numéro du train avec changement de parité.
     * Si le train change de parité, le numéro est sous la forme "XX/Y".
     * Sinon, le numéro est sous la forme "XX".
     * @returns {string} Le numéro du train avec changement de parité.
     */
    get doubleParityNumber() {
        return this.plannedTrain.doubleParityNumber();
    }
}
const REUSES_SHEET = "Réuts";
const REUSES_TABLE = "Réuts";
const REUSES_HEADERS = [[
        "Id",
        "Numéro du train",
        "Jours",
        "Elément Nord",
        "Elément Sud",
        "Train Précédent Nord",
        "Train Précédent Sud",
        "Réutilisation Nord",
        "Réutilisation Sud",
    ]];
const REUSES_COL_KEY = 0;
const REUSES_COL_NUMBER = 1;
const REUSES_COL_DAYS = 2;
const REUSES_COL_TRAIN = 3;
const REUSES_COL_UNIT1 = 4;
const REUSES_COL_UNIT2 = 5;
const REUSES_COL_PREVIOUS1 = 6;
const REUSES_COL_PREVIOUS2 = 7;
const REUSES_COL_REUSE1 = 8;
const REUSES_COL_REUSE2 = 9;
/**
 * Charge les réutilisations à partir du tableau "Réuts" de la feuille "Réuts".
 * Les réutilisations sont stockés dans la table un objet avec comme clés le numéro de train
 * suivi du jour de circulation (numéro du jour ou date) et comme valeur l'objet Réutilisation.
 */
function loadReuses(days = "JW", trains = "") {
}
/**
 * Affiche les réutilisations dans un tableau.
 * Les données sont celles stockées dans l'objet REUSES.
 * @param {string} sheetName - Nom de la feuille de calcul.
 * @param {string} tableName - Nom du tableau.
 * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
 */
function printReuses(sheetName, tableName, startCell = "A1") {
    // Convertir l'objet REUSES en un tableau de données
    const data = Object.values(REUSES).map(reuse => [
        reuse.key,
        reuse.number,
        reuse.day,
        reuse.unit1,
        reuse.unit2,
        reuse.previous1,
        reuse.previous2,
        reuse.reuse1,
        reuse.reuse2
    ]);
    // Imprimer le tableau
    printTable(REUSES_HEADERS, data, sheetName, tableName, startCell);
}
class Stop {
    constructor(stationName, parity = -1, arrivalTime = 0, departureTime = 0, passageTime = 0, track = "", changeNumber = 0) {
        this.stationName = stationName;
        this.station = STATIONS.get(this.stationName);
        this.parity = parity;
        this.arrivalTime = arrivalTime;
        this.departureTime = departureTime;
        this.passageTime = passageTime;
        this.track = track;
        this.changeNumber = changeNumber;
        this.nextStop = null;
    }
    /**
     * Renvoie une clé unique pour l'arrêt, composée du nom de la gare et de la parité (si connue).
     * @returns {string} Clé unique
     */
    get key() {
        return `${this.stationName}${(this.parity >= 0 ? '_' + this.parity : "")}`;
    }
    /**
     * Crée une nouvelle instance de l'arrêt incluant la parité à partir d'une chaîne de caractères.
     * La chaîne de caractères doit être au format "NomDeGare_Parité".
     *
     * @param {string} stopWithParity - Chaîne de caractères contenant le nom de la gare et la parité, séparés par un underscore.
     * @param {number} [arrivalTime=0] - Heure d'arrivée à l'arrêt.
     * @param {number} [departureTime=0] - Heure de départ de l'arrêt.
     * @param {number} [passageTime=0] - Heure de passage à l'arrêt (sans arrêt).
     * @param {string} [track=""] - Voie de l'arrêt.
     * @param {number} [changeNumber=0] - Changement de numérotation.
     * @returns {Stop} Nouvelle instance de l'arrêt avec les informations fournies.
     */
    static newStopIncludingParity(stopWithParity, arrivalTime = 0, departureTime = 0, passageTime = 0, track = "", changeNumber = 0) {
        const [name, parity] = stopWithParity.split("_");
        return new Stop(name, Number(parity), arrivalTime, departureTime, passageTime, track, changeNumber);
    }
    /**
     * Renvoie l'heure d'arrivée, de départ ou de passage.
     * Si l'heure d'arrivée est définie, renvoie cette heure,
     * sinon, renvoie l'heure de départ,
     * sinon, renvoie l'heure de passage.
     * @returns {number} Heure envoyée
     */
    getTime() {
        return this.arrivalTime || this.departureTime || this.passageTime;
    }
}
const STOPS_SHEET = "Arrêts";
const STOPS_TABLE = "Arrêts";
const STOPS_HEADERS = [[
        "Numéro du train",
        "Jour",
        "Gare",
        "Parité",
        "Arrivée",
        "Départ",
        "Passage",
        "Voie"
    ]];
const STOPS_COL_TRAIN_NUMBER = 0;
const STOPS_COL_TRAIN_DAYS = 1;
const STOPS_COL_STATION = 2;
const STOPS_COL_PARITY = 3;
const STOPS_COL_ARRIVAL_TIME = 4;
const STOPS_COL_DEPARTURE_TIME = 5;
const STOPS_COL_PASSAGE_TIME = 6;
const STOPS_COL_TRACK = 7;
/**
 * Charge les arrêts à partir de la feuille "Arrêts" du classeur.
 * Les arrêts sont stockés dans la propriété "stops" des trains correspondants.
 * Si un train n'existe pas, un message d'erreur est affiché.
 */
function loadStops() {
    const data = getDataFromTable(STOPS_SHEET, STOPS_TABLE);
    for (let row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        // Vérifie si le train existe
        const trainNumber = row[STOPS_COL_TRAIN_NUMBER];
        const trainDays = row[STOPS_COL_TRAIN_DAYS];
        const trainKey = trainNumber + "_" + trainDays;
        if (!TRAINS.has(trainKey)) {
            // console.log(`Le train "${trainKey}" n'existe pas !`);
            continue;
        }
        const train = TRAINS.get(trainKey);
        // Extraction des valeurs
        const station = row[STOPS_COL_STATION];
        const parity = row[STOPS_COL_PARITY];
        const arrivalTime = row[STOPS_COL_ARRIVAL_TIME];
        const departureTime = row[STOPS_COL_DEPARTURE_TIME];
        const passageTime = row[STOPS_COL_PASSAGE_TIME];
        const track = row[STOPS_COL_TRACK];
        const stop = new Stop(station, parity, arrivalTime, departureTime, passageTime, track);
        train.addStop(stop);
    }
}
/**
 * Affiche les arrêts des trains dans un tableau.
 * Les données sont celles stockées dans les objets Train et Stop de l'objet TRAINS.
 * @param {string} sheetName - Nom de la feuille de calcul.
 * @param {string} tableName - Nom du tableau.
 * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
 */
function printStops(sheetName, tableName, startCell = "A1") {
    // Filtrer l'objet TRAINS en ne prennant qu'une seule fois les trains ayant la même clé   
    const seenKeys = new Set();
    const uniqueTrains = Array.from(TRAINS.entries())
        .filter(([mapKey, train]) => mapKey === train.key)
        .map(([_, train]) => train);
    // Créer le tableau final avec les données de chaque arrêt pour chaque train
    const data = [];
    for (const train of uniqueTrains) {
        for (const [stationName, stop] of train.stops.entries()) {
            data.push([
                train.number,
                train.days,
                stop.stationName,
                stop.parity,
                stop.arrivalTime,
                stop.departureTime,
                stop.passageTime,
                stop.track
            ]);
        }
    }
    // Imprimer le tableau
    const table = printTable(STOPS_HEADERS, data, sheetName, tableName, startCell);
    const timeColumns = [
        STOPS_COL_ARRIVAL_TIME,
        STOPS_COL_DEPARTURE_TIME,
        STOPS_COL_PASSAGE_TIME
    ];
    for (const col of timeColumns) {
        table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
    }
}
class Station {
    constructor(abbreviation, name, variants_stations, oddTurnaround, evenTurnaround, reverseLineParity) {
        this.abbreviation = abbreviation;
        this.name = name;
        this.variants = [];
        this.variants = [
            ...[abbreviation, ...variants_stations]
                .filter(v => v.trim() !== '')
                .flatMap(v => [v + '_0', v + '_1'])
        ];
        this.oddTurnaround = oddTurnaround;
        this.evenTurnaround = evenTurnaround;
        this.reverseLineParity = reverseLineParity;
    }
}
const STATIONS_SHEET = "Gares";
const STATIONS_TABLE = "Gares";
const STATIONS_HEADERS = [[
        "Abréviation",
        "Nom",
        "Variantes",
        "Gare de rebroussement",
        "Parité de ligne inversée"
    ]];
const STATIONS_COL_ABBR = 0;
const STATIONS_COL_NAME = 1;
const STATIONS_COL_VARIANTS_STATIONS = 2;
const STATIONS_COL_TURNAROUND = 3;
const STATIONS_COL_REVERSE_LINE_PARITY = 4;
/**
 * Charge les gares à partir du tableau "Gares" de la feuille "Gares".
 * Les gares sont stockées dans une Map avec comme clés l'abréviation
 * de la gare et comme valeur l'objet Station.
 * @returns Une Map contenant les gares sous forme de clés (abréviation)
 * et de valeurs (objets Station).
 */
function loadStations(erase = false) {
    // Vérifier que la table à charger existe déjà
    if (STATIONS.size > 0) {
        if (!erase) {
            return;
        }
        STATIONS.clear(); // Vide la map sans changer sa référence
    }
    const data = getDataFromTable(STATIONS_SHEET, STATIONS_TABLE);
    for (const row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        // Extraction des valeurs
        const abbreviation = row[STATIONS_COL_ABBR];
        const name = row[STATIONS_COL_NAME];
        const variants_stations = row[STATIONS_COL_VARIANTS_STATIONS].split(";");
        const oddTurnaround = row[STATIONS_COL_TURNAROUND].indexOf(PARAM.oddLetter) >= 0;
        const evenTurnaround = row[STATIONS_COL_TURNAROUND].indexOf(PARAM.evenLetter) >= 0;
        const reverseLineParity = row[STATIONS_COL_REVERSE_LINE_PARITY];
        // Création de la gare
        const station = new Station(abbreviation, name, variants_stations, oddTurnaround, evenTurnaround, reverseLineParity);
        STATIONS.set(abbreviation, station);
    }
}
/**
 * Affiche les stations dans un tableau.
 * Les données sont celles stockées dans l'objet STATIONS.
 * @param {string} sheetName - Nom de la feuille de calcul.
 * @param {string} tableName - Nom du tableau.
 * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
 */
function printStations(sheetName, tableName, startCell = "A1") {
    // Convertir l'objet STATIONS en un tableau de données
    const data = Object.values(STATIONS).map(station => [
        station.abbreviation,
        station.name,
        station.variants.join(", "),
        station.connectedStationsWithParityChange.join(", ")
    ]);
    // Imprimer le tableau
    printTable(STATIONS_HEADERS, data, sheetName, tableName, startCell);
}
class Connection {
    constructor(from, to, time, withTurnaround, withMovement, changeParity) {
        this.from = from;
        this.to = to;
        this.time = time;
        this.withTurnaround = withTurnaround;
        this.withMovement = withMovement;
        this.changeParity = changeParity;
    }
}
const CONNECTIONS_SHEET = "Param";
const CONNECTIONS_TABLE = "Connexions";
const CONNECTIONS_HEADERS = [[
        "De",
        "Vers",
        "Temps",
        "Rebroussement",
        "Evolution",
        "Changement de parité"
    ]];
const CONNECTIONS_COL_FROM = 0;
const CONNECTIONS_COL_TO = 1;
const CONNECTIONS_COL_TIME = 2;
const CONNECTIONS_COL_TURNAROUND = 3;
const CONNECTIONS_COL_MOVEMENT = 4;
const CONNECTIONS_COL_CHANGE_PARITY = 5;
/**
 * Charge les connexions entre les gares et les variantes de ces gares.
 * Les connexions sont stockées dans un objet avec comme clés les gares de départ
 * et comme valeurs des objets Map où les clés sont les gares d'arrivée et les
 * valeurs des objets contenant le temps de trajet et un booléen indiquant si un
 * retournement est nécessaire.
 * Les variantes sont stockées dans un objet avec comme clés les noms de gares
 * et comme valeurs des tableaux de gares variantes.
 *
 * @param {boolean} [erase=false] - Si true, efface les données existantes
 * avant de charger les nouvelles.
 * Sinon, les données ne sont pas chargée si la table est déjà remplie.
 */
function loadConnections(erase = false) {
    // Vérifier que la table à charger existe déjà
    if (CONNECTIONS.size > 0) {
        if (!erase) {
            return;
        }
        CONNECTIONS.clear(); // Vide la map sans changer sa référence
    }
    loadStations();
    const data = getDataFromTable(CONNECTIONS_SHEET, CONNECTIONS_TABLE);
    for (const row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        // Extraction des valeurs
        const from = row[CONNECTIONS_COL_FROM];
        const to = row[CONNECTIONS_COL_TO];
        const time = row[CONNECTIONS_COL_TIME];
        const withTurnaround = row[CONNECTIONS_COL_TURNAROUND];
        const withMovement = row[CONNECTIONS_COL_MOVEMENT];
        const changeParity = row[CONNECTIONS_COL_CHANGE_PARITY];
        // Création de la connexion
        const connection = new Connection(from, to, time, withTurnaround, withMovement, changeParity);
        if (!CONNECTIONS.has(from)) {
            CONNECTIONS.set(from, new Map());
        }
        CONNECTIONS.get(from).set(to, connection);
    }
}
/**
 * Affiche les connexions entre les gares dans un tableau.
 * Les données sont celles stockées dans l'objet CONNECTIONS.
 * @param {string} sheetName - Nom de la feuille de calcul.
 * @param {string} tableName - Nom du tableau.
 * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
 */
function printConnections(sheetName, tableName, startCell = "A1") {
    // Convertir l'objet CONNECTIONS en un tableau de données
    const data = [];
    for (const [from, connections] of CONNECTIONS) {
        for (const [to, connection] of connections) {
            data.push([
                from,
                to,
                connection.time,
                connection.withTurnaround
            ]);
        }
    }
    // Imprimer le tableau
    printTable(CONNECTIONS_HEADERS, data, sheetName, tableName, startCell);
}
function saveConnectionsTimes(train) {
    train.stops.forEach((stop) => {
        var _a;
        if (stop.nextStop && CONNECTIONS.has(stop.key) && CONNECTIONS.has(stop.nextStop.key)) {
            const connection = (_a = CONNECTIONS.get(stop.key)) === null || _a === void 0 ? void 0 : _a.get(stop.nextStop.key);
            if (connection && stop.nextStop.arrivalTime !== 0 && stop.departureTime !== 0) {
                connection.time = stop.nextStop.arrivalTime - stop.departureTime;
            }
        }
    });
}
