"use strict";
/* Classeur principal. */
var WORKBOOK;
/* Liste des paramètres. */
const PARAM = {
    filled: false,
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
/* Liste des trains associés à un jour donné et leurs réutilisations  */
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
    console.log(daysToNumbers("146"));
    return;
    // loadStations();
    // loadConnections();
    // loadTrains("", "147490");
    // loadStops();
    // printTrains("Test", "Test");
    // console.log(TRAINS.get("147490_2"));
    return;
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
            if (i > 0 && connection.needsTurnaround) {
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
    let distances = new Map();
    let previousNodes = new Map();
    let unvisited = new Set(CONNECTIONS.keys());
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
            if (connexion.needsTurnaround && currentNode !== start) { // Si un changement de sens est nécessaire, ajouter du temps
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
    let filteredVia = via.filter(v => v.trim() !== "");
    // Générer les permutations des gares intermédiaires
    let viaPermutations = permute(filteredVia);
    // Ajouter start au début et end à la fin de chaque permutation
    let routes = viaPermutations.map(permutation => [start, ...permutation, end]);
    // Étendre chaque route pour inclure toutes les variantes possibles
    let allCombinations = routes.flatMap((route) => expandPermutations(route));
    return allCombinations;
}
/**
 * Renvoie toutes les variantes possibles pour une gare.
 * Une variante correspond au sens de passage dans la gare : GARE_1 en impair, GARE_2 en pair.
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
    let first = getAllVariants(permutation[0]);
    let restExpanded = expandPermutations(permutation.slice(1));
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
const PARAM_LINE_MAX_CONNEXIONS_NUMBER = 1;
const PARAM_LINE_TURNAROUND_TIME = 4;
/**
 * Charge les paramètres du tableau "Paramètres" de la feuille "Param".
 * @returns Un objet contenant les paramètres chargés.
 */
function loadParams(erase = false) {
    const data = getDataFromTable(PARAM_SHEET, PARAM_TABLE);
    PARAM.maxConnectionNumber = data[PARAM_LINE_MAX_CONNEXIONS_NUMBER][1];
    PARAM.turnaroundTime = data[PARAM_LINE_TURNAROUND_TIME][1];
    loadWTrainsRegex();
    load4FiguresTrainsRegex();
    loadDays();
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
    if (!PARAM.trains4FiguresRegex)
        load4FiguresTrainsRegex(); // Charge si non encore fait
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
/**
 * Classe Train qui définit un train, plannifié sur un ou plusieurs jours de la semaine,
 * ou sur plusieurs dates précises, avec les mêmes horaires.
 * Plusieurs trains associés à un jour donné et leurs réutilisations y font référence
 */
class Train {
    constructor(number, days, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations) {
        this.number = number;
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
        // Détermine le sens principal pour la ligne C
        const departureStationObj = STATIONS.get(departureStation);
        this.line_C_direction = departureStationObj ? ((this.number + (departureStationObj.lineC_reverse_direction ? 1 : 0)) % 2) : -1;
        // Détermine si la gare d'arrivée change de parité, auquel cas le train a une double parité
        const arrivalStationObj = STATIONS.get(arrivalStation);
        this.changeParity = (departureStationObj && arrivalStationObj) ? (departureStationObj.lineC_reverse_direction != arrivalStationObj.lineC_reverse_direction) : false;
    }
    /**
     * Retourne la clé du train qui est composée du numéro du train
     * suivi de la liste des jours de circulation ou de la première date de circulation.
     * @returns {string} Clé du train plannifié
     */
    get key() {
        return `${this.number}_${this.days.split(';')[0]}`;
    }
    /**
     * Retourne le numéro du train avec changement de parité.
     * Si le train change de parité, le numéro est sous la forme "XX/Y".
     * Sinon, le numéro est sous la forme "XX".
     * @param {boolean} [withNumber4Figures=false] - Si true, le numéro est renommé en 4 chiffres pour les trains commerciaux de la ligne C
     * @returns {string} Le numéro du train avec changement de parité.
     */
    double_parity_number(withNumber4Figures = false) {
        const evenNumber = this.number - (this.number % 2);
        const doubleParityNumber = this.changeParity ? evenNumber + "/" + ((evenNumber + 1) % 10) : this.number.toString();
        return withNumber4Figures ? renameWith4Figures2(doubleParityNumber) : doubleParityNumber;
    }
    /**
     * Retourne le numéro du train utilisé par les opérateurs de la ligne C,
     * avec un numéro de train à 4 chiffres si le train est commercial
     * Si withDoubleParity est true, le numéro est renommé en prenant en compte le changement de parité.
     * @param {boolean} [withDoubleParity=false] - Si true, le numéro est renommé en prenant en compte le changement de parité.
     * @returns {string} Le numéro du train à 4 ou 6 chiffres utilisé par les opérateurs de la ligne C.
     */
    number4Figures(withDoubleParity = false) {
        return withDoubleParity ? this.double_parity_number(true) : renameWith4Figures2(this.number);
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
    getStop(stop) {
        const parity = this.number % 2;
        return this.stops.get(stop + "_" + parity)
            || this.stops.get(stop + "_" + (1 - parity))
            || this.stops.get(stop + "_?")
            || null;
    }
    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée du train,
     * puis génère la liste des arrêts calculés
     * @returns {void}
     */
    findPath() {
        loadStations2();
        // Cherche toutes les combinaisons possibles de départ, d'arrivée et de passages via
        const allCombinations = generateCombinations(this.departureStation, this.arrivalStation, this.viaStations);
        // Trouve le chemin le plus court parmi toutes les combinaisons
        const shortestPath = findShortestPath(allCombinations);
        let previousStop = new Stop("");
        let lastStopWithHour = previousStop;
        let timeSinceLastHour = 0;
        let nbOfReverseSinceLastHour = 0;
        shortestPath.path.array.forEach(stop => {
            // Lecture de l'arrêt : nom de gare et sens
            let thisStop = new Stop(stop.split('_')[0], stop.split('_')[1]);
            // Vérifie si la connexion est un retournement
            if (thisStop.stationName === previousStop.stationName) {
                this.stops[thisStop.stationName].reuse = thisStop.direction - previousStop.direction;
                nbOfReverseSinceLastHour += 1;
            }
            else {
                if (this.stops[thisStop.stationName]) {
                    this.stops[thisStop.stationName].reverse = false;
                }
                this.stops[thisStop.stationName] = new Stop(thisStop.direction, thisStop.stationName, 0, 0, thisStop.direction.toString());
            }
        });
    }
}
const TRAINS_SHEET = "Trains";
const TRAINS_TABLE = "Trains";
const TRAINS_HEADERS = [[
        "Id",
        "Numéro du train",
        "Direction ligne C",
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
const TRAINS_COL_LINE_C_DIRECTION = 2; // Non lue car calculée
const TRAINS_COL_DAYS = 3;
const TRAINS_COL_MISSION_CODE = 4;
const TRAINS_COL_DEPARTURE_TIME = 5;
const TRAINS_COL_DEPARTURE_STATION = 5;
const TRAINS_COL_ARRIVAL_TIME = 6;
const TRAINS_COL_ARRIVAL_STATION = 7;
const TRAINS_COL_VIA_STATIONS = 8;
/**
 * Charge les trains à partir du tableau "Trains" de la feuille "Trains".
 * Les trains sont stockés dans un objet avec comme clés le numéro de train
 * suivi du jour et comme valeur l'objet Train.
 */
function loadTrains(days = "JW", trains = "") {
    loadStations();
    const data = getDataFromTable(TRAINS_SHEET, TRAINS_TABLE);
    const daysToLoadTable = new Set(daysToNumbers(days || "JW"));
    const trainDaysCache = new Map();
    const trainsToLoadTable = new Set(String(trains).split(';'));
    for (let row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        const days = row[TRAINS_COL_DAY];
        // Lecture du cache pour vérifier si la liste des jours du train n'a pas déjà été rencontrée
        if (!trainDaysCache.has(days)) {
            // Dans ce cas, conversion des jours du train en table et enregistrement dans le cache
            const trainDaysTable = daysToNumbers(days);
            const hasCommonDays = trainDaysTable.some(day => daysToLoadTable.has(day));
            trainDaysCache.set(days, { hasCommonDays, trainDaysTable });
        }
        // Récupère les informations des jours depuis le cache
        const { hasCommonDays, trainDaysTable } = trainDaysCache.get(days);
        // Si les jours du train ne correspondent pas aux jours demandés, passer au train suivant
        if (!hasCommonDays) {
            continue;
        }
        const number = row[TRAINS_COL_NUMBER];
        // Analyse si une sélection par numéro de train est demandée
        if (trains !== "") {
            if (trainsToLoadTable && !trainsToLoadTable.has(number.toString())) {
                continue;
            }
        }
        // Extraction des valeurs
        const missionCode = row[TRAINS_COL_MISSION_CODE];
        const departureTime = row[TRAINS_COL_DEPARTURE_TIME];
        const departureStation = row[TRAINS_COL_DEPARTURE_STATION];
        const arrivalTime = row[TRAINS_COL_ARRIVAL_TIME];
        const arrivalStation = row[TRAINS_COL_ARRIVAL_STATION];
        const viaStations = row[TRAINS_COL_VIA_STATIONS];
        // Création des trains
        let train = new Train(number, days, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations);
        TRAINS.set(key, train);
        // Création des trains selon les jours concernés
        trainDaysTable.forEach((day) => {
            let key = number + "_" + day;
        });
    }
}
/**
 * Affiche les trains dans un tableau.
 * @param {string} sheetName - Nom de la feuille de calcul.
 * @param {string} tableName - Nom du tableau.
 * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
 */
function printTrains(sheetName, tableName, startCell = "A1") {
    // Convertir l'objet TRAINS en un tableau de données
    const data = Object.values(TRAINS).map(train => [
        train.key,
        train.number,
        train.line_C_direction,
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
    table.getRange().getColumn(REUSE_COL_DEPARTURE_TIME).setNumberFormat("hh:mm:ss");
    table.getRange().getColumn(REUSE_COL_ARRIVAL_TIME).setNumberFormat("hh:mm:ss");
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
    get double_parity_number() {
        return this.plannedTrain.double_parity_number();
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
     * Renvoie une clé unique pour l'arrêt, composée du nom de la gare et de la parité.
     * Si la parité est inconnue, "X_?" est renvoyé.
     * @returns {string} Clé unique
     */
    get key() {
        return `${this.stationName}_${(this.parity === -1 ? "?" : this.parity)}`;
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
    loadStations();
    const data = getDataFromTable(STOPS_SHEET, STOPS_TABLE);
    const stopDaysCache = new Map();
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
class Station {
    constructor(abbreviation, name, variants_stations, odd_reversal, even_reversal, lineC_reverse_direction) {
        this.abbreviation = abbreviation;
        this.name = name;
        this.variants = [];
        this.variants.push(abbreviation + '_' + 0);
        this.variants.push(abbreviation + '_' + 1);
        variants_stations.forEach(variants_station => {
            this.variants.push(variants_station + '_' + 0);
            this.variants.push(variants_station + '_' + 1);
        });
        this.odd_reversal = odd_reversal;
        this.even_reversal = even_reversal;
        this.lineC_reverse_direction = lineC_reverse_direction;
    }
}
const STATIONS_SHEET = "Gares";
const STATIONS_TABLE = "Gares";
const STATIONS_HEADERS = [[
        "Abréviation",
        "Nom",
        "Variantes",
        "Changements de parité"
    ]];
const STATIONS_COL_ABBR = 0;
const STATIONS_COL_NAME = 1;
const STATIONS_COL_VARIANTS_STATIONS = 2;
const STATIONS_COL_REVERSAL = 3;
const STATIONS_COL_LINEC_REVERSE_DIRECTION = 4;
/**
 * Charge les gares à partir du tableau "Gares" de la feuille "Gares".
 * Les gares sont stockées dans une Map avec comme clés l'abréviation
 * de la gare et comme valeur l'objet Station.
 * @returns Une Map contenant les gares sous forme de clés (abréviation)
 * et de valeurs (objets Station).
 */
function loadStations(erase = false) {
    // Vérifier que la table à charger n'existe pas déjà
    if (Object.keys(STATIONS).length > 0) {
        if (!erase) {
            return;
        }
        STATIONS = new Map();
    }
    const data = getDataFromTable(STATIONS_SHEET, STATIONS_TABLE);
    for (let row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        // Extraction des valeurs
        let abbreviation = row[STATIONS_COL_ABBR];
        let name = row[STATIONS_COL_NAME];
        let variants_stations = row[STATIONS_COL_VARIANTS_STATIONS].split(";");
        let odd_reversal = row[STATIONS_COL_REVERSAL].indexOf('I') >= 0;
        let even_reversal = row[STATIONS_COL_REVERSAL].indexOf('P') >= 0;
        let lineC_reverse_direction = row[STATIONS_COL_LINEC_REVERSE_DIRECTION] === 1;
        // Création de la gare
        let station = new Station(abbreviation, name, variants_stations, odd_reversal, even_reversal, lineC_reverse_direction);
        STATIONS.set(abbreviation, station);
    }
}
/**
 * Affiche les stations dans un tableau.
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
    printTable(headers, data, sheetName, tableName, startCell);
}
class Connection {
    constructor(from, to, time, needsTurnaround) {
        this.from = from;
        this.to = to;
        this.time = time;
        this.needsTurnaround = needsTurnaround;
        let fromParity = parseInt(from.split('_')[1], 10) % 2;
        let toParity = parseInt(to.split('_')[1], 10) % 2;
        this.changeParity = toParity - fromParity;
    }
}
const CONNECTIONS_SHEET = "Param";
const CONNECTIONS_TABLE = "Connexions";
const CONNECTIONS_COL_FROM = 0;
const CONNECTIONS_COL_TO = 1;
const CONNECTIONS_COL_TIME = 2;
const CONNECTIONS_COL_NEEDS_TURNAROUND = 3;
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
    // Vérifier que la table à charger n'existe pas déjà
    if (Object.keys(CONNECTIONS).length > 0) {
        if (!erase) {
            return;
        }
        CONNECTIONS = new Map();
    }
    loadStations();
    const data = getDataFromTable(CONNECTIONS_SHEET, CONNECTIONS_TABLE);
    for (let row of data.slice(1)) {
        // Vérification si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) {
            continue;
        }
        // Extraction des valeurs
        let from = row[CONNECTIONS_COL_FROM];
        let to = row[CONNECTIONS_COL_TO];
        let time = row[CONNECTIONS_COL_TIME];
        let needsTurnaround = row[CONNECTIONS_COL_NEEDS_TURNAROUND];
        // Création de la connexion
        if (!CONNECTIONS.has(from)) {
            CONNECTIONS.set(from, new Map());
        }
        CONNECTIONS.get(from).set(to, new Connection(from, to, time, needsTurnaround));
    }
}
