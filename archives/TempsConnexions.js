"use strict";
const SHEET_TB = "TB";
const TB_CELL_GARE = "B1";
const TB_CELL_JOUR = "B2";
const TB_CELL_HEURE = "B3";
const TB_CELL_ARRIVEE_DEPART = "B4";
const TB_COLS_TRAINPARGARE = "D:F";
var WORKBOOK;
var STATIONS = {};
var CONNECTIONS;
function main(workbook) {
    WORKBOOK = workbook;
    let sheet = WORKBOOK.getActiveWorksheet();
    // Nettoyer les anciennes données dans les colonnes E
    sheet.getRange("E:Z").clear();
    // Récupérer les valeurs de départ, arrivée et via
    let start = sheet.getRange("B1").getValue();
    let end = sheet.getRange("B2").getValue();
    let via = sheet.getRange("B3").getValue().split(";");
    let changeTime = sheet.getRange("B4").getValue(); // Temps de changement de sens
    let rawConnections = sheet.getRange("A5:D400").getValues();
    // Créer les connexions et les variantes
    let { connections, variants } = createConnectionsAndVariants(rawConnections);
    // Générer toutes les combinaisons possibles de parcours
    let allCombinations = generateCombinations(start, end, via, variants);
    // Trouver le chemin le plus court parmi toutes les combinaisons
    let shortestPath = findShortestPath(connections, allCombinations, changeTime);
    // Afficher le chemin le plus court dans la colonne F
    if (shortestPath) {
        sheet.getRange("F1").setValue("Chemin le plus court");
        sheet.getRange("F2").setValue(`Distance totale : ${shortestPath.totalDistance} min`);
        for (let i = 0; i < shortestPath.path.length; i++) {
            sheet.getRange(`F${i + 3}`).setValue(shortestPath.path[i]);
        }
    }
    else {
        sheet.getRange("F1").setValue("Aucun chemin trouvé");
    }
}
/**
 * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
 * @param connections - La carte des connexions entre les gares.
 * @param allCombinations - La liste de toutes les combinaisons de parcours à évaluer.
 * @param changeTime - Temps de changement de sens.
 * @returns Un objet contenant le chemin le plus court et sa distance totale, ou null si aucun chemin n'est trouvé.
 */
function findShortestPath(connections, allCombinations, changeTime) {
    let shortestPath = null;
    for (let combination of allCombinations) {
        // Calculer le chemin complet et la distance totale pour la combinaison actuelle
        let { path, totalDistance } = calculateCompletePath(connections, combination, changeTime);
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
 * @param connections - La carte des connexions entre les gares.
 * @param combination - La liste ordonnée des gares à parcourir.
 * @returns Un objet contenant le chemin complet et la distance totale.
 */
function calculateCompletePath(connections, combination, changeTime) {
    let completePath = [];
    let totalDistance = 0;
    for (let i = 0; i < combination.length - 1; i++) {
        let segmentStart = combination[i];
        let segmentEnd = combination[i + 1];
        // Trouver le chemin le plus court pour le tronçon actuel
        let segmentPath = dijkstra(connections, segmentStart, segmentEnd, changeTime);
        if (segmentPath.length === 0) {
            // Si aucun chemin n'est trouvé pour ce tronçon, retourner un chemin vide
            return { path: [], totalDistance: 0 };
        }
        // Calculer la distance pour ce tronçon
        let segmentDistance = calculatePathTime(connections, segmentPath, changeTime);
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
 * @param connections - La carte des connexions entre les gares, incluant le temps de trajet et l'information sur le besoin de changement de sens.
 * @param path - La liste ordonnée des gares constituant le chemin.
 * @param changeTime - Le temps de changement de sens à ajouter lorsque nécessaire.
 * @returns Le temps total du chemin, incluant les temps de trajet et de changement de sens.
 */
function calculatePathTime(connections, path, changeTime) {
    var _a;
    let totalTime = 0;
    for (let i = 0; i < path.length - 1; i++) {
        let from = path[i];
        let to = path[i + 1];
        let connection = (_a = connections.get(from)) === null || _a === void 0 ? void 0 : _a.get(to);
        if (connection) {
            totalTime += connection.time;
            // Ajouter le temps de changement de sens sauf pour le premier segment
            if (i > 0 && connection.needsTurnaround) {
                totalTime += changeTime;
            }
        }
    }
    return totalTime;
}
/**
 * Cherche le chemin le plus court entre le départ et l'arrivée
 * en appliquant Dijkstra.
 * @param connections - La carte des connexions entre les gares.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @returns Le chemin le plus court.
 */
function dijkstra(connections, start, end, changeTime) {
    let distances = new Map();
    let previousNodes = new Map();
    let unvisited = new Set(connections.keys());
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
        for (let [neighbor, { time, needsTurnaround }] of connections.get(currentNode) || []) {
            let additionalTime = time;
            if (needsTurnaround && currentNode !== start) { // Si un changement de sens est nécessaire, ajouter du temps
                additionalTime += changeTime;
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
 * Les variantes de chaque gare sont incluses en utilisant la Map variants.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @param via - Les gares intermédiaires à passer par.
 * @param variants - La Map qui permet de récupérer les variantes pour chaque gare.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une combinaison de route possible.
 */
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
/**
 * Renvoie toutes les variantes possibles pour une gare.
 * Une variante correspond au sens de passage dans la gare : GARE_1 en impair, GARE_2 en pair.
 * Seules les gares de retournement permettent de passer d'une gare à l'autre
 * Si la gare a déjà un suffixe imposé (_), renvoie uniquement cette gare avec suffixe [gare].
 * Sinon, renvoie toutes les variantes associées.
 * @param gare - La gare dont on cherche les variantes.
 * @param variants - La Map qui permet de récupérer les variantes pour chaque gare.
 * @returns Un tableau contenant toutes les variantes possibles pour la gare.
 */
function getAllVariants(gare, variants) {
    // Si la gare a un suffixe (_), renvoyer uniquement [gare]
    if (gare.includes('_')) {
        return [gare];
    }
    // Sinon, renvoyer toutes les variantes associées
    return variants.get(gare) || [];
}
/**
 * Génère toutes les permutations possibles d'un tableau de chaînes.
 * @param arr - Le tableau de chaînes à permuter.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une permutation possible.
 */
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
/**
 * Étend une permutation de gares pour inclure toutes les variantes possibles.
 * @param permutation - La permutation de gares à étendre.
 * @param variants - La Map qui permet de récupérer les variantes pour chaque gare.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une permutation possible avec toutes les variantes.
 */
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
class Train {
    /**
     * Constructeur d'un train.
     * @param {number} number - Numéro du train
     * @param {number} direction - Direction du train
     * @param {number} day - Jour du train
     * @param {string} missionCode - Code de mission du train
     * @param {number} departureTime - Heure de départ du train
     * @param {string} departureStation - Gare de départ du train
     * @param {number} arrivalTime - Heure d'arrivée du train
     * @param {string} arrivalStation - Gare d'arrivée du train
     * @param {string} viaStations - Gares via
     * @param {string} reuse - Réutilisation du train
     */
    constructor(number, direction, day, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations, reuse) {
        this.number = number;
        this.direction = direction;
        this.day = day;
        this.missionCode = missionCode;
        this.departureTime = departureTime;
        this.departureStation = departureStation;
        this.arrivalTime = arrivalTime;
        this.arrivalStation = arrivalStation;
        this.viaStations = viaStations;
        this.reuse = reuse;
        this.stops = {};
        let stop = new Stop(0, departureStation, 0, departureTime, "0");
        this.addStop(stop);
        stop = new Stop(0, arrivalStation, arrivalTime, 0, "0");
        this.addStop(stop);
    }
    /**
     * Ajoute un arrêt au train.
     * @param {Stop} stop - Arrêt à ajouter
     */
    addStop(stop) {
        this.stops[stop.station] = stop;
    }
}
const SHEET_TRAINS = "Trains";
const TABLE_TRAINS = "Trains";
const TRAINS_COL_NUMBER = 0;
const TRAINS_COL_DIRECTION = 1;
const TRAINS_COL_DAY = 2;
const TRAINS_COL_MISSION_CODE = 3;
const TRAINS_COL_DEPARTURE_TIME = 4;
const TRAINS_COL_DEPARTURE_STATION = 5;
const TRAINS_COL_ARRIVAL_TIME = 6;
const TRAINS_COL_ARRIVAL_STATION = 7;
const TRAINS_COL_VIA_STATIONS = 8;
const TRAINS_COL_REUSE = 9;
/**
 * Charge les trains à partir du tableau "Trains" de la feuille "Trains".
 * Les trains sont stockés dans un objet avec comme clés le numéro de train
 * suivi du jour et comme valeur l'objet Train.
 * @returns Un objet contenant les trains.
 */
function loadTrains() {
    let trains = {};
    let sheet = WORKBOOK.getWorksheet(SHEET_TRAINS);
    if (!sheet) {
        console.log("La feuille " + SHEET_TRAINS + " n'existe pas !");
        return {};
    }
    const table = sheet.getTable(TABLE_TRAINS);
    if (!table) {
        console.log("Le tableau " + TABLE_TRAINS + " n'existe pas !");
        return {};
    }
    let data = table.getRange().getValues();
    for (let i = 0; i < data.length; i++) {
        let number = data[i][TRAINS_COL_NUMBER];
        let direction = data[i][TRAINS_COL_DIRECTION];
        let day = data[i][TRAINS_COL_DAY];
        let missionCode = data[i][TRAINS_COL_MISSION_CODE];
        let departureTime = data[i][TRAINS_COL_DEPARTURE_TIME];
        let departureStation = data[i][TRAINS_COL_DEPARTURE_STATION];
        let arrivalTime = data[i][TRAINS_COL_ARRIVAL_TIME];
        let arrivalStation = data[i][TRAINS_COL_ARRIVAL_STATION];
        let viaStations = data[i][TRAINS_COL_VIA_STATIONS];
        let reuse = data[i][TRAINS_COL_REUSE];
        let train = new Train(number, direction, day, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations, reuse);
        let key = number + "_" + day;
        trains[key] = train;
    }
    return trains;
}
class Stop {
    /**
     * Constructeur d'un arrêt.
     * @param {number} parity - Parité de l'arrêt
     * @param {string} station - Gare de l'arrêt
     * @param {number} arrivalTime - Heure d'arrivée à l'arrêt
     * @param {number} departureTime - Heure de départ à l'arrêt
     * @param {string} track - Voie de l'arrêt
     */
    constructor(parity, station, arrivalTime, departureTime, track) {
        this.parity = parity;
        this.station = station;
        this.arrivalTime = arrivalTime;
        this.departureTime = departureTime;
        this.track = track;
    }
    /**
     * Renvoie l'heure d'arrivée ou de départ à l'arrêt.
     * Si l'heure d'arrivée est définie, renvoie cette heure.
     * Sinon, renvoie l'heure de départ.
     * @returns {number} L'heure d'arrivée ou de départ.
     */
    getTime() {
        return this.arrivalTime ? this.arrivalTime : this.departureTime;
    }
}
const SHEET_STOPS = "Arrêts";
const TABLE_STOPS = "Arrêts";
const STOPS_COL_TRAIN_NUMBER = 0;
const STOPS_COL_PARITY = 1;
const STOPS_COL_DAY = 2;
const STOPS_COL_STATION = 3;
const STOPS_COL_ARRIVAL_TIME = 4;
const STOPS_COL_DEPARTURE_TIME = 5;
const STOPS_COL_TRACK = 6;
function loadStops(trains) {
    let sheet = WORKBOOK.getWorksheet(SHEET_STOPS);
    if (!sheet) {
        console.log("La feuille " + SHEET_STOPS + " n'existe pas !");
        return {};
    }
    const table = sheet.getTable(TABLE_STOPS);
    if (!table) {
        console.log("Le tableau " + TABLE_STOPS + " n'existe pas !");
        return {};
    }
    const data = table.getRange().getValues();
    for (let i = 0; i < data.length; i++) {
        let trainNumber = data[i][STOPS_COL_TRAIN_NUMBER];
        let parity = data[i][STOPS_COL_PARITY];
        let day = data[i][STOPS_COL_DAY];
        let station = data[i][STOPS_COL_STATION];
        let arrivalTime = data[i][STOPS_COL_ARRIVAL_TIME];
        let departureTime = data[i][STOPS_COL_DEPARTURE_TIME];
        let track = data[i][STOPS_COL_TRACK];
        let stop = new Stop(parity, station, arrivalTime, departureTime, track);
        let key = trainNumber + "_" + day;
        if (trains[key]) {
            trains[key].addStop(stop);
        }
    }
    console.log("Arrêts ajoutés !");
}
class Station {
    /**
     * Constructeur d'une gare.
     * @param {string} abbreviation - Abréviation de la gare
     * @param {string} name - Nom de la gare
     * @param {string[]} connectedStationsWithParityChange - Liste des abréviations des gares connectées avec changement de parité
     */
    constructor(abbreviation, name, connectedStationsWithParityChange) {
        this.abbreviation = abbreviation;
        this.name = name;
        this.connectedStationsWithParityChange = connectedStationsWithParityChange;
        this.variants = [];
    }
    /**
     * Vérifie s'il y a un changement de parité en allant vers une autre gare.
     * @param {string} otherStation - Abréviation de l'autre gare
     * @returns {boolean} - True si changement de parité, sinon false
     */
    hasParityChangeWith(otherStation) {
        return this.connectedStationsWithParityChange.includes(otherStation);
    }
}
const SHEET_STATIONS = "Param";
const TABLE_STATIONS = "Gares";
const STATIONS_COL_ABBR = 0;
const STATIONS_COL_NAME = 1;
const STATIONS_COL_PARITY_CHANGE = 2;
/**
 * Charge les gares à partir du tableau "Gares" de la feuille "Param".
 * @returns Un objet contenant les gares sous forme de clés (abréviation) et de valeurs (objets Station).
 */
function loadStations() {
    STATIONS = {};
    let sheet = WORKBOOK.getWorksheet(SHEET_STATIONS);
    if (!sheet) {
        console.log("La feuille " + SHEET_STATIONS + " n'existe pas !");
        return STATIONS;
    }
    const table = sheet.getTable(TABLE_STATIONS);
    if (!table) {
        console.log("Le tableau " + TABLE_STATIONS + " n'existe pas !");
        return STATIONS;
    }
    let data = table.getRange().getValues();
    for (let i = 1; i < data.length; i++) {
        let abbreviation = data[i][STATIONS_COL_ABBR];
        let name = data[i][STATIONS_COL_NAME];
        let parityChangeStr = data[i][STATIONS_COL_PARITY_CHANGE];
        let connectedStationsWithParityChange = parityChangeStr ? parityChangeStr.split(';') : [];
        let station = new Station(abbreviation, name, connectedStationsWithParityChange);
        STATIONS[abbreviation] = station;
    }
    return STATIONS;
}
const SHEET_CONNECTIONS = "Param";
const TABLE_CONNECTIONS = "Connexions";
const CONNECTIONS_COL_FROM = 0;
const CONNECTIONS_COL_TO = 1;
const CONNECTIONS_COL_TIME = 2;
const CONNECTIONS_COL_NEEDS_TURNAROUND = 3;
function createConnectionsAndVariants(rawConnections) {
    let connections = new Map();
    let variants = new Map();
    for (let row of rawConnections) {
        let from = row[0];
        let to = row[1];
        let time = row[2];
        let needsTurnaround = row[3];
        if (!connections.has(from)) {
            connections.set(from, new Map());
        }
        connections.get(from).set(to, { time, needsTurnaround });
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
    CONNECTIONS = connections;
    return { connections, variants };
}
