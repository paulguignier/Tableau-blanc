const SHEET_TB = "TB";
const TB_CELL_GARE = "B1";
const TB_CELL_JOUR = "B2";
const TB_CELL_HEURE = "B3";
const TB_CELL_ARRIVEE_DEPART = "B4";
const TB_COLS_TRAINPARGARE = "D:F";

var WORKBOOK: ExcelScript.Workbook;
var STATIONS: Record<string, Station> = {};
var CONNECTIONS = new Map<string, Connection[]>();

function main(workbook: ExcelScript.Workbook) {
    WORKBOOK = workbook;
    let sheet = WORKBOOK.getActiveWorksheet();

    console.log(CONNECTIONS[]);

    
}




/**
 * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
 * @param connections - La carte des connexions entre les gares.
 * @param allCombinations - La liste de toutes les combinaisons de parcours à évaluer.
 * @param changeTime - Temps de changement de sens.
 * @returns Un objet contenant le chemin le plus court et sa distance totale, ou null si aucun chemin n'est trouvé.
 */
function findShortestPath(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, allCombinations: string[][], changeTime: number): { path: string[], totalDistance: number } | null {
    let shortestPath: { path: string[], totalDistance: number } | null = null;

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
function calculateCompletePath(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, combination: string[], changeTime: number): { path: string[], totalDistance: number } {
    let completePath: string[] = [];
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
function calculatePathTime(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, path: string[], changeTime: number): number {
    let totalTime  = 0;

    for (let i = 0; i < path.length - 1; i++) {
        let from = path[i];
        let to = path[i + 1];
        let connection = connections.get(from)?.get(to);
        if (connection) {
            totalTime += connection.time;
            // Ajouter le temps de changement de sens sauf pour le premier segment
            if (i > 0 && connection.needsTurnaround) {
                totalTime += changeTime;
            }
        }
    }

    return totalTime ;
}

/**
 * Cherche le chemin le plus court entre le départ et l'arrivée
 * en appliquant Dijkstra.
 * @param connections - La carte des connexions entre les gares.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @returns Le chemin le plus court.
 */
function dijkstra(start: string, end: string, changeTime: number): string[] {
    let distances = new Map<string, number>();
    let previousNodes = new Map<string, string | null>();
    let unvisited = new Set<string>();
    let path: string[] = [];

    // Initialisation des distances et des nœuds non visités
    CONNECTIONS.forEach((_, node) => {
        distances.set(node, Infinity);
        previousNodes.set(node, null);
        unvisited.add(node);
    });
    distances.set(start, 0);

    while (unvisited.size > 0) {
        // Sélectionner le nœud avec la plus petite distance
        let currentNode = Array.from(unvisited).reduce((minNode, node) =>
            distances.get(node)! < distances.get(minNode)! ? node : minNode
        );

        if (distances.get(currentNode) === Infinity) break; // Aucun chemin disponible

        unvisited.delete(currentNode);

        // Examiner les voisins
        for (let connection of CONNECTIONS.get(currentNode) || []) {
            let neighbor = connection.to;
            if (!unvisited.has(neighbor)) continue;

            let additionalTime = connection.time;
            if (connection.needsTurnaround && currentNode !== start) {
                additionalTime += changeTime;
            }

            let newDist = distances.get(currentNode)! + additionalTime;
            if (newDist < distances.get(neighbor)!) {
                distances.set(neighbor, newDist);
                previousNodes.set(neighbor, currentNode);
            }
        }
    }

    // Retracer le chemin
    let step = end;
    while (step) {
        path.unshift(step);
        step = previousNodes.get(step)!;
    }

    // Vérifier si le chemin est valide
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
function generateCombinations(start: string, end: string, via: string[], variants: Map<string, string[]>): string[][] {
    // Filtrer les gares intermédiaires pour éliminer les chaînes vides
    let filteredVia = via.filter(v => v.trim() !== "");

    // Générer les permutations des gares intermédiaires
    let viaPermutations = permute(filteredVia);

    // Ajouter start au début et end à la fin de chaque permutation
    let routes = viaPermutations.map(permutation => [start, ...permutation, end]);

    // Étendre chaque route pour inclure toutes les variantes possibles
    let allCombinations: string[][] = routes.flatMap(route => expandPermutations(route, variants));

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
function getAllVariants(gare: string, variants: Map<string, string[]>): string[] {
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
function permute(arr: string[]): string[][] {
    if (arr.length === 0) return [[]];
    if (arr.length === 1) return [[arr[0]]];

    let result: string[][] = [];

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
function expandPermutations(permutation: string[], variants: Map<string, string[]>): string[][] {
    if (permutation.length === 0) return [[]];

    let first = getAllVariants(permutation[0], variants);
    let restExpanded = expandPermutations(permutation.slice(1), variants);

    let result: string[][] = [];
    for (let f of first) {
        for (let r of restExpanded) {
            result.push([f, ...r]);
        }
    }

    return result;
}

class Train {
    number: number;
    direction: number;
    day: number;
    missionCode: string;
    departureTime: number;
    departureStation: string;
    arrivalTime: number;
    arrivalStation: string;
    viaStations: string;
    reuse: string;
    stops: { [abbreviation: string]: Stop };

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
    constructor(number: number, direction: number, day: number, missionCode: string, departureTime: number, departureStation: string, arrivalTime: number, arrivalStation: string, viaStations: string, reuse: string) {
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
    addStop(stop: Stop) {
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
function loadTrains(): Record<string, Train> {
    let trains: Record<string, Train> = {};

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
        let number = data[i][TRAINS_COL_NUMBER] as number;
        let direction = data[i][TRAINS_COL_DIRECTION] as number;
        let day = data[i][TRAINS_COL_DAY] as number;
        let missionCode = data[i][TRAINS_COL_MISSION_CODE] as string;
        let departureTime = data[i][TRAINS_COL_DEPARTURE_TIME] as number;
        let departureStation = data[i][TRAINS_COL_DEPARTURE_STATION] as string;
        let arrivalTime = data[i][TRAINS_COL_ARRIVAL_TIME] as number;
        let arrivalStation = data[i][TRAINS_COL_ARRIVAL_STATION] as string;
        let viaStations = data[i][TRAINS_COL_VIA_STATIONS] as string;
        let reuse = data[i][TRAINS_COL_REUSE] as string;

        let train = new Train(number, direction, day, missionCode, departureTime, departureStation, arrivalTime, arrivalStation, viaStations, reuse);
        let key = number + "_" + day;
        trains[key] = train;
    }

    return trains;
}

class Stop {
    parity: number;
    station: string;
    arrivalTime: number;
    departureTime: number;
    track: string;

    /**
     * Constructeur d'un arrêt.
     * @param {number} parity - Parité de l'arrêt
     * @param {string} station - Gare de l'arrêt
     * @param {number} arrivalTime - Heure d'arrivée à l'arrêt
     * @param {number} departureTime - Heure de départ à l'arrêt
     * @param {string} track - Voie de l'arrêt
     */
    constructor(parity: number, station: string, arrivalTime: number, departureTime: number, track: string) {
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
    getTime(): number {
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

function loadStops(trains: Record<string, Train>) {
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
        let trainNumber = data[i][STOPS_COL_TRAIN_NUMBER] as string;
        let parity = data[i][STOPS_COL_PARITY] as number;
        let day = data[i][STOPS_COL_DAY] as number;
        let station = data[i][STOPS_COL_STATION] as string;
        let arrivalTime = data[i][STOPS_COL_ARRIVAL_TIME] as number;
        let departureTime = data[i][STOPS_COL_DEPARTURE_TIME] as number;
        let track = data[i][STOPS_COL_TRACK] as string;

        let stop = new Stop(parity, station, arrivalTime, departureTime, track);
        let key = trainNumber + "_" + day;

        if (trains[key]) {
            trains[key].addStop(stop);
        }
    }

    console.log("Arrêts ajoutés !");
}

class Station {
    abbreviation: string;
    name: string;
    connectedStationsWithParityChange: string[];
    variants: string[];

    /**
     * Constructeur d'une gare.
     * @param {string} abbreviation - Abréviation de la gare
     * @param {string} name - Nom de la gare
     * @param {string[]} connectedStationsWithParityChange - Liste des abréviations des gares connectées avec changement de parité
     */
    constructor(abbreviation: string, name: string, connectedStationsWithParityChange: string[]) {
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
    hasParityChangeWith(otherStation: string): boolean {
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
function loadStations(): Record<string, Station> {
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
        let abbreviation = data[i][STATIONS_COL_ABBR] as string;
        let name = data[i][STATIONS_COL_NAME] as string;
        let parityChangeStr = data[i][STATIONS_COL_PARITY_CHANGE] as string;
        let connectedStationsWithParityChange = parityChangeStr ? parityChangeStr.split(';') : [];

        let station = new Station(abbreviation, name, connectedStationsWithParityChange);
        STATIONS[abbreviation] = station;
    }

    return STATIONS;
}

class Connection {
    from: string;
    to: string;
    time: number;
    needsTurnaround: boolean;

    /**
     * Constructeur d'une connexion.
     * @param {string} from - Gare de départ
     * @param {string} to - Gare d'arrivée
     * @param {number} time - Temps de trajet
     * @param {boolean} needsTurnaround - Indique si un retournement est nécessaire
     */
    constructor(from: string, to: string, time: number, needsTurnaround: boolean) {
        this.from = from;
        this.to = to;
        this.time = time;
        this.needsTurnaround = needsTurnaround;
    }
}

const SHEET_CONNECTIONS = "Param";
const TABLE_CONNECTIONS = "Connexions";
const CONNECTIONS_COL_FROM = 0;
const CONNECTIONS_COL_TO = 1;
const CONNECTIONS_COL_TIME = 2;
const CONNECTIONS_COL_NEEDS_TURNAROUND = 3;

function createConnectionsAndVariants(rawConnections: (string | number | boolean)[][]): { connections: Map<string, Connection[]>, variants: Map<string, string[]> } {
    CONNECTIONS = new Map<string, Connection[]>();

    // Ensure stations are loaded
    if (STATIONS.size === 0) {
        loadStations()
    }

    let sheet = WORKBOOK.getWorksheet(SHEET_STATIONS);
    if (!sheet) {
        console.log("La feuille " + SHEET_STATIONS + " n'existe pas !");
        return stations;
    }
    const table = sheet.getTable(TABLE_STATIONS);
    if (!table) {
        console.log("Le tableau " + TABLE_STATIONS + " n'existe pas !");
        return stations;
    }
    let data = table.getRange().getValues();

    for (const row of data) {
        const from = row[CONNECTIONS_COL_FROM] as string;
        const to = row[CONNECTIONS_COL_TO] as string;
        const time = row[CONNECTIONS_COL_TIME] as number;
        const needsTurnaround = row[CONNECTIONS_COL_NEEDS_TURNAROUND] as boolean;

        const connection = new Connection(from, to, time, needsTurnaround);

        // Add connection to CONNECTIONS map
        if (!CONNECTIONS.has(from)) {
            CONNECTIONS.set(from, []);
        }
        CONNECTIONS.get(from)!.push(connection);

        // Process variants for 'from' station
        const baseFrom = from.split('_')[0];
        if (STATIONS.has(baseFrom)) {
            if (!STATIONS[baseFrom].variants.includes(from)) {
                STATIONS[baseFrom].variants.push(from);
            }
        }

        // Process variants for 'to' station
        const baseTo = to.split('_')[0];
        if (STATIONS.has(baseTo)) {
            if (!STATIONS[baseTo].variants.includes(to)) {
                STATIONS[baseTo].variants.push(to);
            }
        }
    }

    return CONNECTIONS;
}