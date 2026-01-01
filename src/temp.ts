/**
 * Chargements de trains
 * 
 * Code Excel Automate pour la création et l'utilisation de la base de données des trains.
 * 
 * @author Paul Guignier
 * @version 1.0
 * @package scr\ChargementTrains.ts
 */

// export {loadParams, loadConnections, findShortestPath, generateCombinations};

/* Classeur principal. */
var WORKBOOK: ExcelScript.Workbook;

/* Console pour l'affichage de messages d'informations sur le contenu d'objets. */
var CONSOLE_INFO: Console;
/* Console pour l'affichage de messages d'avertissement. */
var CONSOLE_WARN: Console;
/* Console pour l'affichage de messages de debug. */
var CONSOLE_DEBUG: Console;

function main(workbook: ExcelScript.Workbook) {
    WORKBOOK = workbook;
    CONSOLE_INFO = console;
    CONSOLE_WARN = console;
    CONSOLE_DEBUG = console;
    const sheet = WORKBOOK.getActiveWorksheet();

    const DEBUG_MODE = true;

    // Lance la fonction de tests
    // Si les tests sont actifs, la suite du programme n'est pas exécuté. 
    if (tests(DEBUG_MODE)) return;

    // Lit les paramètres
    loadParams();
    loadConnections();

    // loadTrainPaths("", "147500_J;148504_J;147201_J;148202_J;147402_J;
    //      148402_J;147601_J;148602_J;145801_J;145804_J");
    loadTrainPaths("2", "142446_J");
    CONSOLE_DEBUG.log(TRAIN_PATHS);
  
    loadStops();
    // findPathsOnAllTrainPaths();
    // printTrainPaths("Test", "Trains1");
    return;
    printStops("Test", "Stops1", "A20");

    saveConnectionsTimes();
    printConnections("Test2", "Connections1");


    const allCombinations = generateCombinations("MPU", "ETP", "".split(";"));
    CONSOLE_INFO.log(allCombinations);
    const shortestPath = findShortestPath(allCombinations);
    CONSOLE_INFO.log(shortestPath);

    return;
}



/**
 * Affiche le résultat d'un test avec un symbole de réussite (✔) ou d'échec (✘)
 * @param {string} label Nom du test
 * @param {T} actual Valeur actuelle obtenue
 * @param {T} expected Valeur attendue
 */
function assertDD<T>(
    label: string,
    actual: T,
    expected: T
): void {
    const success = actual === expected;
    CONSOLE_INFO.log(
        `${success ? "✔" : "✘"} ${label} → ${actual} (attendu: ${expected})`
    );
}



/**
 * Fonction de tests pour les différentes parties du code.
 * Lorsqu'elle est appelée, toutes les autres fonctions ne sont pas exécutées.
 * Les tests sont actifs si la constante DEBUG_MODE est à true.
 * @param {boolean} [debugMode=false] Si vrai, les fonctions de test sont lancés,
 *  puis le programme est interrompu. Si faux (par défaut), le programme continue normalement.
 * @returns {boolean} Si les tests sont actifs, la fonction renvoie true, sinon false.
 */
function tests(debugMode = false): boolean {

    if (!debugMode) return false;

    loadParams();
    testParity();
    return true;


    
    /* Fonctions de lecture des feuilles Excel. */
    // const SHEET = "Test";
    // const TABLE = "Trains1";
    // const HEADERS = [[
    //     "Col1",
    //     "Col2"
    // ]];
    // const data = getDataFromTable(SHEET, TABLE);
    // CONSOLE_INFO.log(data);
    // const data2 = [
    //     ["A", "B"],
    //     ["C", "D"]
    // ];
    // const table = printTable(HEADERS, data2, SHEET, TABLE, "A2", false);
    // CONSOLE_INFO.log(checkCellName("A1"));
    // CONSOLE_INFO.log(checkCellName("1A")); // Doit aboutir un une erreur

    /* Fonctions de dates et heures */
    CONSOLE_INFO.log(`22/06/2025 au format jj/mm/aaaa : ${formatDate(45830.94347)}`);
    CONSOLE_INFO.log(`22:38:36 au format hh:mm:ss : ${formatTime(45830.94347)}`);


    loadConnections();
    const parity = new Parity("A", false);
    CONSOLE_INFO.log(!parity);

    /* Lecture des gares et test des variants */
    // CONSOLE_INFO.log(STATIONS);
    // CONSOLE_INFO.log(getAllVariants("VC"));

    // const allCombinations = generateCombinations("MPU", "ETP", "".split(";"));
    // CONSOLE_INFO.log(allCombinations);
    // const shortestPath = findShortestPath(allCombinations);
    // CONSOLE_INFO.log(shortestPath);
    // findShortestPath
    // calculateCompletePath
    // calculatePathTime
    // Dijkstra
    // generateCombinations
    // permute
    // expandPermutations
    // getAllVariants

    /* Fonctions numéro de train */
    // CONSOLE_INFO.log(`146490 est W : ${isWTrain("146490")}`); 
    // CONSOLE_INFO.log(`569907 est W : ${isWTrain("569907")}`); 
    // CONSOLE_INFO.log(`147490 est W : ${isWTrain("147490")}`); 
    // CONSOLE_INFO.log(`146490 renommé : ${abreviateTo4Digits("146490")}`);
    // CONSOLE_INFO.log(`569907 renommé : ${abreviateTo4Digits("569907")}`); 
    // CONSOLE_INFO.log(`147490 renommé : ${abreviateTo4Digits("147490")}`); 

    /* Analyse des jours */
    CONSOLE_INFO.log(`Jours 1Jeudi6 : ${Day.extractFromString("1Jeudi6")}`);
    CONSOLE_INFO.log(`Jours J : ${Day.extractFromString("J")}`);
    CONSOLE_INFO.log(`Jours 14W;24 : ${Day.extractFromString("14W;24")}`);
    CONSOLE_INFO.log(`Jours 14W24 & J: ${Day.extractFromString("14W24","J")}`);


    // const t1 = new TrainPath(569000, 0, "1", "TEST", 12/24, "TRA-PG", 13/24, "PJ", "VFG");
    // CONSOLE_INFO.log(t1.getStop("VC-BG_2",true,true));
    // t1.findPath();
    // CONSOLE_DEBUG.log(t1.getStop("INV_1"));

    /* Test trainPath.getStop */
    // loadTrainPaths("2", "147490");
    // loadStops();
    // const t2 = TRAIN_PATHS.get("147490_2");
    // t2.findPath();
    // // CONSOLE_INFO.log(t2.getStop("VC-BG_2",true,true));
    // CONSOLE_INFO.log(t2);

    // findPathsOnAllTrainPaths();
    // printTrainPaths("Test", "Trains1");
    // printStops("Test", "Stops1", "A10");
    // CONSOLE_INFO.log(TRAIN_PATHS.get("147490_2"));

    return true;
}

function testParity() {

    /* ==========================================================
   TESTS DATA-DRIVEN – CLASSE Parity
   ==========================================================
   Objectifs :
   - Valider l’analyse des valeurs de parité
   - Garantir la cohérence des comportements (update, invert, print…)
   - Centraliser tous les scénarios de test sous forme de données
   ========================================================== */

    /* ==========================================================
    1. CONSTRUCTEUR & analyseValue()
    ----------------------------------------------------------
    Vérifie :
    - Lettres de parité
    - Chiffres de parité
    - Numéros de train
    - Valeurs indéfinies
    - Double parité autorisée / interdite
    ========================================================== */

    const constructorTests = [
        { desc: 'Lettre impair "I"', value: "I", doubleAllowed: false, expected: Parity.odd },
        { desc: 'Lettre pair "P"', value: "P", doubleAllowed: false, expected: Parity.even },
        { desc: 'Chiffre impair 1', value: 1, doubleAllowed: false, expected: Parity.odd },
        { desc: 'Chiffre pair 2', value: 2, doubleAllowed: false, expected: Parity.even },
        { desc: 'Numéro de train impair "12345"', value: "12345", doubleAllowed: false, expected: Parity.odd },
        { desc: 'Numéro de train pair "12346"', value: "12346", doubleAllowed: false, expected: Parity.even },
        { desc: 'Valeur vide', value: "", doubleAllowed: false, expected: Parity.undefined },
        { desc: 'Zéro "0"', value: "0", doubleAllowed: false, expected: Parity.undefined },
        { desc: 'Double parité IP interdite', value: "IP", doubleAllowed: false, expected: Parity.undefined },
        { desc: 'Double parité IP autorisée', value: "IP", doubleAllowed: true, expected: Parity.double },
        { desc: 'Numéro double implicite "1/2"', value: "1/2", doubleAllowed: true, expected: Parity.double }
    ];

    constructorTests.forEach(t => {
        const p = new Parity(t.value, t.doubleAllowed);
        assertDD(
            `new Parity(${JSON.stringify(t.value)}, doubleAllowed=${t.doubleAllowed}) – ${t.desc}`,
            p.value,
            t.expected
        );
    });

    /* ==========================================================
    2. update() & setter value
    ----------------------------------------------------------
    Vérifie :
    - Mise à jour dynamique de la parité
    - Cohérence du setter value
    ========================================================== */

    const updateTests = [
        { desc: 'update impair → pair', start: "I", update: "P", expected: Parity.even },
        { desc: 'update pair → impair', start: "P", update: "I", expected: Parity.odd }
    ];

    updateTests.forEach(t => {
        const p = new Parity(t.start);
        p.update(t.update);
        assertDD(
            `Parity(${t.start}).update(${t.update}) – ${t.desc}`,
            p.value,
            t.expected
        );
    });

    const setterTests = [
        { desc: 'setter value = odd', start: "P", set: Parity.odd, expected: Parity.odd }
    ];

    setterTests.forEach(t => {
        const p = new Parity(t.start);
        p.value = t.set;
        assertDD(
            `Parity(${t.start}).value = ${t.set} – ${t.desc}`,
            p.value,
            t.expected
        );
    });

    /* ==========================================================
    3. equals() & is()
    ----------------------------------------------------------
    Vérifie :
    - Comparaison entre deux instances
    - Comparaison avec une valeur de parité
    ========================================================== */

    const equalsTests = [
        { desc: 'I equals 1', p1: "I", p2: 1, expected: true },
        { desc: 'I not equals P', p1: "I", p2: "P", expected: false }
    ];

    equalsTests.forEach(t => {
        const a = new Parity(t.p1);
        const b = new Parity(t.p2);
        assertDD(
            `Parity(${t.p1}).equals(Parity(${t.p2})) – ${t.desc}`,
            a.equals(b),
            t.expected
        );
    });

    const isTests = [
        { desc: 'is odd', value: "I", parity: Parity.odd, expected: true },
        { desc: 'is not even', value: "I", parity: Parity.even, expected: false }
    ];

    isTests.forEach(t => {
        const p = new Parity(t.value);
        assertDD(
            `Parity(${t.value}).is(${t.parity}) – ${t.desc}`,
            p.is(t.parity),
            t.expected
        );
    });

    /* ==========================================================
    4. invert()
    ----------------------------------------------------------
    Vérifie :
    - Inversion impair ↔ pair
    - Stabilité des états double et indéfini
    ========================================================== */

    const invertTests = [
        { desc: 'impair → pair', value: "I", doubleAllowed: false, expected: Parity.even },
        { desc: 'pair → impair', value: "P", doubleAllowed: false, expected: Parity.odd },
        { desc: 'double reste double', value: "IP", doubleAllowed: true, expected: Parity.double },
        { desc: 'indéfini reste indéfini', value: "", doubleAllowed: false, expected: Parity.undefined }
    ];

    invertTests.forEach(t => {
        const p = new Parity(t.value, t.doubleAllowed).invert();
        assertDD(
            `Parity(${t.value}).invert() – ${t.desc}`,
            p.value,
            t.expected
        );
    });

    /* ==========================================================
    5. printDigit() & printLetter()
    ----------------------------------------------------------
    Vérifie :
    - Sorties texte associées à chaque parité
    - Gestion correcte des états double et indéfini
    ========================================================== */

    const printTests = [
        {
            desc: 'print impair',
            value: "I",
            doubleAllowed: false,
            digit: PARAM.parityDigits.get(Parity.odd),
            letter: PARAM.parityLetters.get(Parity.odd)
        },
        {
            desc: 'print pair',
            value: "P",
            doubleAllowed: false,
            digit: PARAM.parityDigits.get(Parity.even),
            letter: PARAM.parityLetters.get(Parity.even)
        },
        {
            desc: 'print double',
            value: "IP",
            doubleAllowed: true,
            digit: PARAM.parityDigits.get(Parity.double),
            letter:
                PARAM.parityLetters.get(Parity.odd)! +
                PARAM.parityLetters.get(Parity.even)!
        },
        {
            desc: 'print indéfini',
            value: "",
            doubleAllowed: false,
            digit: "",
            letter: ""
        }
    ];

    printTests.forEach(t => {
        const p = new Parity(t.value, t.doubleAllowed);

        assertDD(`printDigit – ${t.desc}`, p.printDigit(), t.digit);
        assertDD(`printLetter – ${t.desc}`, p.printLetter(), t.letter);
    });

    /* ==========================================================
    6. containsParityLetter()
    ----------------------------------------------------------
    Vérifie :
    - Détection correcte des lettres de parité
    - Cas simple et double parité
    ========================================================== */

    const containsTests = [
        { desc: 'contient impair', string: "Train I", parity: Parity.odd, expected: true },
        { desc: 'ne contient pas pair', string: "Train I", parity: Parity.even, expected: false },
        { desc: 'contient double', string: "Train IP", parity: Parity.double, expected: true }
    ];

    containsTests.forEach(t => {
        assertDD(
            `containsParityLetter("${t.string}", ${t.parity}) – ${t.desc}`,
            Parity.containsParityLetter(t.string, t.parity),
            t.expected
        );
    });


}

class WorkbookService {

    /**
     * Renvoie la feuille de calcul Excel correspondant au nom donné.
     * Si la feuille n'existe pas, renvoie null si failOnError est à false,
     *  sinon lance une exception.
     * @param {string} sheetName Nom de la feuille de calcul à chercher.
     * @param {boolean} [failOnError=true] Si vrai (par défaut), lance une exception
     *  si la feuille n'existe pas. Si faux, renvoie null.
     * @returns {ExcelScript.Worksheet | null} Feuille de calcul Excel correspondant au nom donné,
     *  ou null si elle n'existe pas.
     */
    public static getSheetOrFail(
        sheetName: string,
        failOnError = true
    ): ExcelScript.Worksheet | null {
        const sheet = WORKBOOK.getWorksheet(sheetName);
        if (!sheet) {
            const msg = `La feuille "${sheetName}" n'existe pas.`;
            if (failOnError) throw new Error(msg);
            CONSOLE_WARN.log(msg);
            return null;
        }
        return sheet;
    }

}

/**
 * Renvoie le tableau Excel correspondant au nom donné dans la feuille de calcul donnée.
 * Si le tableau n'existe pas, renvoie null si failOnError est à false,
 *  sinon lance une exception.
 * @param {string} sheetName Nom de la feuille de calcul où chercher le tableau.
 * @param {string} tableName Nom du tableau à chercher.
 * @param {boolean} [failOnError=true] Si vrai (par défaut), lance une exception
 *  si le tableau n'existe pas. Si faux, renvoie null.
 * @returns {ExcelScript.Table | null} Tableau Excel correspondant au nom donné,
 *  ou null si il n'existe pas.
 */
function getTableOrFail(
    sheetName: string,
    tableName: string,
    failOnError: boolean = true
): ExcelScript.Table | null {
    const sheet = WorkbookService.getSheetOrFail(sheetName, failOnError);
    const table = sheet.getTable(tableName);
    if (!table) {
        const msg = `Le tableau "${tableName}" n'existe pas dans la feuille "${sheetName}".`;
        if (failOnError) throw new Error(msg);
        CONSOLE_WARN.log(msg);
        return null;
    }
    return table;
}

/**
 * Renvoie les données du tableau Excel correspondant au nom donné
 *  dans la feuille de calcul donnée.
 * Si le tableau n'existe pas, renvoie null si failOnError est à false,
 *  sinon lance une exception.
 * @param {string} sheetName Nom de la feuille de calcul où chercher le tableau.
 * @param {string} tableName Nom du tableau à chercher.
 * @param {boolean} [failOnError=true] Si vrai (par défaut),
 *  lance une exception si le tableau n'existe pas. Si faux, renvoie null.
 * @returns {(string | number | boolean)[][]} Données du tableau Excel
 *  correspondant au nom donné, ou null si il n'existe pas.
 */
function getDataFromTable(
    sheetName: string,
    tableName: string,
    failOnError: boolean = true
): (string | number | boolean)[][] {
    const table = getTableOrFail(sheetName, tableName, failOnError);
    return table.getRange().getValues();
}

/**
 * Vérifie si l'adresse de cellule donnée est valide.
 * Si elle est valide, la renvoie telle quelle.
 * Si elle est invalide, lance une exception si failOnError est à true,
 *  sinon renvoie une chaîne vide.
 * @param {string} cellName Adresse de cellule à vérifier.
 * @param {boolean} [failOnError=true] Si vrai (par défaut), lance une exception
 *  si l'adresse est invalide. Si faux, renvoie une chaîne vide.
 * @returns {string} Adresse de cellule si elle est valide, une chaîne vide sinon.
 */
function checkCellName(cellName: string, failOnError: boolean = true): string {
    // Convertit startCell en majuscules pour éviter les problèmes de casse
    cellName = cellName.toUpperCase();

    // Vérifie si cellName est une adresse de cellule valide
    if (!/^([A-Z]+)(\d+)$/.test(cellName)) {
        const msg = `L'adresse de départ ${cellName} n'est pas valide.`;
        if (failOnError) throw new Error(msg);
        CONSOLE_WARN.log(msg);
        return "";
    }
    return cellName;
}

/**
 * Affiche un tableau avec en-têtes et données dans une feuille de calcul Excel.
 * Combine les en-têtes et les données fournies, puis les insère à partir
 *  de la cellule de départ spécifiée. Efface le contenu existant de la plage
 *  de cellules ciblée et supprime tout tableau existant avec le même nom avant
 *  d'ajouter un nouveau tableau avec les données fournies.
 * @param {string[][]} headers En-têtes du tableau.
 * @param {(string | number)[][]} data Données du tableau.
 * @param {string} sheetName Nom de la feuille de calcul où afficher le tableau.
 * @param {string} tableName Nom du tableau à afficher.
 * @param {string} [startCell="A1"] Cellule où commencer à afficher le tableau
 *  (par défaut: "A1").
 * @param {boolean} [failOnError=true] Si vrai (par défaut), lance une exception
 *  si des erreurs surviennent. Si faux, renvoie null.
 * @returns {ExcelScript.Table | null} Tableau Excel créé, ou null si une erreur survient.
 */
function printTable(
    headers: string[][],
    data: (string | number | boolean)[][],
    sheetName: string,
    tableName: string,
    startCell: string = "A1",
    failOnError: boolean = true
): ExcelScript.Table | null {

    // Combine les en-têtes et les données
    const tableData = headers.concat(data);

    // Vérifie si les données sont non vides
    if (tableData.length === 0 || tableData[0].length === 0) {
        const msg = `Aucune donnée à insérer dans la table "${tableName}".`;
        if (failOnError) throw new Error(msg);
        CONSOLE_WARN.log(msg);
        return;
    }

    // Vérifie si un tableau avec le même nom existe déjà et le supprime si nécessaire
    const sheet = WorkbookService.getSheetOrFail(sheetName, failOnError);
    const existingTable = sheet.getTables().find(table => table.getName() === tableName);
    if (existingTable) existingTable.delete();

    // Détermine la plage où écrire les données
    const startRange = sheet.getRange(checkCellName(startCell));
    const writeRange = startRange
        .getResizedRange(tableData.length - 1, tableData[0].length - 1);

    // Efface le contenu de la plage
    writeRange.clear(ExcelScript.ClearApplyTo.contents);

    // Écrit les données dans la plage
    writeRange.setValues(tableData);

    // Ajoute un nouveau tableau
    const table = sheet.addTable(writeRange.getAddress(), true);
    table.setName(tableName);

    CONSOLE_WARN.log(`Le tableau "${tableName}" a été créé avec succès
                        dans la feuille "${sheetName}".`);

    return table;
}

/**
 * Formatte une date en jour, mois et année.
 * @param {number} dateValue Temps en nombre décimal (en jours depuis 1900).
 * @returns {string} Date au format "jj/mm/aaaa".
 */
function formatDate(dateValue: number): string {
    const excelBase = new Date(Date.UTC(1899, 11, 30)); // base d'Excel
    const days = Math.floor(dateValue); // partie entière
    const date = new Date(excelBase.getTime() + days * 86400000);
    const year = date.getUTCFullYear();
    const month = date.getUTCMonth() + 1;
    const day = date.getUTCDate();
    return `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/${year}`;
}

/**
 * Formatte un temps en heure, minutes et secondes.
 * @param {number} timeValue Temps en nombre décimal (par exemple, 0.5 pour 12h00).
 * @param {boolean} [withSeconds=true] Si vrai (par défaut), affiche les secondes.
 *  Si faux, n'affiche que les heures et les minutes.
 * @returns {string} Temps au format "hh:mm" ou "hh:mm:ss".
 */
function formatTime(timeValue: number, withSeconds: boolean = true): string {
    const totalSeconds = Math.round((timeValue - Math.floor(timeValue)) * 24 * 60 * 60); // secondes dans la journée
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`
        + (withSeconds ? `:${seconds.toString().padStart(2, '0')}` : '');
}

/**
 * Ajuste le temps pour tenir compte de l'heure de changement de journée.
 * Si l'heure est inférieure à l'heure de changement de journée,
 *  on ajoute 1 pour passer à la journée suivante
 *  afin de pouvoir la comparer avec une autre heure de la même journée.
 * Par exemple 1h00 devient 25h00.
 * @param {number} time Temps en nombre décimal (par exemple, 0.5 pour 12h00).
 * @returns {number} Temps ajusté.
 */
function adaptTime(time: number): number {
    return (time < PARAM.rolloverHour) ? time + 1 : time;
}

/**
 * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
 * @param {string[][]} allCombinations Liste de toutes les combinaisons de parcours à évaluer.
 * @returns {path: string[], totalDistance: number} Chemin le plus court et sa distance totale,
 *  ou null si aucun chemin n'est trouvé.
 */
function findShortestPath(allCombinations: string[][])
    : { path: string[], totalDistance: number } | null {

    let shortestPath: { path: string[], totalDistance: number } | null = null;

    for (const combination of allCombinations) {
        // Calcule le chemin complet et la distance totale pour la combinaison actuelle
        const { path, totalDistance } = calculateCompletePath(combination);

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
 * @param {string[]} combination Liste ordonnée des gares à parcourir.
 * @returns {path: string[], totalDistance: number} Chemin complet et la distance totale.
 */
function calculateCompletePath(combination: string[])
    : { path: string[], totalDistance: number } {

    const completePath: string[] = [];
    let totalDistance = 0;

    for (const i = 0; i < combination.length - 1; i++) {
        // Trouve le chemin le plus court pour le tronçon actuel
        const segmentPath = dijkstra(combination[i], combination[i + 1]);

        // Si aucun chemin n'est trouvé pour ce tronçon, retourne un chemin vide
        if (segmentPath.length === 0) return { path: [], totalDistance: 0 };

        // Ajoute la distance du tronçon à la distance totale
        totalDistance += calculatePathTime(segmentPath);

        // Retire la première gare du tronçon qui correspond à la dernière gare
        // du tronçon prédédent
        if (completePath.length > 0) {
            segmentPath.shift();
        }

        // Ajoute le chemin du tronçon au chemin complet
        completePath.push(...segmentPath);
    }

    return { path: completePath, totalDistance };
}

/**
 * Calcule le temps total pour un chemin donné en tenant compte des temps de trajet
 *  et des éventuels temps de rebroussement.
 * @param {string[]} path Liste ordonnée des gares constituant le chemin.
 * @returns {number} Temps total du chemin, incluant les temps de trajet
 *  et les temps de rebroussement.
 */
function calculatePathTime(path: string[]): number {
    let totalTime = 0;

    for (const i = 0; i < path.length - 1; i++) {
        const from = path[i];
        const to = path[i + 1];
        const connection = CONNECTIONS.get(from)?.get(to);
        if (connection) {
            totalTime += connection.time;
            // Ajoute le temps de rebroussement sauf pour le premier segment
            if (i > 0 && connection.withTurnaround) {
                totalTime += PARAM.turnaroundTime;
            }
        } else {
            const msg = `Une connexion manque entre "${from}" et "${to}"
                dans le chemin "${path.join(", ")}".`;
            throw new Error(msg);
        }
    }

    return totalTime;
}

/**
 * Cherche le chemin le plus court entre le départ et l'arrivée
 *  en appliquant Dijkstra.
 * @param {string} start Gare de départ.
 * @param {string} end Gare d'arrivée.
 * @returns {string[]} Chemin le plus court.
 */
function dijkstra(start: string, end: string): string[] {
    const distances = new Map<string, number>();
    const previousNodes = new Map<string, string | null>();
    const unvisited = new Set<string>(CONNECTIONS.keys());
    const path: string[] = [];

    // Initialise les distances
    for (const node of unvisited) {
        distances.set(node, Infinity);
        previousNodes.set(node, null);
    }
    distances.set(start, 0);

    while (unvisited.size > 0) {
        const currentNode = Array.from(unvisited).reduce((minNode, node) =>
            distances.get(node)! < distances.get(minNode)! ? node : minNode
        );

        if (distances.get(currentNode) === Infinity) break; // Aucun chemin

        unvisited.delete(currentNode);

        // Examine les voisins avec les nouveaux attributs
        for (const [neighbor, connexion] of CONNECTIONS.get(currentNode) || []) {
            let additionalTime = connexion.time;
            if (connexion.withTurnaround && currentNode !== start) {
                // Si un rebroussement est nécessaire, ajoute le temps de retournement
                additionalTime += PARAM.turnaroundTime;
            }

            const newDist = distances.get(currentNode)! + additionalTime;
            if (newDist < distances.get(neighbor)!) {
                distances.set(neighbor, newDist);
                previousNodes.set(neighbor, currentNode);
            }
        }
    }

    // Retrace le chemin
    const step: string | null | undefined = end;
    while (step) {
        path.unshift(step);
        step = previousNodes.get(step);
    }

    // Retourne le chemin s'il est valide
    return path[0] === start ? path : [];
}

/**
 * Génère toutes les combinaisons de routes possibles pour aller de start à end
 *  en passant par les gares intermédiaires via, classées dans l'ordre ou non.
 * @param {string} start Gare de départ.
 * @param {string} end Gare d'arrivée.
 * @param {string[]} via Gares intermédiaires à passer par.
 * @param {boolean} [viaSorted=false] Si vrai, les gares intermédiaires sont classées
 *  et l'ordre conservé. Si faux (par défaut), toutes les combinaisons possibles
 *  de gares intermédiaires sont calculées.
 * @returns {string[][]} Combinaisons de routes possibles.
 */
function generateCombinations(start: string, end: string, via: string[],
    viaSorted: boolean = false): string[][] {

    // Filtre les gares intermédiaires pour éliminer les chaînes vides
    const filteredVia = via.filter(v => v.trim() !== "");

    // Supprime les premiers et derniers éléments de via s'ils correspondent à start et end
    if (filteredVia[0] === start) {
        filteredVia.shift();
    }
    if (filteredVia[filteredVia.length - 1] === end) {
        filteredVia.pop();
    }

    // Génère les permutations des gares intermédiaires, sauf si viaSorted est vrai
    const viaPermutations = viaSorted ? permute(filteredVia) : [filteredVia];

    // Ajoute start au début et end à la fin de chaque permutation
    const routes: string[][] = viaPermutations.map(permutation => [start, ...permutation, end]);

    // Étend chaque route pour inclure toutes les variantes possibles
    return routes.flatMap((route: string[]) => expandPermutations(route));
}

/**
 * Génère toutes les permutations possibles d'un tableau de chaînes.
 * @param {string[]} array Chaînes à permuter.
 * @returns {string[][]} Permutations de chaînes possibles.
 */
function permute(array: string[]): string[][] {
    if (array.length === 0) return [[]];
    if (array.length === 1) return [[array[0]]];

    const result: string[][] = [];

    for (const i = 0; i < array.length; i++) {
        const rest = [...array.slice(0, i), ...array.slice(i + 1)];
        const restPermutations = permute(rest);

        for (const perm of restPermutations) {
            result.push([array[i], ...perm]);
        }
    }

    return result;
}

/**
 * Étend une permutation de gares pour inclure toutes les variantes possibles.
 * Une variante correspond au sens de passage dans la gare :
 *  GARE_#pair en pair, GARE_#impair en impair.
 * Seules les gares de retournement permettent de passer d'une variante à l'autre.
 * Si la gare a déjà un suffixe imposé (_), renvoie uniquement cette variante.
 * Sinon, renvoie toutes les variantes associées, incluant celles des gares filles
 *  (gares dont la gare de référence est la gare demandée).
 * @param {string[]} permutation Permutation de gares à étendre.
 * @returns {string[][]} Permutations possibles avec toutes les variantes.
 */
function expandPermutations(permutation: string[]): string[][] {
    if (permutation.length === 0) return [[]];
    const first = getAllVariants(permutation[0]);

    const restExpanded = expandPermutations(permutation.slice(1));

    const result: string[][] = [];
    for (const f of first) {
        for (const r of restExpanded) {
            result.push([f, ...r]);
        }
    }

    return result;
}

/**
 * Renvoie toutes les variantes possibles pour une gare.
 * Une variante correspond au sens de passage dans la gare :
 *  GARE_#pair en pair, GARE_#impair en impair.
 * Seules les gares de retournement permettent de passer d'une variante à l'autre.
 * Si la gare a déjà un suffixe imposé (_), renvoie uniquement cette variante.
 * Sinon, renvoie toutes les variantes associées, incluant celles des gares filles
 *  (gares dont la gare de référence est la gare demandée).
 * @param {string} gare Gare dont on cherche les variantes.
 * @returns {string[]} Variantes possibles pour la gare.
 */
function getAllVariants(gare: string): string[] {
    // Recherche la gare demandée
    const station = STATIONS.get(gare.split('_')[0]);
    if (!station) return [];

    // Si la gare a un suffixe (_), renvoie uniquement [gare]
    if (gare.includes('_')) return [gare];

    // Sinon, renvoie les variantes pour les 2 sens, 
    return [
        ...[station, ...station.childStations]
            .filter(v => v.abbreviation.trim() !== '')
            .map(v => [`${v.abbreviation}_${PARAM.parityDigits.get(Parity.odd)}`,
            `${v.abbreviation}_${PARAM.parityDigits.get(Parity.even)}`])
    ].reduce((acc, curr) => acc.concat(curr), []);
}

/* Interface ParamStructure qui liste les différents paramètres */
interface ParamStructure {
    /* Liste des paramètres a déjà été chargée */
    loaded: boolean;
    /* Nombre de connexions maximum pour une gare */
    maxConnectionNumber: number;
    /* Temps de retournement */
    turnaroundTime: number;
    /* Appelation pour indiquer qu'un arrêt est le terminus */
    terminusName: string;
    /* Heure de changement de journée */
    rolloverHour: number;
    /* Regex des numéros de trains W */
    wTrainsRegex: RegExp;
    /* Regex des numéros de trains que l'on abrège de 6 à 4 chiffres */
    trainsToAbreviateTo4DigitsRegex: RegExp;
    /* Jours de la semaine et leur différentes appelations */
    days: Map<string, Day>;
    /* Lettres décrivant chaque parité */
    parityLetters: Map<number, string>;
    /* Nombres décrivant chaque parité */
    parityDigits: Map<number, number>;
}

/* Liste des paramètres. */
const PARAM: ParamStructure = {
    loaded: false,
    maxConnectionNumber: 0,
    turnaroundTime: 0,
    terminusName: "",
    rolloverHour: 0,
    wTrainsRegex: new RegExp(""),
    trainsToAbreviateTo4DigitsRegex: new RegExp(""),
    days: new Map<string, Day>(),
    parityLetters: new Map<number, string>(),
    parityDigits: new Map<number, number>(),
};

const PARAM_SHEET = "Param";
const PARAM_TABLE = "Paramètres";
const PARAM_ROW_MAX_CONNEXIONS_NUMBER = 1;
const PARAM_ROW_TURNAROUND_TIME = 2;
const PARAM_ROW_TERMINUS_NAME = 3;
const PARAM_ROW_ROLLOVER_HOUR = 4;

/**
 * Charge les paramètres du tableau "Paramètres" de la feuille "Param".
 * Si PARAM.loaded est true et que erase est false, ne fait rien.
 * Charge les paramètres dans l'objet PARAM et met à jour son champ "loaded".
 * Appelle les fonctions loadWTrainsRegex(), load4DigitsAbreviatedTrainsRegex()
 *  et Day.loadFromExcel().
 * @param {boolean} [erase=false] Si vrai, force le rechargement des paramètres.
 *  Si faux (par défaut), ne recharge pas si déjà chargé.
 */
function loadParams(erase: boolean = false) {
    if (PARAM.loaded && !erase) return;
    const data = getDataFromTable(PARAM_SHEET, PARAM_TABLE);

    // Extrait les valeurs
    PARAM.maxConnectionNumber = data[PARAM_ROW_MAX_CONNEXIONS_NUMBER][1] as number;
    PARAM.turnaroundTime = data[PARAM_ROW_TURNAROUND_TIME][1] as number;
    PARAM.terminusName = String(data[PARAM_ROW_TERMINUS_NAME][1]);
    PARAM.rolloverHour = Number(data[PARAM_ROW_ROLLOVER_HOUR][1]) % 1;

    loadWTrainsRegex();
    load4DigitsAbreviatedTrainsRegex();
    Day.loadFromExcel();
    loadParity();

    PARAM.loaded = true;
}

const W_SHEET = "Param";
const W_TABLE = "W";

/**
 * Classe TrainNumber qui défini un numéro de train.
 * Il est alphanumérique, sans ponctuation et sans espaces,
 *  avec un chiffre pour dernier caractère.
 */
class TrainNumber {

    number: string;

    constructor(number: string | number) {
        this.number = number.toString();
    }
}
/**
 * Charge les motifs W depuis la feuille "Param" et les transforme en regex partielles.
 * Ensuite, crée une regex globale combinée des motifs W.
 */
function loadWTrainsRegex() {
    const data = getDataFromTable(W_SHEET, W_TABLE);

    // Transforme chaque motif en regex partielle
    const regexParts: string[] = data
        .flat()
        .filter(v => typeof v === "string" && v.trim() !== "")
        .map(pattern => {
            return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
        });

    // Crée une regex globale combinée
    PARAM.wTrainsRegex = new RegExp(regexParts.join('|'));
}

/**
 * Teste si un train est W (vide voyageur).
 * @param {string} trainNumber Numéro du train à tester.
 * @returns {boolean} Vrai si le train est W, faux sinon.
 */
function isWTrain(trainNumber: string): boolean {
    if (!PARAM.wTrainsRegex) loadWTrainsRegex(); // Charge si non encore fait
    return PARAM.wTrainsRegex.test(trainNumber);
}

const TRAINS_4DIGITS_SHEET = "Param";
const TRAINS_4DIGITS_TABLE = "LigneC4chiffres";

/**
 * Charge les motifs des trains commerciaux que l'on abrège de 6 à 4 chiffres
 *  depuis la feuille "Param" et les transforme en regex partielles.
 * Ensuite, crée une regex globale combinée des motifs de trains abrégeables.
 */
function load4DigitsAbreviatedTrainsRegex() {
    const data = getDataFromTable(TRAINS_4DIGITS_SHEET, TRAINS_4DIGITS_TABLE);

    // Transforme chaque motif en regex partielle.
    const regexParts: string[] = data
        .flat()
        .filter(v => typeof v === "string" && v.trim() !== "")
        .map(pattern => {
            return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
        });

    // Crée une regex globale combinée
    PARAM.trainsToAbreviateTo4DigitsRegex = new RegExp(regexParts.join('|'));
}


/**
 * Abrège un numéro de train de 6 à 4 chiffres pour les trains commerciaux de la ligne C.
 * S'adapte aux numéros à double parité.
 * @param {string | number} trainNumber Numéro du train à transformer.
 * @returns {string | number} Numéro de train abrégé de 6 à 4 chiffres
 *  si le train est commercial, sinon le numéro de train original.
 */
function abreviateTo4Digits(trainNumber: string): string {
    const trainNumberToAnalyse = trainNumber.split("/");
    return (PARAM.trainsToAbreviateTo4DigitsRegex.test(trainNumberToAnalyse[0])
        ? trainNumberToAnalyse[0].substring(2) : trainNumberToAnalyse)
        + (trainNumberToAnalyse.length > 1 ? "/" + trainNumberToAnalyse[1] : "");
}

/**
 * Adapte le numéro de train en fonction de la parité spécifiée.
 * Analyse le dernier chiffre du numéro de train pour déterminer sa parité.
 * Retourne le numéro de train avec un dernier chiffre ajusté selon la parité demandée :
 *  - si la parité demandée est paire (Parity.even), le numéro de train est inchangé
 *   s'il est pair, et décrémenté de 1 s'il est impair.
 *  - si la parité demandée est impaire (Parity.odd), le numéro de train est inchangé
 *   s'il est impair, et incrémenté de 1 s'il est pair.
 *  - si la parité demandée est double (Parity.double), le numéro de train est donné
 *   par sa valeur paire, suivi d'un '/' et du chiffre impair suivant.
 * Si le dernier chiffre du numéro de train est invalide,
 *  un avertissement est enregistré et une chaîne vide est retournée.
 * @param {number} parity Parité demandée (paire, impaire, double).
 * @param {string} trainNumber Numéro du train à adapter.
 * @returns {string} Numéro de train ajusté selon la parité spécifiée.
 */
function renameTrainNumberWithParity(parity: number, trainNumber: string): string {
    // Renvoie le numéro de train si la parité est nulle ou indéfinie
    if (!parity) return trainNumber;
    // Analyse le dernier chiffre du numéro de train
    const trainNumberToAnalyse = trainNumber.split("/")[0];
    const lastDigit = parseInt(trainNumberToAnalyse.slice(-1));
    if (isNaN(lastDigit) || lastDigit <= 0) {
        CONSOLE_WARN.log(`Le dernier chiffre du numéro de train est invalide.`);
        return "";
    }
    const restOfNumber = trainNumberToAnalyse.slice(0, -1);
    const evenLastDigit = lastDigit - (lastDigit % 2);

    // Adapte le numéro de train en fonction de la parité demandée
    switch (parity) {
        case Parity.even:
            return restOfNumber + evenLastDigit;
        case Parity.odd:
            return restOfNumber + (evenLastDigit + 1);
        case Parity.double:
            return restOfNumber + evenLastDigit + '/' + (evenLastDigit + 1);
        default:
            return trainNumber;
    }
}

/**
 * Classe Day qui défini les jours de la semaine individuellement
 *  ou les groupes de jours (JOB du lundi au vendredi, WE pour samedi et dimanche...).
 */
class Day {

    /* Cache pour l'extraction des jours de la semaine depuis une chaine de caractères. */
    private static readonly CACHE = new Map<string, Map<string, number[]>>();

    fullName: string;            // Nom du jour ou du groupe de jours de la semaine
    abreviation: string;         // Abréviation du jour ou du groupe de jours de la semaine
    numbersString: string;       // Numéro(s) concaténés des jours de la semaine
    //     en chaine de caractères (avec ou sans ponctuation)
    number: number;              // Numéro du jour de la semaine (de 1 : lundi à 7 : dimanche,
    //     0 si l'objet est un groupe de jours)

    /**
     * Constructeur de la classe Day.
     * @param {string} numbersString Chaîne de caractères contenant
     *  les numéros de jours de la semaine.
     * @param {string} fullName Nom complet du jour ou du groupe de jours de la semaine.
     * @param {string} abreviation Abréviation du jour ou du groupe de jours de la semaine.
     */
    constructor(numbersString: string, fullName: string, abreviation: string) {
        this.numbersString = Day.cleanAndSortNumbers(numbersString.toString()).join('');
        this.number = parseInt(this.numbersString);
        this.number = this.number > 7 ? 0 : this.number;
        this.fullName = fullName;
        this.abreviation = abreviation;
    }

    /**
     * Nettoie et trie une chaîne de chiffres.
     * Supprime les caractères non numériques et non compris entre 1 et 7,
     * puis trie les chiffres dans l'ordre.
     * @param {string} numbersString Chaîne de caractères contenant des chiffres.
     * @returns {string} Chaîne de caractères contenant les chiffres triés.
     */
    private static cleanAndSortNumbers(numbersString: string): number[] {
            return [...new Set(
                numbersString
                    .replace(/[^1-7]/g, '')     // Supprime les caractères non numériques
                                                //  et non compris entre 1 et 7
                    .split('')                  // Divise la chaîne en un tableau de chiffres
                    .map((x) => Number(x))      // Convertit les caractères en nombres
            )].sort((a, b) => a - b);           // Trie les chiffres dans l'ordre
    }

    /**
     * Extrait les jours d'une ou deux chaînes de noms de jours
     *  en tableau de numéros de jours (1 à 7).
     * Utilise un cache pour éviter de recalculer les résultats pour les mêmes combinaisons.
     * Si deux chaînes sont fournies, retourne l'intersection des jours correspondants.
     * @param {string} input1 Chaîne contenant des noms, numéros ou abréviations de jours
     *  séparés ou non par de la ponctuation  (ex : "lundi;mer7")
     * @param {string} input2 (optionnel) Deuxième chaîne pour calculer
     *  l'intersection des jours.
     * @returns Tableau trié de numéros de jours (sans doublons), ex : [1, 3].
     */
    public static extractFromString(input1: string, input2: string = ''): number[] {
        const key1 = String(input1).toLowerCase();
        const key2 = String(input2).toLowerCase();

        // Vérifie le cache
        if (Day.CACHE.has(key1) && Day.CACHE.get(key1)!.has(key2)) {
            return Day.CACHE.get(key1)!.get(key2)!;
        }

        // Analyse la chaine
        let result: number[] = [];
        if (!key2) {

            // Analyse la chaine pour la transformer en tableau
            let processed = key1;

            // Reconnaissance Regex des noms entiers de jours
            Array.from(PARAM.days.values())
                .sort((a, b) => b.fullName.length - a.fullName.length)
                .forEach(day => {
                    const regex = new RegExp(day.fullName.toLowerCase(), 'g');
                    processed = processed.replace(regex, day.numbersString);
                });

            // Reconnaissance Regex des abréviations de jours
            Array.from(PARAM.days.values())
                .sort((a, b) => b.abreviation.length - a.abreviation.length)
                .forEach(day => {
                    const regex = new RegExp(day.abreviation.toLowerCase(), 'g');
                    processed = processed.replace(regex, day.numbersString);
                });

            // Reconnaissance Regex des numéros de jours
            processed = processed.replace(/[^1-7]/g, '');

            // let processed = key1;
            // PARAM.days.forEach((day) => {

            //     const regex = new RegExp(
            //         day.numbersString + '|' +
            //         day.abreviation.toLowerCase() + '|' +
            //         day.fullName.toLowerCase() ,
            //         'g'
            //       );     
            //     processed = processed.replace(regex, `${day.numbersString}`);
            // });

            result = Day.cleanAndSortNumbers(processed);
        } else {
            // Calcule l'intersection des deux chaines.
            const days1 = Day.extractFromString(key1);
            const days2 = Day.extractFromString(key2);

            result = days1.filter(n => days2.includes(n));
        }

        // Met en cache le resultat pour une utilisation similaire de la fonction
        if (!Day.CACHE.has(key1)) {
            Day.CACHE.set(key1, new Map<string, number[]>());
        }
        Day.CACHE.get(key1)!.set(key2, result);

        return result;
    }

    /* Constantes de lecture du tableau Excel. */
    private static readonly SHEET = "Param";
    private static readonly TABLE = "Jours";
    private static readonly COL_FULL_NAME = 0;
    private static readonly COL_ABBREVIATION = 1;
    private static readonly COL_NUMBERS = 2;

    /**
     * Charge les jours de la semaine à partir du tableau "Jours" de la feuille "Param".
     * Les jours sont stockés dans la structure PARAM.days, sous forme de map, avec
     *  le nom complet et l'abréviation du jour comme clés, et leur numéro correspondant
     *  comme valeur.
     */
    public static loadFromExcel() {

        const data = getDataFromTable(Day.SHEET, Day.TABLE);

        for (const row of data.slice(1)) {
            // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides)
            if (row.every(cell => !cell)) continue;

            // Extrait les valeurs
            const numbersString = String(row[Day.COL_NUMBERS]);
            const fullName = String(row[Day.COL_FULL_NAME]);
            const abreviation = String(row[Day.COL_ABBREVIATION]);

            // Crée l'objet Day
            const day = new Day(numbersString, fullName, abreviation);
            PARAM.days.set(day.numbersString, day);
        }
    }
}


/*
 * Classe Parity qui permet de manipuler la parité
 *  d'un train, d'un sillon ou d'un arrêt.
 */
class Parity {

    /* Constantes de parité/ */
    public static readonly odd: number = 1;         // Parité impaire
    public static readonly even: number = 2;        // Parité paire
    public static readonly double: number = -2;     // Parité double
    public static readonly undefined: number = -1;  // Parité non définie

    private _value: number;                         // Valeur de la parité
    private doubleParityAllowed: boolean;           // Autorise une double parité

    /**
     * Constructeur de la classe Parity.
     * Initialise une instance de parité avec une valeur spécifiée,
     *  qui peut être une lettre de parité, un chiffre de parité, ou un numéro de train.
     * Analyse la valeur donnée pour déterminer la parité.
     * @param {string | number} value Valeur à analyser pour la parité.
     * @param {boolean} [doubleParityAllowed=false] Si vrai, la double parité est autorisée.
     *  Si faux (par défaut), la double parité est impossible.
     */
    constructor(value: string | number = Parity.undefined,
        doubleParityAllowed: boolean = false) {
        this.doubleParityAllowed = doubleParityAllowed;
        this._value = this.analyseValue(value);
    }

    /**
     * Retourne la valeur de la parité.
     * @returns {number} Valeur de la parité.
     */
    get value(): number {
        return this._value;
    }

    /**
     * Modifie la valeur de la parité avec une nouvelle valeur parmi les constantes de parité.
     * Si la valeur est Parity.double et que doubleParityAllowed est faux,
     *  la valeur est modifiée en Parity.undefined.
     * @param {number} value Nouvelle valeur de la parité.
     */
    set value(value: number) {
        this._value = this.analyseValue(value);
    }

    /**
     * Met à jour la valeur de la parité avec une valeur spécifiée,
     *  qui peut être une lettre de parité, un chiffre de parité, ou un numéro de train.
     * Analyse la valeur fournie pour déterminer la parité correspondante.
     * @param {string | number} value Nouvelle valeur à analyser pour la parité.
     */
    public update(value: string | number) {
        this.value = this.analyseValue(value);
    }

    /**
     * Vérifie si la parité est identique à celle d'une autre variable de parité.
     * @param {Parity} other Autre variable de parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont identiques, faux sinon.
     */
    public equals(other: Parity): boolean {
        return this._value === other._value;
    }

    /**
     * Vérifie si la parité est identique à une autre valeur de parité.
     * @param {number} parity Autre valeur de parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont identiques, faux sinon.
     */
    public is(parity: number): boolean {
        return this._value === parity;
    }

    /**
     * Inverse la parité actuelle.
     * Si la parité actuelle est paire, elle devient impaire, et inversement.
     * Si la parité actuelle est double, elle reste double.
     * Si la parité actuelle est indéfinie, elle reste inchangée.
     * @returns {Parity} Parité inversée.
     */
    public invert(): Parity {
        switch (this._value) {
            case Parity.odd:
                this._value = Parity.even;
                break;
            case Parity.even:
                this._value = Parity.odd;
                break;
            case Parity.double:
                this._value = Parity.double;
                break;
            default:
                this._value = Parity.undefined;
                break;
        }
        return this;
    }

    /**
     * Retourne le chiffre de parité correspondant.
     * Si withUnderscores est true, le chiffre est précédé
     *  d'un underscore pour les parités impaires et paires.
     * @param {boolean} [withUnderscores=false] Si vrai, le chiffre est précédé
     *  d'un underscore. Si faux (par défaut), seul le chiffre est retourné.
     * @returns {string | number} Le chiffre de parité, ou une chaîne vide
     *  si la parité est indéfinie.
     */
    public printDigit(withUnderscores: boolean = false): string | number {
        switch (this._value) {
            case Parity.odd:
                return withUnderscores ? '_' : '' + PARAM.parityDigits.get(Parity.odd);
            case Parity.even:
                return withUnderscores ? '_' : '' + PARAM.parityDigits.get(Parity.even);
            case Parity.double:
                return PARAM.parityDigits.get(Parity.double);
            default:
                return "";
        }
    }

    /**
     * Retourne la lettre de parité correspondante
     *  (parité impaire ou paire, concaténation impaire puis paire si parité double).
     * @returns {string} La lettre de parité correspondante, ou une chaîne vide
     *  si la parité est double ou indéfinie.
     */
    public printLetter(): string {
        switch (this._value) {
            case Parity.odd:
                return PARAM.parityLetters.get(Parity.odd);
            case Parity.even:
                return PARAM.parityLetters.get(Parity.even);
            case Parity.double:
                return PARAM.parityLetters.get(Parity.odd)!
                    + PARAM.parityLetters.get(Parity.even)!;
            default:
                return "";
        }
    }

    /**
     * Vérifie si une chaîne de caractères contient la lettre de parité correspondante,
     *  ou les deux lettres si la parité est double.
     * @param {string} string Chaîne de caractères à analyser.
     * @param {number} parity Parité à chercher.
     * @returns {boolean} Vrai si la chaîne de caractères contient la lettre de parité,
     *  faux sinon.
     */
    public static containsParityLetter(string: string, parity: number): boolean {
        switch (parity) {
            case Parity.odd:
                return string.toUpperCase().includes(PARAM.parityLetters.get(Parity.odd)!);
            case Parity.even:
                return string.toUpperCase().includes(PARAM.parityLetters.get(Parity.even)!);
            case Parity.double:
                return string.toUpperCase().includes(PARAM.parityLetters.get(Parity.odd)!)
                    && string.toUpperCase().includes(PARAM.parityLetters.get(Parity.even)!);
            default:
                return false;
        }
    }

    /**
     * Analyse une valeur qui indique la parité, qui peut être :
     *  - la lettre de parité (ou la concaténation des deux lettres sans ordre si double parité)
     *  - le chiffre de parité (format chaîne ou nombre)
     *  - un numéro de train (pair, impair ou double s'il contient un '/')
     * @param {string | number} value Nouvelle valeur de la parité.
     */
    private analyseValue(value: string | number): number {
        // Convertit la valeur en chaine avec majuscules
        const valueToString = value.toString().toUpperCase();
        switch (valueToString) {
            case PARAM.parityLetters.get(Parity.odd):
            case PARAM.parityDigits.get(Parity.odd)!.toString():
                return Parity.odd;
            case PARAM.parityLetters.get(Parity.even):
            case PARAM.parityDigits.get(Parity.even)!.toString():
                return Parity.even;
            case PARAM.parityLetters.get(Parity.even)! + PARAM.parityLetters.get(Parity.odd)!:
            case PARAM.parityLetters.get(Parity.odd)! + PARAM.parityLetters.get(Parity.even)!:
            case PARAM.parityDigits.get(Parity.double)!.toString():
                return this.doubleParityAllowed ? Parity.double : Parity.undefined;
            case '0':
            case "":
                return Parity.undefined;
            default:
                // Si numéro de train avec double parité implicite (contient un '/'),
                //      retourne Parity.double si autorisé, ou Parity.undefined
                if (valueToString.includes('/')) {
                    return this.doubleParityAllowed ? Parity.double : Parity.undefined;
                }
                // Calcule la parité avec le numéro de train
                switch (parseInt(value.toString().slice(-1)) % 2) {
                    case 1:
                        return Parity.odd;
                    case 0:
                        return Parity.even;
                    default:
                        return Parity.undefined;
                }
        }
    }

    /**
     * Adapte le numéro du train en fonction de la parité demandée.
     * Si le numéro du train est pair, il est inchangé si la parité demandée est paire,
     *  et incrémenté de 1 si la parité demandée est impaire.
     * Si le numéro du train est impair, il est décrémenté de 1 si la parité demandée est paire,
     *  et inchangé si la parité demandée est impaire.
     * Si la parité demandée est indéfinie, le numéro du train est inchangé.
     * @param {string} trainNumber Numéro du train, qui peut être un nombre
     *  ou une chaîne de caractères.
     * @param {boolean} with4Digits Si vrai, le numéro du train est abrégé à 4 chiffres.
     *  Si faux, le numéro du train n'est pas abrégé.
     * @returns {string} Numéro du train adapté
     */
    public adaptTrainNumber(trainNumber: string, with4Digits: boolean): string {
        return renameTrainNumberWithParity(this.value,
            with4Digits ? abreviateTo4Digits(trainNumber) : trainNumber);
    }
}

const PARITY_SHEET = "Param";
const PARITY_TABLE = "Parité";
const PARITY_ROW_ODD = 1;
const PARITY_ROW_EVEN = 2;
const PARITY_ROW_DOUBLE = 3;
const PARITY_COL_LETTER = 1;
const PARITY_COL_NUMBER = 2;

/**
 * Charge les paramètres de parité des jours (lettre et chiffre associés
 *  aux jours pairs et impairs) à partir de la feuille "Param".
 * Les paramètres sont stockés dans l'objet PARAM.parity.
 */
function loadParity() {

    const data = getDataFromTable(PARITY_SHEET, PARITY_TABLE);

    PARAM.parityLetters.set(Parity.odd,
        String(data[PARITY_ROW_ODD][PARITY_COL_LETTER]).toUpperCase() || "I");
    PARAM.parityLetters.set(Parity.even,
        String(data[PARITY_ROW_EVEN][PARITY_COL_LETTER]).toUpperCase() || "P");
    PARAM.parityDigits.set(Parity.odd,
        Number(data[PARITY_ROW_ODD][PARITY_COL_NUMBER]) || 1);
    PARAM.parityDigits.set(Parity.even,
        Number([PARITY_ROW_EVEN][PARITY_COL_NUMBER]) || 2);
    PARAM.parityDigits.set(Parity.double,
        Number([PARITY_ROW_DOUBLE][PARITY_COL_NUMBER]) || -2);
}

/**
 * Classe TrainPath qui définit un sillon, avec ses gares et horaires de passage,
 * plannifié sur un ou plusieurs jours de la semaine, ou sur plusieurs dates précises,
 * avec les mêmes horaires.
 * Plusieurs trains circulant un des jours concerné y font référence.
 */
class TrainPath {
    number: string;                     // Numéro du sillon/train
    days: string;                       // Combinaison des jours de circulation du sillon,
    //  ou liste des dates de circulation (commencent par #)
    trainsByDay: Map<number, Train>;    // Liste des trains associés à un jour donné
    doubleParity: boolean;              // Indique si le sillon a une double parité
    lineDirection: Parity;              // Direction du sillon sur la ligne
    //  (donnée par une parité globale)
    missionCode: string;                // Code de mission des trains du sillon
    departureTime: number;              // Heure de départ du sillon (hors évolution)
    departureStation: string;           // Gare de départ du sillon (hors évolution)
    arrivalTime: number;                // Heure d'arrivée du sillon
    arrivalStation: string;             // Gare d'arrivée du sillon
    firstStation: string;               // Gare de départ du sillon incluant les évolutions
    firstStop?: Stop;                   // Arrêt à la gare de départ avec les évolutions
    lastStation: string;                // Gare d'arrivée du sillon incluant les évolutions
    lastStop?: Stop;                    // Arrêt à la gare d'arrivée avec les évolutions
    viaStations: string[];              // Gares intermédiaires du sillon (via)
    stops: Map<string, Stop>            // Arrêts du sillon
    stopsChecked?: number;              // Arrêts vérifiés :
    //  - 1 : uniquement les gares de départ et d'arrivée,
    //  - 2 : arrêts commerciaux dans l'ordre,
    //  - 3 : tous les arrêts et gares de passage du sillon
    //       (suite à findPath)

    constructor(
        number: string | number,
        lineDirection: number = 0,
        days: string = "",
        missionCode: string = "",
        departureTime: number = 0,
        departureStation: string = "",
        arrivalTime: number = 0,
        arrivalStation: string = "",
        firstStation: string = "",
        lastStation: string = "",
        viaStations: string = ""
    ) {
        this.number = number.toString();
        this.lineDirection = new Parity(lineDirection, true);
        this.days = days;
        this.trainsByDay = new Map<number, Train>();
        this.missionCode = missionCode;
        this.doubleParity = false;
        this.departureTime = departureTime;
        this.departureStation = departureStation;
        this.arrivalTime = arrivalTime;
        this.arrivalStation = arrivalStation;
        this.firstStation = firstStation;
        this.lastStation = lastStation;
        this.viaStations = viaStations ? viaStations.split(';') : [];
        this.stops = new Map<string, Stop>();
    }

    /**
     * Vérifie la validité de l'objet TrainPath en envoyant un message d'erreur si :
     *  - la gare de départ ou d'arrivée est vide ou inconnue,
     *  - les heures de départ et d'arrivée sont invalides,
     * @returns {Train | undefined} Objet TrainPath s'il est valide, undefined sinon.
     */
    check(): TrainPath | undefined {

        // Vérifie si les gares de départ et d'arrivée existent
        if (!this.departureStation) {
            CONSOLE_WARN.warn(`Le sillon ${this.number}_${this.days}`
                + ` n'a pas de gare de départ.`);
            return undefined;
        } else if (!this.arrivalStation) {
            CONSOLE_WARN.warn(`Le sillon ${this.number}_${this.days}`
                + ` n'a pas de gare d'arrivée.`);
            return undefined;
        }
        const departureStationObj: Station | undefined = STATIONS.get(this.departureStation);
        const arrivalStationObj: Station | undefined = STATIONS.get(this.arrivalStation);
        if (!departureStationObj) {
            CONSOLE_WARN.warn(`Le sillon ${this.number}_${this.days} a une gare de départ`
                + ` inconnue : ${this.departureStation}.`);
            return undefined;
        } else if (!arrivalStationObj) {
            CONSOLE_WARN.warn(`Le sillon ${this.number}_${this.days} a une gare d'arrivée`
                + ` inconnue : ${this.arrivalStation}.`);
            return undefined;
        }

        // Vérifie si les heures de départ et d'arrivée existent
        if (!this.departureTime) {
            CONSOLE_WARN.warn(`Le sillon ${this.number}_${this.days} n'a pas d'heure`
                + ` de départ à la gare ${this.departureStation}.`);
            return undefined;
        } else if (!this.arrivalTime) {
            CONSOLE_WARN.warn(`Le sillon ${this.number}_${this.days} n'a pas d'heure`
                + ` d'arrivée à la gare ${this.arrivalStation}.`);
            return undefined;
        }

        // Détermine si la gare d'arrivée change de parité,
        //  auquel cas le sillon a une double parité
        const parity = new Parity(this.number, false);
        this.doubleParity = (parity.is(Parity.double))
            || (departureStationObj!.reverseLineDirection
                !== arrivalStationObj!.reverseLineDirection);

        // Détermine le sens principal pour la ligne C s'il n'est pas déjà donné
        if (!this.lineDirection.value) {
            this.lineDirection = new Parity(this.number, false);
            if (departureStationObj!.reverseLineDirection) this.lineDirection.invert();
        }

        return this;
    }

    /**
     * Retourne la clé du sillon qui est composée du numéro du sillon/train
     *  suivi de la liste des jours de circulation ou de la première date de circulation.
     * @returns {string} Clé du sillon plannifié.
     */
    get key(): string {
        return `${this.number}_${this.days.toString().split(';')[0]}`;
    }

    /**
     * Ajoute un arrêt au sillon.
     * Si les trains du sillon sont déjà passés par l'arrêt et que erase est faux,
     *  lance une erreur.
     * @param {Stop} stop Arrêt à ajouter.
     * @param {boolean} [erase=false] Si vrai, remplace l'arrêt s'il existe déjà. Si faux
     *  (par défaut), le nouvel arrêt n'est pas pris en compte.
     * @returns {Stop} L'arrêt ajouté, ou null si une erreur a été levée.
     * @throws {Error} Si les trains du sillon sont déjà passé par l'arrêt
     *  et que erase est faux.
     */
    addStop(stop: Stop, erase: boolean = false): Stop {
        const stopKey = stop.key;
        if (this.stops.has(stop.key) && !erase) {
            const msg = `L'arrêt "${stop.key}" est déjà associé aux trains`
                + ` du sillon ${this.key}. Un même train ne peut pas revenir`
                + ` dans la même gare et avec le même sens.`;
            throw new Error(msg);
        }
        this.stops.set(stop.key, stop);
        if (stop)
            switch (stop.key) {
                case this.firstStation:
                    this.firstStop = stop;
                    break;
                case this.departureStation:
                    if (!this.firstStation) this.firstStop = stop;
                    break;
                case this.lastStation:
                    this.lastStop = stop;
                    break;
                case this.arrivalStation:
                    if (!this.lastStation) this.lastStop = stop;
                    break;
            }
        return stop;
    }

    /**
     * Renvoie l'arrêt associé au sillon et correspondant au nom de gare donné :
     *  - si l'arrêt est un terminus, renvoie null.
     *  - si la parité n'est pas donnée dans la demande, la recherche est faite sans parité,
     *   puis avec la parité du train, puis avec la parité inverse.
     *  - si la parité est donnée dans la demande, la recherche est faite avec la parité,
     *   puis sans.
     * Dans le cas où l'arrêt et trouvé et updateParity est vrai,
     *  la parité de l'arrêt est mise à jour avec celle indiquée dans la demande.
     * Si l'arrêt n'est pas trouvé avec la gare donnée, la recherche est ensuite
     *  effectuée avec la gare de référence, puis avec les gares filles.
     * Si l'arrêt est trouvé avec une gare fille et que updateWithChildStationName est vrai,
     *  le nom de l'arrêt trouvé est mis à jour.
     * Enfin si l'arrêt n'est pas trouvé, renvoie null.
     * @param {string} stopName Nom de la gare de l'arrêt.
     * @param {boolean} [updateParity=false] Si vrai, met à jour la parité de l'arrêt
     *  si elle est fournie. Si faux (par défaut), la parité de l'arrêt n'est pas modifiée.
     * @param {boolean} [updateWithChildStationName=false] Si vrai, met à jour
     *  le nom de l'arrêt avec le nom de la gare fille trouvée. Si faux (par défaut),
     * le nom de l'arrêt n'est pas modifié.
     * @param {boolean} [checkParentStations=true] Si vrai (par défaut), la recherche se fait
     *  aussi avec la gare de référence puis les gares filles. Si faux, seule un arrêt
     *  à la gare stricte est pris en compte.
     * @returns {Stop | null | undefined} Arrêt correspondant ou undefined
     *  si l'arrêt n'est pas trouvé.
     */
    getStop(stopName: string,
        updateParity: boolean = false,
        updateWithChildStationName: boolean = false,
        checkParentStations: boolean = true): Stop | null | undefined {
        // Si l'arrêt est le terminus, renvoie null
        if (stopName = PARAM.terminusName) return null;
        // Si l'arrêt est donné et trouvé avec parité, l'arrêt est renvoyé
        if (this.stops.has(stopName)) return this.stops.get(stopName)!;

        // Recherche du nom de la gare de l'arrêt et de la parité demandée
        const [stationName, parity] = stopName.split("_");
        const station = STATIONS.get(stationName);
        if (!station) return undefined;
        let stop: Stop | null | undefined;
        if (parity === undefined) {
            // Si la parité n'est pas donnée dans la demande, cherche l'arrêt avec parité
            //  en fonction du numéro de train
            const parityFromTrainNumber = new Parity(this.number, false);
            // Si la parité n'est pas trouvée avec le numéro de train (par exemple
            //  si le train a une double parité), cherche d'abord l'arrêt dans le sens pair
            //  (choix arbitraire), puis dans le sens impair
            if (parityFromTrainNumber.is(Parity.undefined)) parityFromTrainNumber
                .update(Parity.odd);
            stop = this.stops.get(stationName + parityFromTrainNumber.printDigit(true))
                || this.stops.get(stationName + parityFromTrainNumber.invert.printDigit(true));
        } else {
            // Si l'arrêt avec parité n'est pas trouvé, cherche l'arrêt sans parité
            stop = this.stops.get(stationName);
            // Met à jour la parité de l'arrêt si updateParity est vrai
            if (stop && updateParity) {
                stop.parity = new Parity(parity, false);
            }
        }
        if (stop || !checkParentStations) return stop;

        // Si l'arrêt n'est pas trouvé, recherche l'arrêt avec la gare de référence
        const referenceStation = station.referenceStation;
        if (referenceStation) {
            stop = this.getStop(referenceStation.abbreviation, updateParity, false, false);
            if (stop) {
                // Met à jour le nom de l'arrêt avec le nom de la gare fille
                //  si updateWithChildStationName est vrai
                if (updateWithChildStationName) {
                    stop.stationName = stopName;
                    stop.station = station;
                }
                return stop;
            }
        }

        // Si l'arrêt n'est pas trouvé, recherche l'arrêt avec les gares filles
        const childStations = station.childStations || [];
        // Recherche l'arrêt dans les gares filles avec la parité demandée
        for (const childStation of childStations) {
            // Si l'arrêt est trouvé avec la parité demandée pour une des gares filles,
            //  l'arrêt est renvoyé
            const childStopToFind = stopName.replace(stationName, childStation.name);
            stop = this.stops.get(childStopToFind);
            if (stop) return stop;
        }
        // Si l'arrêt n'est pas trouvé, recherche l'arrêt dans les gares filles
        //  sans la parité ou avec une parité inverse
        if (!stop) {
            for (const childStation of childStations) {
                stop = this.getStop(childStation.abbreviation, updateParity, false, false)
                    || null;
                if (stop) return stop;
            }
        }

        return stop;
    }

    /**
     * Vérifie que le train a une gare de départ et une gare de terminus,
     *  et qu'il n'y a pas de gare en double.
     * Vérifie que les arrêts sont chaînés : chaque arrêt a un arrêt suivant
     *  et un arrêt précédent, sauf pour le premier et le dernier.
     * Si les arrêts ne sont pas chaînés, trie les arrêts dans l'ordre chronologique
     *  et attribue les arrêts suivant et précédent.
     * @returns {boolean} Vrai si les arrêts sont conformes, faux sinon. 
     */
    checkStops(adjustTimes: boolean = true): boolean {
        // Si les arrêts sont déjà vérifiés, rien à faire
        if (this.stopsChecked) return true;

        // Création des premiers et derniers arrêts s'ils n'existent pas
        if (!this.firstStop) {
            this.firstStop = this.getStop(this.firstStation)
                || this.getStop(this.departureStation)
                || this.addStop(new Stop(this.departureStation, 0, 0, this.departureTime)
                    .check(this.key, this.departureStation)!);
            this.firstStation = this.firstStop.key;
        }
        if (!this.lastStop) {
            this.lastStop = this.getStop(this.lastStation)
                || this.getStop(this.arrivalStation)
                || this.addStop(new Stop(this.arrivalStation, 0, this.arrivalTime)
                    .check(this.key, this.arrivalStation)!);
            this.lastStation = this.lastStop.key;
        }

        // Vérification des premières et dernières gares
        if (!this.firstStop.checkTimes(this.key, adjustTimes, true)) return false;

        // Le sillon n'a qu'un arrêt de départ et un arrêt d'arrivée non chaînés
        if (this.stops.size === 2 && !this.firstStop.nextStop) {
            if (!this.lastStop.checkTimes(this.key, adjustTimes, false, true,
                this.firstStop.departureTime)) return false;
            this.stopsChecked = 1;
            return true;
        }

        // Vérification du chaînage des arrêts
        let lastStop: Stop | null | undefined = this.firstStop;
        let lastTime = this.firstStop!.departureTime;
        let i = 1;
        while (true) {
            const nextStop: Stop | null | undefined = lastStop!.nextStop
                || this.getStop(lastStop!.nextStopName, false, false, false);
            if (!nextStop) break;
            if (nextStop.checkTimes(this.key, adjustTimes, false,
                nextStop.key === this.lastStop!.key, lastTime)) return false;
            lastStop!.nextStop = nextStop;
            lastStop!.nextStopName = nextStop.key;
            lastStop = nextStop;
            lastTime = lastStop.getTime(true);
            i++;
        }

        if (i === this.stops.size && lastStop === this.lastStop) {
            // Chainage correct : tous les arrêts sont chaînés du premier au dernier arrêt
            this.stopsChecked = 3;
            return true;
        }

        // Le chaînage des arrêts est incorrect : tri des arrêts dans l'ordre chronologique
        const stopsArray = Array.from(this.stops.values());
        stopsArray.sort((a, b) => a.getTime() - b.getTime());

        // Vérification des arrêts et attribution des arrêts suivants
        for (const i = 1; i < stopsArray.length; i++) {
            stopsArray[i].checkTimes(this.key, adjustTimes, false,
                i === stopsArray.length - 1, stopsArray[i - 1].getTime());
            stopsArray[i - 1].nextStop = stopsArray[i];
            stopsArray[i - 1].nextStopName = stopsArray[i].key;
            // !!!!!!
        }
        this.firstStop = stopsArray[0];
        // this.firstStop.previousStop = null;
        this.lastStop = stopsArray[stopsArray.length - 1];
        this.lastStop.nextStop = null;
        this.lastStop.nextStopName = "";
    }

    /**
     * Efface la liste des arrêts du train.
     * Supprime également les valeurs de firstStop et lastStop.
     */
    eraseStops() {
        this.stops.clear();
        this.firstStop = undefined;
        this.lastStop = undefined;
        this.stopsChecked = 0;
    }

    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée du sillon,
     * puis génère la liste des arrêts calculés.
     * Le dernier arrêt a pour valeur de changeNumber la valeur Stop.lastStop.
     * Une fois le trajet calculé, this.stopsChecked a pour valeur 3.
     */
    findPath(useIntermediateStops: boolean = true) {


        // Création de la liste des gares intermédiaires
        //  - à partir des arrêts déjà renseignés (classés par ordre chronologique)
        //   s'ils existent et si useIntermediateStops est vrai
        //  - sinon à partir des gares intermédiaires renseignées (non classées)
        const { viaStops, viaSorted }: { viaStops: string[]; viaSorted: boolean } =
            (useIntermediateStops && this.stops.size > 2)
                ? { viaStops: Array.from(this.stops.keys()), viaSorted: true }
                : { viaStops: this.viaStations, viaSorted: false };

        // Cherche toutes les combinaisons possibles de départ, d'arrivée et de passages via
        const allCombinations = generateCombinations(this.firstStop!.key, this.lastStop!.key,
            viaStops, viaSorted);

        // Trouve le chemin le plus court parmi toutes les combinaisons
        const shortestPath = findShortestPath(allCombinations);

        // Quitte la fonction si aucun chemin n'est trouvé
        if (!shortestPath || shortestPath.path.length === 0) return;

        // Crée la nouvelle liste d'arrêts
        const newStops = new Map<string, Stop>();
        CONSOLE_DEBUG.log(newStops);
        let lastTimedStopName: string | null = null;
        let lastTimedTime: number = 0;
        let segmentPath: { stopName: string, time: number }[] = [];
        let segmentTime: number = 0;

        let previousStop: Stop | null = null;
        let previousStopName: string = "";

        // Remplit la liste des arrêts en reprenant les arrêts déjà renseignés
        //  et en ajoutant les gares de passage     
        for (const stopName of shortestPath.path) {
            const currentStop = this.getStop(stopName, true, true)
                || Stop.newStopIncludingParity(this.key, stopName);
            newStops.set(stopName, currentStop);

            // Attribue les données de l'arrêt précédent
            if (previousStop) {
                previousStop.nextStop = currentStop;
                previousStop.nextStopName = stopName;
                // currentStop.previousStop = previousStop;
            } else {
                // this.firstStop = currentStop;
            }

            // Calcule le temps de parcours depuis la dernière gare avec une heure de départ
            //  ou de passage
            if (lastTimedStopName) {

                // Ajoute l'arrêt dans la liste des arrêts intermédiaires
                const connection = CONNECTIONS.get(previousStopName)?.get(stopName);
                if (connection) {
                    segmentPath.push({ stopName, time: connection.time });
                    segmentTime += connection.time;
                } else {
                    const msg = `Une connexion manque entre "${previousStopName}" et "${stopName}" dans le chemin "${shortestPath.path.join(", ")}".`;
                    throw new Error(msg);
                }

                const currentArrivalTime = currentStop.arrivalTime || currentStop.passageTime;
                if (currentArrivalTime) {
                    // Si une heure d'arrivée ou de passage est renseignée, calcule l'heure de passage
                    // pour chaque gare traversée depuis la dernière gare avec une heure de départ ou de passage
                    if (segmentPath.length > 0) {
                        // Répartit le temps aux arrêts intermédiaires
                        let cumulativeTime = 0;
                        for (const i = 0; i < segmentPath.length; i++) {
                            const stop = newStops.get(segmentPath[i].stopName);
                            cumulativeTime += segmentPath[i].time;
                            if (stop && stop.passageTime === 0 && stop.arrivalTime === 0) {
                                const interpolatedTime = lastTimedTime + (cumulativeTime * (currentArrivalTime - lastTimedTime)) / segmentTime;
                                stop.passageTime = interpolatedTime;
                            }
                        }
                    }
                    lastTimedStopName = null;
                    lastTimedTime = 0;
                    segmentPath = [];
                    segmentTime = 0;
                }
            }

            // Sauvegarde le dernier arrêt avec heure de départ ou de passage.
            const currentDepartureTime = currentStop.getTime(true);
            if (currentDepartureTime) {
                lastTimedStopName = stopName;
                lastTimedTime = currentDepartureTime;
            }
            previousStop = currentStop;
            previousStopName = stopName;
        }
        // CONSOLE_DEBUG.log(this.arrivalStation);
        // CONSOLE_DEBUG.log(this.getStop(this.arrivalStation));

        // this.lastStop = previousStop;
        if (this.lastStop) this.lastStop.changeNumber = Stop.lastStop;
        this.stops = newStops;
        // CONSOLE_DEBUG.log(this.stops.get("MPU_1"));
    }

    /**
     * Retourne le numéro du train, éventuellement modifié pour :
     *  - donner le numéro abrégé de 6 à 4 chiffres si with4Digits est vrai
     *  - ajouter la double parité si withDoubleParity est vrai
     * @param {boolean} [with4Digits=false] Si vrai, le numéro est abrégé
     *  de 6 à 4 chiffres pour les trains commerciaux.
     *  Si faux (par défaut), le numéro est renommé.
     * @param {boolean} [withDoubleParity=false] Si vrai, le numéro est renommé
     *  pour indiquer le changement de parité. Si faux (par défaut), le numéro de train
     *  en gare origine est renvoyé.
     * @returns {string} Numéro du train.
     */
    getTrainNumber(with4Digits: boolean = false, withDoubleParity: boolean = false): string {
        const trainNumber = with4Digits ? abreviateTo4Digits(this.number) : this.number;
        return renameTrainNumberWithParity((this.doubleParity && withDoubleParity) ? Parity.double : 0, trainNumber);;
    }

    /**
     * Renvoie le numéro du train à l'arrivée à l'arrêt donné.
     * @returns {number} Numéro du train au départ.
     */
    getArrivalTrainNumberAtStop(station: string, with4Digits: boolean = false): number {
        const stop = this.getStop(station);
        return stop ? stop.arrivalTrainNumber(this.number, with4Digits) : 0;
    }

    /**
     * Renvoie le numéro du train au départ à l'arrêt donné.
     * Si le train est terminus, renvoie 0.
     * Si l'arrêt est un rebroussement, renvoie la parité modifiée.
     * @returns {number} Numéro du train au départ.
     */
    getDepartureTrainNumberAtStop(station: string, with4Digits: boolean = false): number {
        const stop = this.getStop(station);
        return stop ? stop.departureTrainNumber(this.number, with4Digits) : 0;
    }

    /**
     * Compare ce sillon avec un autre pour vérifier s'ils sont identiques
     * en vérifiant le numéro du train, le code de mission,
     * les heures et les gares de départ et d'arrivée,
     * ainsi que les gares intermédiaires.
     * Si compareStops est vrai, compare également chaque arrêt individuel.
     * @param {TrainPath} other Autre chemin de train à comparer.
     * @param {boolean} [compareStops=false] Si vrai, compare également
     *  les arrêts. Si faux (par défaut), compare seulement les éléments du sillon.
     * @returns {boolean} Vrai si les chemins de train sont identiques, faux sinon.
     */

    compareTo(other: TrainPath, compareStops: boolean = false): boolean {
        const sameTrainPath = this.number === other.number
            && this.missionCode === other.missionCode
            && this.departureTime === other.departureTime
            && this.departureStation === other.departureStation
            && this.arrivalTime === other.arrivalTime
            && this.arrivalStation === other.arrivalStation
            && this.viaStations === other.viaStations;

        if (!sameTrainPath || !compareStops) return sameTrainPath;

        this.stops.forEach((stop) => {
            if (!stop.compareTo(other.getStop(stop.key))) return false;
        });
        return true;
    }
}


/* Liste des sillons avec leurs arrêts, plannifiés sur un ou plusieurs jours, avec les mêmes horaires. */
const TRAIN_PATHS = new Map<string, TrainPath>();

const TRAIN_PATHS_SHEET = "Sillons";
const TRAIN_PATHS_TABLE = "Sillons";
const TRAIN_PATHS_HEADERS = [[
    "Id",
    "Numéro du train",
    "Parité de ligne",
    "Jours",
    "Code mission",
    "Heure de départ",
    "Gare de départ",
    "Heure d'arrivée",
    "Gare d'arrivée",
    "Première gare avec évolutions",
    "Dernière gare avec évolutions",
    "Gares intermédiaires"
]];
const TRAIN_PATHS_COL_KEY = 0; // Non lue car calculée
const TRAIN_PATHS_COL_NUMBER = 1;
const TRAIN_PATHS_COL_LINE_PARITY = 2;
const TRAIN_PATHS_COL_DAYS = 3;
const TRAIN_PATHS_COL_MISSION_CODE = 4;
const TRAIN_PATHS_COL_DEPARTURE_TIME = 5;
const TRAIN_PATHS_COL_DEPARTURE_STATION = 6;
const TRAIN_PATHS_COL_ARRIVAL_TIME = 7;
const TRAIN_PATHS_COL_ARRIVAL_STATION = 8;
const TRAIN_PATHS_COL_FIRST_STATION = 9;    // Valeur non lue car affectée lors de la lecture du premier arrêt
const TRAIN_PATHS_COL_LAST_STATION = 10;    // Valeur non lue car affectée lors de la lecture du dernier arrêt
const TRAIN_PATHS_COL_VIA_STATIONS = 11;

/**
 * Charge les sillons de trains à partir du tableau "Sillons" de la feuille "Sillons".
 * Les sillons sont stockés dans un objet avec comme clés le numéro de sillon 
 * suivi du jour et comme valeur l'objet TrainPath.
 * Chaque sillon correspondant à la sélection sera associé avec autant de clés que de jours
 * de circulation, en plus du numéro de sillon suivi du code des jours de circulation
 * (le sillon 123456_J aura pour clés : 123456_J, 123456_1, 123456_2...)
 * @param {string} days Jours pour lesquels les sillons sans jours spécifiques sont demandés.
 * @param {string} trainNumbers Numéros des sillons à charger, avec ou sans jours associés, séparés par des ';'.
 * Si vide, charge tous les trains de la base TRAIN_PATHS.
 * @param {boolean} [erase=false] Si vrai, supprime les trains déjà chargés.
 *  Si faux (par défaut), ne recharge pas si déjà chargé.
 */
function loadTrainPaths(trainDays: string = "JW", trainNumbers: string = "", erase: boolean = false) {

    // Vérifie si la table à charger existe déjà
    if (TRAIN_PATHS.size > 0) {
        if (erase) {
            TRAIN_PATHS.clear(); // Vide la map sans changer sa référence
        }
    }

    loadStations();
    const data = getDataFromTable(TRAIN_PATHS_SHEET, TRAIN_PATHS_TABLE);

    // Map des sillons à charger : numéro → chaîne des jours associés
    // La concaténation des jours peut comporter plusieurs fois le même jour
    const trainNumberMap = new Map<string, string>();
    trainNumbers.split(';').forEach(entry => {
        const [number, days] = entry.split('_');
        const previous = trainNumberMap.get(number) || '';
        trainNumberMap.set(number, previous + (days || trainDays));
    });

    // Parcourt la base de données
    for (const row of data.slice(1)) {
        // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides)
        if (row.every(cell => !cell)) continue;

        const number = String(row[TRAIN_PATHS_COL_NUMBER]);
        const days = String(row[TRAIN_PATHS_COL_DAYS]);
        
        // Vérifie si le sillon est déjà chargé
        if (TRAIN_PATHS.has(`${number}_${days}`)) continue;

        // Vérifie si le sillon est concerné dans la liste des sillons à charger, sauf si aucun filtre n'est fourni
        if (trainNumberMap.size > 0 && !trainNumberMap.has(`${number}`)) continue;

        // Détermine les jours à filtrer
        const filterDays = trainNumberMap.get(`${number}`) || trainDays;

        // Calcule les jours communs entre ceux du sillon et ceux demandés
        const commonDays = Day.extractFromString(days, filterDays);
        if (commonDays.length === 0) continue;

        // Extrait les valeurs
        const lineDirection = row[TRAIN_PATHS_COL_LINE_PARITY] as number;
        const missionCode = String(row[TRAIN_PATHS_COL_MISSION_CODE]);
        const departureTime = row[TRAIN_PATHS_COL_DEPARTURE_TIME] as number;
        const departureStation = String(row[TRAIN_PATHS_COL_DEPARTURE_STATION]);
        const arrivalTime = row[TRAIN_PATHS_COL_ARRIVAL_TIME] as number;
        const arrivalStation = String(row[TRAIN_PATHS_COL_ARRIVAL_STATION]);
        const viaStations = String(row[TRAIN_PATHS_COL_VIA_STATIONS]);

        // Crée l'objet TrainPath
        const trainPath = new TrainPath(
            number,
            lineDirection,
            days,
            missionCode,
            departureTime,
            departureStation,
            arrivalTime,
            arrivalStation,
            viaStations
        );

        // Insert le sillon dans la table avec plusieurs clés d'accès
        //  - une référence pour la clé unique du sillon
        TRAIN_PATHS.set(trainPath.key, trainPath);
        //  - une référence pour chacun des jours demandés
        commonDays.forEach((day) => {
            const key = number + "_" + day;
            if (!TRAIN_PATHS.has(key)) TRAIN_PATHS.set(key, trainPath);
        });
    }
}

/**
 * Affiche les sillons dans un tableau.
 * Les données sont celles stockées dans l'objet TRAIN_PATHS.
 * @param {string} sheetName Nom de la feuille de calcul.
 * @param {string} tableName Nom du tableau.
 * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
 */
function printTrainPaths(sheetName: string, tableName: string, startCell: string = "A1"): void {

    // Filtre l'objet TRAIN_PATHS en ne prennant qu'une seule fois les sillons ayant la même clé   
    const seenKeys = new Set<string>();
    const uniqueTrainPaths: TrainPath[] = Array.from(TRAIN_PATHS.entries())
        .filter(([mapKey, trainPath]) => mapKey === trainPath.key)
        .map(([_, trainPath]) => trainPath);

    // Convertit l'objet TRAIN_PATHS filtré en un tableau de données
    const data: (string | number)[][] = uniqueTrainPaths.map(trainPath => [
        trainPath.key,
        trainPath.number,
        trainPath.lineDirection.printDigit(),
        trainPath.days,
        trainPath.missionCode,
        trainPath.departureTime,
        trainPath.departureStation,
        trainPath.arrivalTime,
        trainPath.arrivalStation,
        trainPath.viaStations.join(';'),
    ]);

    // Imprime le tableau
    const table = printTable(TRAIN_PATHS_HEADERS, data, sheetName, tableName, startCell);

    // Met les heures au format "hh:mm:ss"
    const timeColumns = [
        TRAIN_PATHS_COL_DEPARTURE_TIME,
        TRAIN_PATHS_COL_ARRIVAL_TIME,
    ];

    for (const col of timeColumns) {
        table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
    }
}

/**
 * Cherche les chemins possibles pour tous les sillons de trains stockés 
 * dans l'objet TRAIN_PATHS.
 * Appel la fonction findPath pour chaque sillon de train.
 */
function findPathsOnAllTrainPaths() {
    TRAIN_PATHS.forEach((trainPath, key) => {
        if (key === trainPath.key) trainPath.findPath();
    });
}

/**
 * Classe Train qui définit un train pour un unique jour, étant la réutilisation
 * d'un ou deux trains précédents, et ayant une ou deux réutilisations,
 * en faisant référence à un sillon avec horaires pouvant circuler plusieurs jours par semaine.
 */
class Train {
    number: string;                 // Numéro du train
    trainPath: TrainPath;           // Sillon sur lequel le train est prévu prévu de circuler
    day: number;                    // Jour du train    (1 à 7 = lundi à dimanche, >7 = date précise)
    firstStation?: string;          // Gare de départ si différente de celle du sillon
    lastStation?: string;           // Gare d'arrivée si différente de celle du sillon
    unit1: string;                  // Element 1 Nord (numéro de matériel)
    unit2: string;                  // Element 2 Sud (numéro de matériel)
    previous1: string;              // Clé du train précédent de l'élément 1
    previous2: string;              // Clé du train précédent de l'élément 2
    reuse1?: Train;                 // Train de réutilisation de l'élément 1
    reuse1Key: string;              // Clé du train de réutilisation de l'élément 1
    reuse2?: Train;                 // Train de réutilisation de l'élément 2
    reuse2Key: string;              // Clé du train de réutilisation de l'élément 2

    constructor(
        number: string,
        trainPathKey: string,
        day: number,
        unit1: string = "",
        unit2: string = "",
        previous1: string = "",
        previous2: string = "",
        reuse1Key: string = "",
        reuse2Key: string = ""
    ) {
        this.number = number;
        this.trainPath = TRAIN_PATHS.get(trainPathKey) as TrainPath;
        this.day = day;
        this.unit1 = unit1;
        this.unit2 = unit2;
        this.previous1 = previous1;
        this.previous2 = previous2;
        this.reuse1Key = reuse1Key;
        this.reuse2Key = reuse2Key;
        if (!this.trainPath) {
            CONSOLE_WARN.log(`Train n° ${this.number}_${this.day} : le sillon rattaché est inconnu : ${trainPathKey}.`);
            return;
        }
    }

    /**
     * Vérifie la validité de l'objet Train en envoyant un message d'erreur si :
     *  - le sillon est inconnu.
     * @returns {Train | undefined} Objet Train s'il est valide, undefined sinon.
     */
    check(): Train | undefined {
        if (!this.trainPath) {
            CONSOLE_WARN.log(`Train n° ${this.number}_${this.day} : le sillon rattaché est inconnu : ${this.trainPathKey}.`);
            return undefined;
        }
        return this;
    }

    /**
     * Retourne la clé du train qui est composée du numéro du train
     *  suivi de la liste des jours de circulation.
     * @returns {string} Clé du train
     */
    get key(): string {
        return `${this.number}_${this.day}`;
    }

    /**
     * Retourne le numéro du train, éventuellement modifié pour :
     *  - donner le numéro abrégé de 6 à 4 chiffres si with4Digits est vrai
     *  - ajouter la double parité si withDoubleParity est vrai
     * @param {boolean} [with4Digits=false] Si vrai, le numéro est abrégé
     *  de 6 à 4 chiffres pour les trains commerciaux.
     *  Si faux (par défaut), le numéro est renommé.
     * @param {boolean} [withDoubleParity=false] Si vrai, le numéro est renommé
     *  pour indiquer le changement de parité. Si faux (par défaut), le numéro de train
     *  en gare origine est renvoyé.
     * @returns {string} Numéro du train.
     */
    getTrainNumber(with4Digits: boolean = false, withDoubleParity: boolean = false): string {
        return this.trainPath.getTrainNumber(with4Digits, withDoubleParity);;
    }
}

/* Liste des trains et leurs réutilisations, associés à un jour donné, circulants sur un sillon donné. */
const TRAINS = new Map<string, Train>();

const TRAINS_SHEET = "Réuts";
const TRAINS_TABLE = "Réuts";
const TRAINS_HEADERS = [[
    "Id",
    "Numéro du train",
    "Jours",
    "Sillon",
    "Elément Nord",
    "Elément Sud",
    "Train Précédent Nord",
    "Train Précédent Sud",
    "Réutilisation Nord",
    "Réutilisation Sud",
]];
const TRAINS_COL_KEY = 0;
const TRAINS_COL_NUMBER = 1;
const TRAINS_COL_DAYS = 2;
const TRAINS_COL_TRAIN_PATH = 3;
const TRAINS_COL_UNIT1 = 4;
const TRAINS_COL_UNIT2 = 5;
const TRAINS_COL_PREVIOUS1 = 6;
const TRAINS_COL_PREVIOUS2 = 7;
const TRAINS_COL_REUSE1 = 8;
const TRAINS_COL_REUSE2 = 9;

/**
 * Charge les réutilisations à partir du tableau "Réuts" de la feuille "Réuts".
 * Les réutilisations sont stockés dans la table un objet avec comme clés le numéro de train
 *  suivi du jour de circulation (numéro du jour ou date) et comme valeur l'objet Réutilisation.
 * @param {string} days Jours pour lesquels les sillons sans jours spécifiques sont demandés.
 * @param {string} trainNumbers Numéros des sillons à charger, avec ou sans jours associés,
 *  séparés par des ';'. Si vide, charge tous les trains de la base TRAIN_PATHS.
 * @param {boolean} [erase=false] Si vrai, supprime les trains déjà chargés.
 *  Si faux (par défaut), ne recharge pas si déjà chargé.
 */
function loadTrains(trainDays: string = "JW", trainNumbers: string = "", erase: boolean = false) {

    // Vérifie si la table à charger existe déjà
    if (TRAINS.size > 0) {
        if (erase) {
            TRAINS.clear(); // Vide la map sans changer sa référence
        }
    }
}

/**
 * Affiche les réutilisations dans un tableau.
 * Les données sont celles stockées dans l'objet TRAINS.
 * @param {string} sheetName Nom de la feuille de calcul.
 * @param {string} tableName Nom du tableau.
 * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
 */
function printTrains(sheetName: string, tableName: string, startCell: string = "A1"): void {

    // Convertit l'objet TRAINS en un tableau de données
    const data: (string | number)[][] = Object.values(TRAINS).map(train => [
        train.key,
        train.number,
        train.day,
        train.trainPath.key,
        train.unit1,
        train.unit2,
        train.previous1,
        train.previous2,
        train.reuse1.key,
        train.reuse2.key
    ]);

    // Imprime le tableau
    printTable(TRAINS_HEADERS, data, sheetName, tableName, startCell);
}

class Stop {

    public static readonly lastStop: number = 2;

    station?: Station;          // Gare de l'arrêt
    parity: Parity;             // Parité de l'arrêt à l'arrivée
    arrivalTime: number;        // Heure d'arrivée de l'arrêt
    departureTime: number;      // Heure de départ de l'arrêt
    passageTime: number;        // Heure de passage à l'arrêt (sans arrêt)
    track: string;              // Voie de l'arrêt
    changeNumber: number;       // Changement de numérotation
    //  - 0 = même train,
    //  - 1 = rebroussement pair vers impair,
    //  - 1 = rebroussement impair vers pair,
    //  - Stop.lastStop = réutilisation
    previousStopName?: string;  // Nom de l'arrêt précédent
    nextStop?: Stop | null;     // Arrêt suivant
    nextStopName: string;       // Nom de l'arrêt suivant

    constructor(
        stationName: string,
        parity: string | number = 0,
        arrivalTime: number = 0,
        departureTime: number = 0,
        passageTime: number = 0,
        track: string = "",
        changeNumber: number = 0,
        nextStopName: string = ""
    ) {
        this.station = STATIONS.get(stationName);
        this.parity = new Parity(parity, false);
        this.arrivalTime = arrivalTime;
        this.departureTime = departureTime;
        this.passageTime = passageTime;
        this.track = track;
        this.changeNumber = changeNumber;
        this.nextStopName = nextStopName;
        if (nextStopName = PARAM.terminusName) this.nextStop = null;
    }

    /**
     * Vérifie la validité de l'objet Stop en envoyant un message d'erreur si :
     *  - le nom de la gare est vide,
     *  - la gare est inconnue,
     *  - l'heure d'arrivée et de départ et de passage est vide.
     * @returns {Stop | undefined} Objet Stop s'il est valide, undefined sinon.
     */
    check(trainPathKey: string, stationName: string): Stop | undefined {
        if (!stationName) {
            CONSOLE_WARN.log(`Sillon : ${trainPathKey} Un arrêt ne peut pas avoir`
                + ` pour gare vide.`);
            return undefined;
        } else if (!this.station) {
            CONSOLE_WARN.log(`Sillon : ${trainPathKey} Un arrêt ne peut pas avoir`
                + ` pour gare : "${stationName}" qui est inconnue.`);
            return undefined;
        } else if (!this.arrivalTime && !this.departureTime && !this.passageTime) {
            CONSOLE_WARN.log(`Sillon : ${trainPathKey} Un arrêt doit avoir au moins`
                + ` une heure d'arrivée ou de départ ou de passage.`);
            return undefined;
        }
        return this;
    }

    /**
     * Renvoie une clé unique pour l'arrêt, composée du nom de la gare et de la parité
     *  (si connue).
     * @returns {string} Clé unique
     */
    get key(): string {
        return this.stationName + this.parity.printDigit(true);
    }

    /**
     * Renvoie l'abréviation de la gare associée à cet arrêt.
     * @returns {string} Abréviation de la gare.
     */
    get stationName(): string {
        return this.station!.abbreviation;
    }

    /**
     * Crée une nouvelle instance de l'arrêt incluant la parité à partir d'une chaîne
     *  de caractères. La chaîne de caractères doit être au format "NomDeGare_Parité".
     * @param {string} stopWithParity Chaîne de caractères contenant le nom de la gare
     *  et la parité, séparés par un underscore.
     * @param {number} [arrivalTime=0] Heure d'arrivée à l'arrêt.
     * @param {number} [departureTime=0] Heure de départ de l'arrêt.
     * @param {number} [passageTime=0] Heure de passage à l'arrêt (sans arrêt).
     * @param {string} [track=""] Voie de l'arrêt.
     * @param {number} [changeNumber=0] Changement de numérotation.
     * @param {string} [nextStopName=""] Nom de l'arrêt suivant.
     * @returns {Stop} Nouvelle instance de l'arrêt avec les informations fournies.
     */
    static newStopIncludingParity(
        trainPathKey: string,
        stopWithParity: string,
        arrivalTime: number = 0,
        departureTime: number = 0,
        passageTime: number = 0,
        track: string = "",
        changeNumber: number = 0,
        nextStopName: string = ""
    ): Stop | undefined {
        const [stationName, parity] = stopWithParity.split("_");
        const stop = new Stop(trainPathKey, stationName, parity, arrivalTime, departureTime,
            passageTime, track, changeNumber, nextStopName);
        return stop.check(trainPathKey, stationName);
    }


    /**
     * Renvoie la plus petite des heures d'arrivée, de départ ou de passage à l'arrêt.
     * Si l'heure d'arrivée est lue et que noReadingArrivalTime est vrai,
     * ignore l'heure d'arrivée.
     * @param {boolean} [noReadingArrivalTime=false] Si vrai, ignore l'heure d'arrivée
     *  et préfère l'heure de départ ou de passage. Si faux (par défaut),
     *  c'est d'abord l'heure d'arrivée qui est prise en compte.
     * @returns {number} Heure la plus petite.
     */
    getTime(noReadingArrivalTime: boolean = false): number {
        return (noReadingArrivalTime ? this.arrivalTime : 0)
            || this.departureTime
            || this.passageTime;
    }

    /**
     * Donne le numéro de train à l'arrivée à l'arrêt.
     * Renvoie le numéro du train adapté par la parité,
     *  incrémenté du changement de numérotation.
     * @param {string | number} trainNumber Numéro du train.
     * @returns {number} Numéro du train adapté.
     */
    arrivalTrainNumber(trainNumber: string | number, with4Digits: boolean = false): number {
        const adaptedTrainNumber: number = this.parity.adaptTrainNumber(trainNumber);
        return with4Digits
            ? abreviateTo4Digits(adaptedTrainNumber) as number
            : adaptedTrainNumber;
    }

    /**
     * Donne le numéro de train au départ de l'arrêt.
     * Si l'arrêt est un terminus, renvoie 0.
     * Sinon, renvoie le numéro du train adapté par la parité,
     *  incrémenté du changement de numérotation,
     *  y compris si l'arrêt est un rebroussement.
     * @param {string | number} trainNumber Numéro du train.
     * @returns {number} Numéro du train adapté.
     */
    departureTrainNumber(trainNumber: string | number, with4Digits: boolean = false): number {
        const adaptedTrainNumber: number = this.changeNumber === Stop.lastStop ? 0
            : this.parity.adaptTrainNumber(trainNumber) + this.changeNumber;
        return with4Digits ? abreviateTo4Digits(adaptedTrainNumber) as number
            : adaptedTrainNumber;
    }

    /**
     * Compare cette arrêt avec un autre arrêt,
     *  en vérifiant le nom de gare, la parité,
     *  les heures d'arrivée, de départ et de passage, et la voie.
     * @param {Stop | null | undefined} other Autre arrêt à comparer.
     * @returns {boolean} Vrai si les arrêts sont égaux, faux sinon.
     */
    compareTo(other: Stop | null | undefined): boolean {
        if (!other) return false;
        return this.station === other.station
            && this.parity.equals(other.parity)
            && this.arrivalTime === other.arrivalTime
            && this.departureTime === other.departureTime
            && this.passageTime === other.passageTime
            && this.track === other.track;
    }

    checkTimes(pathTrainKey: string, adjustTimes: boolean = true, isFirstStop: boolean = false,
        isLastStop: boolean = false, departureTimeOfPreviousStop: number = 0): boolean {

        if (isFirstStop) {
            // Vérification de l'arrêt comme premier arrêt du train
            if (this.departureTime) {
                if (this.passageTime) {
                    if (adjustTimes) this.passageTime = 0;
                    CONSOLE_WARN.log(`Le premier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure de passage`
                        + ` (${formatTime(this.passageTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
                if (this.arrivalTime) {
                    if (adjustTimes) this.arrivalTime = 0;
                    CONSOLE_WARN.log(`Le premier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure d'arrivée`
                        + ` (${formatTime(this.arrivalTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else {
                if (adjustTimes && this.passageTime) {
                    this.departureTime = this.passageTime;
                    this.passageTime = 0;
                }
                CONSOLE_WARN.log(`Le premier arrêt ${this.stationName} du sillon`
                    + ` ${pathTrainKey} ne présente pas d'heure de départ.`
                    + this.departureTime
                    ? ` L'heure de passage (${formatTime(this.passageTime)})`
                    + ` a été modifiée en heure de départ.`
                    : "");
                if (!this.departureTime) return false;
            }
            if (departureTimeOfPreviousStop
                && this.departureTime <= departureTimeOfPreviousStop) {
                CONSOLE_WARN.log(`Le premier arrêt ${this.stationName} du sillon ${this.key}`
                    + ` a une heure de départ (${formatTime(this.departureTime)}) inférieure`
                    + ` à l'heure d'arrivée du train prédédent`
                    + ` (${formatTime(departureTimeOfPreviousStop)}).`);
                return false;
            }
            return true;
        } else if (isLastStop) {
            // Vérification de l'arrêt comme dernier arrêt du train
            if (this.arrivalTime) {
                if (this.passageTime) {
                    if (adjustTimes) this.passageTime = 0;
                    CONSOLE_WARN.log(`Le dernier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure de passage`
                        + ` (${formatTime(this.passageTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
                if (this.departureTime) {
                    if (adjustTimes) this.departureTime = 0;
                    CONSOLE_WARN.log(`Le dernier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure de départ`
                        + ` (${formatTime(this.departureTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else {
                if (adjustTimes && this.passageTime) {
                    this.arrivalTime = this.passageTime;
                    this.passageTime = 0;
                }
                CONSOLE_WARN.log(`Le dernier arrêt ${this.stationName} du sillon`
                    + ` ${pathTrainKey} ne présente pas d'heure d'arrivée.`
                    + this.arrivalTime
                    ? ` L'heure de passage (${formatTime(this.passageTime)})`
                    + ` a été modifiée en heure d'arrivée.`
                    : "");
                if (!this.arrivalTime) return false;
            }
        } else {
            // Vérification de l'arrêt comme arrêt intermédiaire du train
            if (this.arrivalTime && this.departureTime) {
                // L'arrêt intermédiaire a une heure d'arrivée et de départ
                if (this.arrivalTime > this.departureTime) {
                    CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${this.key} a une heure de départ`
                        + ` (${formatTime(this.departureTime)}) inférieure à l'heure d'arrivée`
                        + ` (${formatTime(this.arrivalTime)}).`);
                    return false;
                } else if (this.arrivalTime === this.departureTime) {
                    if (adjustTimes) {
                        this.passageTime = this.arrivalTime;
                        this.arrivalTime = 0;
                        this.departureTime = 0;
                    }
                    CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} a des heures d'arrivée et de départ identiques`
                        + ` (${formatTime(this.arrivalTime)}).`
                        + adjustTimes
                        ? ` Elles ont donc été remplacées par une heure de passage.`
                        : "");
                } else if (this.passageTime) {
                    if (adjustTimes) this.passageTime = 0;
                    CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente en plus d'une heure d'arrivée`
                        + ` et de départ, une heure de passage`
                        + ` (${formatTime(this.passageTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else if (this.passageTime) {
                // L'arrêt intermédiaire a une heure de passage
                if (this.arrivalTime) {
                    if (adjustTimes) this.arrivalTime = 0;
                    CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente en plus d'une heure de passage,`
                        + ` une heure d'arrivée (${formatTime(this.arrivalTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
                if (this.departureTime) {
                    if (adjustTimes) this.departureTime = 0;
                    CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente en plus d'une heure de passage,`
                        + ` une heure de départ (${formatTime(this.departureTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else if (this.arrivalTime) {
                // L'arrêt intermédiaire a une heure d'arrivée, mais ni heure de départ,
                //  ni heure de passage
                if (adjustTimes) {
                    this.passageTime = this.arrivalTime;
                    this.arrivalTime = 0;
                }
                CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                    + ` ${pathTrainKey} ne présente qu'une heure d'arrivée`
                    + ` (${formatTime(this.arrivalTime)}).`
                    + adjustTimes ? ` Elle a été modifiée en heure de passage.` : "");
                if (!this.passageTime) return false;
            } else if (this.departureTime) {
                // L'arrêt intermédiaire a une heure de départ, mais ni heure d'arrivée,
                //  ni heure de passage
                if (adjustTimes) {
                    this.passageTime = this.departureTime;
                    this.departureTime = 0;
                }
                CONSOLE_WARN.log(`L'arrêt intermédiaire ${this.stationName} du sillon`
                    + ` ${pathTrainKey} ne présente qu'une heure de départ`
                    + ` (${formatTime(this.departureTime)}).`
                    + adjustTimes ? ` Elle a été modifiée en heure de passage.` : "");
                if (!this.passageTime) return false;
            } else {
                // Pas d'autre cas possible
            }
        }

        // Vérification de l'horaire d'arrivée ou de passage avec l'horaire de départ
        //  ou de passage de l'arrêt prédédent
        if (departureTimeOfPreviousStop && this.departureTime <= departureTimeOfPreviousStop) {
            CONSOLE_WARN.log(`L'arrêt ${this.stationName} du sillon ${this.key}`
                + ` a une heure d'arrivée ou de passage`
                + ` (${formatTime(this.arrivalTime || this.passageTime)}) inférieure`
                + ` à l'heure de départ ou de passage de l'arrêt prédédent`
                + ` (${formatTime(departureTimeOfPreviousStop)}).`);
            return false;
        }

        return true;
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
    "Voie",
    "Changement de numérotation",
    "Arrêt suivant"
]];
const STOPS_COL_TRAIN_NUMBER = 0;
const STOPS_COL_TRAIN_DAYS = 1;
const STOPS_COL_STATION = 2;
const STOPS_COL_PARITY = 3;
const STOPS_COL_ARRIVAL_TIME = 4;
const STOPS_COL_DEPARTURE_TIME = 5;
const STOPS_COL_PASSAGE_TIME = 6;
const STOPS_COL_TRACK = 7;
const STOPS_COL_CHANGE_NUMBER = 8;
const STOPS_COL_NEXT_STOP = 9;

/**
 * Charge les arrêts à partir de la feuille "Arrêts" du classeur.
 * Les arrêts sont stockés dans la propriété "stops" des trains correspondants.
 * Si un train n'existe pas, un message d'erreur est affiché.
 */
function loadStops() {

    const data = getDataFromTable(STOPS_SHEET, STOPS_TABLE);

    // Parcourt la base de données
    for (const row of data.slice(1)) {

        // Vérifie si le train existe
        const trainNumber = String(row[STOPS_COL_TRAIN_NUMBER]);
        const trainDays = String(row[STOPS_COL_TRAIN_DAYS]);
        if (!trainNumber || !trainDays) continue;

        const trainKey = trainNumber + "_" + trainDays;
        if (!TRAIN_PATHS.has(trainKey)) continue;

        const train = TRAIN_PATHS.get(trainKey) as TrainPath;

        // Extrait les valeurs
        const stationName = String(row[STOPS_COL_STATION]);
        if (!stationName) continue;
        const parity = row[STOPS_COL_PARITY] as number;
        const arrivalTime = row[STOPS_COL_ARRIVAL_TIME] as number;
        const departureTime = row[STOPS_COL_DEPARTURE_TIME] as number;
        const passageTime = row[STOPS_COL_PASSAGE_TIME] as number;
        const track = String(row[STOPS_COL_TRACK]);
        const changeNumber = row[STOPS_COL_CHANGE_NUMBER] as number;
        const nextStopName = String(row[STOPS_COL_NEXT_STOP]);

        const stop = new Stop(
            train.key,
            stationName,
            parity,
            arrivalTime,
            departureTime,
            passageTime,
            track,
            changeNumber,
            nextStopName
        );
        if (!stop) continue;

        // Ajoute l'arrêt au train
        train.addStop(stop);
    }

    // Boucle pour reparcourir tous les trains et vérifier leurs arrêts
    for (const train of TRAIN_PATHS.values()) {
        train.checkStops();
    }
}

/**
 * Affiche les arrêts des trains dans un tableau.
 * Les données sont celles stockées dans les objets TrainPath et Stop de l'objet TRAIN_PATHS.
 * @param {string} sheetName Nom de la feuille de calcul.
 * @param {string} tableName Nom du tableau.
 * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
 */
function printStops(sheetName: string, tableName: string, startCell: string = "A1"): void {

    // Filtre l'objet TRAIN_PATHS en ne prennant qu'une seule fois les trains
    //  ayant la même clé   
    const seenKeys = new Set<string>();
    const uniqueTrainPaths: TrainPath[] = Array.from(TRAIN_PATHS.entries())
        .filter(([mapKey, train]) => mapKey === train.key)
        .map(([_, train]) => train);

    // Crée le tableau final avec les données de chaque arrêt pour chaque train
    const data: (string | number)[][] = [];

    for (const train of uniqueTrainPaths) {
        for (const [stationName, stop] of train.stops.entries()) {
            data.push([
                train.number,
                train.days,
                stop.stationName,
                stop.parity.printDigit(),
                stop.arrivalTime || "",
                stop.departureTime || "",
                stop.passageTime || "",
                stop.track,
                stop.changeNumber,
                stop.nextStop == null ? PARAM.terminusName : stop.nextStopName
            ]);
        }
    }

    // Imprime le tableau
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
    abbreviation!: string;              // Abréviation de la gare
    name: string;                       // Nom de la gare
    referenceStationName: string;       // Gare de rattachement
    referenceStation: Station | null;   // Gare de rattachement
    childStations: Station[];           // Sous-gares
    turnaround: Parity;                 // Parité d'un rebroussement possible
    //  (la parité est celle du train avant rebroussement)
    reverseLineDirection: boolean;      // Parité de la ligne inversée sur cette gare

    constructor(
        abbreviation: string,
        name: string,
        referenceStationName: string,
        turnaround: string | number,
        reverseLineDirection: boolean,
    ) {
        this.abbreviation = abbreviation;
        this.name = name;
        this.referenceStationName = referenceStationName;
        this.referenceStation = null;
        this.childStations = [];
        this.turnaround = new Parity(turnaround, true);
        this.reverseLineDirection = reverseLineDirection;
    }

    /**
     * Vérifie la validité de l'objet Station en envoyant un message d'erreur si :
     *  - l'abréviation est vide.
     * @returns {Station | undefined} Objet Station s'il est valide, undefined sinon.
     */
    check(): Station | undefined {
        if (!this.abbreviation) {
            CONSOLE_WARN.log(`Une gare ne peut pas avoir une abréviation vide.`);
            return undefined;
        }
        return this;
    }
}

/* Liste des gares et leurs coordonnées. */
const STATIONS = new Map<string, Station>();

const STATIONS_SHEET = "Gares";
const STATIONS_TABLE = "Gares";
const STATIONS_HEADERS = [[
    "Abréviation",
    "Nom",
    "Gare de rattachement",
    "Gare de rebroussement",
    "Parité de ligne inversée"
]];
const STATIONS_COL_ABBR = 0;
const STATIONS_COL_NAME = 1;
const STATIONS_COL_REFERENCE_STATION = 2;
const STATIONS_COL_TURNAROUND = 3;
const STATIONS_COL_REVERSE_LINE_PARITY = 4;

/**
 * Charge les gares à partir du tableau "Gares" de la feuille "Gares".
 * Les gares sont stockées dans une Map avec comme clés l'abréviation 
 * de la gare et comme valeur l'objet Station.
 * @param {boolean} [erase=false] Si vrai, force le rechargement des gares.
 *  Si faux (par défaut), ne recharge pas si déjà chargé.
 */
function loadStations(erase: boolean = false) {

    // Vérifie si la table à charger existe déjà
    if (STATIONS.size > 0) {
        if (erase) {
            STATIONS.clear(); // Vide la map sans changer sa référence
        } else {
            return;
        }
    }

    const data = getDataFromTable(STATIONS_SHEET, STATIONS_TABLE);

    // Parcourt la base de données
    const referenceStationPairs: [string, string][] = [];
    for (const row of data.slice(1)) {

        // Extrait les valeurs
        const abbreviation = String(row[STATIONS_COL_ABBR]);
        if (!abbreviation) continue;
        const name = String(row[STATIONS_COL_NAME]);
        const referenceStationName = String(row[STATIONS_COL_REFERENCE_STATION]);
        const oddTurnaround =
            Parity.containsParityLetter(String(row[STATIONS_COL_TURNAROUND]), Parity.odd);
        const evenTurnaround =
            Parity.containsParityLetter(String(row[STATIONS_COL_TURNAROUND]), Parity.even);
        const reverseLineDirection = row[STATIONS_COL_REVERSE_LINE_PARITY] as boolean;

        // Crée l'objet Station
        const station = new Station(
            abbreviation,
            name,
            referenceStationName,
            oddTurnaround,
            evenTurnaround,
            reverseLineDirection
        ).check();
        if (!station) continue;

        // Ajoute l'objet Station dans la map
        if (STATIONS.has(abbreviation)) {
            CONSOLE_WARN.log(`La gare ${abbreviation} est présente deux fois`
                + ` dans la base de données.`);
            continue;
        }
        STATIONS.set(abbreviation, station);

        // Mémorise les paires gare/gare de rattachement
        referenceStationPairs.push([abbreviation, referenceStationName]);
    }

    // Parcourt les paires pour ajouter les objets des gares de réference à chaque gare
    for (const [abbreviation, referenceStationName] of referenceStationPairs) {
        const station = STATIONS.get(abbreviation);
        const referenceStation = STATIONS.get(referenceStationName);
        if (station && referenceStation) {
            station.referenceStation = referenceStation;
            referenceStation.childStations.push(station);
        }
    }
}

/**
 * Affiche les stations dans un tableau.
 * Les données sont celles stockées dans l'objet STATIONS.
 * @param {string} sheetName Nom de la feuille de calcul.
 * @param {string} tableName Nom du tableau.
 * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
 */
function printStations(sheetName: string, tableName: string, startCell: string = "A1"): void {

    // Convertit l'objet STATIONS en un tableau de données
    const data: (string | number)[][] = Object.values(STATIONS).map(station => [
        station.abbreviation,
        station.name,
        station.variants.join(", "),
        station.connectedStationsWithParityChange.join(", ")
    ]);

    // Imprime le tableau
    printTable(STATIONS_HEADERS, data, sheetName, tableName, startCell);
}

class Connection {
    from: string;               // Gare de départ
    to: string;                 // Gare d'arrivée
    time: number;               // Temps de trajet
    withTurnaround: boolean;    // Connexion impliquant un rebroussement
    withMovement: boolean;      // Connexion sous régime de l'évolution
    changeParity: boolean;      // Connexion avec changement de parité

    constructor(
        from: string,
        to: string,
        time: number = 1,
        withTurnaround: boolean = false,
        withMovement: boolean = false,
        changeParity: boolean = false
    ) {
        this.from = from;
        this.to = to;
        this.time = time || 1;
        this.withTurnaround = withTurnaround;
        this.withMovement = withMovement;
        this.changeParity = changeParity;
    }

    /**
     * Vérifie la validité de l'objet Connection en envoyant un message d'erreur si :
     *  - les gares de départ et d'arrivée de la connexion sont vides ou identiques,
     *  - la gare de départ ou d'arrivée n'existe pas,
     *  - la connexion existe déjà.
     * @returns {Connection | undefined} Objet Connection s'il est valide, undefined sinon.
     */
    check(): Connection | undefined {
        if (!this.from || !this.to) {
            CONSOLE_WARN.log(`Une connexion ne peut pas avoir des gares de départ`
                + ` et d'arrivée vides.`);
            return undefined;
        } else if (this.from === this.to) {
            CONSOLE_WARN.log(`Une connexion ne peut pas avoir des gares de départ`
                + ` et d'arrivée ${this.from} identiques et sans changement de parité.`);
            return undefined;
        } else if (!STATIONS.has(this.from.split("_")[0])) {
            CONSOLE_WARN.log(`La gare de départ ${this.from} de la connexion n'existe pas.`);
            return undefined;
        } else if (!STATIONS.has(this.to.split("_")[0])) {
            CONSOLE_WARN.log(`La gare d'arrivée ${this.to} de la connexion n'existe pas.`);
            return undefined;
        } else if (CONNECTIONS.has(this.from) && CONNECTIONS.get(this.from)!.has(this.to)) {
            CONSOLE_WARN.log(`La connexion ${this.from} -> ${this.to} est présente`
                + ` deux fois dans la base de données.`);
            return undefined;
        }
        return this;
    }
}

/* Liste des connexions entre les gares, incluant le temps de trajet et l'information
 *  sur le besoin de rebroussement. */
const CONNECTIONS = new Map<string, Map<string, Connection>>();

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
 * @param {boolean} [erase=false] Si vrai, force le rechargement des connections.
 *  Si faux (par défaut), ne recharge pas si déjà chargé.
 */
function loadConnections(erase: boolean = false) {

    // Vérifie si la table à charger existe déjà
    if (CONNECTIONS.size > 0) {
        if (erase) {
            CONNECTIONS.clear(); // Vide la map sans changer sa référence
        } else {
            return;
        }
    }

    loadStations();
    const data = getDataFromTable(CONNECTIONS_SHEET, CONNECTIONS_TABLE);

    // Parcourt la base de données
    for (const row of data.slice(1)) {

        // Extrait des valeurs
        const from = String(row[CONNECTIONS_COL_FROM]);
        const to = String(row[CONNECTIONS_COL_TO]);
        if (!from || !to) continue;
        const time = row[CONNECTIONS_COL_TIME] as number;
        const withTurnaround = row[CONNECTIONS_COL_TURNAROUND] as boolean;
        const withMovement = row[CONNECTIONS_COL_MOVEMENT] as boolean;
        const changeParity = row[CONNECTIONS_COL_CHANGE_PARITY] as boolean;

        // Crée l'objet Connection
        const connection = new Connection(
            from,
            to,
            time,
            withTurnaround,
            withMovement,
            changeParity
        ).check();
        if (!connection) continue;

        // Ajoute l'objet Connection dans la map
        if (!CONNECTIONS.has(from)) {
            CONNECTIONS.set(from, new Map<string, Connection>());
        }
        CONNECTIONS.get(from)!.set(to, connection);
    }
}

/**
 * Affiche les connexions entre les gares dans un tableau.
 * Les données sont celles stockées dans l'objet CONNECTIONS.
 * @param {string} sheetName Nom de la feuille de calcul.
 * @param {string} tableName Nom du tableau.
 * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
 */
function printConnections(sheetName: string, tableName: string, startCell: string = "A1"): void {

    // Convertit l'objet CONNECTIONS en un tableau de données
    const data: (string | number)[][] = [];
    for (const [from, connections] of CONNECTIONS) {
        for (const [to, connection] of connections) {
            data.push([
                from,
                to,
                connection.time,
                connection.withTurnaround ? 1 : 0,
                connection.withMovement ? 1 : 0,
                connection.changeParity ? 1 : 0
            ]);
        }
    }

    // Imprime le tableau
    const table = printTable(CONNECTIONS_HEADERS, data, sheetName, tableName, startCell);

    // Met les heures au format "hh:mm:ss"
    table.getRange().getColumn(CONNECTIONS_COL_TIME).setNumberFormat("hh:mm:ss");
}

/**
 * Sauvegarde les temps de connexions entre les gares dans l'objet CONNECTIONS.
 * Les données sont calculées en fonction des horaires de départ et d'arrivée des trains.
 * @param {string} [trainNumbers=""] Trains à traiter, séparés par des ; . Si vide,
 *  traite tous les trains.
 */
function saveConnectionsTimes(trainNumbers: string = "") {
    if (trainNumbers === "") {
        trainNumbers =
            Array.from(TRAIN_PATHS.keys()).filter(key => key === TRAIN_PATHS.get(key)!.key)
                .join(";");
    }
    trainNumbers.split(";").forEach((trainNumber) => {
        const trainPath = TRAIN_PATHS.get(trainNumber);
        trainPath?.stops.forEach((stop) => {
            if (stop.nextStop && CONNECTIONS.has(stop.key)
                && CONNECTIONS.has(stop.nextStop.key)) {
                const connection = CONNECTIONS.get(stop.key)!.get(stop.nextStop.key);
                if (connection && stop.nextStop.arrivalTime !== 0
                    && stop.departureTime !== 0) {
                    connection.time = stop.nextStop.arrivalTime - stop.departureTime;
                }
            }
        });
    });
}


