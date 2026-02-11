/**
 * Chargements de trains
 * 
 * Code Excel Automate pour la création et l'utilisation de la base de données des trains.
 * 
 * @author Paul Guignier
 * @version 1.0
 * @package scr\ChargementTrains.ts
 */

// export {loadParams, loadConnections, Paths.findShortestPath, Paths.generateCombinations};

/* Variables globales nécessaires dans ExcelScript (pas d'injection possible) */
var WORKBOOK: ExcelScript.Workbook;     // Classeur principal
var CONSOLE: Console;                   // Console pour l'affichage de messages

function main(workbook: ExcelScript.Workbook) {
    WORKBOOK = workbook;
    CONSOLE = console;
    const sheet = WORKBOOK.getActiveWorksheet();

    const testMode = true;
    Log.configure({
        debug: true,
        info: true,
        warn: true
    });

    

    // Lance la fonction de tests
    // Si les tests sont actifs, la suite du programme n'est pas exécuté. 
    if (runAllTests(testMode)) return;

    // Lit les paramètres
    Params.load();
    Connections.load();


    // Paths.load("", "147500_J;148504_J;147201_J;148202_J;147402_J;
    //      148402_J;147601_J;148602_J;145801_J;145804_J");
    // Paths.load("2", "142446_J");
    // Log.debug(Paths.map);
      // const allCombinations = Paths.generateCombinations("MPU", "ETP", "".split(";"));
    // Log.info(allCombinations);
    // const shortestPath = Paths.findShortestPath(allCombinations);
    // Log.info(shortestPath);

    return;
}

/**
 * Fonction de tests pour les différentes parties du code.
 * Lorsqu'elle est appelée, toutes les autres fonctions ne sont pas exécutées.
 * Les tests sont actifs si la constante TEST_MODE est vrai.
 * @param {boolean} [testMode=false] Si vrai, les fonctions de test sont lancés,
 *  puis le programme est interrompu. Si faux (par défaut), le programme continue normalement.
 * @returns {boolean} Si les tests sont actifs, la fonction renvoie true, sinon false.
 */
function runAllTests(testMode: boolean = false): boolean {

    if (!testMode) return false;

    Params.load();
    Connections.load();

    // testWorkbookService({ printSuccess: false, printFailure: true });
    // testDateTime({ printSuccess: false, printFailure: true });
    // testDay({ printSuccess: false, printFailure: true });
    // testParity({ printSuccess: false, printFailure: true });
    // testTrainNumber({ printSuccess: false, printFailure: true });
    // testStations({ printSuccess: false, printFailure: true });
    // testStationWithParity({ printSuccess: false, printFailure: true });
    // testConnections({ printSuccess: false, printFailure: true });

    Log.info(`Fin des tests`);
    Log.info(`-------------`);
    Log.info(`Fin des tests`);
    return true;

  

    /* Lecture des gares et test des variants */

    // Log.info(Paths.getAllVariants("VC"));

    // const allCombinations = Paths.generateCombinations("MPU", "ETP", "".split(";"));
    // Log.info(allCombinations);
    // const shortestPath = Paths.findShortestPath(allCombinations);
    // Log.info(shortestPath);
    // Paths.findShortestPath
    // Paths.calculateCompletePath
    // Paths.calculatePathTime
    // Dijkstra
    // Paths.generateCombinations
    // Paths.permute
    // expandPermutations
    // Paths.getAllVariants




    // const t1 = new Path(569000, 0, "1", "TEST", 12/24, "TRA-PG", 13/24, "PJ", "VFG");
    // Log.info(t1.getStop("VC-BG_2",true,true));
    // t1.findPath();
    // Log.debug(t1.getStop("INV_1"));

    /* Test path.getStop */
    // Paths.load("2", "147490");
    // loadStops();
    // const t2 = Paths.map.get("147490_2");
    // t2.findPath();
    // // Log.info(t2.getStop("VC-BG_2",true,true));
    // Log.info(t2);

    // Paths.findPathsOnAllPaths();
    // Paths.print("Test", "Trains1");
    // printStops("Test", "Stops1", "A10");
    // Log.info(Paths.map.get("147490_2"));

    // return true;
}

/* 
 * Options de l'affichage des logs
 *  - debug: Afficher les messages de debug
 *  - info: Afficher les messages d'information
 *  - warn: Afficher les messages d'avertissement
 */
type LogOptions = {
    debug: boolean;
    info: boolean;
    warn: boolean;
}
/*
 * Classe Logs contenant les trois types de messages du console, et leurs options d'affichage
 */
class Log {

    // Propriétés de la classe Log
    private static options: LogOptions = {  // Options de l'affichage des logs
        debug: true,                        //  Remontée des messages de debug
        info: true,                         //  Remontée des messages d'information
        warn: true                          //  Remontée des messages d'avertissement
    };                                      
    
    /**
     * Vérifie si une valeur est concatenable (null, undefined, string, number, boolean).
     * @param {unknown} value Valeur à vérifier.
     * @returns {boolean} Vrai si la valeur est concatenable, faux sinon.
     */
    private static isConcatable(value: unknown): boolean {
        return (
            value === null ||
            value === undefined ||
            typeof value === "string" ||
            typeof value === "number" ||
            typeof value === "boolean" ||
            typeof value === "bigint" ||
            value instanceof Date
        );
    }

    /*
     * Méthode interne qui écrit un message dans la console.
     * Elle prend en paramètre le niveau du message (debug, info, warn) et
     * un tableau d'arguments qui peuvent être des strings, des numbers, des
     * booleans, des null, des undefined, des objets.
     * Les arguments concaténables sont transformés en string et ajoutés
     * au buffer. Les objets sont ajoutés au tableau output sans
     * modification.
     * Lorsque le buffer contient un objet, il est flush (vide) pour
     * laisser place à l'objet.
     * Enfin, le tableau output est passé à CONSOLE.log pour afficher le
     * message.
     * @param {string} level Niveau du message (debug, info, warn)
     * @param {unknown[]} args Tableau des arguments à afficher
     */
    private static log(level: string, args: unknown[]): void {

        const output: unknown[] = [];
        let buffer = `[${level}]`;
    
        args.forEach(arg => {
            if (Log.isConcatable(arg)) {
                buffer += " " + String(arg);
            } else {
                // On flush le buffer texte avant l'objet
                if (buffer.trim() !== "") {
                    output.push(buffer);
                    buffer = "";
                }
                output.push(arg);
            }
        });
    
        // Flush final
        if (buffer.trim() !== "") {
            output.push(buffer);
        }
    
        CONSOLE.log(...output);
    }    
    
    /**
     * Configure les options de l'affichage des logs.
     * @param {Partial<LogOptions>} options Options de l'affichage des logs.
     *      - debug: Afficher les messages de debug.
     *      - info: Afficher les messages d'information.
     *      - warn: Afficher les messages d'avertissement.
     */
    public static configure(options: Partial<LogOptions>) {
        Object.assign(Log.options, options);
    }

    /**
     * Envoie un message au console avec le niveau "DEBUG".
     * @param {...unknown[]} args Arguments à passer au console.log
     */
    public static debug(...args: unknown[]): void {
        if (!Log.options.debug) return;
        Log.log("DEBUG", args);
    }
    
    /**
     * Envoie un message au console avec le niveau "INFO".
     * @param {...unknown[]} args Arguments à passer au console.log
     */
    public static info(...args: unknown[]): void {
        if (!Log.options.info) return;
        Log.log("INFO", args);
    }
    
    /**
     * Envoie un message au console avec le niveau "WARN".
     * @param {...unknown[]} args Arguments à passer au console.log
     */
    public static warn(...args: unknown[]): void {
        if (!Log.options.warn) return;
        Log.log("WARN", args);
    }
    
}

/*
 * Options de l'affichage des tests
 *  - printSuccess: Afficher le message de succès
 *  - printFailure: Afficher le message d'échec
 */
type AssertDDOptions = {
    printSuccess?: boolean;
    printFailure?: boolean;
}

/* 
 * Classe AssertDD contenant les options et les fonctions de tests Data-Driven
 */
class AssertDD {

    private options: AssertDDOptions;

    private total = 0;
    private success = 0;
    private failure = 0;

    constructor(options: AssertDDOptions = {}) {
        this.options = {
            printSuccess: options.printSuccess ?? true,
            printFailure: options.printFailure ?? true
        };
    }

    /**
     * Réalise le test et l'imprime avec un symbole de réussite (✔) ou d'échec (✘)
     * @param {string} label Nom du test
     * @param {T} actual Valeur actuelle obtenue
     * @param {T} expected Valeur attendue
     * @param {AssertDDOptions} options Options d'affichage des succès et des échecs
     */
    public check<T>(
        label: string,
        actual: T,
        expected: T,
        options: AssertDDOptions = {}
    ): boolean {

        const printSuccess = options.printSuccess ?? this.options.printSuccess;
        const printFailure = options.printFailure ?? this.options.printFailure;

        const ok = expected === actual;

        this.total++;
        ok ? this.success++ : this.failure++;

        if (ok) {
            if (printSuccess) {
                CONSOLE.log(`✔ ${label} | obtenu: ${expected}`);
            }
        } else {
            if (printFailure) {
                CONSOLE.log(
                    `✘ ${label} | attendu: ${expected} | obtenu: ${actual}`
                );
            }
        }

        return ok;
    }

    /**
     * Vérifie que la fonction fn lance une erreur.
     * @param {string} desc Nom du test
     * @param {() => void} fn Fonction à tester
     */
    public throws(desc: string, fn: () => void) {
        let thrown = false;
        try {
            fn();
        } catch {
            thrown = true;
        }
        this.check(desc, thrown, true);
    }

    /**
     * Imprime le resultat des tests
     * @param {string} [title="Résultats des tests"] Titre du test
     */
    public printSummary(title: string = "Résultats des tests", reset: boolean = true): void {
        CONSOLE.log(
            `${title} : ${this.success} / ${this.total} réussis`
            + ` (échecs : ${this.failure})`
        );
        if (reset) this.reset();
    }

    /**
     * Réinitialise le compteur de tests
     */
    public reset(): void {
        this.total = 0;
        this.success = 0;
        this.failure = 0;
    }
}

/*
 * Classe utilitaire de manipulation des feuilles de calcul Excel.
 */
class WorkbookService {

    /**
     * Renvoie la feuille de calcul Excel correspondant au nom donné.
     * Si la feuille n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * Si createIfMissing est vrai, crée la feuille si elle n'existe pas.
     * @param {string} sheetName Nom de la feuille de calcul à chercher.
     * @param {{failOnError?: boolean, createIfMissing?: boolean}} options Options
     *      pour la récupération de la feuille :
     *      - createIfMissing : Si vrai, crée la feuille si elle n'existe pas (faux par défaut).
     *      - failOnError : Si vrai (par défaut), lance une exception si la feuille n'existe pas.
     * @returns Feuille de calcul Excel correspondant au nom donné, ou null si elle n'existe pas.
     */
    public static getSheet(
        sheetName: string,
        options?: {
            createIfMissing?: boolean;  // Faux par défaut
            failOnError?: boolean;      // Vrai par défaut
        }
    ): ExcelScript.Worksheet | null {
    
        const createIfMissing = options?.createIfMissing ?? false;
        const failOnError = options?.failOnError ?? true;
        
        let sheet = WORKBOOK.getWorksheet(sheetName);
    
        if (!sheet) {
            if (createIfMissing) {
                sheet = WORKBOOK.addWorksheet(sheetName);
                Log.info(`Feuille "${sheetName}" créée.`);
                return sheet;
            }
    
            const msg = `La feuille "${sheetName}" n'existe pas.`;
    
            if (failOnError) throw new Error(msg);
    
            Log.warn(msg);
            return null;
        }
    
        return sheet;
    }

    /**
     * Renvoie le tableau Excel correspondant au nom donné dans la feuille de calcul donnée.
     * Si le tableau n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * @param {string} sheetName Nom de la feuille de calcul où chercher le tableau.
     * @param {string} tableName Nom du tableau à chercher.
     * @param {boolean} [failOnError=true] Si vrai (par défaut), lance une exception
     *  si le tableau n'existe pas. Si faux, renvoie null.
     * @returns {ExcelScript.Table | null} Tableau Excel correspondant au nom donné,
     *  ou null si il n'existe pas.
     */
    public static getTable(
        sheetName: string,
        tableName: string,
        failOnError: boolean = true
    ): ExcelScript.Table | null {
        const sheet = WorkbookService.getSheet(sheetName, { failOnError: false });
        if (!sheet) return null;
        const table = sheet.getTable(tableName);
        if (!table) {
            const msg = `Le tableau "${tableName}" n'existe pas dans la feuille "${sheetName}".`;
            if (failOnError) throw new Error(msg);
            Log.warn(msg);
            return null;
        }
        return table;
    }

    /**
     * Renvoie les données du tableau Excel correspondant au nom donné
     *  dans la feuille de calcul donnée.
     * Si le tableau n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * @param {string} sheetName Nom de la feuille de calcul où chercher le tableau.
     * @param {string} tableName Nom du tableau à chercher.
     * @param {boolean} [failOnError=true] Si vrai (par défaut),
     *  lance une exception si le tableau n'existe pas. Si faux, renvoie null.
     * @returns {(string | number | boolean)[][]} Données du tableau Excel
     *  correspondant au nom donné, ou null si il n'existe pas.
     */
    public static getDataFromTable(
        sheetName: string,
        tableName: string,
        failOnError: boolean = true
    ): (string | number | boolean)[][] {
        const table = WorkbookService.getTable(sheetName, tableName, failOnError);
        return table.getRange().getValues();
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne.
     * Si la valeur est null ou undefined, renvoie undefined.
     * Si la valeur est un nombre, le convertit en chaîne.
     * Si la valeur est une chaîne, la renvoie telle quelle, en supprimant les espaces
     *  inutiles.
     * @param {unknown[]} row Ligne contenant la cellule.
     * @param {number} col Colonne contenant la cellule.
     * @returns {string | undefined} Valeur de la cellule sous forme de chaîne,
     *  ou undefined si elle est null ou undefined.
     */
    static getString(row: unknown[], col: number): string | undefined {
        const v = row[col];
        if (v === undefined || v === null) return undefined;
        return String(v).trim() || undefined;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de nombre.
     * Si la valeur est null ou undefined, renvoie undefined.
     * Si la valeur est un nombre, le renvoie tel quel.
     * Si la valeur est une chaîne, essaie de la convertir en nombre en remplaçant les virgules
     *  par des points.
     * Si la conversion échoue, renvoie undefined.
     * @param {unknown[]} row Ligne contenant la cellule.
     * @param {number} col Colonne contenant la cellule.
     * @returns {number | undefined} Valeur de la cellule sous forme de nombre,
     *  ou undefined si la conversion échoue.
     */
    static getNumber(row: unknown[], col: number): number | undefined {
        const v = row[col];
        if (typeof v === "number") return v;
        if (typeof v === "string") {
            const n = Number(v.replace(",", "."));
            return Number.isFinite(n) ? n : undefined;
        }
        return undefined;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de booléen.
     * Si la valeur est null ou undefined, renvoie undefined.
     * Si la valeur est un booléen, le renvoie tel quel.
     * Si la valeur est un nombre, le renvoie converti en booléen
     *  (true si le nombre est différent de 0, false sinon).
     * Si la valeur est une chaîne, essaie de la convertir en booléen en remplaçant
     *  les chaînes "true", "1", "oui" et "yes" par true,
     *  et les chaînes "false", "0", "non" et "no" par false.
     * Si la conversion échoue, renvoie undefined.
     * @param {unknown[]} row Ligne contenant la cellule.
     * @param {number} col Colonne contenant la cellule.
     * @returns {boolean | undefined} Valeur de la cellule sous forme de booléen,
     *  ou undefined si la conversion échoue.
     */
    static getBoolean(row: unknown[], col: number): boolean | undefined {
        const v = row[col];
        if (typeof v === "boolean") return v;
        if (typeof v === "number") return v !== 0;
        if (typeof v === "string") {
            if (v === "") return undefined;
            return ["true", "1", "oui", "yes"].includes(v.toLowerCase());
        }
        return undefined;
    }   

    /**
     * Vérifie si l'adresse de cellule donnée est valide.
     * Si elle est valide, la renvoie telle quelle.
     * Si elle est invalide, lance une exception si failOnError est vrai,
     *  sinon renvoie une chaîne vide.
     * @param {string} cellName Adresse de cellule à vérifier.
     * @param {boolean} [failOnError=true] Si vrai (par défaut), lance une exception
     *  si l'adresse est invalide. Si faux, renvoie une chaîne vide.
     * @returns {string} Adresse de cellule si elle est valide, une chaîne vide sinon.
     */
    public static checkCellName(cellName: string, failOnError: boolean = true): string {
        // Convertit startCell en majuscules pour éviter les problèmes de casse
        cellName = cellName.toUpperCase();

        // Vérifie si cellName est une adresse de cellule valide
        if (!/^([A-Z]+)(\d+)$/.test(cellName)) {
            const msg = `L'adresse de départ ${cellName} n'est pas valide.`;
            if (failOnError) throw new Error(msg);
            Log.warn(msg);
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
    public static printTable(
        headers: string[][],
        data: (string | number | boolean)[][],
        sheetName: string,
        tableName: string,
        startCell: string = "A1",
        failOnError: boolean = true
    ): ExcelScript.Table | null {

        // Combine les en-têtes et les données
        const tableData = headers.concat(data);
// debug
        // const tableData = headers.concat(data.map(row => row.map(cell => cell.toString())));

        // Vérifie si les données sont non vides
        if (tableData.length === 0 || tableData[0].length === 0) {
            const msg = `Aucune donnée à insérer dans la table "${tableName}".`;
            if (failOnError) throw new Error(msg);
            Log.warn(msg);
            return;
        }

        // Vérifie si un tableau avec le même nom existe déjà et le supprime si nécessaire
        const sheet = WorkbookService.getSheet(sheetName, { createIfMissing: true, failOnError: false });
        const existingTable = sheet.getTables().find(table => table.getName() === tableName);
        if (existingTable) existingTable.delete();

        // Détermine la plage où écrire les données
        const startRange = sheet.getRange(WorkbookService.checkCellName(startCell));
        const writeRange = startRange
            .getResizedRange(tableData.length - 1, tableData[0].length - 1);

        // Efface le contenu de la plage
        writeRange.clear(ExcelScript.ClearApplyTo.contents);

        // Écrit les données dans la plage
        writeRange.setValues(tableData);

        // Ajoute un nouveau tableau
        const table = sheet.addTable(writeRange.getAddress(), true);
        table.setName(tableName);

        Log.info(`Le tableau "${tableName}" a été créé avec succès`
            + ` dans la feuille "${sheetName}".`);

        return table;
    }
}

/*
 * Classe utilitaire contenant les paramètres globaux
 */
class Params {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";                // Feuille contenant les paramètres globaux
    private static readonly TABLE = "Paramètres";           // Tableau contenant les paramètres globaux
    private static readonly ROW_MAX_CONNEXIONS_NUMBER = 1;  // Ligne contenant le nombre maximum de connexions
    private static readonly ROW_TURNAROUND_TIME = 2;        // Ligne contenant le temps de retournement (en minutes)
    private static readonly ROW_MAX_TRAIN_UNITS = 3;        // Ligne contenant le nombre maximal d'unités en UM
    private static readonly ROW_TERMINUS_NAME = 5;          // Ligne contenant le nom du terminus

    // Indicateur de chargement
    private static loaded = false;

    // Paramètres globaux
    public static maxConnectionNumber: number;              // Nombre maximum de connexions
    public static turnaroundTime: DateTime;                 // Temps de retournement
    public static maxTrainUnits: number;                    // Nombre maximal d'unités en UM
    public static terminusName: string;                     // Nom du terminus

    /**
     * Chargement global des paramètres
     * @param {boolean} erase Si vrai, force le rechargement des paramètres
     */
    public static load(erase: boolean = false): void {
        if (Params.loaded && !erase) return;

        // Chargement des paramètres des classes utilitaires
        DateTime.load();
        Day.load();
        Parity.load();
        TrainNumber.load();

        // Chargement des autres paramètres
        const data = WorkbookService.getDataFromTable(Params.SHEET, Params.TABLE);

        Params.maxConnectionNumber = WorkbookService.getNumber(data[Params.ROW_MAX_CONNEXIONS_NUMBER], 1) ?? 6;
        const turnaroundTime = WorkbookService.getNumber(data[Params.ROW_TURNAROUND_TIME], 1) ?? 10;
        Params.turnaroundTime = DateTime.from(turnaroundTime / 24 / 60, true)!;

        Params.maxTrainUnits = WorkbookService.getNumber(data[Params.ROW_MAX_TRAIN_UNITS], 1) ?? 2;
        Params.terminusName = WorkbookService.getString(data[Params.ROW_TERMINUS_NAME],1) ?? "Terminus";

        Params.loaded = true;
    }
}

/**
 * Classe utilitaire immutable pour la gestion des dates et horaires Excel.
 *  Si l'heure est inférieure à l'heure de changement de journée,
 *  elle est incrémentée de 1 pour rester comparable aux autres heures de la journée précédente.
 */
class DateTime {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";        // Feuille contenant les paramètres globaux
    private static readonly TABLE = "Paramètres";   // Tableau contenant les paramètres globaux
    private static readonly ROW_ROLLOVER_HOUR = 4;  // Ligne contenant l'heure de changement de journée
    private static readonly MIN_EXCEL_DATE = 2;     // Valeur minimale d'un temps incluant une date

    // Etat de chargement
    private static loaded = false;

    // Heure de changement de journée (fraction de jour Excel)
    public static rolloverHour: number;             // Heure de changement de journée (en temps Excel)

    // Format des dates et heures
    public static readonly DATE_FORMAT_FOR_ID: string = "yymmdd";
    public static readonly DATE_FORMAT_WITH_YEAR: string = "dd/mm/yyyy";
    public static readonly DATE_FORMAT_WITHOUT_YEAR: string = "dd/mm";
    public static readonly TIME_FORMAT_WITH_SECONDS: string = "hh:nn:ss";
    public static readonly TIME_FORMAT_WITHOUT_SECONDS: string = "hh:nn";

    // Propriétés de l'objet DateTime
    public readonly excelValue: number;             // Valeur du temps en format Excel
                                                    //  à partir du 01/01/1900 00:00:00
    public readonly isRelative: boolean = false;    // Indique si le temps est relatif
                                                    //  (différence entre 2 horaires)
    private _dateComputed = false;                  // Indique si les éléments de la date
                                                    //  et de l'heure ont été calculés                                              
    public year: number = 0;                        // Année
    public month: number = 0 ;                      // Mois
    public day: number = 0 ;                        // Jour
    public dayOfWeek: Day | undefined;              // Jour de la semaine
    public hours: number = 0;                       // Heures
    public minutes: number = 0;                     // Minutes
    public seconds: number = 0;                     // Secondes               

    /**
     * Constructeur de l'objet DateTime.
     * @param {number|string} excelValue Valeur Excel du temps en nombre décimal ou en chaîne de caractères.
     * @param {boolean} [isRelative=false] Indique si le temps est relatif (différence entre 2 horaires).
     * @param {boolean} [adaptTime=true] Indique si l'heure doit être adaptée ou non.
     * Si la valeur est inférieure à l'heure de changement de journée et que l'horaire est absolu,
     *  elle est incrémentée de 1.
     * @param {boolean} [isRelative=false] Indique si le temps est relatif (différence entre 2 horaires).
     */
    private constructor(excelValue: number = 0, isRelative: boolean = false, adaptTime: boolean = true) {
        this.isRelative = isRelative;
        this.excelValue = (adaptTime && !this.isRelative) ? DateTime.adaptTime(excelValue) : excelValue;
    }

    /**
     * Crée un objet DateTime à partir d'une valeur.
     * Si la valeur est déjà un objet DateTime, il est retourné tel quel.
     * Sinon, un nouvel objet DateTime est créé avec la valeur fournie.
     * @param {number|string|DateTime|null|undefined} value Valeur du temps en nombre décimal
     *  ou en chaîne de caractères.
     * @param {boolean} [isRelative=false] Indique si le temps est relatif (différence entre 2 horaires).
     * @param {boolean} [adaptTime=true] Indique si l'heure doit être adaptée ou non.
     *  Si la valeur est inférieure à l'heure de changement de journée et que l'horaire est absolu,
     *  elle est incrémentée de 1.
     * @returns {DateTime | undefined} Nouvel objet DateTime égal à la valeur fournie, ou undefined.
     * @throws {Error} Si la valeur est un temps relatif et qu'on cherche à l'affecter à un temps absolu.
     */
    public static from(
        value: number | string | DateTime | null | undefined,
        isRelative: boolean = false,
        adaptTime: boolean = true
    ): DateTime | undefined{
        
        if (value == null || value === "") {
            return undefined;
        }

        if (value instanceof DateTime) {
            if (value.isRelative !== isRelative) {
                throw new Error(
                    `Un temps ${value.isRelative ? "relatif" : "absolu"}`
                    + ` cherche à être affecté à un temps ${isRelative ? "relatif" : "absolu"}.`
                );
            }
            return value;
        }

        const v = value == null ? 0 : Number(value);

        // Un temps absolu doit être supérieur ou égal à 0
        if (!isRelative && v < 0) return undefined;

        return new DateTime(v, isRelative, adaptTime);
    }  

    /**
     * Calcule les éléments de la date et de l'heure.
     * Si le temps est absolu et pas seulement une heure, calcule les éléments de la date.
     * Sinon, calcule uniquement les éléments de l'heure.
     * Les éléments sont stockés dans les propriétés de l'objet DateTime.
     */
    private computeDate(): void {
        if (this._dateComputed) return;

        // Calcul des éléments de la date (si absolu et pas seulement une heure)
        if (!this.isRelative && this.excelValue > DateTime.MIN_EXCEL_DATE) {
            const excelBase = new Date(Date.UTC(1899, 11, 30));
            const days = Math.floor(this.excelValue);
            const date = new Date(excelBase.getTime() + days * 86400000);

            this.year = date.getUTCFullYear();
            this.month = date.getUTCMonth() + 1;
            this.day = date.getUTCDate();
            const jsDay = date.getUTCDay();
            this.dayOfWeek = Day.fromNumber(jsDay);
        }

        // Calcul des éléments de l'heure
        const absValue = Math.abs(this.excelValue);
        const dayFraction = absValue % 1;
        const totalSeconds = Math.round(dayFraction * 86400);
        this.hours = Math.floor(totalSeconds / 3600);
        this.minutes = Math.floor((totalSeconds % 3600) / 60);
        this.seconds = totalSeconds % 60;

        this._dateComputed = true;
    }

    /**
     * Renvoie un nouvel objet DateTime égal au temps courant
     * résolu par rapport à une référence.
     * Si le temps courant est relatif, il est ajouté à la référence.
     * Sinon, le temps courant est renvoyé tel quel.
     * @param {DateTime} reference Référence à utiliser pour résoudre le temps courant.
     * @returns {DateTime} Nouvel objet DateTime égal au temps courant résolu par rapport à la référence.
     * @throws {Error} Si la référence est un temps relatif.
     */
    public resolveAgainst(reference: DateTime): DateTime {
        if (this.isRelative) {
            if (reference.isRelative) {
                throw new Error(`La référence doit être un temps absolu`);
            }
            return new DateTime(
                reference.excelValue + this.excelValue,
                false
            );
        }

        return this;
    }

    /**
     * Renvoie un nouvel objet DateTime égal au temps courant relatif par rapport à une référence.
     * Les deux temps doivent être absolus.
     * @param {DateTime} reference Référence à utiliser pour résoudre le temps courant.
     * @returns {DateTime} Nouvel objet DateTime égal au temps courant relatif par rapport à la référence.
     * @throws {Error} Si l'un des deux temps est relatif.
     */
    public relativeTo(reference: DateTime): DateTime {
        if (this.isRelative || reference.isRelative) {
            throw new Error(`Les deux temps doivent être absolus`);
        }

        return new DateTime(
            this.excelValue - reference.excelValue,
            true
        );
    }

    /**
     * Vérifie si le temps courant est égal à un autre temps.
     * @param {DateTime | null | undefined} other Temps à comparer.
     * @returns {boolean} Vrai si les deux temps sont égaux, faux sinon.
     */
    public equalsTo(other: DateTime | null | undefined): boolean {
        return (
            !! other &&
            this.isRelative === other.isRelative &&
            this.excelValue === other.excelValue
        );
    }

    /**
     * Compare le temps courant avec un autre temps.
     * @param {DateTime} other Temps à comparer.
     * @returns {number} Différence entre les deux temps.
     * @throws {Error} Si les deux temps ont des types différents (relatif ou absolu).
     */
    public compareTo(other: DateTime): number {
        if (this.isRelative !== other.isRelative) {
            throw new Error(`Un temps relatif ne peut pas être comparé avec un temps absolu`);
        }
        return this.excelValue - other.excelValue;
    }
    
    /**
     * Ajoute un temps relatif à un autre temps relatif.
     * @param {DateTime} other Temps relatif à ajouter.
     * @returns {DateTime} Nouvel objet DateTime égal à la somme des deux temps relatifs.
     * @throws {Error} Si l'un des deux temps n'est pas relatif.
     */
    public add(other: DateTime): DateTime {
        if (!this.isRelative || !other.isRelative) {
            throw new Error(`L'addition n'est possible qu'entre temps relatifs`);
        }

        return new DateTime(
            this.excelValue + other.excelValue,
            true
        );
    }

    /**
     * Soustrait au temps relatif un autre temps relatif.
     * @param {DateTime} other Temps relatif à soustraire.
     * @returns {DateTime} Nouvel objet DateTime égal à la différence entre les deux temps relatifs.
     * @throws {Error} Si l'un des deux temps n'est pas relatif.
     */
    public subtract(other: DateTime): DateTime {
        if (!this.isRelative || !other.isRelative) {
            throw new Error(`La soustraction n'est possible qu'entre temps relatifs`);
        }

        return new DateTime(
            this.excelValue - other.excelValue,
            true
        );
    }

    /**
     * Formate la date ou l'heure en fonction du format fourni.
     * @param {string} format Format de la date ou de l'heure.
     * @returns {string} Date ou heure formattée.
     */
    public format(format: string): string {
        this.computeDate();
        let prefix = "";
        if (this.excelValue < 0) prefix = "-";
        const pad = (v: number) => v.toString().padStart(2, "0");
    
        const tokens: Record<string, string> = {
        // Année
        "yyyy": this.year.toString(),
        "yy": pad(this.year % 100),
        // Mois
        "mm": pad(this.month),
        "m": this.month.toString(),
        // Jour de semaine
        "dddd": this.dayOfWeek?.fullName ?? "",
        "ddd": this.dayOfWeek?.abreviation ?? "",
        // Jour
        "dd": pad(this.day),
        "d": this.day.toString(),
        // Heures
        "hh": pad(this.hours),
        "h": this.hours.toString(),
        // Minutes
        "nn": pad(this.minutes),
        "n": this.minutes.toString(),
        // Secondes
        "ss": pad(this.seconds),
        "s": this.seconds.toString(),
        };

        // Création des clés temporaires
        const tempMap: Record<string, string> = {};
        let i = 0;
        let tempFormat = format.toLowerCase();

        Object.keys(tokens)
            .sort((a, b) => b.length - a.length)
            .forEach(token => {
                const tempKey = `__TOKEN${i}__`;
                const re = new RegExp(token, "g");
                tempFormat = tempFormat.replace(re, tempKey);
                tempMap[tempKey] = tokens[token];
                i++;
            });

        // Remplacer toutes les clés temporaires par les valeurs réelles
        Object.entries(tempMap).forEach(([key, val]) => {
            tempFormat = tempFormat.replace(new RegExp(key, "g"), val);
        });
        
        return prefix + tempFormat;
    }

    /**
     * Vérifie si les deux parités sont identiques ou si elles sont toutes les deux undefined.
     * @param {DateTime | undefined} a Première parité à comparer.
     * @param {DateTime | undefined} b Seconde parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont identiques
     *  ou si elles sont toutes les deux undefined, faux sinon.
     */
    public static equalsOrUndefined(
        a?: DateTime,
        b?: DateTime
    ): boolean {
        return a === b || (!!a && !!b && a.equalsTo(b));
    }

    /**
     * Ajuste une heure pour tenir compte du changement de journée.
     * Si l'heure est inférieure à l'heure de changement de journée,
     *  on ajoute 1 pour passer à la journée suivante.
     *  Par exemple : 01:00 → 25:00 si changement de journée à 03:00
     * Cela ne s'applique pas sur les heures datées (valeur > 2).
     * @param {number} time Heure à ajuster.
     * @returns {number} Heure ajustée.
     */
    public static adaptTime(time: number): number {
        return (time < DateTime.rolloverHour && time < 2) ? time + 1 : time;
    }
    
    /**
     * Charge le paramètre de l'heure de changement de journée.
     */
    public static load(erase = false): void {
        if (DateTime.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(DateTime.SHEET, DateTime.TABLE);

        DateTime.rolloverHour = (WorkbookService.getNumber(data[DateTime.ROW_ROLLOVER_HOUR], 1) ?? 0) % 1;

        DateTime.loaded = true;
    }
}

/**
 * Classe utilitaire pour la gestion des jours de la semaine, individuels ou groupés. 
 *  (JOB du lundi au vendredi, WE pour samedi et dimanche...).
 */
class Day {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";        // Feuille contenant les paramètres des jours de la semaine
    private static readonly TABLE = "Jours";        // Tableau contenant les paramètres des jours de la semaine
    private static readonly COL_FULL_NAME = 0;      // Colonne contenant le nom complet du jour de la semaine
    private static readonly COL_ABBREVIATION = 1;   // Colonne contenant l'abréviation du jour de la semaine
    private static readonly COL_NUMBERS = 2;        // Colonne contenant le numéro du jour  

    // Indicateur de chargement
    private static loaded = false;

    // Map des jours de la semaine
    private static readonly daysByNumbers  = new Map<string, Day>();
    
    // Cache pour l'extraction des jours de la semaine depuis une chaine de caractères
    private static readonly cache = new Map<string, Map<string, number[]>>();

    // Propriétés de l'objet Day
    fullName: string;               // Nom du jour ou du groupe de jours de la semaine
    abreviation: string;            // Abréviation du jour ou du groupe de jours de la semaine
    numbersString: string;          // Numéro(s) concaténés des jours de la semaine
                                    //  en chaine de caractères (avec ou sans ponctuation)
    number: number;                 // Numéro du jour de la semaine (de 1 : lundi à 7 : dimanche,
                                    //  0 si l'objet est un groupe de jours)

    /**
     * Constructeur de la classe Day.
     * @param {string} fullName Nom complet du jour ou du groupe de jours de la semaine.
     * @param {string} abreviation Abréviation du jour ou du groupe de jours de la semaine.
     * @param {string} numbersString Chaîne de caractères contenant
     *  les numéros de jours de la semaine.
     */
    constructor(fullName: string, abreviation: string, numbersString: string) {
        if (!fullName)
            throw new Error(`Ligne vide ou non renseignée dans le tableau des jours`);
        this.fullName = fullName;
        if (!abreviation)
            throw new Error(`Groupe de jours ${fullName} : Abbréviation non renseignée`);
        this.abreviation = abreviation;
        const cleanAndSortedNumbersString = Day.cleanAndSortNumbers(numbersString.toString()).join('');
        if (!cleanAndSortedNumbersString)
            throw new Error(`Groupe de jours ${fullName} : Numéros des jours non renseignés`);
        this.numbersString = cleanAndSortedNumbersString;
        const number = parseInt(this.numbersString);
        this.number = number > 7 ? 0 : number;
    }

    /**
     * Renvoie un objet Day correspondant au numéro de jour fourni.
     * Si le numéro de jour n'existe pas, renvoie undefined.
     * Charge les paramètres des jours de la semaine si ce n'est pas déjà fait.
     * @param {number} dayNumber Numéro du jour de la semaine
     *  (de 1 : lundi à 6 : samedi, 0 ou 7 : dimanche).
     * @returns {Day | undefined} Objet Day correspondant au numéro de jour fourni,
     *  ou undefined si le numéro de jour n'existe pas.
     */
    public static fromNumber(dayNumber: number): Day | undefined {
        const key = String(dayNumber === 0 ? 7 : dayNumber);
        const day = Day.daysByNumbers.get(key);
        if (!day) {
            throw new Error(`Jour de semaine invalide : ${dayNumber}`);
        }
        return day;
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
        if (Day.cache.has(key1) && Day.cache.get(key1)!.has(key2)) {
            return Day.cache.get(key1)!.get(key2)!;
        }

        // Analyse la chaine
        let result: number[] = [];
        if (!key2) {

            // Analyse la chaine pour la transformer en tableau
            let processed = key1;

            // Reconnaissance Regex des noms entiers de jours
            Array.from(Day.daysByNumbers .values())
                .sort((a, b) => b.fullName.length - a.fullName.length)
                .forEach(day => {
                    const regex = new RegExp(day.fullName.toLowerCase(), 'g');
                    processed = processed.replace(regex, day.numbersString);
                });

            // Reconnaissance Regex des abréviations de jours
            Array.from(Day.daysByNumbers .values())
                .sort((a, b) => b.abreviation.length - a.abreviation.length)
                .forEach(day => {
                    const regex = new RegExp(day.abreviation.toLowerCase(), 'g');
                    processed = processed.replace(regex, day.numbersString);
                });

            // Reconnaissance Regex des numéros de jours
            processed = processed.replace(/[^1-7]/g, '');
            result = Day.cleanAndSortNumbers(processed);
        } else {
            // Calcule l'intersection des deux chaines.
            const days1 = Day.extractFromString(key1);
            const days2 = Day.extractFromString(key2);

            result = days1.filter(n => days2.includes(n));
        }

        // Met en cache le resultat pour une utilisation similaire de la fonction
        if (!Day.cache.has(key1)) {
            Day.cache.set(key1, new Map<string, number[]>());
        }
        Day.cache.get(key1)!.set(key2, result);

        return result;
    }

    /**
     * Charge les jours de la semaine à partir du tableau "Jours" de la feuille "Param".
     * Les jours sont stockés dans la structure Day.daysByNumbers , sous forme de map, avec
     *  le nom complet et l'abréviation du jour comme clés, et leur numéro correspondant
     *  comme valeur.
     */
    public static load(erase = false): void {

        // Vérifie si la table à charger existe déjà
        if (Day.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(Day.SHEET, Day.TABLE);

        for (const [rowIndex, row] of data.slice(1).entries()) {

            // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides)
            if (row.every((cell: unknown) => !cell)) continue;

            // Calcule le numéro de ligne Excel
            const excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

            // Extrait les valeurs
            const fullName = WorkbookService.getString(row, Day.COL_FULL_NAME) ?? "";
            const abreviation = WorkbookService.getString(row, Day.COL_ABBREVIATION) ?? "";
            const numbersString = WorkbookService.getString(row, Day.COL_NUMBERS) ?? "";

            // Crée l'objet Day
            try {
                const day = new Day(fullName, abreviation, numbersString);
                Day.daysByNumbers.set(day.numbersString, day);
            } catch (e) {
                Log.warn(`Day.load (ligne ${excelRow}) : ${e}`);
                continue;
            }
        }

        // Si non renseignés dans le tableau, charge les jours individuels par défaut
        if (!Day.daysByNumbers.has('1')) Day.daysByNumbers.set('1', new Day('Lundi', 'Lu', '1'));
        if (!Day.daysByNumbers.has('2')) Day.daysByNumbers.set('2', new Day('Mardi', 'Ma', '2'));
        if (!Day.daysByNumbers.has('3')) Day.daysByNumbers.set('3', new Day('Mercredi', 'Me', '3'));
        if (!Day.daysByNumbers.has('4')) Day.daysByNumbers.set('4', new Day('Jeudi', 'Je', '4'));
        if (!Day.daysByNumbers.has('5')) Day.daysByNumbers.set('5', new Day('Vendredi', 'Ve', '5'));
        if (!Day.daysByNumbers.has('6')) Day.daysByNumbers.set('6', new Day('Samedi', 'Sa', '6'));
        if (!Day.daysByNumbers.has('7')) Day.daysByNumbers.set('7', new Day('Dimanche', 'Di', '7'));

        Day.loaded = true;
    }
}

/*
 * Classe utilitaire immutable qui permet de manipuler la parité
 *  d'un train, d'un parcours ou d'un arrêt.
 */
class Parity {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";        // Feuille contenant les paramètres de parité
    private static readonly TABLE = "Parité";       // Tableau contenant les paramètres de parité
    private static readonly ROW_ODD = 1;            // Ligne de la parité impaire
    private static readonly ROW_EVEN = 2;           // Ligne de la parité paire
    private static readonly ROW_DOUBLE = 3;         // Ligne de la parité double
    private static readonly COL_LETTER = 1;         // Colonne des parités exprimées en lettres
    private static readonly COL_NUMBER = 2;         // Colonne des parités exprimées en chiffres

    // Constantes de parité
    public static readonly ODD: number = 1;         // Parité impaire
    public static readonly EVEN: number = 2;        // Parité paire
    public static readonly DOUBLE: number = -2;     // Parité double
    public static readonly UNDEFINED: number = -1;  // Parité non définie

    // Indicateur de chargement
    private static loaded = false;

    // Map des lettres et nombres désignants les parités
    private static letters = new Map<number, string>();
    private static digits = new Map<number, number>();

    // Propriétés de l'objet Parity
    public readonly value: number;                         // Valeur de la parité
    private readonly doubleParityAllowed: boolean;           // Autorise une double parité

    /**
     * Constructeur de la classe Parity.
     * Initialise une instance de parité avec une valeur spécifiée,
     *  qui peut être une lettre de parité, un chiffre de parité, ou un numéro de train.
     * Analyse la valeur donnée pour déterminer la parité.
     * @param {string | number} value Valeur à analyser pour la parité.
     * @param {boolean} [doubleParityAllowed=false] Si vrai, la double parité est autorisée.
     *  Si faux (par défaut), la double parité est impossible.
     */
    private constructor(
        value: string | number | null | undefined,
        doubleParityAllowed: boolean = false
    ) {
        this.doubleParityAllowed = doubleParityAllowed;
        this.value = this.normalizeParityValue(value);
    }

    /**
     * Retourne une instance de Parity à partir d'une valeur qui peut être :
     *  - une lettre de parité (ou la concaténation des deux lettres sans ordre si double parité)
     *  - un chiffre de parité (format chaîne ou nombre)
     *  - un numéro de train (pair, impair ou double s'il contient un '/')
     *  - une instance de Parity (retourne la même instance)
     *  - null ou undefined (retourne une instance de Parity avec valeur Parity.UNDEFINED)
     * @param {string | number | Parity | null | undefined} value Valeur à analyser pour la parité.
     * @returns {Parity} Instance de Parity correspondante.
     */
    public static from(
        value: string | number | Parity | null | undefined,
        doubleParityAllowed: boolean = false
    ): Parity {
        return new Parity(value instanceof Parity ? value.value : value, doubleParityAllowed);
    }    

    /**
     * Analyse une valeur qui indique la parité, qui peut être :
     *  - la lettre de parité (ou la concaténation des deux lettres sans ordre si double parité)
     *  - le chiffre de parité (format chaîne ou nombre)
     *  - un numéro de train (pair, impair ou double s'il contient un '/')
     * @param {string | number | null | undefined} value Nouvelle valeur de la parité.
     */
    private normalizeParityValue(value: string | number | null | undefined): number {

        // 1️⃣ null / undefined
        if (value == null) {
            return Parity.UNDEFINED;
        }
    
        // 2️⃣ NUMBER — traitement prioritaire
        if (typeof value === 'number') {
    
            // Valeurs de Parity explicites
            if (
                value === Parity.UNDEFINED ||
                value === Parity.ODD ||
                value === Parity.EVEN ||
                value === Parity.DOUBLE
            ) {
                return value === Parity.DOUBLE && !this.doubleParityAllowed
                    ? Parity.UNDEFINED
                    : value;
            }
    
            // Nombres <= 0 → undefined
            if (value <= 0) {
                return Parity.UNDEFINED;
            }
    
            // Parité du nombre
            return value % 2 === 0 ? Parity.EVEN : Parity.ODD;
        }

        // 3️⃣ STRING
        const str = value.trim().toUpperCase();
    
        if (str === '' || str === '0') {
            return Parity.UNDEFINED;
        }

        // Double implicite (ex: "12345/6")
        if (str.includes('/')) {
            return this.doubleParityAllowed ? Parity.DOUBLE : Parity.UNDEFINED;
        }
    
        // Tentative de conversion numérique
        const numeric = parseInt(str, 10);
        if (!Number.isNaN(numeric)) {
            return this.normalizeParityValue(numeric);
        }
    
        // Lettres
        switch (str) {
            case Parity.letter(Parity.ODD):
                return Parity.ODD;
            case Parity.letter(Parity.EVEN):
                return Parity.EVEN;
            case Parity.letter(Parity.ODD)! + Parity.letter(Parity.EVEN)!:
            case Parity.letter(Parity.EVEN)! + Parity.letter(Parity.ODD)!:
                return this.doubleParityAllowed ? Parity.DOUBLE : Parity.UNDEFINED;
        }
    
        return Parity.UNDEFINED;
    }

    /**
     * Vérifie si la parité est identique à une valeur de parité donnée.
     * @param {number} parity Autre valeur de parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont identiques, faux sinon.
     */
    public is(parity: number): boolean {
        return this.value === parity;
    }

    /**
     * Vérifie si la parité est définie (différente de Parity.UNDEFINED).
     * @returns {boolean} Vrai si la parité est définie, faux sinon.
     */
    public isDefined(): boolean {
        return this.value !== Parity.UNDEFINED;
    }

    /**
     * Vérifie si deux parités sont opposées (parité impaire versus parité paire).
     * @param {Parity | undefined} other Autre variable de parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont opposées, faux sinon.
     */
    public isOpposedTo(other: Parity | undefined): boolean {
        return !!other
            && (this.value === Parity.ODD && other.value === Parity.EVEN
                || this.value === Parity.EVEN && other.value === Parity.ODD);
    }

    /**
     * Vérifie si la parité est définie et identique à celle d'une autre variable de parité.
     * @param {Parity | undefined} other Autre variable de parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont identiques, faux sinon.
     */
    public equalsTo(other: Parity | undefined): boolean {
        return (
            !! other
                && this.value === other.value
                && this.doubleParityAllowed === other.doubleParityAllowed
        );
    }

    /**
     * Vérifie si la parité inclut une autre valeur de parité.
     * @param {Parity | number | null | undefined} other Autre valeur de parité à inclure.
     * @returns {boolean} Vrai si la parité inclut la valeur de parité, faux sinon.
     */
    public includes(other: string | number | Parity | null | undefined): boolean {
        const requested = Parity.from(other, this.doubleParityAllowed);
    
        if (!requested) {
            return false;
        }
    
        // undefined n'inclut rien
        if (this.value === Parity.UNDEFINED) {
            return false;
        }
    
        // Parité double inclut toutes les parités définies
        if (this.value === Parity.DOUBLE) {
            return requested.value !== Parity.UNDEFINED;
        }
    
        // Sinon : égalité stricte
        return this.value === requested.value;
    }
    
    /**
     * Inverse la parité actuelle.
     * Si la parité actuelle est paire, elle devient impaire, et inversement.
     * Si la parité actuelle est double, elle reste double.
     * Si la parité actuelle est indéfinie, elle reste inchangée.
     * @returns {Parity} Parité inversée.
     */
    public invert(): Parity {
        switch (this.value) {
            case Parity.ODD:
                return Parity.from(Parity.EVEN, this.doubleParityAllowed);
            case Parity.EVEN:
                return Parity.from(Parity.ODD, this.doubleParityAllowed);
            case Parity.DOUBLE:
                return Parity.from(Parity.DOUBLE, this.doubleParityAllowed);
            default:
                return Parity.from(Parity.UNDEFINED, this.doubleParityAllowed);
        }
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
        switch (this.value) {
            case Parity.ODD:
                const oddDigit = Parity.digit(Parity.ODD)!;
                return withUnderscores ? '_' + oddDigit : oddDigit;
            case Parity.EVEN:
                const evenDigit = Parity.digit(Parity.EVEN)!;
                return withUnderscores ? '_' + evenDigit : evenDigit;
            case Parity.DOUBLE:
                return Parity.digit(Parity.DOUBLE)!;
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
        switch (this.value) {
            case Parity.ODD:
                return Parity.letter(Parity.ODD);
            case Parity.EVEN:
                return Parity.letter(Parity.EVEN);
            case Parity.DOUBLE:
                return Parity.letter(Parity.ODD)!
                    + Parity.letter(Parity.EVEN)!;
            default:
                return "";
        }
    }

    /**
     * Vérifie si les deux parités sont identiques ou si elles sont toutes les deux undefined.
     * @param {Parity | undefined} a Première parité à comparer.
     * @param {Parity | undefined} b Seconde parité à comparer.
     * @returns {boolean} Vrai si les deux parités sont identiques
     *  ou si elles sont toutes les deux undefined, faux sinon.
     */
    public static equalsOrUndefined(
        a?: Parity,
        b?: Parity
    ): boolean {
        return a === b || (!!a && !!b && a.equalsTo(b));
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
            case Parity.ODD:
                return string.toUpperCase().includes(Parity.letter(Parity.ODD)!);
            case Parity.EVEN:
                return string.toUpperCase().includes(Parity.letter(Parity.EVEN)!);
            case Parity.DOUBLE:
                return string.toUpperCase().includes(Parity.letter(Parity.ODD)!)
                    && string.toUpperCase().includes(Parity.letter(Parity.EVEN)!);
            default:
                return false;
        }
    }

    /**
     * Retourne la lettre de parité correspondante.
     * @param {number} parity Valeur de la parité.
     * @returns {string} Lettre de parité correspondante, ou une chaîne vide si la parité est undefined.
     */
    public static letter(parity: number): string {
        return Parity.letters.get(parity) ?? "";
    }

    /**
     * Retourne le chiffre de parité correspondant.
     * @param {number} parity Valeur de la parité.
     * @returns {number} Chiffre de parité correspondante, ou 0 si la parité est undefined.
     */
    public static digit(parity: number): number {
        return Parity.digits.get(parity) ?? 0;
    }

    /**
     * Charge les paramètres de parité des jours (lettre et chiffre associés
     *  aux jours pairs et impairs) à partir de la feuille "Param".
     */
    public static load(erase = false): void {
        if (Parity.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(Parity.SHEET, Parity.TABLE);

        const getParityLetter = (
            row: number,
            fallback: string
        ): string =>
            (
                WorkbookService.getString(data[row], Parity.COL_LETTER) ?? fallback
            ).toUpperCase();

        Parity.letters.set(Parity.ODD, getParityLetter(Parity.ROW_ODD, "I"));
        Parity.letters.set(Parity.EVEN, getParityLetter(Parity.ROW_EVEN, "P"));

        const getParityDigit = (
            row: number,
            fallback: number
        ): number =>
            WorkbookService.getNumber(data[row], Parity.COL_NUMBER) ?? fallback;

        Parity.digits.set(Parity.ODD, getParityDigit(Parity.ROW_ODD, 1));
        Parity.digits.set(Parity.EVEN, getParityDigit(Parity.ROW_EVEN, 2));
        Parity.digits.set(Parity.DOUBLE, getParityDigit(Parity.ROW_DOUBLE, -2));

        Parity.loaded = true;
    }
}

/**
 * Classe TrainNumber définissant un numéro de train.
 * Il est alphanumérique, sans ponctuation et sans espaces, avec un chiffre pour dernier caractère.
 * La double parité est marquée par ######/#
 */
class TrainNumber {

    // Constantes de lecture de la base de données Excel
    private static readonly W_SHEET = "Param";                          // Feuille contenant les motifs des trains W
    private static readonly W_TABLE = "W";                              // Tableau contenant les motifs des trains W
    private static readonly TRAINS_4DIGIT_SHEET = "Param";              // Feuille contenant les motifs des trains abrégeables à 4 chiffres
    private static readonly TRAINS_4DIGIT_TABLE = "LigneC4chiffres";    // Tableau contenant les motifs des trains abrégeables à 4 chiffres

    // Indicateur de chargement
    private static loaded = false;

    // Regex globales
    private static wRegex: RegExp;
    private static abbreviate4Regex: RegExp;

    // Propriétés de l'objet TrainNumber
    public readonly value: string;                  // Numéro de train avec parité
                                                    //  (la double parité est marquée par ######/#)

    /**
     * Constructeur de la classe TrainNumber.
     * Garde uniquement les chiffres et lettres mises en majuscules.
     * @param {string | number} value Numéro de train (nombre ou chaine de caractères).
     * @param {boolean} doubleParity Si vrai, force la double parité. Si faux (par défaut),
     *  la double parité est détectée avec la présence de "/" dans le numéro de train.
     */
    constructor(value: string | number, doubleParity: boolean = false) {

        const raw = value.toString();
        const applyDoubleParity = doubleParity || raw.includes("/");

        const normalized = TrainNumber.normalize(raw);

        if (!TrainNumber.isValidTrainNumber(normalized)) {
            Log.warn(`Numéro de train invalide : ${value}`);
            this.value = "";
            return;
        }

        // Force l'éventuelle double parité
        this.value = applyDoubleParity
            ? TrainNumber.applyParity(normalized, Parity.DOUBLE)
            : normalized;
    }

    /**
     * Normalise un numéro de train en supprimant les caractères non-alphanumériques
     * et en ne gardant que la partie précédent un "/".
     * @param {string} value Numéro de train à normaliser.
     * @returns {string} Numéro de train normalisé.
     */
    private static normalize(value: string): string {
        return value
            .split("/")[0]
            .toUpperCase()
            .replace(/[^A-Z0-9]/g, '');
    }

    /**
     * Vérifie si un numéro de train est valide.
     * @param {string} value Numéro de train à vérifier.
     * @returns {boolean} Vrai si le numéro de train est valide, faux sinon.
     */
    private static isValidTrainNumber(value: string): boolean {
        if (!value) return false;
        const lastChar = value.slice(-1);
        return /^[0-9]$/.test(lastChar);
    }

    /**
     * Abrège le numéro de train à 4 chiffres si possible.
     * La méthode teste si le numéro de train correspond à une expression régulière
     * définie dans la classe TrainNumber.
     * Si le numéro de train correspond, il est abrégé en supprimant les 2 premiers
     * chiffres.
     * Si le numéro de train ne correspond pas, il est renvoyé inchangé.
     * @returns {string} Numéro de train abrégé de 6 à 4 chiffres s'il est abrégeable.
     */
    private abbreviateTo4Digits(): string {

        const abbreviated = TrainNumber.abbreviate4Regex?.test(this.value.split("/")[0])
            ? this.value.substring(2)
            : this.value;

        return abbreviated;
    }

    /**
     * Adapte le numéro du train en fonction de la parité demandée.
     * Applique la parité demandée au numéro du train.
     * Si le numéro du train est pair, il est inchangé si la parité demandée est paire,
     *  et incrémenté de 1 si la parité demandée est impaire.
     * Si le numéro du train est impair, il est décrémenté de 1 si la parité demandée est paire,
     *  et inchangé si la parité demandée est impaire.
     * Si la parité demandée est double, le numéro du train est donné
     *  par sa valeur paire, suivi d'un '/' et du chiffre impair suivant.
     * Si la parité demandée est indéfinie, le numéro du train est inchangé.
     * @param {string} value Numéro du train à adapter.
     * @param {number} parity Parité demandée (paire, impaire, double).
     * @returns {string} Numéro du train adapté.
     */
    private static applyParity(value: string, parity: number): string {

        const base = value.split("/")[0];
        const lastDigit = parseInt(base.slice(-1), 10);
        const rest = base.slice(0, -1);
        const even = lastDigit - (lastDigit % 2);

        switch (parity) {
            case Parity.EVEN:
                return rest + even;
            case Parity.ODD:
                return rest + (even + 1);
            case Parity.DOUBLE:
                return rest + even + "/" + (even + 1);
            default:
                return base;
        }
    }

    /**
     * Adapte le numéro du train en fonction de la parité demandée.
     * Si le numéro du train est pair, il est inchangé si la parité demandée est paire,
     *  et incrémenté de 1 si la parité demandée est impaire.
     * Si le numéro du train est impair, il est décrémenté de 1 si la parité demandée est paire,
     *  et inchangé si la parité demandée est impaire.
     * Si la parité demandée est indéfinie, le numéro du train est inchangé.
     * @param {number} parityValue Parité demandée (paire, impaire, double).
     * @param {boolean} abbreviateTo4Digits Si vrai, le numéro du train est abrégé à 4 chiffres.
     *  Si faux, le numéro du train n'est pas abrégé.
     * @returns {string} Numéro du train adapté
     */
    public adaptWithParity(parityValue: number, abbreviateTo4Digits: boolean = false): string {

        if (!parityValue) return this.value;

        const base = abbreviateTo4Digits? this.abbreviateTo4Digits() : this.value;
        return TrainNumber.applyParity(base, parityValue);
    }

    /**
     * Retourne le numéro du train en fonction des paramètres :
     *  - si abbreviateTo4Digits est vrai, le numéro du train est abrégé à 4 chiffres.
     *  - si withoutDoubleParity est vrai, le numéro du train est renommé sans double parité.
     * @param {boolean} [abbreviateTo4Digits=false] Si vrai, le numéro du train est abrégé
     *  de 6 à 4 chiffres pour les trains commerciaux. Si faux (par défaut), le numéro n'est pas abrégé.
     * @param {boolean} [withoutDoubleParity=false] Si vrai, le numéro est renommé
     *  pour ne pas indiquer le changement de parité. Si faux (par défaut), le numéro de train
     *  en gare origine est renvoyé avec double parité si concerné.
     * @returns {string} Numéro du train.
     */
    public toString(
        abbreviateTo4Digits: boolean = false,
        withoutDoubleParity: boolean = false
    ): string {
        let result = this.value;

        if (abbreviateTo4Digits) {
            result = this.abbreviateTo4Digits();
        }
    
        if (withoutDoubleParity) {
            result = result.split('/')[0];
        }
    
        return result;
    }

    /**
     * Teste si le train est W (vide voyageur).
     * @returns {boolean} Vrai si le train est W, faux sinon.
     */
    public isW(): boolean {
        return TrainNumber.wRegex?.test(this.value) ?? false;
    }
   
    /**
     * Charge les regex globales pour
     *  - les numéros de train W,
     *  - les numéros de train abrégeables à 4 chiffres.
     */
    public static load(erase = false): void {
        if (TrainNumber.loaded && !erase) return;

        TrainNumber.loadWRegex();
        TrainNumber.loadAbbreviateRegex();

        TrainNumber.loaded = true;
    }

    /**
     * Charge les motifs des numéros de train W.
     * Les valeurs de la table sont transformées en regex partielles avec les numéros
     *  remplacés par des chiffres, puis combinées en une regex globale unique.
     */
    private static loadWRegex(): void {
        const data = WorkbookService.getDataFromTable(TrainNumber.W_SHEET, TrainNumber.W_TABLE);
    
        // Transforme chaque motif en regex partielle
        const regexParts: string[] = data
            .slice(1)
            .flat()
            .filter((v: unknown) => typeof v === "string" && v.trim() !== "")
            .map((pattern: string) => {
                return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
            });

        // Crée une regex globale combinée
        TrainNumber.wRegex = new RegExp(regexParts.join('|'));
    }

    /**
     * Charge les motifs des numéros de train abrégeables à 4 chiffres.
     * Les valeurs de la table sont transformées en regex partielles avec les numéros
     *  remplacés par des chiffres, puis combinées en une regex globale unique.
     */
    private static loadAbbreviateRegex(): void {
        const data = WorkbookService.getDataFromTable(TrainNumber.TRAINS_4DIGIT_SHEET,
            TrainNumber.TRAINS_4DIGIT_TABLE);
    
        // Transforme chaque motif en regex partielle
        const regexParts: string[] = data
            .slice(1)
            .flat()
            .filter((v: unknown) => typeof v === "string" && v.trim() !== "")
            .map((pattern: string) => {
                return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
            });

        // Crée une regex globale combinée
        TrainNumber.abbreviate4Regex = new RegExp(regexParts.join('|'));
    }
}

/* 
 * Classe Station définissant une gare
 */
class Station {

    // Propriétés de l'objet Station
    public readonly abbreviation!: string;                  // Abréviation de la gare
    public readonly name: string;                           // Nom de la gare
    public readonly referenceStation: Station | null;       // Gare de rattachement
    public childStations: Station[];                        // Sous-gares
    public readonly turnaround: Parity;                     // Parité d'un rebroussement possible
                                                            //  (la parité est celle du train avant rebroussement)
    public readonly reverseLineDirection: boolean;          // Parité de la ligne inversée sur cette gare

    /**
     * Constructeur d'une gare.
     * @param {string} abbreviation Abréviation de la gare
     * @param {string} name Nom de la gare
     * @param {Station} referenceStation Gare de rattachement
     * @param {Parity} turnaround Parité d'un rebroussement possible
     *  (la parité est celle du train avant rebroussement)
     * @param {boolean} reverseLineDirection Parité de la ligne inversée sur cette gare
     */
    constructor(
        abbreviation: string,
        name: string,
        referenceStation: Station | null,
        turnaround: Parity | string | number,
        reverseLineDirection: boolean,
    ) {
        if (!abbreviation) {
            throw new Error(`Une gare ne peut pas avoir une abréviation vide.`);
        }
        this.abbreviation = abbreviation;
        if (!name) {
            throw new Error(`La gare ${abbreviation} ne peut pas avoir un nom vide.`);
        }
        this.name = name;
        this.referenceStation = referenceStation ?? null;
        this.childStations = [];;
        this.turnaround = Parity.from(turnaround, true);
        this.reverseLineDirection = reverseLineDirection;
    }
}

/**
 * Classe Stations contenant la liste des gares
 */
class Stations {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Gares";                // Feuille contenant la liste des gares
    private static readonly TABLE = "Gares";                // Tableau contenant la liste des gares
    private static readonly HEADERS = [[                    // En-têtes du tableau des gares
        "Abréviation",
        "Nom",
        "Gare de rattachement",
        "Gare de rebroussement",
        "Parité de ligne inversée"
    ]];                                             
    private static readonly COL_ABBR = 0;                   // Colonne de l'abréviation de la gare
    private static readonly COL_NAME = 1;                   // Colonne du nom de la gare
    private static readonly COL_REFERENCE_STATION = 2;      // Colonne de la gare de rattachement
    private static readonly COL_TURNAROUND = 3;             // Colonne indiquant si un rebroussement
                                                            //  est possible (pair ou impair)
    private static readonly COL_REVERSE_LINE_PARITY = 4;    // Colonne indiquant si la parité de la ligne est inversée

    // Map des gares indexées par abréviation
    public static readonly map: Map<string, Station> = new Map();

    /**
     * Accesseurs utilitaires
     */
    // Nombre de gares
    public static get size(): number {
        return Stations.map.size;
    }
    // Vérifie si une gare est présente dans la liste
    public static has(abbreviation: string): boolean {
        return Stations.map.has(abbreviation);
    }
    // Récupère une gare par son abréviation
    public static get(abbreviation: string): Station | undefined {
        return Stations.map.get(abbreviation);
    }
    // Accès à toutes les gares
    public static values(): IterableIterator<Station> {
        return Stations.map.values();
    }
    // Efface toutes les gares
    public static clear(): void {
        Stations.map.clear();
    }

    /**
     * Charge les gares à partir du tableau "Gares" de la feuille "Gares".
     * Les gares sont stockées dans une Map avec comme clés l'abréviation 
     * de la gare et comme valeur l'objet Station.
     * @param {boolean} [erase=false] Si vrai, force le rechargement des gares.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà
        if (Stations.map.size > 0) {
            if (erase) Stations.clear(); // Vide la map sans changer sa référence
            else return;
        }

        // Charge la base de données
        const data = WorkbookService.getDataFromTable(Stations.SHEET, Stations.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Stations.load : aucune donnée trouvée dans la table.`);
            return;
        }

        // Parcourt les lignes (hors en-tête)
        const referenceStationPairs: [string, string][] = [];
        for (const [rowIndex, row] of data.slice(1).entries()) {

            // Vérifie si la ligne est vide
            if (row.length === 0) continue;

            // Calcule le numéro de ligne Excel
            const excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
    
            try {

                // Récupère les champs
                const abbreviation = WorkbookService.getString(row, Stations.COL_ABBR) ?? "";
                const name = WorkbookService.getString(row, Stations.COL_NAME) ?? "";
                const referenceStationName = WorkbookService.getString(row, Stations.COL_REFERENCE_STATION) ?? "";
                const turnaroundLetters = WorkbookService.getString(row, Stations.COL_TURNAROUND) ?? "";
                const reverseLineDirection = !!WorkbookService.getBoolean(row, Stations.COL_REVERSE_LINE_PARITY);

                // Instancie les propriétés objets
                const turnaround = Parity.from(turnaroundLetters, true);

                // Instancie l'objet Station
                const station = new Station(
                    abbreviation,
                    name,
                    null,
                    turnaround,
                    reverseLineDirection
                );

                // Ajoute l'objet Station dans la map, indexé par son nom et son abréviation
                if (Stations.map.has(abbreviation)) {
                    throw new Error(`La gare ${abbreviation} est déjà présente`
                        + ` dans la base de données.`);
                } 
                Stations.map.set(abbreviation, station);
                if (Stations.map.has(name)) {
                    throw new Error(`La gare ${name} est déjà présente`
                        + ` dans la base de données.`);
                }
                Stations.map.set(name, station);

                // Mémorise les paires gare/gare de rattachement
                if (referenceStationName) {
                    referenceStationPairs.push([abbreviation, referenceStationName]);
                }

            } catch (e) {
                Log.warn(`Stations.load (ligne ${excelRow}) : ${e}`);
                continue;
            } 
        }

        // Parcourt les paires pour ajouter les objets des gares de réference à chaque gare
        for (const [abbr, refName] of referenceStationPairs) {
            const station = Stations.map.get(abbr);
            const referenceStation = Stations.map.get(refName);

            if (station && referenceStation) {
                station.referenceStation = referenceStation;
                referenceStation.childStations.push(station);
            }
        }
    }

    /**
     * Affiche les stations de la map dans un tableau.
     * @param {string} [sheetName=Stations.SHEET] Nom de la feuille de calcul.
     * @param {string} [tableName=Stations.TABLE] Nom du tableau.
     * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = Stations.SHEET,
        tableName: string = Stations.TABLE,
        startCell: string = "A1"
    ): void {

        // Convertit la map en un tableau de données
        const data: (string | number)[][] = Array
            .from(Stations.map.values())
            .map(station => [
                station.abbreviation,
                station.name,
                station.referenceStation?.abbreviation ?? "",
                station.turnaround.printLetter(),
                station.reverseLineDirection ? 1 : 0
            ]);

        // Imprime le tableau
        WorkbookService.printTable(Stations.HEADERS, data, sheetName, tableName, startCell);
    }
}

/**
 * Classe StationWithParity immutable définissant une gare d'arrêt ou de passage d'un train
 *  et sa parité associée.
 */
class StationWithParity {

    // Propriétés de l'objet StationWithParity
    private readonly _station: Station;
    private readonly _parity: Parity;

    /**
     * Constructeur de la classe StationWithParity.
     * @param {Station | string} stationValue Gare (Station) ou nom de gare avec ou sans suffixe _PARITE
     * @param {Parity | string | number} [parity] Parité associée à la gare d'arrêt ou de passage.
     * Si une parité explicite est fournie, on vérifie la cohérence avec la parité de la gare.
     * Si une erreur est détectée, une exception est levée.
     */
    private constructor(
        stationValue: Station | string,
        parity?: Parity | string | number
    ) {
        const { station, parity: parsedParity } = StationWithParity.parseStationAndParity(stationValue);
        const parityObj = Parity.from(parity, false);

        if (parsedParity.isOpposedTo(parityObj)) {
            throw new Error(`Conflit de parité pour ${stationValue}.`);
        }

        this._station = station;
        this._parity = parsedParity.is(Parity.UNDEFINED) ? parityObj : parsedParity;
    }

    /**
     * Retourne une instance de StationWithParity à partir d'une valeur qui peut être :
     *  - une instance de StationWithParity (retourne la même instance)
     *  - une instance de Station (retourne une nouvelle instance de StationWithParity
     *  avec la parité undefined)
     *  - un nom de gare (retourne une nouvelle instance de StationWithParity avec la parité undefined)
     *  - null ou undefined (lève une erreur)
     * @param {StationWithParity | Station | string | null | undefined} value Valeur à analyser
     *  pour la gare.
     * @param {Parity | string | number} [parity] Parité optionnelle. Si fournie et contradictoire
     *  avec la chaîne, une exception est levée.
     * @returns {StationWithParity} Instance de StationWithParity correspondante.
     * @throws {Error} Si la valeur est null ou undefined.
     */
    public static from(
        value: StationWithParity | Station | string | null | undefined,
        parity?: Parity | string | number
    ): StationWithParity {

        if (value == null) {
            throw new Error(`La gare n'est pas renseignée.`);
        }

        return value instanceof StationWithParity
            ? value
            : new StationWithParity(value, parity);
    }

    /**
     * Renvoie l'identifiant unique de l'objet StationWithParity, qui est
     *  sa représentation sous forme de chaîne.
     * @returns {string} L'identifiant unique de l'objet StationWithParity.
     */
    get key(): string {
        return this.toString();
    }

    /**
     * Renvoie la gare associée à l'objet StationWithParity.
     * @returns {Station} La gare associée.
     */
    get station(): Station {
        return this._station;
    }

    /**
     * Renvoie la parité de l'objet StationWithParity.
     * @returns {Parity} La parité de l'objet StationWithParity.
     */
    get parity(): Parity {
        return this._parity;
    }

    /**
     * Analyse une valeur qui peut être une gare (Station) ou un nom de gare avec ou sans suffixe _PARITE
     * et renvoie un objet avec la gare correspondante et la parité associée.
     * Si la valeur est une instance de Station, la parité est undefined.
     * Si la valeur est un nom de gare, la parité est undefined
     *  si le nom ne contient pas de suffixe _PARITE.
     * Si la valeur contient une erreur (par exemple, si le nom de gare n'existe pas),
     *  une exception est levée.
     * @param {Station | string} value Valeur à analyser pour la gare.
     * @returns {{ station: Station; parity: Parity }} Objet avec la gare et la parité associée.
     */
    private static parseStationAndParity(
        value: Station | string
    ): { station: Station; parity: Parity } {

        if (!value) {
            throw new Error(`La gare ne peut pas être vide.`);
        }

        if (value instanceof Station) {
            return {
                station: value,
                parity: Parity.from(Parity.UNDEFINED, false)
            };
        }

        const [stationName, parityPart] = value.split("_");

        const station = Stations.get(stationName);
        if (!station) {
            throw new Error(`La gare ${stationName} n'existe pas.`);
        }

        return {
            station,
            parity: Parity.from(parityPart ?? Parity.UNDEFINED, false)
        };
    }

    /**
     * Retourne la gare après rebroussement si celui-ci est possible.
     * Si la parité de l'arrivée est donnée, renvoie la gare avec parité opposée.
     * Si elle n'est pas donnée, renvoie la même gare (sans parité).
     * Si le rebroussement n'est pas possible, renvoie undefined.
     * @returns {StationWithParity | undefined} Gare après rebroussement si possible, sinon undefined.
     */
    public stationAfterTurnaround(): StationWithParity | undefined {
    
        // La gare de rebroussement est donnée par l'inversion de parité si définie,
        //  ou sans changement sinon
        const reversedParity = Parity.from(this._parity.value).invert();
        const stationAfterTurnaround = StationWithParity.from(
            this._station,
            reversedParity
        )

        // Si parité définie, rebroussement possible si parité incluse dans la propriété Station.turnaround
        // Si parité non définie, rebroussement considéré comme possible si autorisé depuis au moins un sens
        const canTurnAround = this._parity.isDefined()
            ? this._station.turnaround.includes(this._parity)
            : this._station.turnaround.isDefined();

        return canTurnAround ? stationAfterTurnaround : undefined;
    }

    /**
     * Vérifie si l'objet StationWithParity a la même gare que l'autre.
     * @param other Autre objet StationWithParity à comparer.
     * @returns {boolean} Vrai si les deux objets ont la même gare, faux sinon.
     */
    public hasSameStationTo(other: StationWithParity | null | undefined): boolean {
        return (
            !! other &&
            this._station === other.station
        );
    }

    /**
     * Vérifie si l'objet StationWithParity est identique à l'autre.
     * @param other Autre objet StationWithParity à comparer.
     * @returns {boolean} Vrai si les deux objets sont identiques, faux sinon.
     */
    public equalsTo(other: StationWithParity | null | undefined): boolean {
        return (
            !! other &&
            this._station === other.station &&
            this.parity.equalsTo(other.parity)
        );
    }

    /**
     * Renvoie une chaîne représentant l'objet StationWithParity sous la forme
     *  GARE_PARITE, où GARE est le nom de la gare sans suffixe _PARITE et
     *  PARITE est la parité sous forme de chiffre.
     * @returns {string} Chaîne représentant l'objet StationWithParity.
     */
    public toString(): string {
        return this.parity.isDefined() 
            ? `${this._station.abbreviation}_${this.parity.printDigit()}` 
            : `${this._station.abbreviation}`;
    }
}

/**
 * Classe Connection définissant une connexion orientée entre deux gares
 */
class Connection {

    // Constantes des valeurs par défaut
    static readonly DEFAULT_CONNECTION_TIME= 1; // Durée de connection par défaut en jours
                                                //  (si 0 ou non renseignée)
                                                //  La durée est très importante pour privilégier
                                                //  les connexions avec une durée de connexion
                                                //  déjà évaluée à partir de parcours réels

    // Propriétés de l'objet Connexion
    public readonly from: StationWithParity;    // Gare de départ
    public readonly to: StationWithParity;      // Gare d'arrivée
    private _time: DateTime;                    // Temps de trajet
    public readonly withTurnaround: boolean;    // Connexion impliquant un rebroussement
    public readonly withMovement: boolean;      // Connexion sous régime de l'évolution
    public readonly changeParity: boolean;      // Connexion avec changement de parité

    /**
     * Constructeur d'une connexion.
     * @param {StationWithParity | string} from Gare de départ
     * @param {StationWithParity | string} to Gare d'arrivée
     * @param {DateTime | number | string} [time] Temps de trajet (si 0 ou non renseigné : durée par défaut)
     * @param {boolean} [withTurnaround=false] Indique si la connexion implique un rebroussement
     * @param {boolean} [withMovement=false] Indique si la connexion est sous régime de l'évolution
     * @param {boolean} [changeParity=false] Indique si la connexion implique un changement de parité
     */
    constructor(
        from: StationWithParity | string,
        to: StationWithParity | string,
        time: DateTime | number | string = Connection.DEFAULT_CONNECTION_TIME,
        withTurnaround: boolean = false,
        withMovement: boolean = false,
        changeParity: boolean = false
    ) {
        this.from = StationWithParity.from(from);
        this.to = StationWithParity.from(to);
        if (this.from.equalsTo(this.to)) {
            throw new Error(
                `Une connexion ne peut pas relier ${this.from} à elle-même sans changement de gare ou de parité.`
            );
        }
        this.time = time;
        this.withTurnaround = withTurnaround;
        this.withMovement = withMovement;
        this.changeParity = changeParity;
    }

    /**
     * Renvoie le temps de trajet de la connexion.
     * @returns {DateTime} Le temps de trajet de la connexion.
     */
    get time(): DateTime {
        return this._time;
    }

    /**
     * Modifie le temps de trajet de la connexion.
     * @param {DateTime | number | string} value Nouveau temps de trajet de la connexion.
     * @throws {Error} Si le temps de trajet est inférieur ou égal à 0 ou n'est pas relatif.
     */
    set time(value: DateTime | number | string) {
        const timeObj = DateTime.from(value, true);
        if (timeObj.excelValue <= 0) {
            throw new Error(
                `Le temps de trajet de la connexion ${this.from} -> ${this.to}`
                + ` est inférieur ou égal à 0.`
            );
        }
        if (!timeObj.isRelative) {
            throw new Error(
                `Le temps de trajet de la connexion ${this.from} -> ${this.to}`
                + ` n'est pas relatif.`
            );
        }
        this._time = timeObj;
    }
}

/**
 * Classe Connections contenant la liste des connexions
 */
class Connections {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";            // Feuille contenant la liste des connexions
    private static readonly TABLE = "Connexions";       // Tableau contenant la liste des connexions
    private static readonly HEADERS = [[                // En-têtes du tableau des connexions
        "De",
        "Vers",
        "Durée",
        "Rebroussement",
        "Evolution",
        "Changement de parité"
    ]];                                         
    private static readonly COL_FROM = 0;               // Colonne de la gare de départ
    private static readonly COL_TO = 1;                 // Colonne de la gare d'arrivée
    private static readonly COL_TIME = 2;               // Colonne de la durée de parcours (en minutes)
    private static readonly COL_TURNAROUND = 3;         // Colonne indiquant si la connexion
                                                        //  implique un rebroussement
    private static readonly COL_MOVEMENT = 4;           // Colonne indiquant si la connexion
                                                        //  est sous régime de l'évolution
    private static readonly COL_CHANGE_PARITY = 5;      // Colonne indiquant si la connexion
                                                        //  implique un changement de parité

    // Map des gares indexées par abréviation
    public static readonly map: Map<string, Map<string, Connection>> = new Map();

    /**
     * Accesseurs utilitaires
     */
    // Nombre de connexions
    public static get size(): number {
        let count = 0;
        for (const m of Connections.map.values()) count += m.size;
        return count;
    }
    // Vérifie si une connexion est présente dans la liste
    public static has(from: string, to?: string): boolean {
        if (!Connections.map.has(from)) return false;
        return to ? Connections.map.get(from)!.has(to) : true;
    }
    // Récupère une connexion par son oririne et sa destination
    public static get(from: string, to: string): Connection | undefined {
        return Connections.map.get(from)?.get(to);
    }
    // Efface toutes les connexions
    public static clear(): void {
        Connections.map.clear();
    }

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
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà
        if (Connections.map.size > 0) {
            if (erase) Connections.clear(); // Vide la map sans changer sa référence
            else return;
        }

        // Charge les gares si elles n'ont pas encore été chargées
        Stations.load(); 

        // Charge la base de données
        const data = WorkbookService.getDataFromTable(Connections.SHEET, Connections.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Connections.load : aucune donnée trouvée dans la table.`);
            return;
        }

        // Parcourt les lignes (hors en-tête)
        for (const [rowIndex, row] of data.slice(1).entries()) {

            // Vérifie si la ligne est vide
            if (row.length === 0) continue;

            // Calcule le numéro de ligne Excel
            const excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

            try {

                // Récupère les champs
                const from = WorkbookService.getString(row, Connections.COL_FROM);
                const to = WorkbookService.getString(row, Connections.COL_TO);
                if (!from || !to) continue;
                const timeInMinutes = WorkbookService.getNumber(row, Connections.COL_TIME);
                const withTurnaround = !!WorkbookService.getBoolean(row, Connections.COL_TURNAROUND);
                const withMovement = !!WorkbookService.getBoolean(row, Connections.COL_MOVEMENT);
                const changeParity = !!WorkbookService.getBoolean(row, Connections.COL_CHANGE_PARITY);

                // Instancie les propriétés objets (si 0 ou non renseignée : valeur par défaut)
                const excelTime = timeInMinutes
                    ? timeInMinutes / 24 / 60
                    : Connection.DEFAULT_CONNECTION_TIME;
                const timeObj = DateTime.from(excelTime, true);
                
                // Instancie l'objet Connection
                const connection = new Connection(
                    from,
                    to,
                    timeObj,
                    withTurnaround,
                    withMovement,
                    changeParity
                );

                // Ajoute l'objet Connection dans la map
                if (!Connections.map.has(from)) {
                    Connections.map.set(from, new Map());
                }
                const targets = Connections.map.get(from)!;
                if (targets.has(to)) {
                    throw new Error(`La connection de ${from} vers ${to}`
                        + ` est déjà présente dans la base de données.`);           
                }
                targets.set(to, connection);

            } catch (e) {
                Log.warn(`Connections.load (ligne ${excelRow}) : ${e}`);
                continue;
            }
        }
    }

    /**
     * Affiche les connexions entre les gares dans un tableau.
     * @param {string} [sheetName=Connections.SHEET] Nom de la feuille de calcul.
     * @param {string} [tableName=Connections.TABLE] Nom du tableau.
     * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = Connections.SHEET,
        tableName: string = Connections.TABLE,
        startCell: string = "A1"
    ): void {

        // Convertit la map en un tableau de données
        const data: (string | number)[][] = [];

        for (const [from, targets] of Connections.map) {
            for (const [to, connection] of targets) {
                data.push([
                    connection.from.key,
                    connection.to.key,
                    connection.time.excelValue * 24 * 60,
                    connection.withTurnaround ? 1 : 0,
                    connection.withMovement ? 1 : 0,
                    connection.changeParity ? 1 : 0
                ]);
            }
        }

        // Imprime le tableau
        const table = WorkbookService.printTable(
            Connections.HEADERS,
            data,
            sheetName,
            tableName,
            startCell
        );

        // Met les durées de parcours au format "hh:mm:ss"
        table.getRange()
            .getColumn(Connections.COL_TIME)
            .setNumberFormat("hh:mm:ss");
    }

    /**
     * Sauvegarde les temps de connexions entre les gares dans la map.
     * Les données sont calculées en fonction des horaires de départ et d'arrivée des trains.
     * @param {string} [trainNumbers=""] Trains à traiter, séparés par des ; . Si vide,
     *  traite tous les trains.
     */
    // public static saveConnectionsTimes(trainNumbers: string = "") {
    //     if (trainNumbers === "") {
    //         trainNumbers =
    //             Array.from(Paths.map.keys()).filter(key => key === Paths.map.get(key)!.key)
    //                 .join(";");
    //     }
    //     trainNumbers.split(";").forEach((trainNumber) => {
    //         const path = Paths.map.get(trainNumber);
    //         path?.stops.forEach((stop) => {
    //             if (stop.nextStop && Connections.has(stop.key)
    //                 && Connections.has(stop.nextStop.key)) {
    //                 const connection = Connections.get(stop.key, stop.nextStop.key);
    //                 if (connection && stop.nextStop.arrivalTime !== 0
    //                     && stop.departureTime !== 0) {
    //                     connection.time = stop.nextStop.arrivalTime - stop.departureTime;
    //                 }
    //             }
    //         });
    //     });
    // }
}

/*
 * Classe Stop définissant l'arrêt ou le passage d'un train dans une gare
 */
class Stop {

    // Propriétés de l'objet Stop
    public readonly station: StationWithParity;     // Gare de l'arrêt
    private _withTurnaround: boolean = false;        // Arrêt avec rebroussement
    private _arrivalTime?: DateTime;                 // Temps / Heure d'arrivée de l'arrêt
    private _departureTime?: DateTime;               // Temps / Heure de départ de l'arrêt
    private _passageTime?: DateTime;                 // Temps / Heure de passage à l'arrêt (sans arrêt)
    private _tracks: string[];                       // Voies de l'arrêt
    
    constructor(
        station : StationWithParity | Station | string,
        stationAfterTurnaround?: StationWithParity | string,
        arrivalTime?: DateTime | number | string,
        departureTime?: DateTime | number | string,
        passageTime?: DateTime | number | string,
        areRelativeTimes: boolean = false,
        tracks: string[] | string = [],
    ) {
        // Détermination de la gare d'arrêt
        this.station = StationWithParity.from(station);

        // Détermination du rebroussement
        this._withTurnaround = this.canTurnaroundTo(stationAfterTurnaround);

        // Détermination des horaires de l'arrêt
        this.setTimes(arrivalTime, departureTime, passageTime, areRelativeTimes);

        // Détermination des voies de l'arrêt
        this._tracks = tracks instanceof Array ? tracks : Stop.getTracksFromString(tracks);
    }

    /**
     * Renvoie une clé unique pour l'arrêt, composée du nom de la gare et de la parité
     *  (si connue).
     * @returns {string} Clé unique
     */
    get key(): string {
        return this.station.key;
    }

    /**
     * Renvoie le nom de la gare associée à cet arrêt.
     * @returns {string} Nom de la gare.
     */
    get stationName(): string {
        return this.station!.station.name;
    }

    /**
     * Renvoie l'abréviation de la gare associée à cet arrêt.
     * @returns {string} Abréviation de la gare.
     */
    get stationAbbreviation(): string {
        return this.station.station.abbreviation;
    }

    /**
     * Renvoie true si l'arrêt à un rebroussement possible, false sinon.
     * @returns {boolean} Vrai si l'arrêt a un rebroussement possible, faux sinon.
     */
    get withTurnaround(): boolean {
        return this._withTurnaround;
    }

    /**
     * Retourne l'heure d'arrivée à l'arrêt, si connue.
     * @returns {DateTime | undefined} Heure d'arrivée à l'arrêt, ou undefined si non connue.
     */
    get arrivalTime(): DateTime | undefined {
        return this._arrivalTime;
    }

    /**
     * Retourne l'heure de départ à l'arrêt, si connue.
     * @returns {DateTime | undefined} Heure de départ à l'arrêt, ou undefined si non connue.
     */
    get departureTime(): DateTime | undefined {
        return this._departureTime;
    }

    /**
     * Retourne l'heure de passage à l'arrêt, si connue.
     * @returns {DateTime | undefined} Heure de passage à l'arrêt, ou undefined si non connue.
     */
    get passageTime(): DateTime | undefined {
        return this._passageTime;
    }

    /**
     * Retourne le tableau des voies de l'arrêt.
     * @returns {string[]} Tableau des voies de l'arrêt.
     */
    get tracks(): readonly string[] {
        return this._tracks;
    }

    /**
     * Retourne la gare associée à l'objet StationWithParity et la parité opposée,
     * si le rebroussement est possible (connection existante).
     * @returns {StationWithParity | undefined} La StationWithParity avec la gare de l'objet StationWithParity
     * et la parité opposée, ou undefined si la parité est indéfinie ou le rebroussement n'est pas possible.
     */
    get stationAfterTurnaround(): StationWithParity | undefined {
        return this.withTurnaround ? this.station.stationAfterTurnaround() : undefined;
    }

    /**
     * Vérifie si un rebroussement est possible avec stationAfterTurnaround comme gare après rebroussement.
     * Un rebroussement est possible si la gare après rebroussement correspond à la gare calculée.
     *  
     * @param {StationWithParity | string} stationAfterTurnaround Gare après rebroussement.
     * @returns {boolean} True si le rebroussement est possible, false sinon.
     */
    private canTurnaroundTo(stationAfterTurnaround : StationWithParity | string | undefined): boolean {

        // Vérifie si la gare après rebroussement demandée est connue
        const stationAfterTurnaroundObj = StationWithParity.from(stationAfterTurnaround);
        if (!stationAfterTurnaroundObj) return false;

        // Calcule la gare théorique après rebroussement si celui-ci est possible
        const calculated = this.station.stationAfterTurnaround();
        if (!calculated) {
            Log.warn(`Un rebroussement n'est pas autorisé à la gare de ${this.station.key}.`
            + ` Il ne sera pas pris en compte.`);
            return false;
        }

        // Comparaison des gares théoriques et demandées
        if (!stationAfterTurnaroundObj.equalsTo(calculated)) {
            Log.warn(`Le rebroussement à la gare de ${this.station.key} ne sera pas pris en compte,`
                + ` car la gare après rebroussement demandée ${stationAfterTurnaroundObj.key} ne correspond pas.`);
            return false
        }

        return true;
    }

    public setTimes(
        arrivalTime?: DateTime | number | string,
        departureTime?: DateTime | number | string,
        passageTime?: DateTime | number | string,
        areRelativeTimes: boolean = false
    ) {
        this._arrivalTime = DateTime.from(arrivalTime, areRelativeTimes);
        this._departureTime = DateTime.from(departureTime, areRelativeTimes);
        this._passageTime = (!arrivalTime && !departureTime)
            ? DateTime.from(passageTime, areRelativeTimes)
            : undefined;
            if (!this._arrivalTime && !this._departureTime && !this._passageTime) {
            throw new Error(`L'arrêt ${this.station.key} n'a pas d'heure d'arrivée,`
                + ` d'heure de départ ou d'heure de passage.`);
        }
        if (this._arrivalTime && this._departureTime) {
            const timeDiff = this._departureTime.compareTo(this._arrivalTime);
            if (timeDiff <= 0) {
                if (timeDiff === 0) {
                    const t = DateTime.from(1);
                    Log.warn(`L'heure d'arrivée`
                        + ` ${this._arrivalTime!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                        + ` à l'arrêt ${this.station} est identique à l'heure de départ.`
                        + ` Cette heure est donc renseignée comme heure de passage.`);
                } else {
                    Log.warn(`L'heure d'arrivée`
                        + ` ${this._arrivalTime!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                        + ` à l'arrêt ${this.station} est supérieure à l'heure de départ`
                        + ` ${this._departureTime!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}.`
                        + ` Seule l'heure d'arrivée sera prise en compte comme heure de passage.`);
                }
                this._passageTime = this._arrivalTime;
                this._arrivalTime = undefined;
                this._departureTime = undefined;
            }
        }
        if (this._withTurnaround && !(this._arrivalTime && this._departureTime)) {
            Log.warn(`Le rebroussement à la gare ${this.station}`
                + ` ne peut pas avoir lieu que si l'arrêt présente`
                + ` une heure de départ ultérieure à l'heure d'arrivée.`
                + ` Le rebroussement ne sera pas pris en compte.`);
            this._withTurnaround = false;
        }
    }

    /**
     * Retourne un tableau de chaînes de caractères correspondant à
     *  une liste de voies séparées par des points-virgules.
     * @param {string} tracksString Chaîne de caractères contenant la liste de voies.
     * @returns {string[]} Tableau de chaînes de caractères correspondant à la liste de voies.
     */
    private static getTracksFromString(tracksString: string): string[] {
        return tracksString
            .split(";")
            .map((t) => t.trim())
            .filter((t) => Boolean(t));
    }

    /**
     * Renvoie la plus petite des heures d'arrivée, de départ ou de passage à l'arrêt.
     * Si ignoreArrival est vrai, lit plutôt l'heure de départ ou de passage.
     * @param {boolean} [ignoreArrival=false] Si vrai, ignore l'heure d'arrivée
     *  et préfère l'heure de départ ou de passage. Si faux (par défaut),
     *  c'est d'abord l'heure d'arrivée qui est prise en compte.
     * @param {DateTime} [reference] Heure de référence pour les heures relatives.
     * @returns {DateTime | undefined} Heure la plus petite, ou undefined si
     *  l'heure d'arrivée est lue et que noReadingArrivalTime est faux.
     */
    public getTime(ignoreArrival: boolean = false, reference?: DateTime): DateTime | undefined {
        let time = this._arrivalTime;

        if (ignoreArrival || !this._arrivalTime) {
            time = this._departureTime ?? this._passageTime;
        }
        return (time && time!.isRelative && reference) ? time.resolveAgainst(reference) : time;
    }

    /**
     * Convertit les heures d'arrivée, de départ et de passage
     *  en temps relatifs par rapport à une référence.
     * @param {DateTime} reference Référence à utiliser pour convertir les heures.
     */
    public convertToRelativeTime(reference: DateTime): void {
        if (this._arrivalTime) this._arrivalTime = this._arrivalTime.relativeTo(reference);
        if (this._departureTime) this._departureTime = this._departureTime.relativeTo(reference);
        if (this._passageTime) this._passageTime = this._passageTime.relativeTo(reference);
    }
    
    /**
     * Compare cette arrêt avec un autre arrêt,
     *  en vérifiant la gare avec parité, le rebroussement,
     *  les heures d'arrivée, de départ et de passage.
     *  La comparaison ignore les voies.
     * @param {Stop | null | undefined} other Autre arrêt à comparer.
     * @returns {boolean} Vrai si les arrêts sont égaux, faux sinon.
     */
    public equalsTo(other: Stop | null | undefined): boolean {
        return (
            !! other &&
            this.station.equalsTo(other.station) &&
            this._withTurnaround === other.withTurnaround &&
            DateTime.equalsOrUndefined(this._arrivalTime, other.arrivalTime) &&
            DateTime.equalsOrUndefined(this._departureTime, other.departureTime) &&
            DateTime.equalsOrUndefined(this._passageTime, other.passageTime)
        );
    }

    /**
     * Ajoute une voie à l'arrêt si elle n'y est pas déjà.
     * Si la voie n'est pas déjà dans la liste des voies, l'ajoute et trie la liste.
     * @param {string} track Voie à ajouter.
     */
    public addTrack(track: string): void {
        if (!this._tracks.includes(track)) {
            this._tracks.push(track);
            this._tracks.sort();
        }
    }
}

/**
 * Classe Stops contenant la liste des arrêts
 */
class Stops {
        
    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Arrêts";               // Feuille contenant la liste des arrêts
    private static readonly TABLE = "Arrêts";               // Tableau contenant la liste des arrêts
    private static readonly HEADERS = [[                    // En-têtes du tableau des arrêts
        "Parcours",
        "Gare",
        "Parité",
        "Parité après rebroussement",
        "Arrivée",
        "Départ",
        "Passage",
        "Voie",
        "Gare suivante"
    ]];                                            
    private static readonly COL_PATH_KEY = 0;                   // Colonne du numéro de train
    private static readonly COL_STATION = 1;                    // Colonne de la gare avec parité
    private static readonly COL_STATION_AFTER_TURNAROUND = 2;   // Colonne de la gare après rebroussement
    private static readonly COL_ARRIVAL_TIME = 3;               // Colonne de l'heure d'arrivée
    private static readonly COL_DEPARTURE_TIME = 4;             // Colonne de l'heure de départ
    private static readonly COL_PASSAGE_TIME = 5;               // Colonne de l'heure de passage
    private static readonly COL_TRACK = 6;                      // Colonne de la voie
    private static readonly COL_NEXT_STATION = 7;               // Colonne de la gare suivante

    // Constantes de lecture du tableau d'importation
    private static readonly IMPORT_SHEET = "Import arrêts";     // Feuille d'import des arrêts
    private static readonly IMPORT_TABLE = "Import_arrêts";     // Tableau d'import des arrêts
    private static readonly IMPORT_HEADERS = [[                 // En-têtes du tableau d'import des arrêts
        "N° origine",
        "Date",
        "Service",
        "Jours de circulation",
        "Gare",
        "Parité",
        "Arrivée",
        "Départ",
        "Passage",
        "Voie"
    ]];                                            
    private static readonly COL_IMPORT_TRAIN_NUMBER = 0;        // Colonne du numéro de train
    private static readonly COL_IMPORT_DATE = 1;                // Colonne de la date
    private static readonly COL_IMPORT_SERVICE = 2;             // Colonne du service
    private static readonly COL_IMPORT_DAYS = 3;                // Colonne des jours de circulation
    private static readonly COL_IMPORT_STATION = 4;             // Colonne de la gare
    private static readonly COL_IMPORT_DEPARTURE_TIME = 5;      // Colonne de l'heure de départ
    private static readonly COL_IMPORT_PASSAGE_TIME = 6;        // Colonne de l'heure de passage
    private static readonly COL_IMPORT_TRACK = 7;               // Colonne de la voie
    private static readonly COL_IMPORT_NEXT_STATION = 8;        // Colonne de la gare suivante

    /**
     * Charge les arrêts à partir du tableau "Arrêts" de la feuille "Arrêts".
     * Les gares sont stockées dans une Map avec comme clés l'abréviation 
     * Les arrêts sont stockés dans la propriété "stops" des trains et parcours correspondants.
     * Si un train n'existe pas, un message d'erreur est affiché.
     */
    public static load(): void {

        // Charge la base de données
        const data = WorkbookService.getDataFromTable(Stops.SHEET, Stops.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Stops.load : aucune donnée trouvée dans la table.`);
            return;
        }
    
        // Parcourt les lignes (hors en-tête)
        for (const [rowIndex, row] of data.slice(1).entries()) {

            // Vérifie si la ligne est vide
            if (row.length === 0) continue;

            // Calcule le numéro de ligne Excel
            const excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
            
            try {

                // Récupère les champs
                const pathKey = WorkbookService.getString(row, Stops.COL_PATH_KEY);
                if (!pathKey) throw new Error(`pathKey manquant.`);
                const station = WorkbookService.getString(row, Stops.COL_STATION) || "";
                const stationAfterTurnaround =
                    WorkbookService.getString(row, Stops.COL_STATION_AFTER_TURNAROUND);
                const arrivalTime =
                    WorkbookService.getNumber(row, Stops.COL_ARRIVAL_TIME);
                const departureTime =
                    WorkbookService.getNumber(row, Stops.COL_DEPARTURE_TIME);
                const passageTime =
                    WorkbookService.getNumber(row, Stops.COL_PASSAGE_TIME);
                const tracks = WorkbookService.getString(row, Stops.COL_TRACK);
                const nextStation =
                    WorkbookService.getString(row, Stops.COL_NEXT_STATION);

                // Instancie l'objet Stop
                const stop = new Stop(
                    station,
                    stationAfterTurnaround,
                    arrivalTime,
                    departureTime,
                    passageTime,
                    true,
                    tracks,
                    nextStation
                );

            } catch (e) {
                Log.warn(`Stops.load (ligne ${excelRow}) : ${e}`);
                continue;
            }
    
            // const path = Paths.map.get(pathKey);
            // if (!path) {
            //     Log.warn(`Stops.load (ligne ${excelRow}) : parcours "${pathKey}" inexistant.`);
            //     continue;
            // }
    
            // // Ajout de l'arrêt au parcours
            // path.addStop(stop);
        }
    
        // Vérification métier finale
        // for (const train of Paths.map.values()) {
        //     train.checkStops();
        // }
    }
    
    /**
     * Affiche les arrêts des trains dans un tableau.
     * Les données sont celles stockées dans les objets Path et Stop de l'objet Paths.map.
     * @param {string} [sheetName=Stops.SHEET] Nom de la feuille de calcul.
     * @param {string} [tableName=Stops.TABLE] Nom du tableau.
     * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = Stops.SHEET,
        tableName: string = Stops.TABLE,
        startCell: string = "A1"
    ): void {
    
        // Filtre l'objet Paths.map en ne prenant qu'une seule fois les trains
        // ayant la même clé
        const uniquePaths: Path[] = Array.from(Paths.map.values())
            .filter((train, index, self) => self.findIndex(t => t.key === train.key) === index);
    
        // Crée le tableau final avec les données de chaque arrêt pour chaque train
        const data: (string | number)[][] = [];
    
        for (const path of Paths.map.values()) {
            for (const [stationName, stop] of path.stops.entries()) {
                data.push([
                    path.key,
                    stop.station.key,
                    stop.stationAfterTurnaround ? stop.stationAfterTurnaround.key : "",
                    stop.arrivalTime ? stop.arrivalTime.excelValue : "",
                    stop.departureTime ? stop.departureTime.excelValue : "",
                    stop.passageTime ? stop.passageTime.excelValue : "",
                    stop.tracks.join(";"),
                    stop.nextStation ? stop.nextStation.key : ""
                ]);
            }
        }
    
        // Imprime le tableau
        const table = WorkbookService.printTable(
            Stops.HEADERS,
            data,
            sheetName,
            tableName,
            startCell
        );
    
        // Met les horaires au format "hh:mm:ss"
        const timeColumns = [
            Stops.COL_ARRIVAL_TIME,
            Stops.COL_DEPARTURE_TIME,
            Stops.COL_PASSAGE_TIME
        ];
        for (const col of timeColumns) {
            table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
        }
    }

    public static import(): void {

        // Charge la base de données
        const data = WorkbookService.getDataFromTable(Stops.IMPORT_SHEET, Stops.IMPORT_TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Stops.load : aucune donnée trouvée dans la table.`);
            return;
        }
    
        // Parcourt les lignes (hors en-tête)
        for (const [rowIndex, row] of data.slice(1).entries()) {

            // Vérifie si la ligne est vide
            if (row.length === 0) continue;

            // Calcule le numéro de ligne Excel
            const excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
            
            try {

                // Récupère les champs
                const trainNumber = WorkbookService.getString(row, Stops.COL_IMPORT_TRAIN_NUMBER);
                const date = WorkbookService.getNumber(row, Stops.COL_IMPORT_DATE);
                const pathKey = WorkbookService.getString(row, Stops.COL_IMPORT_PATH_KEY);
                if (!pathKey) throw new Error(`pathKey manquant.`);
                const station = WorkbookService.getString(row, Stops.COL_IMPORT_STATION) || "";
                const stationAfterTurnaround =
                    WorkbookService.getString(row, Stops.COL_IMPORT_STATION_AFTER_TURNAROUND);
                const arrivalTime =
                    WorkbookService.getNumber(row, Stops.COL_IMPORT_ARRIVAL_TIME);
                const departureTime =
                    WorkbookService.getNumber(row, Stops.COL_IMPORT_DEPARTURE_TIME);
                const passageTime =
                    WorkbookService.getNumber(row, Stops.COL_IMPORT_PASSAGE_TIME);
                const tracks = WorkbookService.getString(row, Stops.COL_IMPORT_TRACK);
                const nextStation =
                    WorkbookService.getString(row, Stops.COL_IMPORT_NEXT_STATION);

                // Instancie l'objet Stop
                const stop = new Stop(
                    station,
                    stationAfterTurnaround,
                    arrivalTime,
                    departureTime,
                    passageTime,
                    true,
                    tracks,
                    nextStation
                );

            } catch (e) {
                Log.warn(`Stops.load (ligne ${excelRow}) : ${e}`);
                continue;
            }
    
            // const path = Paths.map.get(pathKey);
            // if (!path) {
            //     Log.warn(`Stops.load (ligne ${excelRow}) : parcours "${pathKey}" inexistant.`);
            //     continue;
            // }
    
            // // Ajout de l'arrêt au parcours
            // path.addStop(stop);
        }
    
        // Vérification métier finale
        // for (const train of Paths.map.values()) {
        //     train.checkStops();
        // }
    }
}

/**
 * Classe Path définissant le parcours d'un train, avec ses gares et temps de passage
 *  par rapport à la gare origine
 */
class Path {

    // Résultats de la vérification du parcours
    public static readonly  UNCHECKED = 0;
    public static readonly  ONLY_FROM_AND_TO_STOPS = 1;
    public static readonly  WITH_VIA_STOPS = 2;
    public static readonly  FIND_PATH_OK = 3;
    public static readonly  ERROR_WITH_STOPS = -1;
    
    // Propriétés de l'objet Path
    public key: string;                             // Clé du parcours
    public parity: Parity;                          // Parité du parcours
                                                    //  (synthèse des parités pour chaque gare)
    public lineDirection: Parity;                   // Direction du parcours sur la ligne
                                                    //  (donnée par une parité globale)
    public missionCode: string;                     // Code de mission des trains du parcours
    public name: string;                            // Nom du parcours (facultatif)
    public viaStations: string[];                   // Gares définissant le parcours
                                                    //  (gares précédées de @
                                                    //  si l'ordre de passage n'est pas imposé)
    public stops: Stop[] = [];                      // Gares d'arrêt ou gares de passage du parcours
    private _stopIndex = new Map<string, Stop | null>();   // Dictionnaire des gares d'arrêt du parcours
                                                    //  - gares référencées par leur clé
                                                    //  et l'abbréviation de gare
                                                    //  - null pour l'abréviation de gare
                                                    //  si la gare est desservie dans les 2 sens
    public stopsChecked?: number = 0;               // Résultat de la vérification du parcours
                                                    //  (0 si non vérifié)

    constructor(
        key: string = "",
        parityValue: number = Parity.UNDEFINED,
        lineDirection: number = Parity.UNDEFINED,
        missionCode: string = "",
        name: string = "",
        departureStation: string = "",
        arrivalStation: string = "",
        viaStations: string = ""
    ) {
        this.key = key;
        this.parity = Parity.from(parityValue, true);
        this.lineDirection = Parity.from(lineDirection, true);
        this.missionCode = missionCode;
        this.name = name;
        if (departureStation) this.stops.push(new Stop(key, departureStation));
        if (arrivalStation) this.stops.push(new Stop(key, arrivalStation));
        this.viaStations = viaStations ? viaStations.split(';') : [];
    }

    /**
     * Ajoute un arrêt au parcours.
     * Si les trains du parcours sont déjà passés par l'arrêt et que erase est faux,
     *  lance une erreur.
     * @param {Stop} stop Arrêt à ajouter.
     * @param {boolean} [erase=false] Si vrai, remplace l'arrêt s'il existe déjà. Si faux
     *  (par défaut), le nouvel arrêt n'est pas pris en compte.
     * @returns {Stop | null} L'arrêt ajouté, ou null si une erreur a été levée.
     * @throws {Error} Si les trains du parcours sont déjà passé par l'arrêt
     *  et que erase est faux.
     */
    public addStop(stop: Stop, erase: boolean = false): void {

        const hasParityDefined = stop.station.parity.isDefined();
        const stationAfterTurnaround = stop.stationAfterTurnaround;

        // Le parcours a été calculé => contient des arrêts avec parité
        if (this.stopsChecked === Path.FIND_PATH_OK) {
            if (!hasParityDefined) {
                throw new Error(`Le parcours calculé ${this.key} ne doit comporter`
                    + `que des arrêts avec parité définie.`);
            }
            if (this._stopIndex.has(stop.key)) {
                if (!erase) {
                    throw new Error(`L'arrêt "${stop.key}" est déjà associé aux trains`
                        + ` du parcours ${this.key}. Un même train ne peut pas revenir`
                        + ` dans la même gare et avec le même sens.`);
                }
                this.stops.splice(this.stops.indexOf(this._stopIndex.get(stop.key)!), 1);
            }

        // Le parcours n'a pas été calculé => ne contient pas d'arrêts avec parité
        } else {
            if (hasParityDefined) {
                Log.warn(`Le parcours ${this.key} n'a pas été calculé. Il ne peut donc pas `
                    + ` contenir d'arrêts avec parité. L'arrêt ${stop.key} sera donc pris en compte`
                    + ` sans parité.`);
                stop.station = StationWithParity.from(stop.station, Parity.UNDEFINED); 
            }
            if (this._stopIndex.has(stop.key)) {
                if (!erase) {
                    throw new Error(`L'arrêt "${stop.key}" est déjà associé aux trains`
                        + ` du parcours ${this.key}. Si le train dessert une gare dans les deux sens,`
                        + ` il est nécessaire de calculer les parités de passage en gare.`
                        + ` L'arrêt ne sera pas pris en compte.`);
                }
                this.stops.splice(this.stops.indexOf(this._stopIndex.get(stop.key)!), 1);
            }
        }

        // Ajout dans le tableau des arrêts
        this.stops.push(stop);
        this.orderStops();
        this._stopIndex.set(stop.key, stop);
        if (stationAfterTurnaround && hasParityDefined) {
            this._stopIndex.set(stationAfterTurnaround.key, stop);
        }
        return;
    }    

    /**
     * Trie les arrêts du parcours par ordre chronologique.
     * Les arrêts sans heure de passage sont placés en fin de liste.
     * @returns {void}
     */
    public orderStops(): void {
        this.stops.sort((a: Stop, b: Stop) => {
            const aTime = a.getTime();
            const bTime = b.getTime();
            return !bTime ? 1 : !aTime ? -1 : aTime.compareTo(bTime);
        });
    }

    /**
     * Retourne l'arrêt du parcours associé à une gare.
     * Si la gare a une parité définie, renvoie l'arrêt correspondant.
     * Sinon, cherche l'arrêt dans le sens pair, puis dans le sens impair.
     * Si les deux arrêts sont trouvés, renvoie le premier arrêt chronologique.
     * Sinon, renvoie l'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     * @param {StationWithParity | Station | string} station - La gare à chercher
     * @returns {Stop | undefined} - L'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé
     */
    public getStop(station: StationWithParity | Station | string): Stop | undefined {

        let stationObj = StationWithParity.from(station);

        // Le parcours a été calculé => contient des arrêts avec parité
        if (this.stopsChecked === Path.FIND_PATH_OK) {
            if (stationObj.parity.isDefined()) {
                return this._stopIndex.get(stationObj.key) ?? undefined;
            }
            const oddStop = this._stopIndex.get(StationWithParity.from(stationObj, Parity.ODD).key);
            const evenStop = this._stopIndex.get(StationWithParity.from(stationObj, Parity.EVEN).key);
            if (oddStop && evenStop) {
                const firstStop = oddStop.getTime()!.compareTo(evenStop.getTime()!) < 0 ? oddStop : evenStop;
                Log.warn(`Le parcours ${this.key} a un arrêt dans chaque sens dans la gare ${stationObj.key}.`
                    + ` C'est le premier arrêt ${firstStop.key} qui est renvoyé.`);
                return firstStop;
            }
            return oddStop ?? evenStop ?? undefined;
        }

        // Le parcours n'a pas été calculé => ne contient pas d'arrêts avec parité
        return this._stopIndex.get(stationObj.key) ?? undefined;
    }

    /**
     * Efface la liste des arrêts du train.
     * Supprime également les valeurs de firstStop et lastStop.
     */
    public eraseStops() {
        this.stops = [];
        this.stopsChecked = 0;
    }

    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée du parcours,
     * puis génère la liste des arrêts calculés.
     * Une fois le trajet calculé, this.stopsChecked a pour valeur 3.
     */
    public findPath(useIntermediateStops: boolean = true) {

        // A reprendre

    }
  
}

class Paths {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Parcours";             // Feuille contenant la liste des parcours  
    private static readonly TABLE = "Parcours";             // Tableau contenant la liste des parcours
    private static readonly HEADERS = [[                    // En-têtes du tableau des parcours
        "Clé",
        "Parité du parcours",
        "Parité de ligne",
        "Code mission",
        "Nom",
        "Gare de départ",
        "Gare d'arrivée",
        "Gares intermédiaires"
    ]];
    private static readonly COL_KEY = 0;                    // Colonne de la clé du parcours
    private static readonly COL_PARITY = 1;                 // Colonne de la parité du parcours
    private static readonly COL_LINE_PARITY = 2;            // Colonne de la parité de ligne du parcours
    private static readonly COL_MISSION_CODE = 3;           // Colonne du code de mission
    private static readonly COL_NAME = 4;                   // Colonne du nom du parcours
    private static readonly COL_DEPARTURE_STATION = 5;      // Colonne de la gare de départ
    private static readonly COL_ARRIVAL_STATION = 6;        // Colonne de la gare d'arrivée
    private static readonly COL_VIA_STATIONS = 7;           // Colonne des gares intermédiaires

    // Map des parcours indexés par leur clé
    public static readonly map: Map<string, Path> = new Map();

    // /**
    //  * Charge les parcours de trains à partir du tableau "Sillons" de la feuille "Sillons".
    //  * Les parcours sont stockés dans un objet avec comme clés le numéro de parcours 
    //  * suivi du jour et comme valeur l'objet Path.
    //  * Chaque parcours correspondant à la sélection sera associé avec autant de clés que de jours
    //  * de circulation, en plus du numéro de parcours suivi du code des jours de circulation
    //  * (le parcours 123456_J aura pour clés : 123456_J, 123456_1, 123456_2...)
    //  * @param {string} days Jours pour lesquels les parcours sans jours spécifiques sont demandés.
    //  * @param {string} trainNumbers Numéros des parcours à charger, avec ou sans jours associés, séparés par des ';'.
    //  * Si vide, charge tous les trains de la base Paths.map.
    //  * @param {boolean} [erase=false] Si vrai, supprime les trains déjà chargés.
    //  *  Si faux (par défaut), ne recharge pas si déjà chargé.
    //  */
    // public static load(trainDays: string = "JW", trainNumbers: string = "", erase: boolean = false) {

    //     // Vérifie si la table à charger existe déjà
    //     if (Paths.map.size > 0) {
    //         if (erase) {
    //             Paths.map.clear(); // Vide la map sans changer sa référence
    //         }
    //     }

    //     Stations.load(); // Charge les gares si elles ne sont pas encore chargées
    //     const data = WorkbookService.getDataFromTable(Paths.SHEET, Paths.TABLE);

    //     // Map des parcours à charger : numéro → chaîne des jours associés
    //     // La concaténation des jours peut comporter plusieurs fois le même jour
    //     const trainNumberMap = new Map<string, string>();
    //     trainNumbers.split(';').forEach(entry => {
    //         const [number, days] = entry.split('_');
    //         const previous = trainNumberMap.get(number) || '';
    //         trainNumberMap.set(number, previous + (days || trainDays));
    //     });

    //     // Parcourt la base de données
    //     for (const row of data.slice(1)) {
    //         // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides)
    //         if (row.every(cell => !cell)) continue;

    //         const number = String(row[Paths.COL_NUMBER]);
    //         const days = String(row[Paths.COL_DAYS]);
            
    //         // Vérifie si le parcours est déjà chargé
    //         if (Paths.map.has(`${number}_${days}`)) continue;

    //         // Vérifie si le parcours est concerné dans la liste des parcours à charger, sauf si aucun filtre n'est fourni
    //         if (trainNumberMap.size > 0 && !trainNumberMap.has(`${number}`)) continue;

    //         // Détermine les jours à filtrer
    //         const filterDays = trainNumberMap.get(`${number}`) || trainDays;

    //         // Calcule les jours communs entre ceux du parcours et ceux demandés
    //         const commonDays = Day.extractFromString(days, filterDays);
    //         if (commonDays.length === 0) continue;

    //         // Extrait les valeurs
    //         const lineDirection = row[Paths.COL_LINE_PARITY] as number;
    //         const missionCode = String(row[Paths.COL_MISSION_CODE]);
    //         const departureTime = row[Paths.COL_DEPARTURE_TIME] as number;
    //         const departureStation = String(row[Paths.COL_DEPARTURE_STATION]);
    //         const arrivalTime = row[Paths.COL_ARRIVAL_TIME] as number;
    //         const arrivalStation = String(row[Paths.COL_ARRIVAL_STATION]);
    //         const viaStations = String(row[Paths.COL_VIA_STATIONS]);

    //         // Crée l'objet Path
    //         const path = new Path(
    //             number,
    //             lineDirection,
    //             days,
    //             missionCode,
    //             departureTime,
    //             departureStation,
    //             arrivalTime,
    //             arrivalStation,
    //             viaStations
    //         );

    //         // Insert le parcours dans la table avec plusieurs clés d'accès
    //         //  - une référence pour la clé unique du parcours
    //         Paths.map.set(path.key, path);
    //         //  - une référence pour chacun des jours demandés
    //         commonDays.forEach((day) => {
    //             const key = number + "_" + day;
    //             if (!Paths.map.has(key)) Paths.map.set(key, path);
    //         });
    //     }
    // }

    // /**
    //  * Affiche les parcours dans un tableau.
    //  * Les données sont celles stockées dans l'objet Paths.map.
    //  * @param {string} [sheetName=Paths.SHEET] Nom de la feuille de calcul.
    //  * @param {string} [tableName=Paths.TABLE] Nom du tableau.
    //  * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
    //  */
    // public static print(
    //     sheetName: string = Paths.SHEET,
    //     tableName: string = Paths.SHEET,
    //     startCell: string = "A1"
    // ): void {

    //     // Filtre l'objet Paths.map en ne prennant qu'une seule fois les parcours ayant la même clé   
    //     const seenKeys = new Set<string>();
    //     const uniquePaths: Path[] = Array.from(Paths.map.entries())
    //         .filter(([mapKey, path]) => mapKey === path.key)
    //         .map(([_, path]) => path);

    //     // Convertit l'objet Paths.map filtré en un tableau de données
    //     const data: (string | number)[][] = uniquePaths.map(path => [
    //         path.key,
    //         path.number,
    //         path.lineDirection.printDigit(),
    //         path.days,
    //         path.missionCode,
    //         path.departureTime,
    //         path.departureStation,
    //         path.arrivalTime,
    //         path.arrivalStation,
    //         path.viaStations.join(';'),
    //     ]);

    //     // Imprime le tableau
    //     const table = WorkbookService.printTable(Paths.HEADERS, data, sheetName, tableName, startCell);

    //     // Met les heures au format "hh:mm:ss"
    //     const timeColumns = [
    //         Paths.COL_DEPARTURE_TIME,
    //         Paths.COL_ARRIVAL_TIME,
    //     ];

    //     for (const col of timeColumns) {
    //         table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
    //     }
    // }

}


/**
 * Classe Train définissant un train, pour un unique jour, étant la réutilisation
 * d'un ou deux trains précédents, et ayant une ou deux réutilisations,
 * en faisant référence à un sillon avec horaires pouvant circuler plusieurs jours par semaine.
 */
class Train {

    // Propriétés de l'objet Train
    public readonly number: TrainNumber;            // Numéro du train
    public readonly path: Path;                     // Parcours sur lequel le train circule
    public readonly day: DateTime;                  // Jour du train    (1 à 7 = lundi à dimanche, >7 = date précise)
    public readonly service: string;                // Service auquel le train est rattaché
    public readonly firstStation?: string;          // Gare de départ si différente de celle du sillon
    public readonly lastStation?: string;           // Gare d'arrivée si différente de celle du sillon
    public readonly unit1: string;                  // Element 1 Nord (numéro de matériel)
    public readonly unit2: string;                  // Element 2 Sud (numéro de matériel)
    public readonly previous1: string;              // Clé du train précédent de l'élément 1
    public readonly previous2: string;              // Clé du train précédent de l'élément 2
    public readonly reuse1?: Train;                 // Train de réutilisation de l'élément 1
    public readonly reuse1Key: string;              // Clé du train de réutilisation de l'élément 1
    public readonly reuse2?: Train;                 // Train de réutilisation de l'élément 2
    public readonly reuse2Key: string;              // Clé du train de réutilisation de l'élément 2

    constructor(
        number: string,
        pathKey: string,
        day: number,
        service: string,
        unit1: string = "",
        unit2: string = "",
        previous1: string = "",
        previous2: string = "",
        reuse1Key: string = "",
        reuse2Key: string = ""
    ) {
        this.number = new TrainNumber(number);
        this.path = Paths.map.get(pathKey) as Path;
        this.day = day;
        this.service = service;
        this.unit1 = unit1;
        this.unit2 = unit2;
        this.previous1 = previous1;
        this.previous2 = previous2;
        this.reuse1Key = reuse1Key;
        this.reuse2Key = reuse2Key;
        if (!this.path) {
            Log.warn(`Train n° ${this.number}_${this.day} : le sillon rattaché est inconnu : ${pathKey}.`);
            return;
        }
    }

    /**
     * Vérifie la validité de l'objet Train en envoyant un message d'erreur si :
     *  - le sillon est inconnu.
     * @returns {Train | undefined} Objet Train s'il est valide, undefined sinon.
     */
    check(): Train | undefined {
        if (!this.path) {
            Log.warn(`Train n° ${this.number}_${this.day} : le sillon rattaché est inconnu : ${this.pathKey}.`);
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
        return this.path.getTrainNumber(with4Digits, withDoubleParity);;
    }
}

/**
 * Classe Trains contenant la liste des trains
 */
class Trains {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Trains";               // Feuille contenant la liste des trains
    private static readonly TABLE = "Trains";               // Tableau contenant la liste des trains
    private static readonly HEADERS = [[                    // En-têtes du tableau des trains
        "Id",
        "Numéro du train",
        "Jours",
        "Parcours",
        "Elément Nord",
        "Elément Sud",
        "Train Précédent Nord",
        "Train Précédent Sud",
        "Réutilisation Nord",
        "Réutilisation Sud",
    ]];
    private static readonly COL_KEY = 0;                    // Colonne de la clé du train
    private static readonly COL_NUMBER = 1;                 // Colonne du numéro du train
    private static readonly COL_DAYS = 2;                   // Colonne des jours de circulation
    private static readonly COL_TRAIN_PATH = 3;
    private static readonly COL_UNIT1 = 4;
    private static readonly COL_UNIT2 = 5;
    private static readonly COL_PREVIOUS1 = 6;
    private static readonly COL_PREVIOUS2 = 7;
    private static readonly COL_REUSE1 = 8;
    private static readonly COL_REUSE2 = 9;

    // Constantes de classe
    public static readonly UNKNOWN_UNIT = "?";
    
    // Map des trains indexées par abréviation
    public static readonly map: Map<string, Train> = new Map();
    
    /**
     * Charge les arrêts à partir du tableau "Arrêts" de la feuille "Arrêts".
     * Les gares sont stockées dans une Map avec comme clés l'abréviation 
     * Les arrêts sont stockés dans la propriété "stops" des trains et parcours correspondants.
     * Si un train n'existe pas, un message d'erreur est affiché.
     */
    public static load(erase: boolean = false): void {

        const data = WorkbookService.getDataFromTable(Stops.SHEET, Stops.TABLE);
        
        // Parcourt la base de données
        for (const row of data.slice(1)) {
    // private static readonly COL_KEY = 0;
    // private static readonly COL_NUMBER = 1;
    // private static readonly COL_DAYS = 2;
    // private static readonly COL_TRAIN_PATH = 3;
    // private static readonly COL_UNIT1 = 4;
    // private static readonly COL_UNIT2 = 5;
    // private static readonly COL_PREVIOUS1 = 6;
    // private static readonly COL_PREVIOUS2 = 7;
    // private static readonly COL_REUSE1 = 8;
    // private static readonly COL_REUSE2 = 9;
            // Vérifie si le train existe
            const trainNumber = String(row[COL_TRAIN_NUMBER]);
            const trainDays = String(row[COL_TRAIN_DAYS]);
            if (!trainNumber || !trainDays) continue;

            const trainKey = trainNumber + "_" + trainDays;
            if (!Paths.map.has(trainKey)) continue;

            const train = Paths.map.get(trainKey) as Path;

            // Extrait les valeurs
            const stationName = String(row[COL_STATION]);
            if (!stationName) continue;
            const parity = row[COL_PARITY] as number;
            const arrivalTime = row[COL_ARRIVAL_TIME] as number;
            const departureTime = row[COL_DEPARTURE_TIME] as number;
            const passageTime = row[COL_PASSAGE_TIME] as number;
            const tracks = String(row[COL_TRACK]);
            const changeNumber = row[COL_CHANGE_NUMBER] as number;
            const nextStopName = String(row[COL_NEXT_STOP]);

            const stop = new Stop(
                train.key,
                stationName,
                parity,
                arrivalTime,
                departureTime,
                passageTime,
                tracks,
                changeNumber,
                nextStopName
            );
            if (!stop) continue;

            // Ajoute l'arrêt au train
            train.addStop(stop);
        }

        // Boucle pour reparcourir tous les trains et vérifier leurs arrêts
        for (const train of Paths.map.values()) {
            train.checkStops();
        }
    }

    /**
     * Affiche les trains de la map dans un tableau.
     * @param {string} [sheetName=Trains.SHEET] Nom de la feuille de calcul.
     * @param {string} [tableName=Trains.TABLE] Nom du tableau.
     * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = Trains.SHEET,
        tableName: string = Trains.TABLE,
        startCell: string = "A1"
    ): void {

        // Convertit la map en un tableau de données
        const data: (string | number)[][] = Array
            .from(Trains.map.values())
            .map((train: Train) => [
                train.key,
                train.number.toString(false, true),
                train.day,
                train.path.key,
                train.unit1,
                train.unit2,
                train.previous1,
                train.previous2,
                train.reuse1.key,
                train.reuse2.key
            ]);

        WorkbookService.printTable(Trains.HEADERS, data, sheetName, tableName, startCell);
    }
}

function testWorkbookService(options: Partial<AssertDDOptions> = {}) {
    const assert = new AssertDD(options);
    const testSheetName = "testWorkbookService";
    const testTableName = "testTable";

    // 1️⃣ Test getSheet : création si inexistante
    let sheet = WorkbookService.getSheet(testSheetName, { createIfMissing: true });
    assert.check("Création d'une nouvelle feuille", sheet.getName(), testSheetName);

    // 2️⃣ Test getSheet : récupération feuille existante
    const sheet2 = WorkbookService.getSheet(testSheetName);
    assert.check("Récupération feuille existante", sheet2.getName(), testSheetName);

    // 3️⃣ Test checkCellName : valide
    assert.check("Cellule valide A1", WorkbookService.checkCellName("A1"), "A1");

    // 4️⃣ Test checkCellName : invalide
    assert.check(
        "Cellule invalide 123",
        WorkbookService.checkCellName("123", false),
        ""
    );

    // 5️⃣ Test printTable : création tableau avec données simples
    const headers = [["ColStr", "ColNum", "ColBool"]];
    const data = [
        ["Paris", 42, true],
        ["", "12", "FALSE"],
        [undefined, "abc", undefined]
    ];
    const table = WorkbookService.printTable(headers, data, testSheetName, testTableName, "A1", true);
    assert.check("Création tableau", table?.getName(), testTableName);

    // 6️⃣ Test getTable : récupération tableau existant
    const table2 = WorkbookService.getTable(testSheetName, testTableName, true);
    assert.check("Récupération tableau existant", table2?.getName(), testTableName);

    // 7️⃣ Test getDataFromTable : vérifie les données brutes
    const tableData = WorkbookService.getDataFromTable(testSheetName, testTableName, true);
    assert.check("Lecture donnée brute [Paris]", tableData[1][0], "Paris");

    // --- TESTS I/O WorkbookService ------------------------------------

    const row1 = tableData[1]; // ["Paris", 42, true]
    const row2 = tableData[2]; // ["", "12", "FALSE"]
    const row3 = tableData[3]; // [undefined, "abc", ""]

    // 8️⃣ getString
    assert.check("getString normal", WorkbookService.getString(row1, 0), "Paris");
    assert.check("getString chaîne vide", WorkbookService.getString(row2, 0), undefined);
    assert.check("getString undefined", WorkbookService.getString(row3, 0), undefined);

    // 9️⃣ getNumber
    assert.check("getNumber number", WorkbookService.getNumber(row1, 1), 42);
    assert.check("getNumber string numérique", WorkbookService.getNumber(row2, 1), 12);
    assert.check("getNumber string invalide", WorkbookService.getNumber(row3, 1), undefined);

    // 🔟 getBoolean
    assert.check("getBoolean true", WorkbookService.getBoolean(row1, 2), true);
    assert.check("getBoolean 'FALSE'", WorkbookService.getBoolean(row2, 2), false);
    assert.check("getBoolean undefined", WorkbookService.getBoolean(row3, 2), undefined);

    // 1️⃣1️⃣ Nettoyage : supprime la feuille de test
    WorkbookService.getSheet(testSheetName)?.delete();
    assert.check(
        "Suppression feuille test",
        WorkbookService.getSheet(testSheetName, { failOnError: false }),
        null
    );

    // 1️⃣2️⃣ Résumé
    assert.printSummary("Tests WorkbookService");
}

function testDateTime(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    DateTime.load();

    /* ==========================================================
    1. CONSTRUCTION via DateTime.from
    ----------------------------------------------------------
    Vérifie :
    - Rollover appliqué ou non
    - Temps relatifs
    - Parsing number / string
    ========================================================== */

    const constructorTests = [
        {
            desc: 'Heure après rollover (04:00)',
            value: 4 / 24,
            isRelative: false,
            expected: 4 / 24
        },
        {
            desc: 'Heure avant rollover (01:00 → 25:00)',
            value: 1 / 24,
            isRelative: false,
            expected: 1 / 24 + 1
        },
        {
            desc: 'Minuit (00:00 → 24:00)',
            value: 0,
            isRelative: false,
            expected: 1
        },
        {
            desc: 'Valeur string "0.5" (12:00)',
            value: "0.5",
            isRelative: false,
            expected: 0.5
        },
        {
            desc: 'Durée relative (01:00)',
            value: 1 / 24,
            isRelative: true,
            expected: 1 / 24
        },
    ];

    constructorTests.forEach(t => {
        const dt = DateTime.from(t.value, t.isRelative);

        assert.check(
            `DateTime.from(${t.value}, ${t.isRelative}) → excelValue (${t.desc})`,
            dt?.excelValue,
            t.expected
        );
    });

    /* ==========================================================
    2. DateTime.from()
    ----------------------------------------------------------
    Règle :
    - null | undefined → undefined
    - DateTime → même instance (si type compatible)
    ========================================================== */

    const base = DateTime.from(10 / 24)!;

    assert.check(
        'DateTime.from(DateTime) retourne la même instance',
        DateTime.from(base) === base,
        true
    );

    assert.check(
        'DateTime.from(number)',
        DateTime.from(10 / 24)?.excelValue,
        10 / 24
    );

    assert.check(
        'DateTime.from(string)',
        DateTime.from("0.5")?.excelValue,
        0.5
    );

    assert.check(
        'DateTime.from(undefined) → undefined',
        DateTime.from(undefined),
        undefined
    );

    assert.check(
        'DateTime.from(null) → undefined',
        DateTime.from(null),
        undefined
    );

    assert.check(
        'DateTime.from("") → undefined',
        DateTime.from(""),
        undefined
    );

    // ---- from() incohérence relatif / absolu ----
    let fromErrorCaught = false;
    try {
        const relative = DateTime.from(1 / 24, true)!;
        DateTime.from(relative, false);
    } catch (e) {
        fromErrorCaught = e.message.includes('relatif');
    }

    assert.check(
        'DateTime.from() erreur relatif → absolu',
        fromErrorCaught,
        true
    );

    /* ==========================================================
       3. format() — heure
       ========================================================== */

    const formatTimeTests = [
        {
            desc: '04:30:00',
            value: 4.5 / 24,
            format: DateTime.TIME_FORMAT_WITH_SECONDS,
            expected: '04:30:00'
        },
        {
            desc: '04:30',
            value: 4.5 / 24,
            format: DateTime.TIME_FORMAT_WITHOUT_SECONDS,
            expected: '04:30'
        },
        {
            desc: '-04:30',
            value: -4.5 / 24,
            format: DateTime.TIME_FORMAT_WITHOUT_SECONDS,
            expected: '-04:30'
        }
    ];

    formatTimeTests.forEach(t => {
        const dt = DateTime.from(t.value, true);
        assert.check(
            `format("${t.format}") (${t.desc})`,
            dt.format(t.format),
            t.expected
        );
    });

    /* ==========================================================
       4. format() — date
       ========================================================== */

    const formatDateTests = [
        {
            desc: 'Date Excel valide (22/06/2025)',
            value: 45830,
            format: DateTime.DATE_FORMAT_WITH_YEAR,
            expected: '22/06/2025'
        },
        {
            desc: 'Date sans année',
            value: 45830,
            format: DateTime.DATE_FORMAT_WITHOUT_YEAR,
            expected: '22/06'
        },
        {
            desc: 'Date avec heure',
            value: 45830.94347,
            format: DateTime.DATE_FORMAT_WITH_YEAR,
            expected: '22/06/2025'
        },
        {
            desc: 'Date avec jour de la semaine',
            value: 45830.94347,
            format: "dddd dd/mm/yyyy",
            expected: 'Dimanche 22/06/2025'
        }
    ];

    formatDateTests.forEach(t => {
        const dt = DateTime.from(t.value);
        assert.check(
            `format("${t.format}") (${t.desc})`,
            dt.format(t.format),
            t.expected
        );
    });

    /* ==========================================================
       5. format() — ID
       ========================================================== */

    const formatIdTests = [
        {
            desc: 'Date Excel valide',
            value: 45830,
            expected: '250622'
        },
        {
            desc: 'Date avec heure',
            value: 45830.75,
            expected: '250622'
        }
    ];

    formatIdTests.forEach(t => {
        const dt = DateTime.from(t.value);
        assert.check(
            `format(DATE_FORMAT_FOR_ID) (${t.desc})`,
            dt.format(DateTime.DATE_FORMAT_FOR_ID),
            t.expected
        );
    });

    /* ==========================================================
       6. resolveAgainst(), relativeTo(), equalsTo(), compareTo()
       ========================================================== */

    const testCases = [
        {
            desc: 'Temps relatif positif (+3h)',
            ref: 45830 + 10 / 24,
            rel: 3 / 24,
            expectedResolve: 45830 + 13 / 24
        },
        {
            desc: 'Temps relatif négatif (-2h)',
            ref: 45830 + 15 / 24,
            rel: -2 / 24,
            expectedResolve: 45830 + 13 / 24
        },
        {
            desc: 'Temps absolu inchangé',
            ref: 45830 + 15 / 24,
            rel: 0,
            expectedResolve: 45830 + 15 / 24
        }
    ];

    testCases.forEach(t => {

        const reference = DateTime.from(t.ref)!;
        const relative  = DateTime.from(t.rel, true)!;
        const absolute  = DateTime.from(t.ref)!;

        assert.check(
            `resolveAgainst() (${t.desc})`,
            relative.resolveAgainst(reference).excelValue,
            t.expectedResolve
        );

        assert.check(
            `relativeTo() (${t.desc})`,
            absolute.relativeTo(reference).excelValue,
            0
        );

        assert.check(
            `equalsTo() (${t.desc})`,
            absolute.equalsTo(DateTime.from(t.ref)),
            true
        );

        assert.check(
            `compareTo() (${t.desc})`,
            absolute.compareTo(DateTime.from(t.ref)),
            0
        );
    });

    /* ==========================================================
       7. add(), subtract() – temps relatifs uniquement
       ========================================================== */

    const round = (v: number) => Math.round(v * 1e10) / 1e10;

    const relativeTestCases = [
        { a: 2 / 24, b: 3 / 24, add: 5 / 24, sub: -1 / 24 },
        { a: 3 / 24, b: -2 / 24, add: 1 / 24, sub: 5 / 24 },
        { a: -1 / 24, b: -2 / 24, add: -3 / 24, sub: 1 / 24 },
    ];

    relativeTestCases.forEach(t => {

        const A = DateTime.from(t.a, true)!;
        const B = DateTime.from(t.b, true)!;

        assert.check(
            'add()',
            round(A.add(B).excelValue),
            round(t.add)
        );

        assert.check(
            'subtract()',
            round(A.subtract(B).excelValue),
            round(t.sub)
        );
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary('testDateTime');
}

function testDay(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    Day.load();

    /* ==========================================================
       1. Constructeur
       ========================================================== */

    const constructorTests = [
        {
            desc: "Jour simple lundi",
            input: "1",
            expectedNumber: 1,
            expectedString: "1"
        },
        {
            desc: "Groupe semaine 12345",
            input: "5-4-3-2-1",
            expectedNumber: 0,
            expectedString: "12345"
        },
        {
            desc: "Valeurs invalides ignorées",
            input: "a9b1c7",
            expectedNumber: 0,
            expectedString: "17"
        }
    ];

    constructorTests.forEach(t => {
        const d = new Day("x", "x", t.input);

        assert.check(
            `new Day("${t.input}") → numbersString (${t.desc})`,
            d.numbersString,
            t.expectedString
        );

        assert.check(
            `new Day("${t.input}") → number (${t.desc})`,
            d.number,
            t.expectedNumber
        );
    });

    /* ==========================================================
       2. extractFromString
       ========================================================== */

    const extractTests = [
        { desc: "Nom complet", input: "lundi", expected: [1] },
        { desc: "Abréviation", input: "ma", expected: [2] },
        { desc: "Numéros mélangés", input: "7;1;3", expected: [1, 3, 7] },
        { desc: "Texte mixte", input: "lumeven", expected: [1, 3, 5] },
        { desc: "Mot clé groupe", input: "J", expected: [1, 2, 3, 4, 5] }
    ];

    extractTests.forEach(t => {
        const result = Day.extractFromString(t.input);
        assert.check(
            `Day.extractFromString("${t.input}") (${t.desc})`,
            JSON.stringify(result),
            JSON.stringify(t.expected)
        );
    });

    /* ==========================================================
       3. extractFromString — intersection
       ========================================================== */

    const intersectionTests = [
        {
            desc: "Intersection simple",
            input1: "lundi;mercredi",
            input2: "mer",
            expected: [3]
        },
        {
            desc: "Intersection groupe / jour",
            input1: "JOB",
            input2: "samedi;dimanche",
            expected: []
        },
        {
            desc: "Intersection multiple",
            input1: "JOB",
            input2: "mar-mer",
            expected: [2, 3]
        }
    ];

    intersectionTests.forEach(t => {
        const result = Day.extractFromString(t.input1, t.input2);
        assert.check(
            `Day.extractFromString("${t.input1}", "${t.input2}") (${t.desc})`,
            JSON.stringify(result),
            JSON.stringify(t.expected)
        );
    });

    /* ==========================================================
       4. fromNumber
       ========================================================== */

    const fromNumberTests = [
        { input: 1, expected: "Lundi" },
        { input: 7, expected: "Dimanche" },
        { input: 0, expected: "Dimanche" }
    ];

    fromNumberTests.forEach(t => {
        const day = Day.fromNumber(t.input);
        assert.check(
            `Day.fromNumber(${t.input})`,
            day?.fullName,
            t.expected
        );
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary('testDay');
}

function testParity(options: Partial<AssertDDOptions> = {}) {

    Parity.load();
    const assert = new AssertDD(options);

    /* ==========================================================
       TESTS DATA-DRIVEN - CLASSE Parity
       ========================================================== */

    /* ==========================================================
       1. CONSTRUCTEUR & normalizeParityValue()
       ========================================================== */

    const constructorTests = [
        { desc: 'Lettre impair "I"', value: "I", doubleAllowed: false, expected: Parity.ODD },
        { desc: 'Lettre pair "P"', value: "P", doubleAllowed: false, expected: Parity.EVEN },
        { desc: 'Chiffre impair 1', value: 1, doubleAllowed: false, expected: Parity.ODD },
        { desc: 'Chiffre pair 2', value: 2, doubleAllowed: false, expected: Parity.EVEN },
        { desc: 'Numéro de train impair', value: "12345", doubleAllowed: false, expected: Parity.ODD },
        { desc: 'Numéro de train pair', value: "12346", doubleAllowed: false, expected: Parity.EVEN },
        { desc: 'Valeur vide', value: "", doubleAllowed: false, expected: Parity.UNDEFINED },
        { desc: 'Zéro "0"', value: "0", doubleAllowed: false, expected: Parity.UNDEFINED },
        { desc: 'Double IP interdite', value: "IP", doubleAllowed: false, expected: Parity.UNDEFINED },
        { desc: 'Double IP autorisée', value: "IP", doubleAllowed: true, expected: Parity.DOUBLE },
        { desc: 'Double implicite "1/2"', value: "1/2", doubleAllowed: true, expected: Parity.DOUBLE }
    ];

    constructorTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleAllowed);
        assert.check(
            `Parity.from(${JSON.stringify(t.value)}, ${t.doubleAllowed}) – ${t.desc}`,
            p.value,
            t.expected
        );
    });

    /* ==========================================================
       2. null / undefined
       ========================================================== */

       const nullTests = [
        { value: null },
        { value: undefined }
    ];

    nullTests.forEach(t => {
        const p = Parity.from(t.value);
        assert.check(
            `Parity.from(${t.value}) → undefined`,
            p.value,
            Parity.UNDEFINED
        );
    });

    /* ==========================================================
       3. is() / isDefined() / isOpposedTo()
       ========================================================== */


    const isTests = [
        { value: "I", parity: Parity.ODD, expected: true },
        { value: "I", parity: Parity.EVEN, expected: false }
    ];

    isTests.forEach(t => {
        assert.check(
            `is ${t.value} === ${t.parity}`,
            Parity.from(t.value).is(t.parity),
            t.expected
        );
    });

    const isDefinedTests = [
        { parity: Parity.UNDEFINED, expected: false },
        { parity: Parity.EVEN, expected: true }
    ];

    isDefinedTests.forEach(t => {
        assert.check(
            `isDefined ${t.parity}`,
            Parity.from(t.parity).isDefined(),
            t.expected
        );
    });

    const isOpposedTests = [
        { a: "I", b: "P", expected: true },
        { a: "P", b: "I", expected: true },
        { a: "I", b: "I", expected: false },
        { a: undefined, b: "I", expected: false },
        { a: undefined, b: undefined, expected: false }
    ];

    isOpposedTests.forEach(t => {
        const a = t.a !== undefined ? Parity.from(t.a) : undefined;
        const b = t.b !== undefined ? Parity.from(t.b) : undefined;
        assert.check(
            `isOpposedTo ${t.a} / ${t.b}`,
            a?.isOpposedTo(b) ?? false,
            t.expected
        );
    });

    /* ==========================================================
       4. equalsTo() / 
       ========================================================== */

    const equalsTests = [
        { a: "I", b: 1, expected: true },
        { a: "I", b: "P", expected: false }
    ];

    equalsTests.forEach(t => {
        assert.check(
            `equalsTo ${t.a} / ${t.b}`,
            Parity.from(t.a).equalsTo(Parity.from(t.b)),
            t.expected
        );
    });

    const pRef = Parity.from("I");
    const pCopy = Parity.from(pRef);
    
    assert.check(
        'Parity.from recrée une instance équivalente',
        pCopy.equalsTo(pRef),
        true
    );
    
    assert.check(
        'Parity.from ne conserve pas la même instance',
        pCopy === pRef,
        false
    );

    /* ==========================================================
        5. Parity.includes()
        ----------------------------------------------------------
        Vérifie :
        - parité simple vs simple
        - parité double
        - undefined
        - valeurs non valides
        ========================================================== */

    const includesTests = [
        // --- parité simple ---
        { a: "I", b: "I", expected: true },
        { a: "I", b: 1,   expected: true },   // odd / odd
        { a: "I", b: "P", expected: false },
        { a: "P", b: "I", expected: false },

        // --- parité double ---
        { a: "IP", b: "I", expected: true },
        { a: "IP", b: "P", expected: true },
        { a: "IP", b: 1,   expected: true },
        { a: "IP", b: 2,   expected: true },

        // --- simple n'inclut pas double ---
        { a: "I",  b: "IP", expected: false },
        { a: "P",  b: "IP", expected: false },

        // --- undefined ---
        { a: null, b: "I", expected: false },
        { a: "I",  b: null, expected: false },
        { a: null, b: null, expected: false },
    ];

    includesTests.forEach(t => {
        assert.check(
            `includes ${t.a} ⊇ ${t.b}`,
            Parity.from(t.a, true).includes(t.b),
            t.expected
        );
    });

    /* ==========================================================
       6. invert()
       ========================================================== */

    const invertTests = [
        { value: "I", doubleAllowed: false, expected: Parity.EVEN },
        { value: "P", doubleAllowed: false, expected: Parity.ODD },
        { value: "IP", doubleAllowed: true, expected: Parity.DOUBLE },
        { value: "", doubleAllowed: false, expected: Parity.UNDEFINED }
    ];

    invertTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleAllowed).invert();
        assert.check(
            `invert ${t.value}`,
            p.value,
            t.expected
        );
    });

    /* ==========================================================
       7. printDigit() / printLetter()
       ========================================================== */

    const printTests = [
        { value: "I", digit: Parity.digit(Parity.ODD), letter: Parity.letter(Parity.ODD) },
        { value: "P", digit: Parity.digit(Parity.EVEN), letter: Parity.letter(Parity.EVEN) },
        {
            value: "IP",
            doubleAllowed: true,
            digit: Parity.digit(Parity.DOUBLE),
            letter: Parity.letter(Parity.ODD) + Parity.letter(Parity.EVEN)
        },
        { value: "", digit: "", letter: "" }
    ];

    printTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleAllowed);
        assert.check(`printDigit ${t.value}`, p.printDigit(), t.digit);
        assert.check(`printLetter ${t.value}`, p.printLetter(), t.letter);
    });

    /* ==========================================================
       8. printDigit(withUnderscores)
       ========================================================== */

    const underscoreTests = [
        { value: "I", expected: "_" + Parity.digit(Parity.ODD) },
        { value: "P", expected: "_" + Parity.digit(Parity.EVEN) },
        { value: "IP", doubleAllowed: true, expected: Parity.digit(Parity.DOUBLE) },
        { value: "", expected: "" }
    ];

    underscoreTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleAllowed);
        assert.check(
            `printDigit underscore ${t.value}`,
            p.printDigit(true),
            t.expected
        );
    });

    /* ==========================================================
       9. containsParityLetter()
       ========================================================== */

    const containsTests = [
        { text: "Train I", parity: Parity.ODD, expected: true },
        { text: "Train I", parity: Parity.EVEN, expected: false },
        { text: "Train IP", parity: Parity.DOUBLE, expected: true }
    ];

    containsTests.forEach(t => {
        assert.check(
            `containsParityLetter "${t.text}"`,
            Parity.containsParityLetter(t.text, t.parity),
            t.expected
        );
    });

    /* ==========================================================
       10. static letter() / digit()
       ========================================================== */

    const staticTests = [
        { method: 'letter', parity: Parity.ODD, type: 'string' },
        { method: 'letter', parity: Parity.EVEN, type: 'string' },
        { method: 'digit', parity: Parity.ODD, type: 'number' },
        { method: 'digit', parity: Parity.EVEN, type: 'number' },
        { method: 'digit', parity: 999, expected: 0 }
    ];

    staticTests.forEach(t => {
        const result =
            t.method === 'letter'
                ? Parity.letter(t.parity)
                : Parity.digit(t.parity);

        if ('expected' in t) {
            assert.check(`${t.method}(${t.parity})`, result, t.expected);
        } else {
            assert.check(`${t.method}(${t.parity}) type`, typeof result, t.type);
        }
    });

    assert.printSummary("testParity");
}

function testTrainNumber(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    TrainNumber.load(true);

    // ------------------------------------------------------------
    // Constructeur
    // ------------------------------------------------------------

    const constructorTests = [
        { desc: "Nombre simple", input: 146490, expected: "146490" },
        { desc: "Chaîne avec slash", input: "146490/91", expected: "146490/1" },
        { desc: "Minuscules + parasites", input: "w-14a6490", expected: "W14A6490" }
    ];

    constructorTests.forEach(t => {
        const tn = new TrainNumber(t.input);
        assert.check(
            `new TrainNumber(${JSON.stringify(t.input)}) (${t.desc})`,
            tn.value,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // isW()
    // ------------------------------------------------------------

    const wTests = [
        { value: 146490, expected: true },
        { value: 569907, expected: true },
        { value: 147490, expected: false },
        { value: 165470, expected: false }
    ];

    wTests.forEach(t => {
        const tn = new TrainNumber(t.value);
        assert.check(`isW(${t.value})`, tn.isW(), t.expected);
    });

    // ------------------------------------------------------------
    // Tests toString()
    // ------------------------------------------------------------

    const printTests = [
        {
            desc: "Sans abréviation, sans double parité",
            value: 146490,
            doubleParity: false,
            abbreviate: false,
            withoutDoubleParity: false,
            expected: "146490"
        },
        {
            desc: "Abréviation à 4 chiffres",
            value: 146490,
            doubleParity: false,
            abbreviate: true,
            withoutDoubleParity: false,
            expected: "6490"
        },
        {
            desc: "Non abrégeable à 4 chiffres",
            value: 569907,
            doubleParity: false,
            abbreviate: true,
            withoutDoubleParity: false,
            expected: "569907"
        },
        {
            desc: "Double parité implicite conservée",
            value: "146490/1",
            doubleParity: false,
            abbreviate: false,
            withoutDoubleParity: false,
            expected: "146490/1"
        },
        {
            desc: "Double parité forcée par constructeur",
            value: 146490,
            doubleParity: true,
            abbreviate: false,
            withoutDoubleParity: false,
            expected: "146490/1"
        },
        {
            desc: "Double parité avec abréviation",
            value: 146490,
            doubleParity: true,
            abbreviate: true,
            withoutDoubleParity: false,
            expected: "6490/1"
        },
        {
            desc: "Double parité masquée",
            value: 146490,
            doubleParity: true,
            abbreviate: true,
            withoutDoubleParity: true,
            expected: "6490"
        }
    ];

    printTests.forEach(t => {
        const tn = new TrainNumber(t.value, t.doubleParity);

        assert.check(
            `TrainNumber(${t.value}, doubleParity=${t.doubleParity}).toString(${t.abbreviate}, ${t.withoutDoubleParity}) (${t.desc})`,
            tn.toString(t.abbreviate, t.withoutDoubleParity),
            t.expected
        );
    });

    // ------------------------------------------------------------
    // adaptWithParity()
    // ------------------------------------------------------------

    const parityTests = [
        { value: 146491, parity: Parity.EVEN, expected: "146490" },
        { value: 146490, parity: Parity.ODD, expected: "146491" },
        { value: 146490, parity: Parity.DOUBLE, expected: "146490/1" },
        { value: 146490, parity: Parity.DOUBLE, abbreviate: true, expected: "6490/1" }
    ];

    parityTests.forEach(t => {
        const tn = new TrainNumber(t.value);
        assert.check(
            `adaptWithParity(${t.value}, ${t.parity})`,
            tn.adaptWithParity(t.parity, t.abbreviate),
            t.expected
        );
    });

    assert.printSummary("testTrainNumber");
}

function testStations(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    Stations.load(true);

    /* ==========================================================
       TESTS DATA-DRIVEN - CLASSE Stations
       ==========================================================
       Objectifs :
       - Valider le chargement des gares
       - Vérifier la cohérence des rattachements
       - Garantir l’accès via la Map statique
       - Tester l’impression dans une feuille de test
       ========================================================== */

    /* ==========================================================
       1. load() & état global
       ----------------------------------------------------------
       Vérifie :
       - Chargement des gares
       - Non-duplication sans erase
       - Vidage et rechargement avec erase
       ========================================================== */


    assert.check(
        "Stations.load(true) - au moins une gare chargée",
        Stations.size > 0,
        true
    );

    const sizeAfterFirstLoad = Stations.size;

    Stations.load(false);

    assert.check(
        "Stations.load(false) - pas de rechargement supplémentaire",
        Stations.size,
        sizeAfterFirstLoad
    );

    Stations.load(true);

    assert.check(
        "Stations.load(true) - rechargement après erase",
        Stations.size,
        sizeAfterFirstLoad
    );

    /* ==========================================================
       2. Accès Map & get()
       ----------------------------------------------------------
       Vérifie :
       - Présence des clés
       - Cohérence entre get() et map.get()
       ========================================================== */

    const firstStation = Stations.map.values().next().value as Station;

    assert.check(
        "Stations.map contient au moins une Station",
        firstStation instanceof Station,
        true
    );

    if (firstStation) {
        assert.check(
            `Stations.get("${firstStation.abbreviation}") retourne la même instance`,
            Stations.get(firstStation.abbreviation),
            firstStation
        );
    }

    /* ==========================================================
       3. Rattachements parent / enfants
       ----------------------------------------------------------
       Vérifie :
       - Cohérence referenceStation → childStations
       - Cohérence childStations → referenceStation
       ========================================================== */

    const attachmentTests: {
        desc: string;
        valid: boolean;
    }[] = [];

    for (const station of Stations.values()) {

        if (station.referenceStation) {
            attachmentTests.push({
                desc:
                    `${station.abbreviation} référencée dans `
                    + ` ${station.referenceStation.abbreviation}.childStations`,
                valid: station.referenceStation.childStations.includes(station)
            });
        }

        for (const child of station.childStations) {
            attachmentTests.push({
                desc:
                    `${child.abbreviation}.referenceStation === `
                    + station.abbreviation,
                valid: child.referenceStation === station
            });
        }
    }

    attachmentTests.forEach(t => {
        assert.check(
            `Rattachement - ${t.desc}`,
            t.valid,
            true
        );
    });

    /* ==========================================================
       4. Données métier de base
       ----------------------------------------------------------
       Vérifie :
       - Abréviation non vide
       - Parité valide
       ========================================================== */

    const dataTests = Array.from(Stations.values()).map(station => ({
        desc: `Station ${station.abbreviation} - abréviation non vide`,
        value: station.abbreviation !== ""
    }));

    dataTests.forEach(t => {
        assert.check(t.desc, t.value, true);
    });

    /* ==========================================================
       5. print() - feuille et tableau de test
       ----------------------------------------------------------
       Vérifie :
       - Exécution sans erreur
       - Création d’un tableau test indépendant
       ========================================================== */

    let printSucceeded = true;

    try {
        Stations.print(
            "testGares",
            "testGares",
            "A1"
        );
    } catch (e) {
        printSucceeded = false;
    }

    assert.check(
        'Stations.print() - impression dans la feuille "testGares"',
        printSucceeded,
        true
    );

    // Synthèse finale
    assert.printSummary("testStations");
}

function testStationWithParity(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    /* ==========================================================
       1. Construction valide
       ----------------------------------------------------------
       Vérifie :
       - Initialisation sans parité explicite
       - Initialisation avec parité dans la chaîne
       - Initialisation avec parité explicite
       ========================================================== */

    const validTests = [
        {
            desc: "Station sans parité explicite",
            value: "PZB",
            parity: undefined,
            expectedStation: "PZB",
            expectedParity: Parity.UNDEFINED
        },
        {
            desc: "Station avec parité dans la chaîne",
            value: "PZB_1",
            parity: undefined,
            expectedStation: "PZB",
            expectedParity: Parity.ODD
        },
        {
            desc: "Station avec parité explicite",
            value: "PZB",
            parity: Parity.EVEN,
            expectedStation: "PZB",
            expectedParity: Parity.EVEN
        }
    ];

    validTests.forEach(t => {
        const s = StationWithParity.from(t.value, t.parity);

        assert.check(
            `${t.desc} - station`,
            s.station.abbreviation,
            t.expectedStation
        );

        assert.check(
            `${t.desc} - parity`,
            s.parity.is(t.expectedParity),
            true
        );
    });

    /* ==========================================================
       2. Initialisation de la parité (formes acceptées)
       ========================================================== */

    const parityInitTests = [
        { desc: "parité en Parity", value: "PZB", parity: Parity.ODD, expected: Parity.ODD },
        { desc: "parité en string", value: "PZB", parity: "2", expected: Parity.EVEN },
        { desc: "parité en number", value: "PZB", parity: 1, expected: Parity.ODD }
    ];

    parityInitTests.forEach(t => {
        const s = StationWithParity.from(t.value, t.parity);

        assert.check(
            `${t.desc}`,
            s.parity.is(t.expected),
            true
        );
    });

    /* ==========================================================
       3. Conflit de parité
       ----------------------------------------------------------
       Vérifie :
       - Exception levée si conflit entre chaîne et paramètre
       ========================================================== */

    assert.throws(
        "Conflit de parité entre chaîne et paramètre",
        () => StationWithParity.from("PZB_1", Parity.EVEN)
    );

    /* ==========================================================
       4. from()
       ----------------------------------------------------------
       Vérifie :
       - Retourne la même instance si StationWithParity
       - Crée une instance depuis Station ou string
       - Lève une erreur si null / undefined
       ========================================================== */

    const s1 = StationWithParity.from("PZB_1");

    assert.check(
        "from(StationWithParity) retourne la même instance",
        StationWithParity.from(s1) === s1,
        true
    );

    assert.check(
        "from(string)",
        StationWithParity.from("PZB_2").parity.is(Parity.EVEN),
        true
    );

    assert.throws(
        "from(null) lève une erreur",
        () => StationWithParity.from(null)
    );

    assert.throws(
        "from(undefined) lève une erreur",
        () => StationWithParity.from(undefined)
    );

    /* ==========================================================
       5. Egalité
       ----------------------------------------------------------
       Vérifie :
       - equalsTo() vrai si même station et même parité
       - equalsTo() faux sinon
       ========================================================== */

    const a = StationWithParity.from("PZB_1");
    const b = StationWithParity.from("PZB", Parity.ODD);
    const c = StationWithParity.from("PZB_2");
    const d = StationWithParity.from("BFM_1");

    assert.check(
        "equalsTo() vrai",
        a.equalsTo(b),
        true
    );

    assert.check(
        "equalsTo() faux",
        a.equalsTo(c),
        false
    );

    assert.check(
        "equalsTo(undefined) faux",
        a.equalsTo(undefined),
        false
    );

    assert.check(
        "hasSameStationTo() vrai",
        a.hasSameStationTo(c),
        true
    );

    assert.check(
        "hasSameStationTo() faux",
        a.hasSameStationTo(d),
        false
    );

    /* ==========================================================
       6. stationAfterTurnaround()
       ----------------------------------------------------------
       Vérifie :
       - Retourne une nouvelle instance
       - Parité inversée
       ========================================================== */

       const turned = a.stationAfterTurnaround();

       assert.check(
           "stationAfterTurnaround() retourne une instance si rebroussement possible",
           turned !== undefined,
           true
       );
       
       assert.check(
           "stationAfterTurnaround() retourne une nouvelle instance",
           turned !== a,
           true
       );
       
       assert.check(
           "stationAfterTurnaround() conserve la gare",
           turned?.station === a.station,
           true
       );
       
       assert.check(
           "stationAfterTurnaround() inverse la parité",
           turned?.parity.is(Parity.EVEN),
           true
       );
       

    /* ==========================================================
       7. toString()
       ----------------------------------------------------------
       Vérifie :
       - Format GARE_PARITE
       ========================================================== */

    assert.check(
        "toString()",
        StationWithParity.from("PZB_2").toString(),
        "PZB_2"
    );

    // Synthèse finale
    assert.printSummary("testStationWithParity");
}
function testConnections(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    /* ==========================================================
       TESTS DATA-DRIVEN - CLASSE Connections
       ==========================================================
       Objectifs :
       - Vérifier le chargement des connexions
       - Garantir l’unicité et la cohérence from → to
       - Tester la cohérence StationWithParity ↔ clés Map
       - Tester l’impression
       ========================================================== */

    /* ==========================================================
       1. load()
       ========================================================== */

    Connections.load(true);

    assert.check(
        "Connections.load(true) - au moins une connexion chargée",
        Connections.size > 0,
        true
    );

    const sizeAfterLoad = Connections.size;

    Connections.load(false);

    assert.check(
        "Connections.load(false) - pas de rechargement",
        Connections.size,
        sizeAfterLoad
    );

    /* ==========================================================
       2. Accès Map & get()
       ========================================================== */

    const firstFrom = Connections.map.keys().next().value as string;
    const firstTo =
        Connections.map.get(firstFrom)!.keys().next().value as string;

    const c = Connections.get(firstFrom, firstTo);

    assert.check(
        `Connections.get(${firstFrom}, ${firstTo}) retourne une Connection`,
        c instanceof Connection,
        true
    );

    /* ==========================================================
       3. Cohérence métier & structure
       ========================================================== */

    const coherenceTests: { desc: string; value: boolean }[] = [];

    for (const [fromKey, targets] of Connections.map) {
        for (const [toKey, connection] of targets) {

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : from est StationWithParity`,
                value: connection.from instanceof StationWithParity
            });

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : to est StationWithParity`,
                value: connection.to instanceof StationWithParity
            });

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : cohérence clé from`,
                value: connection.from.toString() === fromKey
            });

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : cohérence clé to`,
                value: connection.to.toString() === toKey
            });

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : from ≠ to`,
                value: !connection.from.equalsTo(connection.to)
            });

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : temps > 0`,
                value: connection.time.excelValue > 0
            });

            coherenceTests.push({
                desc: `${fromKey} → ${toKey} : temps relatif`,
                value: connection.time.isRelative === true
            });
        }
    }

    coherenceTests.forEach(t =>
        assert.check(`Cohérence - ${t.desc}`, t.value, true)
    );

    /* ==========================================================
       4. print()
       ========================================================== */

    let printOk = true;

    try {
        Connections.print(
            "testConnexions",
            "testConnexions",
            "A1"
        );
    } catch {
        printOk = false;
    }

    assert.check(
        'Connections.print() - impression feuille "testConnexions"',
        printOk,
        true
    );

    // Synthèse finale
    assert.printSummary("testConnections");
}
