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
 * @returns {boolean} Vrai si les tests sont actifs, faux sinon.
 */
function runAllTests(testMode: boolean = false): boolean {

    if (!testMode) return false;

    Params.load();
    Connections.load();

    // testWorkbookService({ printSuccess: false, printFailure: true });
    testDateTime({ printSuccess: false, printFailure: true });
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
     *  (vrai si le nombre est différent de 0, faux sinon).
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

/*
 * Classe utilitaire immuable contenant les valeurs d'une date Excel
 */
class DateExcel {

    // Constante de valeur initiale (epoch) des dates Excel
    public static readonly EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30));
        
    // Propriétés de l'objet DateTime
    public readonly value: number;
    public readonly year: number;
    public readonly month: number;
    public readonly day: number;
    public readonly dayOfWeek: Day | undefined;

    /**
     * Constructeur de l'objet DateExcel.
     * @param {number} excelValue Valeur Excel du jour, qui représente le nombre de jours
     *  écoulés depuis le 30 décembre 1899.
     */
    constructor(excelValue: number) {

        this.value = Math.floor(excelValue);

        const ms = DateExcel.EXCEL_EPOCH.getTime() + this.value * 86400000;
        const d = new Date(ms);

        this.year = d.getUTCFullYear();
        this.month = d.getUTCMonth() + 1;
        this.day = d.getUTCDate();
        this.dayOfWeek = Day.fromNumber(d.getUTCDay());
    }
}

/*
 * Classe utilitaire immuable contenant les valeurs d'une heure Excel
 */
class TimeExcel {

    // Propriétés de l'objet DateTime
    public readonly value: number;
    public readonly hour: number;
    public readonly minute: number;
    public readonly second: number;

    /**
     * Constructeur de l'objet TimeExcel.
     * @param {number} excelValue Valeur Excel du temps, dont la fraction de jour représente l'heure.
     */
    constructor(excelValue: number) {
        this.value = excelValue;
        const abs = Math.abs(this.value);
        const totalSeconds = Math.round(abs * 86400);
        this.hour = Math.floor(totalSeconds / 3600);
        this.minute = Math.floor((totalSeconds % 3600) / 60);
        this.second = totalSeconds % 60;
    }
}

/**
 * Classe utilitaire immutable pour la gestion des dates et horaires Excel.
 *  Si le temps est absolu et non daté, et que l'heure est inférieure à l'heure de changement de journée,
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
    private _computed = false;                      // Indique si les éléments de la date sont calculés

    // Valeurs des éléments
    private _realDate: DateExcel | undefined;       // Date réelle
    private _adaptedDate: DateExcel | undefined;    // Date adaptée si l'heure de la date est inférieure
                                                    //  à l'heure de changement de jour

    private _time: TimeExcel | undefined;           // Heure de la journée

    /**
     * Constructeur privé de l'objet DateTime.
     * @param {number} [excelValue=0] Valeur du temps en format Excel
     *  à partir du 01/01/1900 00:00:00.
     * @param {boolean} [isRelative=false] Indique si le temps est relatif
     *  (différence entre 2 horaires).
     * @param {boolean} [adaptTime=true] Indique si le temps doit être adapté pour
     *  tenir compte de l'heure de changement de jour (temps absolu non daté uniquement).
     *  Si la valeur est inférieure à l'heure de changement de journée
     *  et que l'horaire est absolu, elle est incrémentée de 1.
     *  Si le temps est daté, il sera adapté dans la méthode compute.
     *  Si le temps est relatif, il ne peut pas être adapté.
     */
    private constructor(excelValue: number = 0, isRelative: boolean = false, adaptTime: boolean = true) {
        this.isRelative = isRelative;
        this.excelValue = (!isRelative && adaptTime) ? DateTime.adaptTime(excelValue) : excelValue;
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
    ): DateTime | undefined {
        
        if (value == null || value === "") return undefined;

        if (value instanceof DateTime) {
            if (value.isRelative !== isRelative) {
                throw new Error(
                    `Un temps ${value.isRelative ? "relatif" : "absolu"}`
                    + ` cherche à être affecté à un temps ${isRelative ? "relatif" : "absolu"}.`
                );
            }
            return value;
        }

        const v = Number(value);

        // Un temps absolu doit être supérieur ou égal à 0
        if (!isRelative && v < 0) return undefined;

        return new DateTime(v, isRelative, adaptTime);
    }  

    /**
     * Accesseurs utilitaires
     */
    // Objet Date (selon s'il est demandé adapté ou non)
    private getDateObj(adapted: boolean): DateExcel | undefined {
        return (adapted && this._adaptedDate) ? this._adaptedDate : this._realDate;
    }
    // Date
    public getDate(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.value ?? 0;
    }
    // Année
    public getYear(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.year ?? 0;
    }
    // Mois
    public getMonth(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.month ?? 0;
    }
    // Jour
    public getDay(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.day ?? 0;
    }
    // Jour de la semaine
    public getDayOfWeek(adaptedValue: boolean = true): Day | undefined {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.dayOfWeek;
    }
    // Temps
    //  Si le temps est daté, il correspond à la fraction d'une journée.
    //  Si la temps est adapté, il est incrémenté de 1 (ex : 25h00)
    public getTime(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        if (!timeObj) return 0;
        if (!this.isRelative && !adaptedValue) {
            return timeObj.value % 1;
        }
        return timeObj.value;
    }
    // Heure
    //  Si le temps est adapté, l'heure est incrémentée de 24 (ex : 25h00)
    public getHours(adaptedValue: boolean = false): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        if (!timeObj) return 0;
        if (!this.isRelative && !adaptedValue) {
            return timeObj.hour % 24;
        }
        return timeObj.hour;
    }
    // Minute
    public getMinutes(): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        return timeObj?.minute ?? 0;
    }
    // Seconde
    public getSeconds(): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        return timeObj?.second ?? 0;
    }

    /**
     * Calcule les éléments de la date et de l'heure de la journée.
     * Si le temps est relatif, seule l'heure est calculée.
     * Si le temps est absolu et non daté (<1), il a déjà été adapté dans le constructeur,
     *  seule l'heure est donc calculée.
     * Si le temps est absolu et daté, l'heure et la date sont calculées, et également adaptées
     *  si l'heure est inférieure à l'heure de changement de journée. Dans ce cas la date adaptée
     *  correspond à la date du jour précédent, et l'heure est incrémentée de 1 (+24h).
     */
    private compute(): void {
        if (this._computed) return;

        // Récupère la valeur du temps, inchangé si le temps est relatif ou absolu non daté
        let timeOfDay = this.excelValue;

        // Calcul des éléments de la date (si temps absolu et daté)
        if (!this.isRelative && this.excelValue > DateTime.MIN_EXCEL_DATE) {
            this._realDate = new DateExcel(this.excelValue);
            timeOfDay = this.excelValue % 1;
            if (timeOfDay < DateTime.rolloverHour) {
                this._adaptedDate = new DateExcel(this.excelValue - 1);
                timeOfDay += 1;
            }
        }

        // Calcul des éléments de l'heure de la journée
        //  - si le temps est relatif, l'heure correspond au temps total, positif ou négatif,
        //  - si le temps est absolu, l'heure est la fraction de la journée,
        //     adaptée si l'heure est inférieure à l'heure de changement de jour,
        //     d'une valeur comprise entre 0 et 1, ou dépassant 1 si adaptée
        this._time = new TimeExcel(timeOfDay);
        
        this._computed = true;
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
     * Vérifie si les deux temps sont identiques ou s'ils sont tous les deux undefined.
     * @param {Parity | undefined} a Premier temps à comparer.
     * @param {Parity | undefined} b Second temps à comparer.
     * @returns {boolean} Vrai si les deux temps sont identiques
     *  ou s'ils sont tous les deux undefined, faux sinon.
     */
    public static equalsOrUndefined(
        a?: DateTime,
        b?: DateTime
    ): boolean {
        return a === b || (!!a && !!b && a.equalsTo(b));
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
    public format(format: string, adaptTime: boolean = true): string {
        this.compute();
        let prefix = "";
        if (this.excelValue < 0) prefix = "-";
        const pad = (v: number) => v.toString().padStart(2, "0");
    
        const tokens: Record<string, string> = {
        // Année
        "yyyy": this.getYear().toString(),
        "yy": pad(this.getYear(adaptTime) % 100),
        // Mois
        "mm": pad(this.getMonth()),
        "m": this.getMonth(adaptTime).toString(),
        // Jour
        "dd": pad(this.getDay()),
        "d": this.getDay(adaptTime).toString(),
        // Jour de semaine
        "dddd": this.getDayOfWeek(adaptTime)?.fullName ?? "",
        "ddd": this.getDayOfWeek(adaptTime)?.abreviation ?? "",
        // Heure
        "hh": pad(this.getHours(adaptTime)),
        "h": this.getHours().toString(),
        // Minute
        "nn": pad(this.getMinutes()),
        "n": this.getMinutes().toString(),
        // Seconde
        "ss": pad(this.getSeconds()),
        "s": this.getSeconds().toString(),
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
     * Ajuste une heure pour tenir compte du changement de journée.
     * Si l'heure est inférieure à l'heure de changement de journée,
     *  on ajoute 1 pour passer à la journée suivante.
     *  Par exemple : 01:00 → 25:00 si changement de journée à 03:00
     * Cela ne s'applique que sur les heures non datées (valeur < 1).
     * @param {number} time Heure à ajuster.
     * @returns {number} Heure ajustée.
     */
    public static adaptTime(time: number): number {
        return (time < DateTime.rolloverHour) ? time + 1 : time;
    }
    
    /**
     * Charge les paramètres des dates et heures
     *  - heure de changement de journée
     *  à partir de la feuille "Param".
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
     * Renvoie l'objet Day correspondant au numéro de jour fourni,
     * en chargeant les paramètres des jours de la semaine si ce n'est pas déjà fait.
     * @param {string} n Numéro du jour de la semaine (de 1 : lundi à 6 : samedi, 0 ou 7 : dimanche).
     * @returns {Day} Objet Day correspondant au numéro de jour fourni.
     */
    private static getDayConst(n: string): Day {
        // if (!Day.loaded) Day.load();
        return Day.daysByNumbers.get(n)!;
    }

    /**
     * Accesseurs des jours de la semaine
     */
    public static get MONDAY(): Day { return Day.getDayConst('1'); }
    public static get TUESDAY(): Day { return Day.getDayConst('2'); }
    public static get WEDNESDAY(): Day { return Day.getDayConst('3'); }
    public static get THURSDAY(): Day { return Day.getDayConst('4'); }
    public static get FRIDAY(): Day { return Day.getDayConst('5'); }
    public static get SATURDAY(): Day { return Day.getDayConst('6'); }
    public static get SUNDAY(): Day { return Day.getDayConst('7'); }

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
     * Les jours sont stockés dans la structure Day.daysByNumbers sous forme de map, avec
     *  comme clé le nom complet et l'abréviation du jour, et comme valeur leur numéro correspondant.
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
    public readonly value: number;                          // Valeur de la parité
    private readonly doubleParityAllowed: boolean;          // Autorise une double parité

    /**
     * Constructeur privé de la classe Parity.
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
                return Parity.even(this.doubleParityAllowed);
            case Parity.EVEN:
                return Parity.odd(this.doubleParityAllowed);
            case Parity.DOUBLE:
            case Parity.UNDEFINED:
            default:
                return this;
        }
    }
    
    /**
     * Combine une parité avec une autre en les aditionnant.
     * Si la parité de départ n'autorise pas les parités doubles, il est impossible de combiner
     * cette parité avec une autre. Le résultat est forcément une parité qui accepte les parités doubles.
     * Si la parité de départ n'est pas définie, on utilise la parité fournie en paramètre.
     * Si la parité fournie en paramètre n'est pas définie, on utilise la parité de départ.
     * Si les deux parités sont identiques, on retourne la parité de départ.
     * Sinon, on combine ces deux parités en une parité double.
     * @param {Parity} other Parité à combiner avec la parité actuelle.
     * @returns {Parity} Parité combinée.
     */
    public combineWith(other: Parity): Parity {

        if (!this.doubleParityAllowed) throw new Error(`Il n'est pas possible de combiner une`
            + ` parité à une autre si celle de départ n'autorise pas les parités doubles.`
            + ` Le résultat est forcément une parité qui accepte les parités doubles.`);
    
        if (!this.isDefined()) return other.isDefined()
            ? Parity.from(other.value, true)
            : this;

        if (!other.isDefined() || this.value === other.value) {
            return this;
        }
        
        return Parity.double();
    }    
    
    /**
     * Retourne le chiffre de parité correspondant.
     * Si withUnderscores est vrai, le chiffre est précédé
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
     * Crée une parité qui n'a pas de valeur définie.
     * @param {boolean} [doubleParityAllowed=false] Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} La parité sans valeur définie.
     */
    public static undefined(doubleParityAllowed: boolean = false): Parity {
        return new Parity(Parity.UNDEFINED, doubleParityAllowed);
    }

    /**
     * Crée une parité qui correspond à une parité impaire.
     * @param {boolean} [doubleParityAllowed=false] Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} La parité impaire.
     */
    public static odd(doubleParityAllowed: boolean = false): Parity {
        return new Parity(Parity.ODD, doubleParityAllowed);
    }

    /**
     * Crée une parité qui correspond à une parité paire.
     * @param {boolean} [doubleParityAllowed=false] Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} La parité paire.
     */
    public static even(doubleParityAllowed: boolean = false): Parity {
        return new Parity(Parity.EVEN, doubleParityAllowed);
    }

    /**
     * Crée une parité qui correspond à une parité double.
     * Elle est représentée par le chiffre -2.
     * @returns {Parity} La parité double.
     */
    public static double(): Parity {
        return new Parity(Parity.DOUBLE, true);
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
     * Charge les paramètres de parité des jours :
     *  - lettres et chiffres associés
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
     * Charge les paramètres des numéros de train
     *  - regex des numéros de train W,
     *  - regex des numéros de train abrégeables à 4 chiffres.
     *  à partir de la feuille "Param".
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
    public referenceStation: Station | null;                // Gare de rattachement
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
     * Les gares sont stockées dans une map avec comme clé l'abréviation 
     *  de la gare et comme valeur l'objet Station.
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
     * Sauvegardeles stations de la map dans un tableau.
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
     * Constructeur privé de la classe StationWithParity.
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
                parity: Parity.undefined()
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

    // Propriétés de l'objet Connection
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
        this._time = DateTime.from(time, true) ?? DateTime.from(Connection.DEFAULT_CONNECTION_TIME, true)!;
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
        if (!timeObj) {
            throw new Error(
                `Le temps de trajet de la connexion ${this.from} -> ${this.to}`
                + ` est invalide.`
            );
        }
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
     * Les connexions sont stockées dans une map à deux niveaux avec comme première clé la gare de départ,
     *  comme deuxième clé la gare d'arrivée et comme valeur l'objet Connection.
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
     * Sauvegardeles connexions entre les gares dans un tableau.
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
    private _withTurnaround: boolean = false;       // Arrêt avec rebroussement
    private _arrivalTime?: DateTime;                // Temps / Heure d'arrivée de l'arrêt
    private _departureTime?: DateTime;              // Temps / Heure de départ de l'arrêt
    private _passageTime?: DateTime;                // Temps / Heure de passage à l'arrêt (sans arrêt)
    private _tracks: string[];                      // Voies de l'arrêt
    
    /**
     * Constructeur d'un arrêt.
     * @param {StationWithParity | Station | string} station - Gare de l'arrêt.
     * @param {StationWithParity | string} [stationAfterTurnaround] - Gare de rebroussement.
     * @param {DateTime | number | string} [arrivalTime] - Temps / Heure d'arrivée de l'arrêt.
     * @param {DateTime | number | string} [departureTime] - Temps / Heure de départ de l'arrêt.
     * @param {DateTime | number | string} [passageTime] - Temps / Heure de passage à l'arrêt (sans arrêt).
     * @param {boolean} [areRelativeTimes=false] - Indique si les horaires sont relatives (par exemple, par rapport à un autre arrêt).
     * @param {string[] | string} [tracks=[]] - Voies de l'arrêt.
     */
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
     * Renvoie une clé unique pour l'arrêt, composée du nom de la gare et de la parité (si connue).
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
        return this.station!.station.abbreviation;
    }

    /**
     * Renvoie vrai si l'arrêt à un rebroussement possible, faux sinon.
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
     * @returns {boolean} Vrai si le rebroussement est possible, faux sinon.
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
     * Indique si l'arrêt est un arrêt intermédiaire,
     * avec une heure d'arrivée et une heure de départ, ou une heure de passage.
     * @returns {boolean} Vrai si l'arrêt est un arrêt intermédiaire, faux sinon.
     */
    public isIntermediateStop(): boolean {
        return (!!this._arrivalTime && !!this._departureTime) || !!this._passageTime;
    }

    /**
     * Convertit les heures d'arrivée, de départ et de passage
     *  en temps relatifs par rapport à une référence.
     * Lève une erreur si le temps de référence est déjà relatif.
     * Lève un avertissement si les temps à convertir sont déjà relatifs.
     * Cependant pas d'erreur levée si le temps de référence et toutes les heures sont déjà relatives.
     * @param {DateTime} reference Référence à utiliser pour convertir les heures.
     */
    public convertToRelativeTime(reference: DateTime, throwErrorIfAlreadyRelative: boolean = false): void {
        
        // Temps de référence déjà relatif : pas de conversion possible
        // Vérifie simplement que les temps soient déjà relatifs
        if (reference.isRelative){
            const arrivalTimeIsAbsolute = this._arrivalTime && !this._arrivalTime.isRelative;
            const departureTimeIsAbsolute = this._departureTime && !this._departureTime.isRelative;
            const passageTimeIsAbsolute = this._passageTime && !this._passageTime.isRelative;

            if (arrivalTimeIsAbsolute || departureTimeIsAbsolute || passageTimeIsAbsolute) {
                if (throwErrorIfAlreadyRelative) {
                    throw new Error(`Le temps de référence`
                        + ` ${reference.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                        + ` est déjà relatif. Les horaires de l'arrêt ${this.key} qui sont absolus`
                        + ` ne peuvent donc pas être convertis en temps relatifs.`);
                }
            }
            return;
        }

        // Temps de référence absolu : conversion possible
        // Vérifie si les temps sont bien absolus avant de les convertir
        if (this._arrivalTime) {
            if (this._arrivalTime.isRelative) {
                Log.warn(`L'heure d'arrivée à l'arrêt ${this.key}`
                    + ` ${this._arrivalTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` est déjà relative. Elle ne sera donc pas convertie.`);
            } else {
                this._arrivalTime = this._arrivalTime.relativeTo(reference);
            }
        }
        if (this._departureTime) {
            if (this._departureTime.isRelative) {
                Log.warn(`L'heure de départ à l'arrêt ${this.key}`
                    + ` ${this._departureTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` est déjà relative. Elle ne sera donc pas convertie.`);
            } else {
                this._departureTime = this._departureTime.relativeTo(reference);
            }
        }
        if (this._passageTime) {
            if (this._passageTime.isRelative) {
                Log.warn(`L'heure de passage à l'arrêt ${this.key}`
                    + ` ${this._passageTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` est déjà relative. Elle ne sera donc pas convertie.`);
            } else {
                this._passageTime = this._passageTime.relativeTo(reference);
            }
        }
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
     * Les arrêts sont stockés dans la propriété "stops" des parcours correspondants.
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

                // Instancie l'objet Stop
                const stop = new Stop(
                    station,
                    stationAfterTurnaround,
                    arrivalTime,
                    departureTime,
                    passageTime,
                    true,
                    tracks
                );

                // Ajoute l'arrêt au parcours
                const path = Paths.map.get(pathKey);
                if (!path) {
                    throw new Error(`Parcours "${pathKey}" inexistant.`);
                }
                path.stops.push(stop);

            } catch (e) {
                Log.warn(`Stops.load (ligne ${excelRow}) : ${e}`);
                continue;
            }
        }
    }
    
    /**
     * Sauvegardeles arrêts des trains dans un tableau.
     * Les données sont celles stockées dans les objets Path de la map Paths.map.
     * @param {string} [sheetName=Stops.SHEET] Nom de la feuille de calcul.
     * @param {string} [tableName=Stops.TABLE] Nom du tableau.
     * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = Stops.SHEET,
        tableName: string = Stops.TABLE,
        startCell: string = "A1"
    ): void {

        // Crée le tableau final avec les données de chaque arrêt pour chaque train
        const data: (string | number)[][] = [];
    
        for (const path of Paths.map.values()) {
            for (const [stationName, stop] of path.stops.entries()) {
                data.push([
                    path.key,
                    stop.key,
                    stop.stationAfterTurnaround ? stop.stationAfterTurnaround.key : "",
                    stop.arrivalTime ? stop.arrivalTime.excelValue : "",
                    stop.departureTime ? stop.departureTime.excelValue : "",
                    stop.passageTime ? stop.passageTime.excelValue : "",
                    stop.tracks.join(";"),
                    path.nextStop(stop.key) ? path.nextStop(stop.key)!.key : ""
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
                const trainNumber = WorkbookService.getString(row, Stops.COL_IMPORT_TRAIN_NUMBER) || "";
                const date = WorkbookService.getNumber(row, Stops.COL_IMPORT_DATE) || 0;
                const service = WorkbookService.getString(row, Stops.COL_IMPORT_SERVICE) || "";
                const days = WorkbookService.getString(row, Stops.COL_IMPORT_DAYS) || "";
                
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

                // Instancie l'objet Stop
                const stop = new Stop(
                    station,
                    stationAfterTurnaround,
                    arrivalTime,
                    departureTime,
                    passageTime,
                    true,
                    tracks
                );

            } catch (e) {
                Log.warn(`Stops.load (ligne ${excelRow}) : ${e}`);
                continue;
            }
        }

    }
}

/**
 * Classe Path définissant le parcours d'un train, avec ses gares et temps de passage
 *  par rapport à la gare origine
 */
class Path {

    // Résultats de la vérification du parcours
    public static readonly  UNCHECKED = 0;          // Parcours non vérifié
    public static readonly  ONLY_FROM_AND_TO = 1;   // Parcours avec uniquement les gares origine et destination
    public static readonly  WITH_VIA_STOPS = 2;     // Parcours avec gares intermédiaires
    public static readonly  FULL_PATH = 3;          // Parcours complet calculé par chainage de connexions 
    public static readonly  ERROR_WITH_STOPS = -1;  // Parcours avec erreur
    
    // Propriétés de l'objet Path
    public key: string;                             // Clé du parcours
    public parity: Parity;                          // Parité du parcours
                                                    //  (synthèse des parités pour chaque gare)
    public lineDirection: Parity;                   // Direction du parcours sur la ligne
                                                    //  (donnée par une parité globale)
    public missionCode: string;                     // Code de mission des trains du parcours (facultatif)
    public name: string;                            // Nom du parcours (facultatif)
    private _signature: string;                     // Signature du parcours : gares définissant le parcours
                                                    //  séparées par '>' pour les arrêts ordonnés et
                                                    //  par ';' si leur ordre de parcours est laissé libre
    /** cache lazy du tableau */
    private _routeStations?: string[][];                                                
    public stops: Stop[] = [];                      // Gares d'arrêt ou gares de passage du parcours
    private _stopsIndex: Map<string, Stop> = new Map();   // Dictionnaire des arrêts référencés
                                                    //  par leur clé (abbréviation_parité)
    private _stopPosition: Map<string, number> = new Map();    // Dictionnaire de la position des arrêts
                                                    //  dans le parcours (référencés par leur clé)
    public stopsChecked: number = Path.UNCHECKED;   // Résultat de la vérification du parcours
                                                    //  (0 si non vérifié)

    /**
     * Constructeur d'un parcours.
     * @param {string} [key=""] Clé du parcours
     * @param {Parity|string/number} [parityValue=Parity.UNDEFINED] Parité du parcours
     * @param {Parity|string/number} [lineDirection=Parity.UNDEFINED] Direction du parcours sur la ligne
     * @param {string} [missionCode=""] Code de mission des trains du parcours
     * @param {string} [name=""] Nom du parcours
     * @param {string} [signature=""] Signature du parcours : gares définissant le parcours
     * @param {number} [stopsChecked=Path.UNCHECKED] Résultat de la vérification du parcours
     */
    constructor(
        key: string = "",
        parityValue: Parity | string | number = Parity.UNDEFINED,
        lineDirection: Parity | string | number = Parity.UNDEFINED,
        missionCode: string = "",
        name: string = "",
        signature: string = "",
        stopsChecked: number = Path.UNCHECKED
    ) {
        this.key = key;
        this.parity = Parity.from(parityValue, true);
        this.lineDirection = Parity.from(lineDirection, true);
        this.missionCode = missionCode;
        this.name = name;
        this._signature = signature;
        this.stopsChecked = stopsChecked;
    }

    /**
     * Crée un parcours Path à partir des gares d'origine et de destination,
     *  ainsi que de leur heures de départ et d'arrivée.
     * @param {string} from - Nom de la gare d'origine
     * @param {DateTime} departureTime - Heure de départ à la gare d'origine
     * @param {string} to - Nom de la gare de destination
     * @param {DateTime} arrivalTime - Heure d'arrivée à la gare de destination
     * @param {string} [missionCode=""] - Code de mission des trains du parcours (facultatif)
     * @param {string} [name=""] - Nom du parcours (facultatif)
     * @returns {Path} - Un objet Path représentant le parcours
     */
    public static fromTerminals(
        from: string,
        departureTime: DateTime,
        to: string,
        arrivalTime: DateTime,
        missionCode?: string,
        name?: string
    ): Path {

        const path = new Path("", undefined, undefined, missionCode, name);
    
        const s1 = new Stop(from, undefined, undefined, departureTime, undefined, departureTime.isRelative);  
        const s2 = new Stop(to, undefined, arrivalTime, undefined, undefined, arrivalTime.isRelative);

        path.stops = [s1, s2];
        path.buildSignatureFromStops();
        path.rebuildStopIndex();
        path.rebuildStopPosition();
    
        path.stopsChecked = Path.ONLY_FROM_AND_TO;
    
        return path;
    }


    /**
     * Renvoie l'arrêt d'origine du parcours.
     * @returns {Stop | undefined} L'arrêt d'origine, ou undefined si le parcours n'a pas d'arrêt.
     */
    public get origin(): Stop | undefined {
        return this.stops[0];
    }

    /**
     * Renvoie l'arrêt de destination du parcours.
     * @returns {Stop | undefined} L'arrêt de destination, ou undefined si le parcours n'a pas d'arrêt de destination.
     */
    public get destination(): Stop | undefined {
        return this.stops.at(-1);
    }

    /**
     * Renvoie la signature du parcours, qui est la concaténation
     *  des noms des gares d'arrêt du parcours, précédés de "@"
     *  si l'ordre de passage n'est pas imposé.
     * @returns {string} La signature du parcours
     */
    public get signature(): string {
        return this._signature;
    }

    /**
     * Renvoie le tableau des gares d'arrêt du parcours.
     * Le tableau est construit à partir de la signature du parcours.
     * Chaque élément du tableau est ordonné et correspond à une gare d'arrêt du parcours,
     *  ou à un groupe de gares à parcourir dans un ordre indifférent, séparées par un ";".
     *  Chaque gare ou ensemble de gares est parcouru dans l'ordre du tableau, et séparé par un ">".
     * @returns {string[][]} Le tableau des gares d'arrêt du parcours
     */
    public get routeStations(): string[][] {

        if (!this._routeStations) {
    
            this._routeStations = this._signature
                .split(">")
                .map(group => group.split(";"));
        }
    
        return this._routeStations;
    }

    /**
     * Renvoie le radical de la clé du parcours constitué de
     *  origine_destination_codeMission_nomDuParcours (si ces valeurs existent)
     * @returns {string} Radical de la clé du parcours
     */
    public buildRadical(): string {
        const origin = this.origin?.stationAbbreviation ?? "";
        const dest = this.destination?.stationAbbreviation ?? "";
    
        const parts = [origin, dest];

        if (this.missionCode) parts.push(this.missionCode);
        if (this.name) parts.push(this.name);

        return parts.join("_");
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
        if (this.stopsChecked === Path.FULL_PATH) {
            if (!hasParityDefined) {
                Log.warn(`Le parcours calculé ${this.key} ne doit comporter`
                    + ` que des arrêts avec parité définie.`
                    + ` L'arrêt ${stop.key} ne sera donc pas pris en compte.`);
                return;
            }
            if (this._stopsIndex.has(stop.key)) {
                if (!erase) {
                    Log.warn(`L'arrêt "${stop.key}" est déjà associé aux trains`
                        + ` du parcours ${this.key}. Un même train ne peut pas revenir`
                        + ` dans la même gare et avec le même sens.`
                        + ` Le deuxième arrêt ne sera donc pas pris en compte.`);                
                    return;
                }
                this.stops.splice(this.stops.indexOf(this._stopsIndex.get(stop.key)!), 1);
            }
            // Mise à jour des parités (parité sur l'ensemble du parcours et parité de ligne)
            this.parity = this.parity.combineWith(stop.station.parity);
            this.lineDirection = this.lineDirection.combineWith(
                stop.station.station.reverseLineDirection
                    ? stop.station.parity.invert()
                    : stop.station.parity
            );

        // Le parcours n'a pas été calculé => ne contient pas d'arrêts avec parité
        } else {
            if (hasParityDefined) {
                Log.warn(`Le parcours ${this.key} n'a pas été calculé.`
                    + ` Il ne peut donc pas contenir d'arrêts avec parité.`
                    + ` L'arrêt ${stop.key} ne sera donc pas pris en compte.`);
                return; 
            }
            if (this._stopsIndex.has(stop.key)) {
                if (!erase) {
                    Log.warn(`L'arrêt "${stop.key}" est déjà associé aux trains`
                        + ` du parcours ${this.key}. Si le train dessert une gare dans les deux sens,`
                        + ` il est nécessaire de calculer les parités de passage en gare.`
                        + ` Le deuxième arrêt ne sera donc pas pris en compte.`);
                    return;
                }
                this.stops.splice(this.stops.indexOf(this._stopsIndex.get(stop.key)!), 1);
            }
        }

        // Ajout dans le tableau des arrêts
        this.stops.push(stop);
        this.orderStops();
        this._stopsIndex.set(stop.key, stop);
        if (!!stationAfterTurnaround && hasParityDefined) {
            this._stopsIndex.set(stationAfterTurnaround.key, stop);
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
        this.rebuildStopPosition();
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
        if (this.stopsChecked === Path.FULL_PATH) {
            if (stationObj.parity.isDefined()) {
                return this._stopsIndex.get(stationObj.key) ?? undefined;
            }
            const oddStop = this._stopsIndex.get(StationWithParity.from(stationObj, Parity.ODD).key);
            const evenStop = this._stopsIndex.get(StationWithParity.from(stationObj, Parity.EVEN).key);
            if (oddStop && evenStop) {
                const firstStop = oddStop.getTime()!.compareTo(evenStop.getTime()!) < 0 ? oddStop : evenStop;
                Log.warn(`Le parcours ${this.key} a un arrêt dans chaque sens dans la gare ${stationObj.key}.`
                    + ` C'est le premier arrêt ${firstStop.key} qui est renvoyé.`);
                return firstStop;
            }
            return oddStop ?? evenStop ?? undefined;
        }

        // Le parcours n'a pas été calculé => ne contient pas d'arrêts avec parité
        return this._stopsIndex.get(stationObj.key) ?? undefined;
    }

    /**
     * Retourne l'arrêt suivant de la gare spécifiée.
     * Si la gare spécifiée est la dernière de la liste, renvoie undefined.
     * @param {StationWithParity | Station | string} station - La gare à chercher
     * @returns {Stop | undefined} - L'arrêt suivant, ou undefined si la gare est la dernière
     */
    public nextStop(
        station: StationWithParity | Station | string
    ): Stop | undefined {
        
        const stop = this.getStop(station);
        if (!stop) return undefined;
    
        const index = this._stopPosition.get(stop.key);
        if (index === undefined || index === this.stops.length - 1) return undefined;
    
        return this.stops[index + 1];
    }

    /**
     * Retourne l'arrêt précédent de la gare spécifiée.
     * Si la gare spécifiée est la première de la liste, renvoie undefined.
     * @param {StationWithParity | Station | string} station - La gare à chercher
     * @returns {Stop | undefined} - L'arrêt précédent, ou undefined si la gare est la première
     */
    public previousStop(
        station: StationWithParity | Station | string
    ): Stop | undefined {
    
        const stop = this.getStop(station);
        if (!stop) return undefined;
    
        const index = this._stopPosition.get(stop.key);
        if (index === undefined || index === 0) return undefined;
    
        return this.stops[index - 1];
    }

    /**
     * Efface la liste des arrêts du train.
     * Supprime également les valeurs de firstStop et lastStop.
     */
    public eraseStops() {
        this.stops = [];
        this.stopsChecked = 0;
        this._stopsIndex.clear();
        this._stopPosition.clear();
    }

    /**
     * Vérifie si deux parcours ont les mêmes arrêts.
     * Les arrêts sont comparés en fonction de leur gare et de leur heure de passage.
     * @param {Path} other Le parcours à comparer.
     * @returns {boolean} Vrai si les deux parcours ont les mêmes arrêts, faux sinon.
     */
    public equalsStops(other: Path): boolean {
        if (this.stops.length !== other.stops.length) return false;

        for (let i = 0; i < this.stops.length; i++) {
            if (!this.stops[i].equalsTo(other.getStop(this.stops[i].station))) return false;
        }
        return true;
    }

    /**
     * Convertit les heures d'arrivée, de départ et de passage des arrêts
     *  en temps relatifs par rapport à l'heure de départ du premier arrêt.
     *  Si un arrêt a déjà un horaire relatif, une erreur est levée.
     */
    public convertStopsToRelative(): void {
        if (this.stops.length === 0) return;

        const t0 = this.stops[0].departureTime;
        if (!t0) throw new Error(`Le premier arrêt du parcours ${this.key}`
            + ` n'a pas d'heure de départ. Les horaires ne peuvent donc pas`
            + ` être convertis en horaires relatifs.`);
        for (const stop of this.stops) {
            stop.convertToRelativeTime(t0);
        }
    }

    /**
     * Construit un index des arrêts en fonction de leur clés.
     * Les clés sont utilisées pour accéder rapidement à un arrêt.
     * L'index est mis à jour automatiquement lorsque la liste des arrêts change.
     */
    private rebuildStopIndex(): void {
        this._stopsIndex.clear();
        for (const stop of this.stops) {
            this._stopsIndex.set(stop.key, stop);
            if (!!stop.stationAfterTurnaround && stop.station.parity.isDefined()) {
                this._stopsIndex.set(stop.stationAfterTurnaround.key, stop);
            }
        }
    }

    /**
     * Reconstruit l'index des arrêts en fonction de leur position dans le parcours.
     * Les clés sont utilisées pour accéder rapidement à la position d'un arrêt.
     * L'index est mis à jour automatiquement lorsque la liste des arrêts change.
     */
    private rebuildStopPosition(): void {
        this._stopPosition.clear();
        for (let i = 0; i < this.stops.length; i++) {
            const stop = this.stops[i];
            this._stopPosition.set(stop.key, i);
            if (!!stop.stationAfterTurnaround && stop.station.parity.isDefined()) {
                this._stopPosition.set(stop.stationAfterTurnaround.key, i);
            }
        }
    }

    /**
     * Construit la signature du parcours en fonction de la liste des arrêts.
     * La signature est une chaîne de caractères qui identifie de manière unique
     * le parcours. Elle est utilisée pour chercher les connexions entre les
     * différents parcours.
     */
    public buildSignatureFromStops(): void {
        this._signature = this.stops
            .map(s => s.key).join(">");
        this._routeStations = undefined;
    }

    // ===== connexions depuis signature =====
    public buildConnectionsFromStops(): Connection[] {

        const connections: Connection[] = [];
    
        if (this.stops.length < 2 || this.stopsChecked !== Path.FULL_PATH) {
            return connections;
        }
    
        for (let i = 0; i < this.stops.length - 1; i++) {
    
            const fromStop = this.stops[i];
            const fromStation = fromStop.station;
            const fromStationAfterTurnaround = fromStop.stationAfterTurnaround;
            const toStation = this.stops[i + 1].station;

            if (fromStationAfterTurnaround) {

                const connection1 = Connections.get(fromStation.key, fromStationAfterTurnaround.key);
                if (!connection1) throw new Error(`Connection introuvable`
                    + ` entre ${fromStation.key} et ${fromStationAfterTurnaround.key}`);
                connections.push(connection1);
    
                const connection2 = Connections.get(fromStationAfterTurnaround.key, toStation.key);
                if (!connection2) throw new Error(`Connection introuvable`
                    + ` entre ${fromStationAfterTurnaround.key} et ${toStation.key}`);
                connections.push(connection2);

            } else {

                const connection = Connections.get(fromStation.key, toStation.key);
                if (!connection) throw new Error(`Connection introuvable`
                    + ` entre ${fromStation.key} et ${toStation.key}`);
                connections.push(connection);
            }
        }
    
        return connections;
    }

    // ===== stops depuis connexions =====
    public buildStopsFromConnections(connexions: Connection[]): void {

        this.stops = [];

        let currentTime = 0;

        for (const c of connexions) {
            const stop = new Stop(c.fromStation);
            stop.setRelativeTime(currentTime);
            this.stops.push(stop);

            currentTime += c.duration;
        }

        const last = connexions[connexions.length - 1];
        if (last) {
            const stop = new Stop(last.toStation);
            stop.setRelativeTime(currentTime);
            this.stops.push(stop);
        }
        
        this.rebuildStopIndex();
        this.rebuildStopPosition();
    }

    public check(): void {
  
        // Pas de vérification si le parcours a une erreur
        switch (this.stopsChecked) {
            case Path.ERROR_WITH_STOPS:
            case Path.UNCHECKED:
                return;
        }
        
        try {

            this.checkTerminals();

            this.checkSignature();

            // Test valide si parcours avec gares origine et destination uniquement
            if (this.stopsChecked === Path.ONLY_FROM_AND_TO) {
                return;
            }

            this.checkTimes();

            // Test valide si parcours avec gares intermédiaires non calculé
            if (this.stopsChecked === Path.WITH_VIA_STOPS) {
                return;
            }

            this.checkConnections();
            return;

        } catch (e) {
            this.stopsChecked = Path.ERROR_WITH_STOPS;
            throw new Error(`Parcours ${this.key} : ${e}`);
        }
    }

    /**
     * Vérifie les gares et horaires de départ et d'arrivée.
     * @throws {Error} Si une erreur est détectée
     */
    private checkTerminals(): void {
        
        // Vérifie l'existence d'une gare de départ
        const firstStop = this.stops[0];
        if (!firstStop) {
            throw new Error(`Il n'y a pas de gare de départ.`);
        }
        // Vérifie l'existence d'une gare d'arrivée
        const lastStop = this.stops[this.stops.length - 1];
        if (!lastStop) {
            throw new Error(`Il n'y a pas de gare d'arrivée.`);
        }
        // Vérifie l'existence d'une heure de départ
        const departureTime = this.stops[0].departureTime;
        if (!departureTime) {
            throw new Error(`Le premier arrêt n'a pas d'heure de départ.`);
        }
        // Vérifie l'existence d'une heure d'arrivée
        const arrivalTime = this.stops[this.stops.length - 1].arrivalTime;
        if (!arrivalTime) {
            throw new Error(`Le dernier arrêt n'a pas d'heure d'arrivée.`);
        }
        // Vérifie l'absence d'heure d'arrivée dans le premier arrêt
        if (firstStop.isIntermediateStop()) {
            throw new Error(`Le premier arrêt ne peut pas contenir d'heure d'arrivée`
                + ` mais uniquement une heure de départ.`);
        }
        // Vérifie l'absence d'heure de départ dans le dernier arrêt
        if (lastStop.isIntermediateStop()) {
            throw new Error(`Le dernier arrêt ne peut pas contenir d'heure de départ`
                + ` mais uniquement une heure d'arrivée.`);
        }
        // Vérifie la concordance entre les heures de départ et d'arrivée
        //  (toutes deux absolues ou relatives)
        if (arrivalTime.isRelative !== departureTime.isRelative) {
            throw new Error(`Les deux heures de départ et d'arrivée`
                + ` doivent être toutes deux absolues ou relatives.`);
        }
        // Vérifie que l'heure de départ est nulle si relative
        //  (l'heure de départ est une référence pour la suite du parcours)
        if (departureTime.isRelative && departureTime.excelValue !== 0) {
            throw new Error(`Une heure de départ relative doit avoir pour valeur 0.`);
        }
        // Vérifie que l'heure d'arrivée est postérieure à l'heure de départ
        if (arrivalTime.compareTo(departureTime) <= 0) {
            throw new Error(`L'heure d'arrivée ${arrivalTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                + ` doit être supérieure`
                + ` à l'heure de départ ${departureTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}.`);
        }
        // Vérifie que le départ ne contient pas d'arrêt après retournement
        if (!!firstStop.stationAfterTurnaround) {
            throw new Error(`L'heure de départ ${departureTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                + ` ne doit pas contenir d'arrêt aprés retournement.`);
        }
    }

    /**
     * Vérifie la signature, et la présence des gares de départ et d'arrivée.
     * @throws {Error} Si une erreur est détectée
     */
    private checkSignature() {
        
        // Vérifie l'existance de la signature, ou la constitue si inexistante
        //  dans le cas où le parcours n'a pas été calculé
        const sigStations = this.routeStations;
        if (!sigStations) {
            switch (this.stopsChecked) {
                case Path.ONLY_FROM_AND_TO:
                case Path.WITH_VIA_STOPS:
                    this.buildSignatureFromStops();
                    break
                case Path.FULL_PATH:
                    throw new Error(`Il n'y a pas de signature.`);
            }
        }
        // Vérifie que la signature contient au moins la gare de départ et d'arrivée
        if (sigStations.length < 2) {
            throw new Error(`La signature  ${this.signature} ne comporte pas au minimum une gare départ`
                + ` et une gare d'arrivée, séparées par '>'.`);
        }
        // Vérifie que la liste des arrêts contient au moins la gare de départ et d'arrivée
        if (this.stops.length < 2) {
            throw new Error(`Le parcours doit comporter au moins 2 arrêts`
                + ` à la gare de départ et la gare d'arrivée.`);
        }
        // Vérifie que la gare de départ est isolée (ne peut pas être dans un ordre quelconque avec d'autres gares)
        if (sigStations[0].length !== 1){
            throw new Error(`La gare de départ dans la signature ${this.signature}`
                + ` est forcément suivie de '>'.`);
        }
        // Vérifie que la gare de départ correspond à la première gare de la signature
        if (sigStations[0][0].split('_')[0] !== this.stops[0].stationAbbreviation){
            throw new Error(`La gare de départ dans la signature ${sigStations[0][0]}`
                + ` ne correspond pas à la première gare du parcours ${this.stops[0].key}.`);
        }
        // Vérifie que la gare d'arrivée est isolée (ne peut pas être dans un ordre quelconque avec d'autres gares)
        if (sigStations[sigStations.length - 1].length !== 1){
            throw new Error(`La gare d'arrivée dans la signature ${this.signature}`
                + ` est forcément précédée de '>'.`);
        }
        // Vérifie que la gare d'arrivée correspond à la dernière gare de la signature
        if (sigStations[sigStations.length - 1][0].split('_')[0] !== this.stops[this.stops.length - 1].stationAbbreviation){
            throw new Error(`La gare d'arrivée dans la signature ${sigStations[sigStations.length - 1][0]}`
                + ` ne correspond pas à la dernière gare du parcours ${this.stops[this.stops.length - 1].key}.`);
        }
    }
    
    /**
     * Vérifie que tous les arrêts intermédiaires sont corrects, en vérifiant
     * que les heures de passage sont concordantes et que les gares intermédiaires
     * correspondent aux gares de la signature.
     * @throws {Error} Si une erreur est détectée
     */
    private checkTimes() {

        const areTimesRelative = this.stops[0].getTime()!.isRelative;
        const sigStations = this.routeStations;
        let j = 1;
        let stopFromSigToFind = new Map<string, string>();

        for (let i = 1; i < this.stops.length; i++) {
            switch (this.stopsChecked) {
                case Path.ERROR_WITH_STOPS:
                case Path.UNCHECKED:
                    return;
                case Path.ONLY_FROM_AND_TO:
                case Path.WITH_VIA_STOPS:
                    // Parcours non calculé : tous les arrêts de la liste des arrêts du parcours 
                    //  doivent être présents dans la signature, sans être dans des groupes d'arrêts
                    //  non ordonnés (séparées par '>'). Les arrêts non ordonnés ne sont pris en compte
                    //  que dans le calcul du parcours. Les arrêts de la signature non trouvés
                    //  ou faisant partie d'un ensemble d'arrêts sont sautés
                    //  jusqu'à trouver dans la signature l'arrêt en cours.
                    while (sigStations[j].length !== 1
                        || sigStations[j][0].split('_')[0] !== this.stops[i].stationAbbreviation) {
                        j++;
                        if (j >= sigStations.length) {
                            throw new Error(`La gare ${this.stops[i].stationAbbreviation}`
                                + `n'est pas reprise dans la signature.`);
                        }
                    }
                    break;
                case Path.FULL_PATH:
                    // Parcours calculé : tous les arrêts de la signature doivent être présents
                    //  dans la liste des arrêts du parcours. Chaque arrêt ou ensemble d'arrêts
                    //  non ordonnés (séparées par ';') sont ajoutés dans un cache stopFromSigToFind,
                    //  dont tous les arrêts doivent être trouvés avant de passer au (groupe) suivant                 
                    if (stopFromSigToFind.size === 0) {
                        sigStations[j].reduce((map, value) => {
                            map.set(value, value);
                            return map;
                        }, stopFromSigToFind);
                    }
                    if (stopFromSigToFind.has(this.stops[i].key)) {
                        stopFromSigToFind.delete(this.stops[i].key);
                        if (stopFromSigToFind.size === 0) j++;
                    }
                    break;
            }
            // Vérifie si l'arrêt comporte des horaires (arrivée, départ ou passage)
            const stopTime = this.stops[i].getTime();
            if (!stopTime) {
                throw new Error(`L'heure de passage à la gare de ${this.stops[i].key}`
                    + ` n'est pas renseignée.`);
            }
            // Vérifie la concordance des horaires (tous absolus ou relatives)
            if (stopTime.isRelative !== areTimesRelative) {
                throw new Error(`L'heure de passage à la gare de ${this.stops[i].key}`
                    + ` doit être ${areTimesRelative ? "relative" : "absolue"}`
                    + ` comme la gare origine.`);
            }
            // Vérifie que l'arrêt est une gare intermédiaire
            if ((i < this.stops.length - 1) || !this.stops[i].isIntermediateStop()) {
                throw new Error(`L'arrêt à la gare de ${this.stops[i].key} doit comporter`
                    + ` une heure d'arrivée et une heure de départ, ou une heure de passage.`);
            }
            // Vérifie que l'heure de passage est postérieure au passage précedent
            if (this.stops[i].getTime()!.compareTo(this.stops[i - 1].getTime(true)!) <= 0) {
                throw new Error(`L'heure d'arrivée ou de passage`
                    + ` ${this.stops[i].getTime()!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` à la gare de ${this.stops[i].key}`
                    + ` doit être postérieure à l'heure de passage ou de départ`
                    + ` ${this.stops[i - 1].getTime(true)!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` à la gare de ${this.stops[i - 1].key}.`);
            }

        }
        // Vérifie que tous les arrêts de la signature ont été trouvés
        if (this.stopsChecked === Path.FULL_PATH && j < sigStations.length) {
            throw new Error(`L'arrêt ${sigStations[j]} dans la signature n'a pas été trouvé dans la liste des arrêts.`);
        }
    }

    /**
     * Vérifie si une connexion existe entre chaque gare de la liste des arrêts
     * @throws {Error} Si une connexion est inexistante
     */
    private checkConnections() {

        for (let i = 1; i < this.stops.length; i++) {

            // Vérifie si une connexion existe entre la gare précédente et la gare actuelle
            if (this.stopsChecked === Path.FULL_PATH) {
                const lastStop = this.stops[i - 1].stationAfterTurnaround
                    ? this.stops[i - 1].stationAfterTurnaround?.key
                    : this.stops[i - 1].key;
                if (!Connections.has(lastStop!, this.stops[i].key)) {
                    throw new Error(`Il n'y a pas de connexion`
                        + ` entre la gare ${lastStop} et la gare ${this.stops[i].key}.`);
                }
            }
        }
    }

    // ===== findPath avec cache signature =====
    public findPath(): void {

        let connections: Connection[];
    
        const ref = Paths.signatureIndex.get(this.signature);
    
        if (ref) {
            connections = ref.buildConnectionsFromStops();
        } else {
            connections = this.shortestPathThrough();
            Paths.signatureIndex.set(this.signature, this);
        }
    
        this.buildStopsFromConnections(connections);
    
        this.interpolateTimes();
    
        this.stopsChecked = Path.FULL_PATH;
    }

    private shortestPathThrough(): Connection[] {

        if (!this.routeStations || this.routeStations.length < 2) {
            throw new Error("RouteStations invalide");
        }

        const connections = Connections.shortestPathWithGroups(
            this.routeStations
        );

        if (!connections.length) {
            throw new Error(`Impossible de calculer le parcours ${this.signature}`);
        }

        return connections;
    }

    private interpolateTimes(): void {
        
    }

    /**
     * Génère toutes les permutations possibles d'un tableau.
     * @param arr Le tableau à permuter.
     * @returns Un tableau contenant toutes les permutations du tableau d'origine.
     * @example
     */
    private static permutations<T>(arr: T[]): T[][] {

        if (arr.length <= 1) return [arr];
    
        const result: T[][] = [];
    
        for (let i = 0; i < arr.length; i++) {
    
            const rest = arr.slice(0, i).concat(arr.slice(i + 1));
    
            for (const p of Path.permutations(rest)) {
                result.push([arr[i], ...p]);
            }
        }
    
        return result;
    }

    /**
     * Génère toutes les permutations possibles de la liste des gares
     * en prenant en compte les groupes de gares intermédiaires.
     * @returns Un tableau contenant toutes les permutations possibles de la liste des gares.
     */
    public expandRoutes(): string[][] {

        let result: string[][] = [[]];
    
        for (const group of this.routeStations) {
    
            const perms = Path.permutations(group);
    
            const newResult: string[][] = [];
    
            for (const base of result) {
                for (const perm of perms) {
                    newResult.push([...base, ...perm]);
                }
            }
    
            result = newResult;
        }

        
    
        return result;
    }
    
  
}

/**
 * Classe Paths contenant la liste des parcours.
 */
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
        "Route",
        "Etat de vérification"
    ]];
    private static readonly COL_KEY = 0;                    // Colonne de la clé du parcours
    private static readonly COL_PARITY = 1;                 // Colonne de la parité du parcours
    private static readonly COL_LINE_PARITY = 2;            // Colonne de la parité de ligne du parcours
    private static readonly COL_MISSION_CODE = 3;           // Colonne du code de mission
    private static readonly COL_NAME = 4;                   // Colonne du nom du parcours
    private static readonly COL_SIGNATURE = 5;              // Colonne de la signature du parcours
    private static readonly COL_STOP_CHECKED = 6;           // Colonne de l'état de vérification du parcours

    // Map des parcours indexés par clé
    public static readonly map: Map<string, Path> = new Map();

    // Map des parcours indexés par radical, puis par suffixe alphabétique, puis par suffixe numérique
    public static readonly structure:
        Map<string, Map<string, Map<number, Path>>> = new Map();

    // Map des parcours indexés par signature pour optimiser le calcul Dijkstra
    public static readonly signatureIndex: Map<string, Path> = new Map();

    /**
     * Accesseurs utilitaires
     */
    // Nombre de parcours
    public static get size(): number {
        return this.map.size;
    }

    /**
     * Renvoie le parcours correspondant à la clé donnée.
     * @param {string} key Clé du parcours.
     * @returns {Path | undefined} Parcours correspondant, ou undefined si la clé n'existe pas.
     */
    public static get(key: string): Path | undefined {
        return this.map.get(key);
    }

    /**
     * Insère un nouveau parcours dans la base de données, avec génération de la clé.
     * @param {Path} path Parcours à insérer.
     * @returns {Path} Parcours inséré avec sa clé.
     */
    public static insert(path: Path): Path {

        const radical = path.buildRadical();
        const signature = path.signature;
    
        let radicalMap = this.structure.get(radical);
    
        // Nouveau radical
        if (!radicalMap) {
            radicalMap = new Map();
            this.structure.set(radical, radicalMap);
    
            const numberMap = new Map<number, Path>();
            numberMap.set(0, path);
    
            // Par convention, le premier parcours d'un radical différent
            //  n'a pas de suffixe lettre => représenté par ""
            radicalMap.set("", numberMap);
    
            path.key = radical;
            this.map.set(path.key, path);
    
            return path;
        }
    
        // Radical existant : recherche de l'existance de la signature
        let letterKey = this.findLetterBySignature(radicalMap, signature);
    
        // Nouvelle signature
        if (letterKey === null) {
            letterKey = this.nextLetter(radicalMap);
    
            const numberMap = new Map<number, Path>();
            numberMap.set(0, path);
    
            radicalMap.set(letterKey, numberMap);
    
            path.key = this.buildKey(radical, letterKey, 0);
            this.map.set(path.key, path);
    
            return path;
        }
    
        // Signature existante : recherche de l'existance d'un parcours identique (mêmes horaires)
        const numberMap = radicalMap.get(letterKey)!;
    
        for (const existing of numberMap.values()) {
            if (existing.equalsStops(path)) {
                return existing;
            }
        }
    
        // Nouveau parcours
        const number = this.nextNumber(numberMap);
    
        numberMap.set(number, path);
    
        path.key = this.buildKey(radical, letterKey, number);
        this.map.set(path.key, path);
    
        return path;
    }

    /**
     * Supprime un parcours de la structure interne.
     * Si le parcours n'existe pas, cette fonction ne fait rien.
     * @param {Path} path Le parcours à supprimer.
     */
    public static delete(path: Path): void {

        const radical = path.buildRadical();
        const letter = Paths.extractLetter(path.key);
        const number = Paths.extractNumber(path.key);
    
        const radicalMap = this.structure.get(radical);
        if (!radicalMap) return;
    
        const numberMap = radicalMap.get(letter);
        if (!numberMap) return;
    
        numberMap.delete(number);
        this.map.delete(path.key);
    
        // nettoyer étage nombre
        if (numberMap.size === 0) {
            radicalMap.delete(letter);
        }
    
        // nettoyer étage lettre
        if (radicalMap.size === 0) {
            this.structure.delete(radical);
        }
    }
    
    /**
     * Cherche le prochain suffixe lettre libre dans la liste des suffixes utilisés.
     * Si un seul élément existe déjà (donc sans suffixe, valeur "" dans la map),
     *  atribue le suffixe "A" à cet élément et au nouvel élément le suffixe "B".
     * Sinon, cherche le premier suffixe lettre non utilisé.
     * Les suffixes lettre sont précédés de "~".
     * @param {Map<number, Path>} numberMap Map des suffixes déjà utilisés.
     * @returns {number} Le prochain suffixe lettre libre dans la map.
     */
    private static nextLetter(
        radicalMap: Map<string, Map<number, Path>>
    ): string {
    
        // Si un seul élément existe déjà (donc sans suffixe), donne à cet élément le suffixe "A"
        // et au nouvel élément le suffixe "B"
        if (radicalMap.size === 1 && radicalMap.has("")) {
    
            const numberMap = radicalMap.get("")!;
            const radical = Paths.extractRadical(numberMap.values().next().value!.key)!;
    
            radicalMap.delete("");
            radicalMap.set("A", numberMap);

            for (const path of numberMap.values()) {
                const number = Paths.extractNumber(path.key);
                this.map.delete(path.key);
                path.key = this.buildKey(radical, "A", number);
                this.map.set(path.key, path);
            }
    
            return "B";
        }
    
        // Si plusieurs éléments existent déjà (donc avec suffixes),
        //  cherche le premier suffixe lettre non utilisé
        const used = new Set(radicalMap.keys());
    
        let index = 0;
    
        while (true) {
            const candidate = this.indexToLetters(index);
            if (!used.has(candidate)) return candidate;
            index++;
        }
    }

    /**
     * Convertit un index en une chaîne de lettres.
     * Par exemple, 0 donnera "A", 1 donnera "B", 25 donnera "Z", 26 donnera "AA", etc.
     * @param {number} index L'index à convertir.
     * @returns {string} La chaîne de lettres correspondante.
     */
    private static indexToLetters(index: number): string {

        let s = "";
        index += 1;
    
        while (index > 0) {
            index--;
            s = String.fromCharCode(65 + (index % 26)) + s;
            index = Math.floor(index / 26);
        }
    
        return s;
    }

    /**
     * Cherche le prochain suffixe numérique libre dans la liste des suffixes utilisés.
     * Si un seul élément existe déjà (donc sans suffixe, valeur 0 dans la map),
     *  atribue le suffixe "1" à cet élément et au nouvel élément le suffixe "2".
     * Sinon, cherche le premier suffixe numérique non utilisé.
     * Les suffixes numériques sont précédés de "#".
     * @param {Map<number, Path>} numberMap Map des suffixes déjà utilisés.
     * @returns {number} Le prochain suffixe numérique libre dans la map.
     */
    private static nextNumber(
        numberMap: Map<number, Path>
    ): number {
    
        // Si un seul élément existe déjà (donc sans suffixe), donne à cet élément le suffixe "1"
        // et au nouvel élément le suffixe "2"
        if (numberMap.size === 1 && numberMap.has(0)) {
    
            const firstPath = numberMap.get(0)!;
    
            numberMap.delete(0);
            numberMap.set(1, firstPath);

            this.map.delete(firstPath.key);
            firstPath.key = firstPath.key + "#1";
            this.map.set(firstPath.key, firstPath)

            return 2;
        }
    
        // Si plusieurs éléments existent déjà (donc avec suffixes),
        // cherche le premier suffixe numérique non utilisé
        let n = 1;
        while (numberMap.has(n)) n++;
    
        return n;
    }

    /**
     * Extrait le radical de la clé d'un parcours
     *  (chaîne de la forme "X~Y#Z" où X est le radical et Y et Z sont des suffixes).
     * @param {string} key Clé du parcours.
     * @returns {string} Radical de la clé (ou une chaîne vide si la clé n'a pas de radical).
     */
    private static extractRadical(key: string): string {
        return key.split("~")[0].split("#")[0];
    }

    /**
     * Extrait la lettre de la clé d'un parcours (chaîne de la forme "~X" où X est la lettre du suffixe).
     * @param {string} key Clé du parcours.
     * @returns {string} Lettre du suffixe (ou une chaîne vide si la clé n'a pas de suffixe lettre).
     */
    private static extractLetter(key: string): string {
        const m = key.match(/~([A-Z]+)/);
        return m ? m[1] : "";
    }
    
    /**
     * Extrait le numéro de la clé d'un parcours (chaîne de la forme "#X" où est le numéro du suffixe numérique).
     * @param {string} key Clé du parcours.
     * @returns {number} Numéro du suffixe numérique (ou 0 si la clé n'a pas de suffixe numérique).
     */
    private static extractNumber(key: string): number {
        const m = key.match(/#(\d+)/);
        return m ? Number(m[1]) : 0;
    }

    /**
     * Construit une clé de parcours à partir d'un radical, d'une lettre de suffixe et d'un numéro de suffixe.
     * La clé est composée de la forme "radical~lettre#nombre" avec
     *  un suffixe lettre optionnel précédé de "~"
     *  et un suffixe numérique optionnel précédé de "#".
     * @param {string} radical Radical de la clé.
     * @param {string} letter Lettre de suffixe (ou une chaîne vide si pas de suffixe lettre).
     * @param {number} number Numéro de suffixe (ou 0 si pas de suffixe numérique).
     * @returns {string} Clé de parcours avec les suffixes appropriés.
     */
    private static buildKey(
        radical: string,
        letter: string,
        number: number
    ): string {
    
        let key = radical;
    
        if (letter) key += `~${letter}`;
        if (number > 0) key += `#${number}`;
    
        return key;
    }

    /**
     * Cherche si un parcours existe déjà avec un même radical et une même signature
     *  Si oui donne le suffixe lettre de ce parcours.
     *  Sinon renvoie null.
     * @param {Map<string, Map<number, Path>>} radicalMap Map des parcours ayant le même radical
     *  que celui du parcours pour lequel la recherche est faite.
     * @param {string} signature Signature du parcours à chercher.
     * @returns {string | null} Lettre du suffixe de la clé du parcours trouvé
     *  (même radical et même signature).
     */
    private static findLetterBySignature(
        radicalMap: Map<string, Map<number, Path>>,
        signature: string
    ): string | null {
    
        for (const [letter, numberMap] of radicalMap.entries()) {
    
            // récupérer un seul Path (le premier)
            const firstPath = numberMap.values().next().value as Path;
    
            if (firstPath.signature === signature) {
                return letter;
            }
        }
    
        return null;
    }

    /**
     * Charge les parcours de trains à partir du tableau "Parcours" de la feuille "Parcours".
     * Les parcours sont stockés dans une map avec comme clé la clé du parcours
     *  et comme valeur l'objet Path.
     * @param {boolean} [erase=false] Si vrai, force le rechargement des parcours.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false) {

        // Vérifie si la table à charger existe déjà
        if (Paths.map.size > 0) {
            if (erase) {
                Paths.map.clear(); // Vide la map sans changer sa référence
                Paths.structure.clear();
                Paths.signatureIndex.clear();
            }
        }

        // Charge les connexions si elles ne sont pas encore chargées
        Connections.load(); // Charge les connexions si elles ne sont pas encore chargées

        // Charge la base de données
        const data = WorkbookService.getDataFromTable(Paths.SHEET, Paths.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Paths.load : aucune donnée trouvée dans la table.`);
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
                const key = WorkbookService.getString(row, Paths.COL_KEY) ?? "";
                const parityLetter = WorkbookService.getString(row, Paths.COL_PARITY) ?? "";
                const lineDirectionLetter = WorkbookService.getString(row, Paths.COL_LINE_PARITY) ?? "";
                const missionCode = WorkbookService.getString(row, Paths.COL_MISSION_CODE) ?? "";
                const name = WorkbookService.getString(row, Paths.COL_NAME) ?? "";
                const signature = WorkbookService.getString(row, Paths.COL_SIGNATURE) ?? "";
                const stopChecked = WorkbookService.getNumber(row, Paths.COL_STOP_CHECKED) ?? 0;

                // Instancie l'objet Station
                const path = new Path(
                    key,
                    parityLetter,
                    lineDirectionLetter,
                    missionCode,
                    name,
                    signature,
                    stopChecked
                );

                // Ajoute l'objet Path dans la map, indexé par sa clé
                if (Paths.map.has(key)) {
                    throw new Error(`Le parcours ${key} est déjà présent`
                        + ` dans la base de données.`);
                } 
                Paths.map.set(key, path);

                // Ajoute l'objet Path dans l'index par signature, si pas encore présent
                //  (parcours calculé uniquement)
                if (path.stopsChecked === Path.FULL_PATH && !Paths.signatureIndex.has(signature)) {
                    Paths.signatureIndex.set(signature, path);
                } 

                // Ajoute l'objet Path dans la structure des radicaux et suffixes
                const radical = Paths.extractRadical(key);
                if (!Paths.structure.has(radical)) {
                    Paths.structure.set(radical, new Map());
                }
                const letter = Paths.extractLetter(key);
                if (!Paths.structure.get(radical)!.has(letter)) {
                    Paths.structure.get(radical)!.set(letter, new Map());
                }
                const number = Paths.extractNumber(key);
                if (!Paths.structure.get(radical)!.get(letter)!.has(number)) {
                    Paths.structure.get(radical)!.get(letter)!.set(number, path);
                }

            } catch (e) {
                Log.warn(`Stations.load (ligne ${excelRow}) : ${e}`);
                continue;
            } 
        }

        // Charge les arrêts des parcours
        Stops.load();

        // Vérifie si les parcours sont valides
        for (const path of this.map.values()) {
            try {
                path.check();
            } catch (e) {
                Log.warn(`Paths.load : ${e}`);
                continue;
            }
        }
    }

    /**
     * Sauvegardeles parcours dans un tableau.
     * Les données sont celles stockées dans l'objet Paths.map.
     * @param {string} [sheetName=Paths.SHEET] Nom de la feuille de calcul.
     * @param {string} [tableName=Paths.TABLE] Nom du tableau.
     * @param {string} [startCell="A1"] Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = Paths.SHEET,
        tableName: string = Paths.SHEET,
        startCell: string = "A1"
    ): void {

        // Convertit la map en un tableau de données
        const data: (string | number)[][] = Array
        .from(Paths.map.values())
        .map(path => [
            path.key,
            path.parity.printLetter(),
            path.lineDirection.printLetter(),
            path.missionCode,
            path.name,
            path.signature,
            path.stopsChecked
        ]);

        // Imprime le tableau
        WorkbookService.printTable(Paths.HEADERS, data, sheetName, tableName, startCell);

        Stops.print();
    }
}

/**
 * Classe Train définissant un train, pour un unique jour, étant la réutilisation
 * d'un ou deux trains précédents, et ayant une ou deux réutilisations,
 * en faisant référence à un sillon avec horaires pouvant circuler plusieurs jours par semaine.
 */
class Train {

    // Constantes des éléments
    public static readonly NORTH: number = 0;
    public static readonly SOUTH: number = 1;

    // Propriétés de l'objet Train
    public readonly number: TrainNumber;            // Numéro du train
    public readonly path: Path;                     // Parcours sur lequel le train circule
    public readonly date: DateTime;                 // Date et heure de départ du train
    public readonly service: string;                // Service auquel le train est rattaché
    public readonly units: string[];                // Eléments (numéro de matériel)
    public readonly previous: TrainNumber[];        // Trains précédents
    public readonly reuses: TrainNumber[];          // Trains de réutilisations
    public readonly reuseKeys: string[];            // Clés des trains de réutilisations


    constructor(
        number: string,
        pathKey: string,
        date: number,
        departureTime: number,
        service: string,
        units: string = "",
        previous: string = "",
        reuseKeys: string = "",
    ) {
        this.number = new TrainNumber(number);
        const path = Paths.map.get(pathKey) as Path;
        this.path = Paths.map.get(pathKey) as Path;
        const date = DateTime.from(date + departureTime, false);
        if (!this.path) {
            throw new Error(`Train n° ${this.number} : le sillon rattaché est inconnu : ${pathKey}.`);
        }
        this.date = DateTime.from(date + departureTime, false);
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
     * Charge les trains à partir du tableau "Trains" de la feuille "Trains".
     * Les trains sont stockées dans une map avec comme clé la clé du train
     *  et comme valeur l'objet Train.
     * @param {boolean} [erase=false] Si vrai, force le rechargement des trains.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
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




    
    // const number = String(row[Paths.COL_NUMBER]);
    // const days = String(row[Paths.COL_DAYS]);
    
    // // Vérifie si le parcours est déjà chargé
    // if (Paths.map.has(`${number}_${days}`)) continue;

    // // Vérifie si le parcours est concerné dans la liste des parcours à charger, sauf si aucun filtre n'est fourni
    // if (trainNumberMap.size > 0 && !trainNumberMap.has(`${number}`)) continue;

    // // Détermine les jours à filtrer
    // const filterDays = trainNumberMap.get(`${number}`) || trainDays;

    // // Calcule les jours communs entre ceux du parcours et ceux demandés
    // const commonDays = Day.extractFromString(days, filterDays);
    // if (commonDays.length === 0) continue;


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
     * Sauvegardeles trains de la map dans un tableau.
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

/**
 * Classe représentant un train.
 */
class TrainPath {

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
 * Classe TrainPaths contenant la liste des sillons
 */
class TrainPaths {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Sillons";               // Feuille contenant la liste des trains
    private static readonly TABLE = "Sillons";               // Tableau contenant la liste des trains
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
     * Charge les sillons à partir du tableau "Sillons" de la feuille "Sillons".
     * Les sillons sont stockées dans une map avec comme clé la clé du sillon
     *  et comme valeur l'objet TrainPath.
     * @param {boolean} [erase=false] Si vrai, force le rechargement des sillons.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
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
     * Sauvegardeles sillons de la map dans un tableau.
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

    const round = (v: number) => Math.round(v * 1e10) / 1e10;

    /* ==========================================================
    1. CONSTRUCTION & ROLLOVER
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
            desc: 'Durée relative (01:00)',
            value: 1 / 24,
            isRelative: true,
            expected: 1 / 24
        },
    ];

    constructorTests.forEach(t => {
        const dt = DateTime.from(t.value, t.isRelative);
        assert.check(
            `from(${t.value}, ${t.isRelative}) → excelValue (${t.desc})`,
            round(dt?.excelValue ?? 0),
            round(t.expected)
        );
    });

    /* ==========================================================
    2. GETTERS HEURE
    ========================================================== */

    const time = DateTime.from(4.5 / 24)!; // 04:30

    assert.check('getHours()', time.getHours(), 4);
    assert.check('getMinutes()', time.getMinutes(), 30);
    assert.check('getSeconds()', time.getSeconds(), 0);

    /* ==========================================================
    3. GETTERS DATE & ADAPTATION
    ========================================================== */

    // 22/06/2025 01:00 → adapté = 21/06/2025
    const dtAdapt = DateTime.from(45830 + 1/24)!;

    assert.check('getDay(adapted)', dtAdapt.getDay(true), 21);
    assert.check('getDay(real)', dtAdapt.getDay(false), 22);

    assert.check(
        'getDayOfWeek(adapted)',
        dtAdapt.getDayOfWeek(true)?.number,
        Day.SATURDAY.number
    );

    assert.check(
        'getDayOfWeek(real)',
        dtAdapt.getDayOfWeek(false)?.number,
        Day.SUNDAY.number
    );

    /* ==========================================================
    4. getTime() adapté / non adapté
    ========================================================== */

    const t = DateTime.from(1/24)!; // 01:00

    assert.check(
        'getTime(adapted) → 25:00',
        round(t.getTime(true)),
        round(1/24 + 1)
    );

    assert.check(
        'getTime(real) → 01:00',
        round(t.getTime(false)),
        round(1/24)
    );

    /* ==========================================================
    5. format() heure
    ========================================================== */

    const formatTimeTests = [
        { value: 4.5/24, fmt: DateTime.TIME_FORMAT_WITH_SECONDS, exp: '04:30:00' },
        { value: 4.5/24, fmt: DateTime.TIME_FORMAT_WITHOUT_SECONDS, exp: '04:30' },
        { value: -4.5/24, fmt: DateTime.TIME_FORMAT_WITHOUT_SECONDS, exp: '-04:30' },
    ];

    formatTimeTests.forEach(t => {
        const dt = DateTime.from(t.value, true)!;
        assert.check(
            `format("${t.fmt}")`,
            dt.format(t.fmt),
            t.exp
        );
    });

    /* ==========================================================
    6. format() date
    ========================================================== */

    const dt = DateTime.from(45830.75)!; // 22/06/2025

    assert.check(
        'format DATE_WITH_YEAR',
        dt.format(DateTime.DATE_FORMAT_WITH_YEAR),
        '22/06/2025'
    );

    assert.check(
        'format DATE_WITHOUT_YEAR',
        dt.format(DateTime.DATE_FORMAT_WITHOUT_YEAR),
        '22/06'
    );

    assert.check(
        'format DATE_WITH_DAY',
        dt.format('dddd dd/mm/yyyy'),
        'Dimanche 22/06/2025'
    );

    assert.check(
        'format DATE_ID',
        dt.format(DateTime.DATE_FORMAT_FOR_ID),
        '250622'
    );

    /* ==========================================================
    7. resolveAgainst / relativeTo / equalsTo / compare
    ========================================================== */

    const ref = DateTime.from(45830 + 10/24)!;
    const rel = DateTime.from(3/24, true)!;
    const abs = DateTime.from(45830 + 10/24)!;

    assert.check(
        'resolveAgainst',
        round(rel.resolveAgainst(ref).excelValue),
        round(45830 + 13/24)
    );

    assert.check(
        'relativeTo',
        abs.relativeTo(ref).excelValue,
        0
    );

    assert.check(
        'equalsTo',
        abs.equalsTo(DateTime.from(45830 + 10/24)),
        true
    );

    assert.check(
        'compareTo',
        abs.compareTo(DateTime.from(45830 + 10/24)!),
        0
    );

    /* ==========================================================
    8. equalsOrUndefined()
    ========================================================== */

    const dt1 = DateTime.from(45830 + 10/24)!;
    const dt2 = DateTime.from(45830)!;

    const equalsOrUndefinedTests = [
        { a: undefined, b: undefined, expected: true },
        { a: dt1, b: undefined, expected: false },
        { a: dt1, b: dt1, expected: false },
        { a: dt1, b: dt2, expected: true },
    ];

    equalsOrUndefinedTests.forEach((t, index) => {
        assert.check(
            `equalsOrUndefined test #${index + 1}`,
            DateTime.equalsOrUndefined(t.a, t.b),
            t.expected
        );
    });

    /* ==========================================================
    9. add / subtract relatifs
    ========================================================== */

    const A = DateTime.from(2/24, true)!;
    const B = DateTime.from(3/24, true)!;

    assert.check('add', round(A.add(B).excelValue), round(5/24));
    assert.check('subtract', round(A.subtract(B).excelValue), round(-1/24));

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
    5. Constantes Day.*
    ========================================================== */

    const constantsTests = [
        { const: Day.MONDAY,    num: 1, name: "Lundi" },
        { const: Day.TUESDAY,   num: 2, name: "Mardi" },
        { const: Day.WEDNESDAY, num: 3, name: "Mercredi" },
        { const: Day.THURSDAY,  num: 4, name: "Jeudi" },
        { const: Day.FRIDAY,    num: 5, name: "Vendredi" },
        { const: Day.SATURDAY,  num: 6, name: "Samedi" },
        { const: Day.SUNDAY,    num: 7, name: "Dimanche" },
    ];

    constantsTests.forEach(t => {

        assert.check(
            `Day constant number (${t.name})`,
            t.const.number,
            t.num
        );

        assert.check(
            `Day constant fullName (${t.name})`,
            t.const.fullName,
            t.name
        );

        assert.check(
            `Day constant identity fromNumber (${t.name})`,
            Day.fromNumber(t.num) === t.const,
            true
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
       3. is() / isDefined()
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

    /* ==========================================================
       4. isOpposedTo()
       ========================================================== */

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

    assert.check(
        "DOUBLE n'est opposé à rien",
        Parity.double().isOpposedTo(Parity.odd()),
        false
    );

    assert.check(
        "UNDEFINED n'est opposé à rien",
        Parity.undefined().isOpposedTo(Parity.even()),
        false
    );


    /* ==========================================================
       5. equalsTo() / 
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

    const oddSimple = Parity.odd(false);
    const oddDoubleAllowed = Parity.odd(true);

    assert.check(
        "equalsTo faux si doubleParityAllowed différent",
        oddSimple.equalsTo(oddDoubleAllowed),
        false
    );

    /* ==========================================================
        6. Parity.includes()
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

    const simpleParity = Parity.odd(false);

    assert.check(
        "includes refuse double si non autorisée",
        simpleParity.includes("IP"),
        false
    );
   

    /* ==========================================================
       7. invert()
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

    const pDouble = Parity.double();
    const pUndefined = Parity.undefined();

    assert.check(
        "invert DOUBLE retourne la même instance",
        pDouble.invert() === pDouble,
        true
    );

    assert.check(
        "invert UNDEFINED retourne la même instance",
        pUndefined.invert() === pUndefined,
        true
    );

    /* ==========================================================
       8. combineWith()
       ========================================================== */

       const combineTests = [
        // undefined + odd → odd
        {
            a: Parity.undefined(true),
            b: Parity.odd(),
            expected: Parity.ODD
        },

        // odd + undefined → odd
        {
            a: Parity.odd(true),
            b: Parity.undefined(),
            expected: Parity.ODD
        },

        // odd + odd → odd
        {
            a: Parity.odd(true),
            b: Parity.odd(),
            expected: Parity.ODD
        },

        // even + even → even
        {
            a: Parity.even(true),
            b: Parity.even(),
            expected: Parity.EVEN
        },

        // odd + even → double
        {
            a: Parity.odd(true),
            b: Parity.even(),
            expected: Parity.DOUBLE
        }
    ];

    combineTests.forEach(t => {
        const result = t.a.combineWith(t.b);
        assert.check(
            `combineWith ${t.a.value} + ${t.b.value}`,
            result.value,
            t.expected
        );
    });

    // Test erreur si double non autorisé
    assert.throws(
        "combineWith interdit si doubleParityAllowed = false",
        () => Parity.odd(false).combineWith(Parity.even())
    );

    // Test immutabilité
    const original = Parity.odd(true);
    const combined = original.combineWith(Parity.even());

    assert.check(
        "combineWith ne modifie pas l'instance d'origine",
        original.value,
        Parity.ODD
    );

    assert.check(
        "combineWith retourne une nouvelle instance si changement",
        combined !== original,
        true
    );

    /* ==========================================================
       9. printDigit() / printLetter()
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
       10. printDigit(withUnderscores)
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
       11. static factories
       ========================================================== */

    assert.check(
        "Parity.odd() crée une parité impaire",
        Parity.odd().value,
        Parity.ODD
    );

    assert.check(
        "Parity.even() crée une parité paire",
        Parity.even().value,
        Parity.EVEN
    );

    assert.check(
        "Parity.double() crée une parité double",
        Parity.double().value,
        Parity.DOUBLE
    );

    assert.check(
        "Parity.undefined() crée une parité undefined",
        Parity.undefined().value,
        Parity.UNDEFINED
    );

    assert.check(
        "Parity.double() autorise toujours doubleParityAllowed",
        Parity.double().combineWith(Parity.odd()).value,
        Parity.DOUBLE
    );

    /* ==========================================================
       12. equalsOrUndefined()
       ========================================================== */

    const equalsOrUndefinedTests = [
        { a: undefined, b: undefined, expected: true },
        { a: Parity.odd(), b: undefined, expected: false },
        { a: undefined, b: Parity.even(), expected: false },
        { a: Parity.odd(), b: Parity.odd(), expected: true },
        { a: Parity.even(), b: Parity.even(), expected: true },
        { a: Parity.odd(), b: Parity.even(), expected: false },
        { a: Parity.double(), b: Parity.double(), expected: true }
    ];

    equalsOrUndefinedTests.forEach((t, index) => {
        assert.check(
            `equalsOrUndefined test #${index + 1}`,
            Parity.equalsOrUndefined(t.a, t.b),
            t.expected
        );
    });

    /* ==========================================================
       13. static containsParityLetter()
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
       14. static letter() / digit()
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
