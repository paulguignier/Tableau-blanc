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
function runAllTests(testMode = false): boolean {

    if (!testMode) return false;

    Params.load();

    // testWorkbookService({ printSuccess: false, printFailure: true });
    // testDateTime({ printSuccess: false, printFailure: true });
    // testDay({ printSuccess: false, printFailure: true });
    // testParity({ printSuccess: false, printFailure: true });
    // testTrainNumber({ printSuccess: false, printFailure: true });
    // testStations({ printSuccess: false, printFailure: true });
    testConnections({ printSuccess: false, printFailure: true });
    Log.debug("Test des connections : ", Connections.map);
    Log.debug("Fin des tests");
    return true;

    loadConnections();
    const parity = new Parity("A", false);
    Log.info(!parity);

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

    return true;
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
     * Imprime le resultat des tests
     * @param {string} [title="Résultats des tests"] Titre du test
     */
    public printSummary(title = "Résultats des tests", reset: boolean = true): void {
        CONSOLE.log(
            `${title} : ${this.success} / ${this.total} réussis `
            + `(échecs : ${this.failure})`
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

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Param";                // Feuille contenant les paramètres globaux
    private static readonly TABLE = "Paramètres";           // Tableau contenant les paramètres globaux
    private static readonly ROW_MAX_CONNEXIONS_NUMBER = 1;  // Ligne contenant le nombre maximum de connexions
    private static readonly ROW_TURNAROUND_TIME = 2;        // Ligne contenant le temps de retournement
    private static readonly ROW_MAX_TRAIN_UNITS = 3;        // Ligne contenant le nombre maximal d'unités en UM
    private static readonly ROW_TERMINUS_NAME = 5;          // Ligne contenant le nom du terminus

    // Indicateur de chargement
    private static loaded = false;

    // Paramètres globaux
    public static maxConnectionNumber: number = 0;
    public static turnaroundTime: number = 0;
    public static maxTrainUnits: number = 0;
    public static terminusName: string = "";

    /**
     * Chargement global des paramètres
     * @param {boolean} erase Si vrai, force le rechargement des paramètres
     */
    public static load(erase = false): void {
        if (Params.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(Params.SHEET, Params.TABLE);

        // Extrait les valeurs
        Params.maxConnectionNumber = data[Params.ROW_MAX_CONNEXIONS_NUMBER][1] as number;
        Params.turnaroundTime = data[Params.ROW_TURNAROUND_TIME][1] as number;
        Params.maxTrainUnits = data[Params.ROW_MAX_TRAIN_UNITS][1] as number;
        Params.terminusName = String(data[Params.ROW_TERMINUS_NAME][1]);

        DateTime.load();
        Day.load();
        Parity.load();
        TrainNumber.load();
        
        Params.loaded = true;
    }
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
    return (time < Params.rolloverHour) ? time + 1 : time;
}

/**
 * Formatte une date en jour, mois et année.
 * Ne renvoie rien si la date est inférieure au 2 janvier 1900 (date avec uniquement une heure)
 * @param {number} dateValue Temps en nombre décimal (en jours depuis 1900).
 * @returns {string} Date au format "jj/mm/aaaa".
 */
function formatDate(dateValue: number): string {
    if (dateValue < 2) return "";
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
 * Classe utilitaire pour la gestion des dates et horaires Excel.
 *  Si l'heure est inférieure à l'heure de changement de journée,
 *  elle est incrémentée de 1 pour rester comparable aux autres heures de la journée précédente.
 */
class DateTime {

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Param";                // Feuille contenant les paramètres globaux
    private static readonly TABLE = "Paramètres";           // Tableau contenant les paramètres globaux
    private static readonly ROW_ROLLOVER_HOUR = 4;          // Ligne contenant l'heure de changement dejournée
       
    // Etat de chargement
    private static readonly loaded = false;

    // Heure de changement de journée (fraction de jour Excel)
    public static readonly rolloverHour: number = 0;

    // Propriété de l'objet DateTime
    public readonly value: number;      // Valeur du temps en nombre décimal
    
    /**
     * Constructeur de l'objet DateTime.
     * @param {number|string} value Valeur du temps en nombre décimal ou en chaîne de caractères.
     * Si la valeur est inférieure à l'heure de changement de journée, elle est incrémentée de 1.
     */
    constructor(value: number | string) {
        const v = Number(value);
        this.value = (v < DateTime.rolloverHour) ? v + 1 : v;
    }

    /**
     * Ajuste une heure pour tenir compte du changement de journée.
     * Si l'heure est inférieure à l'heure de changement de journée,
     *  on ajoute 1 pour passer à la journée suivante.
     * @param {number} time Heure à ajuster.
     * @returns {number} Heure ajustée.
     */
    public static applyRollover(time: number): number {
        if (!DateTime.loaded) DateTime.load();
        return (time < DateTime.rolloverHour) ? time + 1 : time;
    }

    /**
     * Ajuste une heure pour tenir compte du changement de journée.
     * Exemple : 01:00 → 25:00 si changement de journée à 03:00
     */
    public static adaptTime(time: number): number {
        if (!this.loaded) DateTime.load();
        return (time < this.rolloverHour) ? time + 1 : time;
    }

    /**
     * Formatte la date en jj/mm/aaaa (jours depuis 1900)
     */
    public formatDate(): string {
        if (this.value < 2) return "";

        const excelBase = new Date(Date.UTC(1899, 11, 30));
        const days = Math.floor(this.value);
        const date = new Date(excelBase.getTime() + days * 86400000);

        const year = date.getUTCFullYear();
        const month = date.getUTCMonth() + 1;
        const day = date.getUTCDate();

        return `${day.toString().padStart(2, '0')}/`
             + `${month.toString().padStart(2, '0')}/`
             + `${year}`;
    }

    /**
     * Formatte la date en aammjj pour être utilisée comme ID.
     */
    public formatDateForId(): string {
        if (this.value < 2) return "";

        const excelBase = new Date(Date.UTC(1899, 11, 30));
        const days = Math.floor(this.value);
        const date = new Date(excelBase.getTime() + days * 86400000);

        const year = date.getUTCFullYear();
        const month = date.getUTCMonth() + 1;
        const day = date.getUTCDate();

        return `${year.toString().slice(-2).padStart(2, '0')}`
             + `${month.toString().padStart(2, '0')}`
             + `${day.toString().padStart(2, '0')}`
    }

    /**
     * Formatte l'heure en hh:nn:ss (par défault) ou hh:nn.
     */
    public formatTime(withSeconds: boolean = true): string {

        const totalSeconds = Math.round(
            (this.value - Math.floor(this.value)) * 86400
        );

        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;

        return `${hours.toString().padStart(2, '0')}:`
             + `${minutes.toString().padStart(2, '0')}`
             + (withSeconds
                ? `:${seconds.toString().padStart(2, '0')}`
                : '');
    }

    /**
     * Charge le paramètre de l'heure de changement de journée.
     */
    public static load(erase = false): void {
        if (DateTime.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(
            DateTime.SHEET,
            DateTime.TABLE
        );

        DateTime.rolloverHour =
            Number(data[DateTime.ROW_ROLLOVER_HOUR][1]) % 1;

        DateTime.loaded = true;
    }
}

/**
 * Classe utilitaire pour la gestion des jours de la semaine, individuels ou groupés. 
 *  (JOB du lundi au vendredi, WE pour samedi et dimanche...).
 */
class Day {

    // Constantes de lecture du tableau Excel
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
        if (Day.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(Day.SHEET, Day.TABLE);

        for (const row of data.slice(1)) {
            // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides)
            if (row.every(cell => !cell)) continue;

            // Extrait les valeurs
            const numbersString = String(row[Day.COL_NUMBERS]);
            const fullName = String(row[Day.COL_FULL_NAME]);
            const abreviation = String(row[Day.COL_ABBREVIATION]);

            // Crée l'objet Day
            const day = new Day(numbersString, fullName, abreviation);
            Day.daysByNumbers .set(day.numbersString, day);
        }

        Day.loaded = true;
    }
}


/*
 * Classe utilitaire qui permet de manipuler la parité
 *  d'un train, d'un sillon ou d'un arrêt.
 */
class Parity {

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Param";        // Feuille contenant les paramètres de parité
    private static readonly TABLE = "Parité";       // Tableau contenant les paramètres de parité
    private static readonly ROW_ODD = 1;            // Ligne de la parité impaire
    private static readonly ROW_EVEN = 2;           // Ligne de la parité paire
    private static readonly ROW_DOUBLE = 3;         // Ligne de la parité double
    private static readonly COL_LETTER = 1;         // Colonne des parités exprimées en lettres
    private static readonly COL_NUMBER = 2;         // Colonne des parités exprimées en chiffres

    // Constantes de parité
    public static readonly odd: number = 1;         // Parité impaire
    public static readonly even: number = 2;        // Parité paire
    public static readonly double: number = -2;     // Parité double
    public static readonly undefined: number = -1;  // Parité non définie

    // Indicateur de chargement
    private static loaded = false;

    // Map des lettres et nombres désignants les parités
    private static letters = new Map<number, string>();
    private static digits = new Map<number, number>();

    // Propriétés de l'objet Parity
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
                const oddDigit = Parity.digit(Parity.odd)!;
                return withUnderscores ? '_' + oddDigit : oddDigit;
            case Parity.even:
                const evenDigit = Parity.digit(Parity.even)!;
                return withUnderscores ? '_' + evenDigit : evenDigit;
            case Parity.double:
                return Parity.digit(Parity.double)!;
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
                return Parity.letter(Parity.odd);
            case Parity.even:
                return Parity.letter(Parity.even);
            case Parity.double:
                return Parity.letter(Parity.odd)!
                    + Parity.letter(Parity.even)!;
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
                return string.toUpperCase().includes(Parity.letter(Parity.odd)!);
            case Parity.even:
                return string.toUpperCase().includes(Parity.letter(Parity.even)!);
            case Parity.double:
                return string.toUpperCase().includes(Parity.letter(Parity.odd)!)
                    && string.toUpperCase().includes(Parity.letter(Parity.even)!);
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
            case Parity.letter(Parity.odd):
            case Parity.digit(Parity.odd)!.toString():
                return Parity.odd;
            case Parity.letter(Parity.even):
            case Parity.digit(Parity.even)!.toString():
                return Parity.even;
            case Parity.letter(Parity.even)! + Parity.letter(Parity.odd)!:
            case Parity.letter(Parity.odd)! + Parity.letter(Parity.even)!:
            case Parity.digit(Parity.double)!.toString():
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
     * @param {string | number} trainNumber Numéro du train, qui peut être un nombre
     *  ou une chaîne de caractères.
     * @param {boolean} with4Digits Si vrai, le numéro du train est abrégé à 4 chiffres.
     *  Si faux, le numéro du train n'est pas abrégé.
     * @returns {string} Numéro du train adapté
     */
    public adaptTrainNumber(trainNumber: string | number, abbreviateTo4Digits: boolean = false): string {
        const trainNumberObject = new TrainNumber(trainNumber);
        return trainNumberObject.adaptWithParity(this.value, abbreviateTo4Digits);
    }

    /**
     * Charge les paramètres de parité des jours (lettre et chiffre associés
     *  aux jours pairs et impairs) à partir de la feuille "Param".
     */
    public static load(erase = false): void {
        if (Parity.loaded && !erase) return;

        const data = WorkbookService.getDataFromTable(Parity.SHEET, Parity.TABLE);

        Parity.letters.set(Parity.odd,
            String(data[Parity.ROW_ODD][Parity.COL_LETTER]).toUpperCase() || "I");
        Parity.letters.set(Parity.even,
            String(data[Parity.ROW_EVEN][Parity.COL_LETTER]).toUpperCase() || "P");
        Parity.digits.set(Parity.odd,
            Number(data[Parity.ROW_ODD][Parity.COL_NUMBER]) || 1);
        Parity.digits.set(Parity.even,
            Number([Parity.ROW_EVEN][Parity.COL_NUMBER]) || 2);
        Parity.digits.set(Parity.double,
            Number([Parity.ROW_DOUBLE][Parity.COL_NUMBER]) || -2);
    
        Parity.loaded = true;
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
}

/**
 * Classe TrainNumber définissant un numéro de train.
 * Il est alphanumérique, sans ponctuation et sans espaces,
 *  avec un chiffre pour dernier caractère.
 */
class TrainNumber {

    // Constantes de lecture du tableau Excel
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
    public readonly value: string;                  // Numéro de train (sans double parité)

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
            ? TrainNumber.applyParity(normalized, Parity.double)
            : normalized;
    }

    /**
     * Normalise un numéro de train en supprimant les caractères non-alphanumériques
     * et en remplaçant les "/" par des espaces.
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
            case Parity.even:
                return rest + even;
            case Parity.odd:
                return rest + (even + 1);
            case Parity.double:
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
    public adaptWithParity(parityValue: number, abbreviateTo4Digits = false): string {

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
    public print(abbreviateTo4Digits: boolean = false, withoutDoubleParity: boolean = false): string {
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
            .filter(v => typeof v === "string" && v.trim() !== "")
            .map(pattern => {
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
        const data = WorkbookService.getDataFromTable(TrainNumber.TRAINS_4DIGIT_SHEET, TrainNumber.TRAINS_4DIGIT_TABLE);
    
        // Transforme chaque motif en regex partielle
        const regexParts: string[] = data
            .slice(1)
            .flat()
            .filter(v => typeof v === "string" && v.trim() !== "")
            .map(pattern => {
                return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
            });

        // Crée une regex globale combinée
        TrainNumber.abbreviate4Regex = new RegExp(regexParts.join('|'));
    }
}

/**
 * Classe Path définissant le parcours d'un train, avec ses gares et temps de passage par rapport à la gare origine
 */
class Path {

    // Propriétés de l'objet Path
    key: string;                        // Clé du parcours
    parity: Parity;                     // Parité du parcours (synthèse des parités pour chaque gare)
    lineDirection: Parity;              // Direction du sillon sur la ligne
                                        //  (donnée par une parité globale)
    missionCode: string;                // Code de mission des trains du sillon
    viaStations: string[];              // Gares intermédiaires du parcours (via)
                                        //  (gares précédées de @ si l'ordre de passage doit être respecté)
    stops: Map<string, Stop>            // Gares d'arrêt ou gares de passage du parcours
    stopsChecked?: number;              // Arrêts vérifiés :
    //  - 1 : uniquement les gares de départ et d'arrivée,
    //  - 2 : arrêts commerciaux dans l'ordre,
    //  - 3 : tous les arrêts et gares de passage du sillon
    //       (suite à findPath)

    constructor(
        key: string = "",
        parityValue: number = Parity.undefined,
        lineDirection: number = Parity.undefined,
        missionCode: string = "",
        viaStations: string = ""
    ) {
        this.key = key;
        this.parity = new Parity(parityValue, true);
        this.lineDirection = new Parity(lineDirection, true);
        this.missionCode = missionCode;
        this.viaStations = viaStations ? viaStations.split(';') : [];
        this.stops = new Map<string, Stop>();
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
        if (stopName = Params.terminusName) return null;
        // Si l'arrêt est donné et trouvé avec parité, l'arrêt est renvoyé
        if (this.stops.has(stopName)) return this.stops.get(stopName)!;

        // Recherche du nom de la gare de l'arrêt et de la parité demandée
        const [stationName, parity] = stopName.split("_");
        const station = Stations.get(stationName);
        if (!station) return undefined;
        let stop: Stop | null | undefined;
        if (parity === undefined) {
            // Si la parité n'est pas donnée dans la demande, cherche l'arrêt avec parité
            //  en fonction du numéro de train
            const parityFromTrainNumber = new Parity(this.trainNumber.value, false);
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
     * Efface la liste des arrêts du train.
     * Supprime également les valeurs de firstStop et lastStop.
     */
    eraseStops() {
        this.stops.clear();
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
        const allCombinations = Paths.generateCombinations(this.firstStop!.key, this.lastStop!.key,
            viaStops, viaSorted);

        // Trouve le chemin le plus court parmi toutes les combinaisons
        const shortestPath = Paths.findShortestPath(allCombinations);

        // Quitte la fonction si aucun chemin n'est trouvé
        if (!shortestPath || shortestPath.path.length === 0) return;

        // Crée la nouvelle liste d'arrêts
        const newStops = new Map<string, Stop>();
        Log.debug(newStops);
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
                const connection = Connections.get(previousStopName, stopName);
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
        // Log.debug(this.arrivalStation);
        // Log.debug(this.getStop(this.arrivalStation));

        // this.lastStop = previousStop;
        if (this.lastStop) this.lastStop.changeNumber = Stop.lastStop;
        this.stops = newStops;
        // Log.debug(this.stops.get("MPU_1"));
    }
}

class Paths {

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Parcours";
    private static readonly TABLE = "Parcours";
    private static readonly HEADERS = [[
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
    private static readonly COL_KEY = 0; // Non lue car calculée
    private static readonly COL_NUMBER = 1;
    private static readonly COL_LINE_PARITY = 2;
    private static readonly COL_DAYS = 3;
    private static readonly COL_MISSION_CODE = 4;
    private static readonly COL_DEPARTURE_TIME = 5;
    private static readonly COL_DEPARTURE_STATION = 6;
    private static readonly COL_ARRIVAL_TIME = 7;
    private static readonly COL_ARRIVAL_STATION = 8;
    private static readonly COL_FIRST_STATION = 9;    // Valeur non lue car affectée lors de la lecture du premier arrêt
    private static readonly COL_LAST_STATION = 10;    // Valeur non lue car affectée lors de la lecture du dernier arrêt
    private static readonly COL_VIA_STATIONS = 11;

    // Map des parcours indexées par leur clé
    public static readonly map: Map<string, Path> = new Map();


    /**
     * Charge les sillons de trains à partir du tableau "Sillons" de la feuille "Sillons".
     * Les sillons sont stockés dans un objet avec comme clés le numéro de sillon 
     * suivi du jour et comme valeur l'objet Path.
     * Chaque sillon correspondant à la sélection sera associé avec autant de clés que de jours
     * de circulation, en plus du numéro de sillon suivi du code des jours de circulation
     * (le sillon 123456_J aura pour clés : 123456_J, 123456_1, 123456_2...)
     * @param {string} days Jours pour lesquels les sillons sans jours spécifiques sont demandés.
     * @param {string} trainNumbers Numéros des sillons à charger, avec ou sans jours associés, séparés par des ';'.
     * Si vide, charge tous les trains de la base Paths.map.
     * @param {boolean} [erase=false] Si vrai, supprime les trains déjà chargés.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(trainDays: string = "JW", trainNumbers: string = "", erase: boolean = false) {

        // Vérifie si la table à charger existe déjà
        if (Paths.map.size > 0) {
            if (erase) {
                Paths.map.clear(); // Vide la map sans changer sa référence
            }
        }

        Stations.load(); // Charge les gares si elles ne sont pas encore chargées
        const data = WorkbookService.getDataFromTable(Paths.SHEET, Paths.TABLE);

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

            const number = String(row[Paths.COL_NUMBER]);
            const days = String(row[Paths.COL_DAYS]);
            
            // Vérifie si le sillon est déjà chargé
            if (Paths.map.has(`${number}_${days}`)) continue;

            // Vérifie si le sillon est concerné dans la liste des sillons à charger, sauf si aucun filtre n'est fourni
            if (trainNumberMap.size > 0 && !trainNumberMap.has(`${number}`)) continue;

            // Détermine les jours à filtrer
            const filterDays = trainNumberMap.get(`${number}`) || trainDays;

            // Calcule les jours communs entre ceux du sillon et ceux demandés
            const commonDays = Day.extractFromString(days, filterDays);
            if (commonDays.length === 0) continue;

            // Extrait les valeurs
            const lineDirection = row[Paths.COL_LINE_PARITY] as number;
            const missionCode = String(row[Paths.COL_MISSION_CODE]);
            const departureTime = row[Paths.COL_DEPARTURE_TIME] as number;
            const departureStation = String(row[Paths.COL_DEPARTURE_STATION]);
            const arrivalTime = row[Paths.COL_ARRIVAL_TIME] as number;
            const arrivalStation = String(row[Paths.COL_ARRIVAL_STATION]);
            const viaStations = String(row[Paths.COL_VIA_STATIONS]);

            // Crée l'objet Path
            const path = new Path(
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
            Paths.map.set(path.key, path);
            //  - une référence pour chacun des jours demandés
            commonDays.forEach((day) => {
                const key = number + "_" + day;
                if (!Paths.map.has(key)) Paths.map.set(key, path);
            });
        }
    }

    /**
     * Affiche les sillons dans un tableau.
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

        // Filtre l'objet Paths.map en ne prennant qu'une seule fois les sillons ayant la même clé   
        const seenKeys = new Set<string>();
        const uniquePaths: Path[] = Array.from(Paths.map.entries())
            .filter(([mapKey, path]) => mapKey === path.key)
            .map(([_, path]) => path);

        // Convertit l'objet Paths.map filtré en un tableau de données
        const data: (string | number)[][] = uniquePaths.map(path => [
            path.key,
            path.number,
            path.lineDirection.printDigit(),
            path.days,
            path.missionCode,
            path.departureTime,
            path.departureStation,
            path.arrivalTime,
            path.arrivalStation,
            path.viaStations.join(';'),
        ]);

        // Imprime le tableau
        const table = WorkbookService.printTable(Paths.HEADERS, data, sheetName, tableName, startCell);

        // Met les heures au format "hh:mm:ss"
        const timeColumns = [
            Paths.COL_DEPARTURE_TIME,
            Paths.COL_ARRIVAL_TIME,
        ];

        for (const col of timeColumns) {
            table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
        }
    }

    /**
     * Cherche les chemins possibles pour tous les sillons de trains stockés 
     * dans l'objet Paths.map.
     * Appel la fonction findPath pour chaque sillon de train.
     */
    public static findPathsOnAllPaths() {
        Paths.map.forEach((path, key) => {
            if (key === path.key) path.findPath();
        });
    }

    /**
     * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
     * @param {string[][]} allCombinations Liste de toutes les combinaisons de parcours à évaluer.
     * @returns {path: string[], totalDistance: number} Chemin le plus court et sa distance totale,
     *  ou null si aucun chemin n'est trouvé.
     */
    public static findShortestPath(allCombinations: string[][])
        : { path: string[], totalDistance: number } | null {

        let shortestPath: { path: string[], totalDistance: number } | null = null;

        for (const combination of allCombinations) {
            // Calcule le chemin complet et la distance totale pour la combinaison actuelle
            const { path, totalDistance } = Paths.calculateCompletePath(combination);

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
    public static calculateCompletePath(combination: string[])
        : { path: string[], totalDistance: number } {

        const completePath: string[] = [];
        let totalDistance = 0;

        for (const i = 0; i < combination.length - 1; i++) {
            // Trouve le chemin le plus court pour le tronçon actuel
            const segmentPath = Paths.dijkstra(combination[i], combination[i + 1]);

            // Si aucun chemin n'est trouvé pour ce tronçon, retourne un chemin vide
            if (segmentPath.length === 0) return { path: [], totalDistance: 0 };

            // Ajoute la distance du tronçon à la distance totale
            totalDistance += Paths.calculatePathTime(segmentPath);

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
    public static calculatePathTime(path: string[]): number {
        let totalTime = 0;

        for (const i = 0; i < path.length - 1; i++) {
            const from = path[i];
            const to = path[i + 1];
            const connection = Connections.get(from, to);
            if (connection) {
                totalTime += connection.time;
                // Ajoute le temps de rebroussement sauf pour le premier segment
                if (i > 0 && connection.withTurnaround) {
                    totalTime += Params.turnaroundTime;
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
    public static dijkstra(start: string, end: string): string[] {
        const distances = new Map<string, number>();
        const previousNodes = new Map<string, string | null>();
        const unvisited = new Set<string>(Connections.keys());
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
            for (const [neighbor, connexion] of Connections.map.get(currentNode) || []) {
                let additionalTime = connexion.time;
                if (connexion.withTurnaround && currentNode !== start) {
                    // Si un rebroussement est nécessaire, ajoute le temps de retournement
                    additionalTime += Params.turnaroundTime;
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
    public static generateCombinations(start: string, end: string, via: string[],
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
        const viaPermutations = viaSorted ? Paths.permute(filteredVia) : [filteredVia];

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
    public static permute(array: string[]): string[][] {
        if (array.length === 0) return [[]];
        if (array.length === 1) return [[array[0]]];

        const result: string[][] = [];

        for (const i = 0; i < array.length; i++) {
            const rest = [...array.slice(0, i), ...array.slice(i + 1)];
            const restPermutations = Paths.permute(rest);

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
    public static expandPermutations(permutation: string[]): string[][] {
        if (permutation.length === 0) return [[]];
        const first = Paths.getAllVariants(permutation[0]);

        const restExpanded = Paths.expandPermutations(permutation.slice(1));

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
    public static getAllVariants(gare: string): string[] {
        // Recherche la gare demandée
        const station = Stations.get(gare.split('_')[0]);
        if (!station) return [];

        // Si la gare a un suffixe (_), renvoie uniquement [gare]
        if (gare.includes('_')) return [gare];

        // Sinon, renvoie les variantes pour les 2 sens, 
        return [
            ...[station, ...station.childStations]
                .filter(v => v.abbreviation.trim() !== '')
                .map(v => [`${v.abbreviation}_${Parity.digit(Parity.odd)}`,
                `${v.abbreviation}_${Parity.digit(Parity.even)}`])
        ].reduce((acc, curr) => acc.concat(curr), []);
    }
}

/**
 * Classe Train définissant un train, pour un unique jour, étant la réutilisation
 * d'un ou deux trains précédents, et ayant une ou deux réutilisations,
 * en faisant référence à un sillon avec horaires pouvant circuler plusieurs jours par semaine.
 */
class Train {

    // Propriétés de l'objet Train
    number: TrainNumber;            // Numéro du train
    path: Path;           // Parcours sur lequel le train est prévu prévu de circuler
    day: number;                    // Jour du train    (1 à 7 = lundi à dimanche, >7 = date précise)
    service: string;                // Service auquel le train est rattaché
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

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Trains";               // Feuille contenant la liste des arrêts
    private static readonly TABLE = "Trains";               // Tableau contenant la liste des arrêts
    private static readonly HEADERS = [[
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
    private static readonly COL_KEY = 0;
    private static readonly COL_NUMBER = 1;
    private static readonly COL_DAYS = 2;
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
            const trainNumber = String(row[STOPS_COL_TRAIN_NUMBER]);
            const trainDays = String(row[STOPS_COL_TRAIN_DAYS]);
            if (!trainNumber || !trainDays) continue;

            const trainKey = trainNumber + "_" + trainDays;
            if (!Paths.map.has(trainKey)) continue;

            const train = Paths.map.get(trainKey) as Path;

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
            .map(train => [
                train.key,
                train.number.print(false, true),
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
 * Charge les réutilisations à partir du tableau "Réuts" de la feuille "Réuts".
 * Les réutilisations sont stockés dans la table un objet avec comme clés le numéro de train
 *  suivi du jour de circulation (numéro du jour ou date) et comme valeur l'objet Réutilisation.
 * @param {string} days Jours pour lesquels les sillons sans jours spécifiques sont demandés.
 * @param {string} trainNumbers Numéros des sillons à charger, avec ou sans jours associés,
 *  séparés par des ';'. Si vide, charge tous les trains de la base Paths.map.
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

/*
 * Classe Stop définissant l'arrêt ou le passage d'un train dans une gare
 */
class Stop {

    public static readonly lastStop: number = 2;

    station?: Station;          // Gare de l'arrêt
    parity: Parity;             // Parité de l'arrêt à l'arrivée
    arrivalTime: number;        // Heure / Temps d'arrivée de l'arrêt
    departureTime: number;      // Heure / Temps de départ de l'arrêt
    passageTime: number;        // Heure / Temps de passage à l'arrêt (sans arrêt)
    track: string;              // Voie de l'arrêt
    changeNumber: number;       // Changement de numérotation en gare
    //  - 0 = même train,
    //  - +1 = rebroussement pair vers impair,
    //  - -1 = rebroussement impair vers pair,
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
        this.station = Stations.get(stationName);
        this.parity = new Parity(parity, false);
        this.arrivalTime = arrivalTime;
        this.departureTime = departureTime;
        this.passageTime = passageTime;
        this.track = track;
        this.changeNumber = changeNumber;
        this.nextStopName = nextStopName;
        if (nextStopName = Params.terminusName) this.nextStop = null;
    }

    /**
     * Vérifie la validité de l'objet Stop en envoyant un message d'erreur si :
     *  - le nom de la gare est vide,
     *  - la gare est inconnue,
     * @returns {Stop | undefined} Objet Stop s'il est valide, undefined sinon.
     */
    check(pathKey: string, stationName: string): Stop | undefined {
        if (!stationName) {
            Log.warn(`Parcours : ${pathKey} Un arrêt ne peut pas avoir de gare 
                avec un nom vide.`);
            return undefined;
        } else if (!this.station) {
            Log.warn(`Parcours : ${pathKey} Un arrêt ne peut pas avoir`
                + ` pour gare : "${stationName}" qui est inconnue.`);
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
    public static newStopIncludingParity(
        pathKey: string,
        stopWithParity: string,
        arrivalTime: number = 0,
        departureTime: number = 0,
        passageTime: number = 0,
        track: string = "",
        changeNumber: number = 0,
        nextStopName: string = ""
    ): Stop | undefined {
        const [stationName, parity] = stopWithParity.split("_");
        const stop = new Stop(pathKey, stationName, parity, arrivalTime, departureTime,
            passageTime, track, changeNumber, nextStopName);
        return stop.check(pathKey, stationName);
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
                    Log.warn(`Le premier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure de passage`
                        + ` (${formatTime(this.passageTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
                if (this.arrivalTime) {
                    if (adjustTimes) this.arrivalTime = 0;
                    Log.warn(`Le premier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure d'arrivée`
                        + ` (${formatTime(this.arrivalTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else {
                if (adjustTimes && this.passageTime) {
                    this.departureTime = this.passageTime;
                    this.passageTime = 0;
                }
                Log.warn(`Le premier arrêt ${this.stationName} du sillon`
                    + ` ${pathTrainKey} ne présente pas d'heure de départ.`
                    + this.departureTime
                    ? ` L'heure de passage (${formatTime(this.passageTime)})`
                    + ` a été modifiée en heure de départ.`
                    : "");
                if (!this.departureTime) return false;
            }
            if (departureTimeOfPreviousStop
                && this.departureTime <= departureTimeOfPreviousStop) {
                Log.warn(`Le premier arrêt ${this.stationName} du sillon ${this.key}`
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
                    Log.warn(`Le dernier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure de passage`
                        + ` (${formatTime(this.passageTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
                if (this.departureTime) {
                    if (adjustTimes) this.departureTime = 0;
                    Log.warn(`Le dernier arrêt ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente une heure de départ`
                        + ` (${formatTime(this.departureTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else {
                if (adjustTimes && this.passageTime) {
                    this.arrivalTime = this.passageTime;
                    this.passageTime = 0;
                }
                Log.warn(`Le dernier arrêt ${this.stationName} du sillon`
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
                    Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
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
                    Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} a des heures d'arrivée et de départ identiques`
                        + ` (${formatTime(this.arrivalTime)}).`
                        + adjustTimes
                        ? ` Elles ont donc été remplacées par une heure de passage.`
                        : "");
                } else if (this.passageTime) {
                    if (adjustTimes) this.passageTime = 0;
                    Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente en plus d'une heure d'arrivée`
                        + ` et de départ, une heure de passage`
                        + ` (${formatTime(this.passageTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
            } else if (this.passageTime) {
                // L'arrêt intermédiaire a une heure de passage
                if (this.arrivalTime) {
                    if (adjustTimes) this.arrivalTime = 0;
                    Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
                        + ` ${pathTrainKey} présente en plus d'une heure de passage,`
                        + ` une heure d'arrivée (${formatTime(this.arrivalTime)}) qui `
                        + adjustTimes ? `a été supprimée.` : `ne sera pas prise en compte.`);
                }
                if (this.departureTime) {
                    if (adjustTimes) this.departureTime = 0;
                    Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
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
                Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
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
                Log.warn(`L'arrêt intermédiaire ${this.stationName} du sillon`
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
            Log.warn(`L'arrêt ${this.stationName} du sillon ${this.key}`
                + ` a une heure d'arrivée ou de passage`
                + ` (${formatTime(this.arrivalTime || this.passageTime)}) inférieure`
                + ` à l'heure de départ ou de passage de l'arrêt prédédent`
                + ` (${formatTime(departureTimeOfPreviousStop)}).`);
            return false;
        }

        return true;
    }
}

/**
 * Classe Stops contenant la liste des arrêts
 */
class Stops {

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Arrêts";                // Feuille contenant la liste des arrêts
    private static readonly TABLE = "Arrêts";                // Tableau contenant la liste des arrêts
    private static readonly HEADERS = [[
        "Parcours",
        "Jour",
        "Gare",
        "Parité",
        "Arrivée",
        "Départ",
        "Passage",
        "Voie",
        "Changement de numérotation",
        "Arrêt suivant"
    ]];                                             // En-têtes du tableau des arrêts
    private static readonly COL_TRAIN_NUMBER = 0;
    private static readonly COL_TRAIN_DAYS = 1;
    private static readonly COL_STATION = 2;
    private static readonly COL_PARITY = 3;
    private static readonly COL_ARRIVAL_TIME = 4;
    private static readonly COL_DEPARTURE_TIME = 5;
    private static readonly COL_PASSAGE_TIME = 6;
    private static readonly COL_TRACK = 7;
    private static readonly COL_CHANGE_NUMBER = 8;
    private static readonly COL_NEXT_STOP = 9;
    
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

            // Vérifie si le train existe
            const trainNumber = String(row[STOPS_COL_TRAIN_NUMBER]);
            const trainDays = String(row[STOPS_COL_TRAIN_DAYS]);
            if (!trainNumber || !trainDays) continue;

            const trainKey = trainNumber + "_" + trainDays;
            if (!Paths.map.has(trainKey)) continue;

            const train = Paths.map.get(trainKey) as Path;

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
        for (const train of Paths.map.values()) {
            train.checkStops();
        }
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

        // Filtre l'objet Paths.map en ne prennant qu'une seule fois les trains
        //  ayant la même clé   
        const seenKeys = new Set<string>();
        const uniquePaths: Path[] = Array.from(Paths.map.entries())
            .filter(([mapKey, train]) => mapKey === train.key)
            .map(([_, train]) => train);

        // Crée le tableau final avec les données de chaque arrêt pour chaque train
        const data: (string | number)[][] = [];

        for (const train of uniquePaths) {
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
                    stop.nextStop == null ? Params.terminusName : stop.nextStopName
                ]);
            }
        }

        // Imprime le tableau
        const table = WorkbookService.printTable(STOPS_HEADERS, data, sheetName, tableName, startCell);

        const timeColumns = [
            Stops.COL_ARRIVAL_TIME,
            Stops.COL_DEPARTURE_TIME,
            Stops.COL_PASSAGE_TIME
        ];

        for (const col of timeColumns) {
            table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
        }
    }
}

/* 
 * Classe Station définissant une gare
 */
class Station {

    // Propriétés de l'objet Station
    abbreviation!: string;              // Abréviation de la gare
    name: string;                       // Nom de la gare
    referenceStationName: string;       // Gare de rattachement
    referenceStation: Station | null;   // Gare de rattachement
    childStations: Station[];           // Sous-gares
    turnaround: Parity;                 // Parité d'un rebroussement possible
                                        //  (la parité est celle du train avant rebroussement)
    reverseLineDirection: boolean;      // Parité de la ligne inversée sur cette gare

    /**
     * Constructeur d'une gare.
     * @param {string} abbreviation - Abréviation de la gare
     * @param {string} name - Nom de la gare
     * @param {string} referenceStationName - Nom de la gare de rattachement
     * @param {string|number} turnaround - Parité d'un rebroussement possible
     *  (la parité est celle du train avant rebroussement)
     * @param {boolean} reverseLineDirection - Parité de la ligne inversée sur cette gare
     */
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
            Log.warn(`Une gare ne peut pas avoir une abréviation vide.`);
            return undefined;
        }
        return this;
    }
}

/**
 * Classe Stations contenant la liste des gares
 */
class Stations {

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Gares";                // Feuille contenant la liste des gares
    private static readonly TABLE = "Gares";                // Tableau contenant la liste des gares
    private static readonly HEADERS = [[
        "Abréviation",
        "Nom",
        "Gare de rattachement",
        "Gare de rebroussement",
        "Parité de ligne inversée"
    ]];                                             // En-têtes du tableau des gares
    private static readonly COL_ABBR = 0;                   // Colonne de l'abréviation de la gare
    private static readonly COL_NAME = 1;                   // Colonne du nom de la gare
    private static readonly COL_REFERENCE_STATION = 2;      // Colonne de la gare de rattachement
    private static readonly COL_TURNAROUND = 3;             // Colonne indiquant si un rebroussement est possible (pair ou impair)
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
            if (erase) {
                Stations.clear(); // Vide la map sans changer sa référence
            } else {
                return;
            }
        }

        const data = WorkbookService.getDataFromTable(Stations.SHEET, Stations.TABLE);
        
        // Parcourt la base de données
        const referenceStationPairs: [string, string][] = [];
        for (const row of data.slice(1)) {

            // Extrait les valeurs
            const abbreviation = String(row[Stations.COL_ABBR]);
            if (!abbreviation) continue;

            const name = String(row[Stations.COL_NAME]);
            const referenceStationName = String(row[Stations.COL_REFERENCE_STATION]);
            const turnaround = String(row[Stations.COL_TURNAROUND]);
            const reverseLineDirection = row[Stations.COL_REVERSE_LINE_PARITY] as boolean;

            // Crée l'objet Station
            const station = new Station(
                abbreviation,
                name,
                referenceStationName,
                turnaround,
                reverseLineDirection
            ).check();

            if (!station) continue;

            // Ajoute l'objet Station dans la map
            if (Stations.map.has(abbreviation)) {
                Log.warn(`La gare ${abbreviation} est présente deux fois`
                    + ` dans la base de données.`);
                continue;
            }
            Stations.map.set(abbreviation, station);

            // Mémorise les paires gare/gare de rattachement
            referenceStationPairs.push([abbreviation, referenceStationName]);
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

        WorkbookService.printTable(Stations.HEADERS, data, sheetName, tableName, startCell);
    }
}

/**
 * Classe Connection définissant une connexion orientée entre deux gares
 */
class Connection {

    // Propriétés de l'objet Connexion
    from: string;               // Gare de départ
    to: string;                 // Gare d'arrivée
    time: number;               // Temps de trajet
    withTurnaround: boolean;    // Connexion impliquant un rebroussement
    withMovement: boolean;      // Connexion sous régime de l'évolution
    changeParity: boolean;      // Connexion avec changement de parité


    /**
     * Constructeur d'une connexion.
     * @param {string} from - Gare de départ
     * @param {string} to - Gare d'arrivée
     * @param {number} [time=1] - Temps de trajet
     * @param {boolean} [withTurnaround=false] - Indique si la connexion implique un rebroussement
     * @param {boolean} [withMovement=false] - Indique si la connexion est sous régime de l'évolution
     * @param {boolean} [changeParity=false] - Indique si la connexion implique un changement de parité
     */
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
    public check(): Connection | undefined {
        if (!this.from || !this.to) {
            Log.warn(`Une connexion ne peut pas avoir des gares de départ`
                + ` et d'arrivée vides.`);
            return undefined;
        } else if (this.from === this.to) {
            Log.warn(`Une connexion ne peut pas avoir des gares de départ`
                + ` et d'arrivée ${this.from} identiques et sans changement de parité.`);
            return undefined;
        } else if (!Stations.has(this.from.split("_")[0])) {
            Log.warn(`La gare de départ ${this.from} de la connexion n'existe pas.`);
            return undefined;
        } else if (!Stations.has(this.to.split("_")[0])) {
            Log.warn(`La gare d'arrivée ${this.to} de la connexion n'existe pas.`);
            return undefined;
        } else if (Connections.has(this.from) && Connections.map.get(this.from)!.has(this.to)) {
            Log.warn(`La connexion ${this.from} -> ${this.to} est présente`
                + ` deux fois dans la base de données.`);
            return undefined;
        }
        return this;
    }
}

/**
 * Classe Connections contenant la liste des connexions
 */
class Connections {

    // Constantes de lecture du tableau Excel
    private static readonly SHEET = "Param";            // Feuille contenant la liste des connexions
    private static readonly TABLE = "Connexions";       // Tableau contenant la liste des connexions
    private static readonly HEADERS = [[
        "De",
        "Vers",
        "Durée",
        "Rebroussement",
        "Evolution",
        "Changement de parité"
    ]];                                         // En-têtes du tableau des connexions
    private static readonly COL_FROM = 0;               // Colonne de la gare de départ
    private static readonly COL_TO = 1;                 // Colonne de la gare d'arrivée
    private static readonly COL_TIME = 2;               // Colonne de la durée de parcours
    private static readonly COL_TURNAROUND = 3;         // Colonne indiquant si la connexion implique un rebroussement
    private static readonly COL_MOVEMENT = 4;           // Colonne indiquant si la connexion est sous régime de l'évolution
    private static readonly COL_CHANGE_PARITY = 5;      // Colonne indiquant si la connexion implique un changement de parité

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

        Stations.load(); // Charge les gares si elles n'ont pas encore été chargées
        const data = WorkbookService.getDataFromTable(Connections.SHEET, Connections.TABLE);

        // Parcourt la base de données
        for (const row of data.slice(1)) {

            // Extrait des valeurs
            const from = String(row[Connections.COL_FROM]);
            const to = String(row[Connections.COL_TO]);
            if (!from || !to) continue;
            const time = row[Connections.COL_TIME] as number;
            const withTurnaround = row[Connections.COL_TURNAROUND] as boolean;
            const withMovement = row[Connections.COL_MOVEMENT] as boolean;
            const changeParity = row[Connections.COL_CHANGE_PARITY] as boolean;
    
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
            if (!Connections.map.has(from)) {
                Connections.map.set(from, new Map());
            }

            Connections.map.get(from)!.set(to, connection);
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
        const table = WorkbookService.printTable(
            Connections.HEADERS,
            data,
            sheetName,
            tableName,
            startCell
        );

        // Met les heures au format "hh:mm:ss"
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
    public static saveConnectionsTimes(trainNumbers: string = "") {
        if (trainNumbers === "") {
            trainNumbers =
                Array.from(Paths.map.keys()).filter(key => key === Paths.map.get(key)!.key)
                    .join(";");
        }
        trainNumbers.split(";").forEach((trainNumber) => {
            const path = Paths.map.get(trainNumber);
            path?.stops.forEach((stop) => {
                if (stop.nextStop && Connections.has(stop.key)
                    && Connections.has(stop.nextStop.key)) {
                    const connection = Connections.get(stop.key, stop.nextStop.key);
                    if (connection && stop.nextStop.arrivalTime !== 0
                        && stop.departureTime !== 0) {
                        connection.time = stop.nextStop.arrivalTime - stop.departureTime;
                    }
                }
            });
        });
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

    // 4️⃣ Test checkCellName : invalide (renvoie chaîne vide si failOnError=false)
    assert.check(
        "Cellule invalide 123",
        WorkbookService.checkCellName("123", false),
        ""
    );

    // 5️⃣ Test printTable : création tableau avec données simples
    const headers = [["Col1", "Col2"]];
    const data = [
        [1, 2],
        [3, 4]
    ];
    const table = WorkbookService.printTable(headers, data, testSheetName, testTableName, "A1", true);
    assert.check("Création tableau", table?.getName(), testTableName);

    // 6️⃣ Test getTable : récupération tableau existant
    const table2 = WorkbookService.getTable(testSheetName, testTableName, true);
    assert.check("Récupération tableau existant", table2?.getName(), testTableName);

    // 7️⃣ Test getDataFromTable : vérifie les données
    const tableData = WorkbookService.getDataFromTable(testSheetName, testTableName, true);
    assert.check("Lecture données tableau", tableData[1][0], 1); // valeur 1 en ligne 2, colonne 1 (données, pas en-tête)

    // 8️⃣ Nettoyage : supprime la feuille de test
    WorkbookService.getSheet(testSheetName)?.delete();
    assert.check("Suppression feuille test", WorkbookService.getSheet(testSheetName, { failOnError: false }), null);

    // 9️⃣ Résumé des tests
    assert.printSummary("Tests WorkbookService");
}


function testDateTime(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    DateTime.load();

    /* ==========================================================
       1. CONSTRUCTEUR
       ----------------------------------------------------------
       Vérifie :
       - Pas de rollover
       - Rollover appliqué
       - Valeurs string / number
       ========================================================== */

    const constructorTests = [
        {
            desc: 'Heure après rollover (04:00)',
            value: 4 / 24,
            expected: 4 / 24
        },
        {
            desc: 'Heure avant rollover (01:00 → 25:00)',
            value: 1 / 24,
            expected: 1 / 24 + 1
        },
        {
            desc: 'Minuit (00:00 → 24:00)',
            value: 0,
            expected: 1
        },
        {
            desc: 'Valeur string "0.5" (12:00)',
            value: "0.5",
            expected: 0.5
        }
    ];

    constructorTests.forEach(t => {
        const dt = new DateTime(t.value);
        assert.check(
            `new DateTime(${Number(t.value).toFixed(3)}) - ${t.desc}`,
            dt.value,
            t.expected
        );
    });

    /* ==========================================================
       2. formatTime()
       ----------------------------------------------------------
       Vérifie :
       - hh:mm:ss
       - hh:mm sans secondes
       - rollover conservé
       ========================================================== */

    const formatTimeTests = [
        {
            desc: '04:30:00',
            value: 4.5 / 24,
            withSeconds: true,
            expected: '04:30:00'
        },
        {
            desc: '04:30',
            value: 4.5 / 24,
            withSeconds: false,
            expected: '04:30'
        },
        {
            desc: 'Rollover 01:00 → 25:00',
            value: 1 / 24,
            withSeconds: false,
            expected: '01:00'
        }
    ];

    formatTimeTests.forEach(t => {
        const dt = new DateTime(t.value);
        assert.check(
            `formatTime(${t.value}) - ${t.desc}`,
            dt.formatTime(t.withSeconds),
            t.expected
        );
    });

    /* ==========================================================
       3. formatDate()
       ----------------------------------------------------------
       Vérifie :
       - Date valide
       - Date < 2 → chaîne vide
       ========================================================== */

    const formatDateTests = [
        {
            desc: 'Date Excel valide (22/06/2025)',
            value: 45830,
            expected: '22/06/2025'
        },
        {
            desc: 'Date avec heure (22/06/2025)',
            value: 45830.94347,
            expected: '22/06/2025'
        },
        {
            desc: 'Valeur < 2 → vide',
            value: 1.9,
            expected: ''
        }
    ];

    formatDateTests.forEach(t => {
        const dt = new DateTime(t.value);
        assert.check(
            `formatDate(${t.value}) - ${t.desc}`,
            dt.formatDate(),
            t.expected
        );
    });

    /* ==========================================================
       4. formatDateForId()
       ----------------------------------------------------------
       Vérifie :
       - Date valide
       - Date < 2 → chaîne vide
       ========================================================== */

       const formatDateForIdTests = [
        {
            desc: 'Date Excel valide (22/06/2025)',
            value: 45830,
            expected: '250622'
        },
        {
            desc: 'Date avec heure (22/06/2025)',
            value: 45830.94347,
            expected: '250622'
        },
        {
            desc: 'Valeur < 2 → vide',
            value: 1.9,
            expected: ''
        }
    ];

    formatDateForIdTests.forEach(t => {
        const dt = new DateTime(t.value);
        assert.check(
            `formatDateForId(${t.value}) - ${t.desc}`,
            dt.formatDateForId(),
            t.expected
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
   
    // -----------------------------------------------------------------
    // Tests du constructeur
    // -----------------------------------------------------------------

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
        const d = new Day(t.input, "x", "x");

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

    // -----------------------------------------------------------------
    // Tests extractFromString (simple)
    // -----------------------------------------------------------------

    const extractTests = [
        {
            desc: "Nom complet",
            input: "lundi",
            expected: [1]
        },
        {
            desc: "Abréviation",
            input: "ma",
            expected: [2]
        },
        {
            desc: "Numéros mélangés",
            input: "7;1;3",
            expected: [1, 3, 7]
        },
        {
            desc: "Texte mixte",
            input: "lumeven",
            expected: [1, 3, 5]
        },
        {
            desc: "Mot clé groupe",
            input: "J",
            expected: [1, 2, 3, 4, 5]
        }
    ];

    extractTests.forEach(t => {
        const result = Day.extractFromString(t.input);

        assert.check(
            `Day.extractFromString("${t.input}") (${t.desc})`,
            JSON.stringify(result),
            JSON.stringify(t.expected)
        );
    });

    // -----------------------------------------------------------------
    // Tests extractFromString avec intersection
    // -----------------------------------------------------------------

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
    SYNTHÈSE
    ========================================================== */

    assert.printSummary('testDay');
}


function testParity(options: Partial<AssertDDOptions> = {}) {

    Parity.load();
    const assert = new AssertDD(options);

    /* ==========================================================
   TESTS DATA-DRIVEN - CLASSE Parity
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
        assert.check(
            `new Parity(${JSON.stringify(t.value)}, doubleAllowed=${t.doubleAllowed}), ${t.desc}`,
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
        assert.check(
            `Parity(${t.start}).update(${t.update}) - ${t.desc}`,
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
        assert.check(
            `Parity(${t.start}).value = ${t.set} - ${t.desc}`,
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
        assert.check(
            `Parity(${t.p1}).equals(Parity(${t.p2})) - ${t.desc}`,
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
        assert.check(
            `Parity(${t.value}).is(${t.parity}) - ${t.desc}`,
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
        assert.check(
            `Parity(${t.value}).invert() - ${t.desc}`,
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
            digit: Parity.digit(Parity.odd),
            letter: Parity.letter(Parity.odd)
        },
        {
            desc: 'print pair',
            value: "P",
            doubleAllowed: false,
            digit: Parity.digit(Parity.even),
            letter: Parity.letter(Parity.even)
        },
        {
            desc: 'print double',
            value: "IP",
            doubleAllowed: true,
            digit: Parity.digit(Parity.double),
            letter:
                Parity.letter(Parity.odd)! +
                Parity.letter(Parity.even)!
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

        assert.check(`printDigit - ${t.desc}`, p.printDigit(), t.digit);
        assert.check(`printLetter - ${t.desc}`, p.printLetter(), t.letter);
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
        assert.check(
            `containsParityLetter("${t.string}", ${t.parity}) - ${t.desc}`,
            Parity.containsParityLetter(t.string, t.parity),
            t.expected
        );
    });

    /* ==========================================================
    7. printDigit(withUnderscores)
    ----------------------------------------------------------
    Vérifie :
    - Ajout de l’underscore pour impair / pair
    - Absence d’underscore pour double
    - Gestion de l’indéfini
    ========================================================== */

    const printUnderscoreTests = [
        {
            desc: 'impair avec underscore',
            value: "I",
            doubleAllowed: false,
            withUnderscores: true,
            expected: "_" + Parity.digit(Parity.odd)
        },
        {
            desc: 'pair avec underscore',
            value: "P",
            doubleAllowed: false,
            withUnderscores: true,
            expected: "_" + Parity.digit(Parity.even)
        },
        {
            desc: 'double sans underscore',
            value: "IP",
            doubleAllowed: true,
            withUnderscores: true,
            expected: Parity.digit(Parity.double)
        },
        {
            desc: 'indéfini',
            value: "",
            doubleAllowed: false,
            withUnderscores: true,
            expected: ""
        }
    ];

    printUnderscoreTests.forEach(t => {
        const p = new Parity(t.value, t.doubleAllowed);
        assert.check(
            `printDigit(${t.withUnderscores}) - ${t.desc}`,
            p.printDigit(t.withUnderscores),
            t.expected
        );
    });

    /* ==========================================================
    8. adaptTrainNumber()
    ----------------------------------------------------------
    Vérifie :
    - Adaptation du dernier chiffre selon la parité
    - Gestion de la double parité
    - Respect du flag abbreviateTo4Digits
    ========================================================== */

    const adaptTrainTests = [
        {
            desc: 'pair demandé sur train impair',
            train: "142345",
            parity: Parity.even,
            doubleAllowed: false,
            expected: "142344"
        },
        {
            desc: 'impair demandé sur train pair',
            train: "142346",
            parity: Parity.odd,
            doubleAllowed: false,
            expected: "142347"
        },
        {
            desc: 'double parité',
            train: "142345",
            parity: Parity.double,
            doubleAllowed: true,
            expected: "142344/5"
        },
        {
            desc: 'parité indéfinie',
            train: "142345",
            parity: Parity.undefined,
            doubleAllowed: false,
            expected: "142345"
        }
    ];

    adaptTrainTests.forEach(t => {
        const p = new Parity(t.parity, t.doubleAllowed);
        assert.check(
            `adaptTrainNumber("${t.train}") - ${t.desc}`,
            p.adaptTrainNumber(t.train),
            t.expected
        );
    });

    /* ==========================================================
    9. letter() & digit() statiques
    ----------------------------------------------------------
    Vérifie :
    - Accès correct aux maps internes
    - Valeurs par défaut si parité inconnue
    ========================================================== */

    const staticAccessTests = [
        { desc: 'letter odd', method: 'letter', parity: Parity.odd, expected: Parity.letter(Parity.odd) },
        { desc: 'letter even', method: 'letter', parity: Parity.even, expected: Parity.letter(Parity.even) },
        { desc: 'digit odd', method: 'digit', parity: Parity.odd, expected: Parity.digit(Parity.odd) },
        { desc: 'digit even', method: 'digit', parity: Parity.even, expected: Parity.digit(Parity.even) },
        { desc: 'digit undefined', method: 'digit', parity: 999, expected: 0 }
    ];

    staticAccessTests.forEach(t => {
        const result =
            t.method === 'letter'
                ? Parity.letter(t.parity)
                : Parity.digit(t.parity);

        assert.check(
            `${t.method}(${t.parity}) - ${t.desc}`,
            result,
            t.expected
        );
    });

    // Synthèse finale
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
            tn.doubleParity ? tn.value : tn.value,
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
    // Tests print()
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
            `TrainNumber(${t.value}, doubleParity=${t.doubleParity}).print(${t.abbreviate}, ${t.withoutDoubleParity}) (${t.desc})`,
            tn.print(t.abbreviate, t.withoutDoubleParity),
            t.expected
        );
    });

    // ------------------------------------------------------------
    // adaptWithParity()
    // ------------------------------------------------------------

    const parityTests = [
        { value: 146491, parity: Parity.even, expected: "146490" },
        { value: 146490, parity: Parity.odd, expected: "146491" },
        { value: 146490, parity: Parity.double, expected: "146490/1" },
        { value: 146490, parity: Parity.double, abbreviate: true, expected: "6490/1" }
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
                    + `${station.referenceStation.abbreviation}.childStations`,
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

function testConnections(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    /* ==========================================================
       TESTS DATA-DRIVEN - CLASSE Connections
       ==========================================================
       Objectifs :
       - Vérifier le chargement des connexions
       - Garantir l’unicité et la cohérence from → to
       - Tester l’accès Map statique
       - Tester l’impression dans une feuille de test
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
       3. Cohérence métier
       ========================================================== */

    const coherenceTests: { desc: string; value: boolean }[] = [];

    for (const [from, targets] of Connections.map) {
        for (const [to, connection] of targets) {

            coherenceTests.push({
                desc: `${from} → ${to} : from cohérent`,
                value: connection.from === from
            });

            coherenceTests.push({
                desc: `${from} → ${to} : to cohérent`,
                value: connection.to === to
            });

            coherenceTests.push({
                desc: `${from} → ${to} : temps > 0`,
                value: connection.time > 0
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
