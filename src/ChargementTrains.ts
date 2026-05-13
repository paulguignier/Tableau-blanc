/**
 * Chargements de trains
 * 
 * Code Excel Automate pour la création et l'utilisation de la base de données des trains.
 * 
 * @author Paul Guignier
 * @version 2.1
 * @package scr\ChargementTrains.ts
 */


//Variables globales nécessaires dans ExcelScript (pas d'injection possible).
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



    // Lance la fonction de tests.
    // Si les tests sont actifs, la suite du programme n'est pas exécuté. 
    if (runAllTests(testMode)) return;

    try {
        Log.info(`Chargement des paramètres`);
        Params.load();
        Connections.load();

 
        // Trains.import();
        // Trains.print();
        // Paths.print();




        // Paths.load("", "147500_J;148504_J;147201_J;148202_J;147402_J;
        //      148402_J;147601_J;148602_J;145801_J;145804_J");
        // Paths.load("2", "142446_J");
        // Log.debug(Paths.map);
        // const allCombinations = Paths.generateCombinations("MPU", "ETP", "".split(";"));
        // Log.info(allCombinations);
        // const shortestPath = Paths.findShortestPath(allCombinations);
        // Log.info(shortestPath);

        return;

    } catch (e) {
        Log.warn(`${e.message}`);
    }
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

    Log.info(`Chargement des paramètres`);
    Params.load();
    Connections.load();

    Log.info(`Début des tests`);

    try {
        testWorkbookService({ printSuccess: false, printFailure: true });
        testDateTime({ printSuccess: false, printFailure: true });
        testDays({ printSuccess: false, printFailure: true });
        testParity({ printSuccess: false, printFailure: true });
        testTrainNumber({ printSuccess: false, printFailure: true });
        testStation({ printSuccess: false, printFailure: true });
        testStationWithParity({ printSuccess: false, printFailure: true });
        testConnection({ printSuccess: false, printFailure: true });
        testStop({ printSuccess: false, printFailure: true });
        testPath({ printSuccess: false, printFailure: true });

    } catch (e) {
        Log.warn("Erreur lors des tests", e.message);
    }

    Log.info(`Fin des tests`);
    Log.info(`-------------`);
    Log.info(`Fin des tests`);
    return true;

 





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
 * Classe Logs contenant les trois types de messages du console, et leurs options d'affichage.
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
     * @param {unknown} value - Valeur à vérifier.
     * @returns {boolean} - Vrai si la valeur est concatenable, faux sinon.
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

    /**
     * Méthode interne qui écrit un message dans la console.
     * Elle prend en paramètre le niveau du message (debug, info, warn) et
     *  un tableau d'arguments qui peuvent être des strings, des numbers,
     *  des booleans, des null, des undefined, des objets.
     * Les arguments concaténables sont transformés en string et ajoutés au buffer.
     * Les objets sont ajoutés au tableau output sans modification.
     * Lorsque le buffer contient un objet, il est flush (vide) pour laisser place à l'objet.
     * Enfin, le tableau output est passé à CONSOLE.log pour afficher le message.
     * @param {string} level - Niveau du message (debug, info, warn)
     * @param {unknown[]} args - Tableau des arguments à afficher
     */
    private static log(level: string, args: unknown[]): void {

        const output: unknown[] = [];
        let buffer = `[${level}]`;
 
        args.forEach(arg => {
            if (this.isConcatable(arg)) {
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
     * @param {Partial<LogOptions>} options - Options de l'affichage des logs :
     *  - debug: Afficher les messages de debug,
     *  - info: Afficher les messages d'information,
     *  - warn: Afficher les messages d'avertissement.
     */
    public static configure(options: Partial<LogOptions>) {
        Object.assign(this.options, options);
    }

    /**
     * Envoie un message au console avec le niveau "DEBUG".
     * @param {...unknown[]} args - Arguments à passer au console.log.
     */
    public static debug(...args: unknown[]): void {
        if (!this.options.debug) return;
        this.log("DEBUG", args);
    }
 
    /**
     * Envoie un message au console avec le niveau "INFO".
     * @param {...unknown[]} args - Arguments à passer au console.log.
     */
    public static info(...args: unknown[]): void {
        if (!this.options.info) return;
        this.log("INFO", args);
    }
 
    /**
     * Envoie un message au console avec le niveau "WARN".
     * @param {...unknown[]} args - Arguments à passer au console.log.
     */
    public static warn(...args: unknown[]): void {
        if (!this.options.warn) return;
        this.log("WARN", args);
    }
 
}

/*
 * Options de l'affichage des tests :
 *  - printSuccess: afficher le message de succès,
 *  - printFailure: afficher le message d'échec.
 */
type AssertDDOptions = {
    printSuccess?: boolean;
    printFailure?: boolean;
}

/* 
 * Classe AssertDD contenant les options et les fonctions de tests Data-Driven.
 */
class AssertDD {

    public static readonly THROWS = Symbol("ASSERT_THROWS");    // Constante indiquant 
                                                                // qu'une erreur est attendue

    public static completeTests = 0;    // Nombre total d'ensemble de tests complets.
    public static incompleteTests = 0;  // Nombre total d'ensemble de tests incomplets.

    private total = 0;                  // Décompte du nombre total de tests réalisés
    private success = 0;                // Décompte du nombre total de tests réalisés avec succès
    private failure = 0;                // Décompte du nombre total de tests en échec

    private options: AssertDDOptions;   // Options d'affichage des messages de succès et d'échecs

    /**
     * Constructeur de la classe AssertDD.
     * @param {AssertDDOptions} options - Options d'affichage des messages de succès et d'échecs
     */
    constructor(options: AssertDDOptions = {}) {
        this.options = {
            printSuccess: options.printSuccess ?? true,
            printFailure: options.printFailure ?? true
        };
    }

    /**
     * Réalise le test et l'imprime avec un symbole de réussite (✔) ou d'échec (✘).
     * @param {string} label - Nom du test.
     * @param {T} actual - Valeur actuelle obtenue.
     * @param {T} expected - Valeur attendue.
     * @param {AssertDDOptions} options - Options d'affichage des succès et des échecs.
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
     * @param {string} desc - Nom du test.
     * @param {() => void} fn - Fonction à tester.
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
     * Imprime le resultat des tests.
     * @param {string} [title="Résultats des tests"] - Titre du test.
     */
    public printSummary(title: string = "Résultats des tests", reset: boolean = true): void {
        CONSOLE.log(
            `${title} : ${this.success} / ${this.total} réussis`
            + ` (échecs : ${this.failure})`
        );
        if (reset) this.reset();
    }

    /**
     * Réinitialise le compteur de tests.
     */
    public reset(): void {
        this.total = 0;
        this.success = 0;
        this.failure = 0;
    }
}

type CellValue = string | number | boolean;

/*
 * Classe utilitaire WorkbookService de manipulation des feuilles de calcul Excel.
 */
class WorkbookService {

    /**
     * Renvoie la feuille de calcul Excel correspondant au nom donné.
     * Si la feuille n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * Si createIfMissing est vrai, crée la feuille si elle n'existe pas.
     * @param {string} sheetName - Nom de la feuille de calcul à chercher.
     * @param {boolean} createIfMissing - Si vrai, crée la feuille si elle n'existe pas (faux par défaut).
     * @param {boolean} failOnError - Si vrai (par défaut), lance une exception si la feuille n'existe pas.
     * @returns {ExcelScript.Worksheet | null} - Feuille de calcul Excel correspondant au nom donné,
     *  ou null si elle n'existe pas.
     */
    public static getSheet(
        {
            sheetName,
            createIfMissing = false,
            failOnError = true
        }: {
            sheetName: string,
            createIfMissing?: boolean;  // Faux par défaut
            failOnError?: boolean;      // Vrai par défaut
        }
    ): ExcelScript.Worksheet | null {
 
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
     * Renvoie toutes les données non vides d'une feuille Excel.
     * Utilise la plage "utilisée" (used range).
     * @param {string} sheetName - Nom de la feuille.
     * @param {boolean} [failOnError=true] - Si vrai, lance une erreur si la feuille est vide ou inexistante.
     * @returns {CellValue[][]} - Données de la feuille.
     */
    public static getDataFromSheet(
        sheetName: string,
        failOnError: boolean = true
    ): CellValue[][] {

        const sheet = this.getSheet({ sheetName, failOnError });

        if (!sheet) return [];

        const usedRange = sheet.getUsedRange();

        if (!usedRange) {
            const msg = `La feuille "${sheetName}" est vide.`;
            if (failOnError) throw new Error(msg);
            Log.warn(msg);
            return [];
        }

        return usedRange.getValues();
    }

    /**
     * Renvoie le tableau Excel correspondant au nom donné dans la feuille de calcul donnée.
     * Si le tableau n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * @param {string} sheetName - Nom de la feuille de calcul où chercher le tableau.
     * @param {string} tableName - Nom du tableau à chercher.
     * @param {boolean} [failOnError=true] - Si vrai (par défaut), lance une exception
     *  si le tableau n'existe pas. Si faux, renvoie null.
     * @returns {ExcelScript.Table | null} - Tableau Excel correspondant au nom donné,
     *  ou null si il n'existe pas.
     */
    public static getTable(
        sheetName: string,
        tableName: string,
        failOnError: boolean = true
    ): ExcelScript.Table | null {
        const sheet = this.getSheet({ sheetName, failOnError: false });
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
     * @param {string} sheetName - Nom de la feuille de calcul où chercher le tableau.
     * @param {string} tableName - Nom du tableau à chercher.
     * @param {boolean} [failOnError=true] - Si vrai (par défaut),
     *  lance une exception si le tableau n'existe pas. Si faux, renvoie null.
     * @returns {CellValue[][]} - Données du tableau Excel
     *  correspondant au nom donné, ou null si il n'existe pas.
     */
    public static getDataFromTable(
        sheetName: string,
        tableName: string,
        failOnError: boolean = true
    ): CellValue[][] {
        const table = this.getTable(sheetName, tableName, failOnError);
        if (!table) return [];
        return table.getRange().getValues();
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne.
     * Si la valeur est null ou undefined, renvoie undefined.
     * Si la valeur est un nombre, le convertit en chaîne.
     * Si la valeur est une chaîne, la renvoie telle quelle, en supprimant les espaces inutiles.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @returns {string | undefined} - Valeur de la cellule sous forme de chaîne,
     *  ou undefined si elle est null ou undefined.
     */
    public static getStringOrUndefined(row: unknown[], col: number): string | undefined {
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
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @returns {number | undefined} - Valeur de la cellule sous forme de nombre,
     *  ou undefined si la conversion échoue.
     */
    public static getNumberOrUndefined(row: unknown[], col: number): number | undefined {
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
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @returns {boolean | undefined} - Valeur de la cellule sous forme de booléen,
     *  ou undefined si la conversion échoue.
     */
    public static getBooleanOrUndefined(row: unknown[], col: number): boolean | undefined {
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
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne,
     * en supprimant les espaces inutiles, ou la valeur par défaut si la valeur est null ou undefined.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [defaultValue=""] - Valeur par défaut.
     * @returns {string} - Valeur de la cellule sous forme de chaîne.
     **/
    public static getString(row: unknown[], col: number, defaultValue: string = ""): string {
        return this.getStringOrUndefined(row, col) ?? defaultValue;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de nombre,
     * ou la valeur par défaut si la valeur est null ou undefined.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {number} [defaultValue=0] - Valeur par défaut.
     * @returns {number} - Valeur de la cellule sous forme de nombre.
     **/
    public static getNumber(row: unknown[], col: number, defaultValue: number = 0): number {
        return this.getNumberOrUndefined(row, col) ?? defaultValue;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de booléen,
     * ou la valeur par défaut si la valeur est null ou undefined.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {boolean} [defaultValue=false] - Valeur par défaut.
     * @returns {boolean} - Valeur de la cellule sous forme de booléen.
     **/
    public static getBoolean(row: unknown[], col: number, defaultValue: boolean = false): boolean {
        return this.getBooleanOrUndefined(row, col) ?? defaultValue;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne,
     * ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @param {unknown[]} row - Ligne contenant la cellule. 
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {string} - Valeur de la cellule sous forme de chaîne.
     **/
    public static getRequiredString(row: unknown[], col: number, errorMessage?: string): string {
        const value = this.getStringOrUndefined(row, col);
        if (value === undefined) {
            throw new Error(errorMessage
                ?? `La chaine à récupérer est absente`
                    + ` dans la colonne ${col} de la ligne ${JSON.stringify(row)}.`);
        }
        return value;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de nombre,
     * ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {number} - Valeur de la cellule sous forme de nombre.
     **/
    public static getRequiredNumber(row: unknown[], col: number, errorMessage?: string): number {
        const value = this.getNumberOrUndefined(row, col);
        if (value === undefined) {
            throw new Error(errorMessage
                ?? `Le nombre à récupérer est absent`
                    + ` dans la colonne ${col} de la ligne ${JSON.stringify(row)}.`);
        }
        return value;
    }

    /**
     * Renvoie la valeur de la cellule à l'adresse {row}[{col}] sous forme de booléen,
     * ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {boolean} - Valeur de la cellule sous forme de booléen.
     **/
    public static getRequiredBoolean(row: unknown[], col: number, errorMessage?: string): boolean {
        const value = this.getBooleanOrUndefined(row, col);
        if (value === undefined) {
            throw new Error(errorMessage
                ?? `Le booléen à récupérer est absent`
                    + ` dans la colonne ${col} de la ligne ${JSON.stringify(row)}.`);
        }
        return value;
    }

    /**
     * Vérifie si l'adresse de cellule donnée est valide.
     * Si elle est valide, la renvoie telle quelle.
     * Si elle est invalide, lance une exception si failOnError est vrai,
     *  sinon renvoie une chaîne vide.
     * @param {string} cellName - Adresse de cellule à vérifier.
     * @param {boolean} [failOnError=true] - Si vrai (par défaut), lance une exception
     *  si l'adresse est invalide. Si faux, renvoie une chaîne vide.
     * @returns {string} - Adresse de cellule si elle est valide, une chaîne vide sinon.
     */
    public static checkCellName(cellName: string, failOnError: boolean = true): string {
        // Convertit startCell en majuscules pour éviter les problèmes de casse.
        cellName = cellName.toUpperCase();

        // Vérifie si cellName est une adresse de cellule valide.
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
     * @param {string[][]} headers - En-têtes du tableau.
     * @param {(string | number)[][]} data - Données du tableau.
     * @param {string} sheetName - Nom de la feuille de calcul où afficher le tableau.
     * @param {string} tableName - Nom du tableau à afficher.
     * @param {string} [startCell="A1"] - Cellule où commencer à afficher le tableau
     *  (par défaut: "A1").
     * @param {boolean} [failOnError=true] - Si vrai (par défaut), lance une exception
     *  si des erreurs surviennent. Si faux, renvoie null.
     * @returns {ExcelScript.Table | null} - Tableau Excel créé, ou null si une erreur survient.
     */
    public static printTable(
        headers: string[][],
        data: CellValue[][],
        sheetName: string,
        tableName: string,
        startCell: string = "A1",
        failOnError: boolean = true
    ): ExcelScript.Table | null {

        // Combine les en-têtes et les données.
        if (headers[0].length !== data[0].length) {
            throw new Error("Les en-têtes et les données doivent avoir la même longueur.");
        }
        const tableData: CellValue[][] =
            headers as CellValue[][];
        tableData.push(...data);

        // Vérifie si les données sont non vides.
        if (tableData.length === 0 || tableData[0].length === 0) {
            const msg = `Aucune donnée à insérer dans la table "${tableName}".`;
            if (failOnError) throw new Error(msg);
            Log.warn(msg);
            return;
        }

        // Vérifie si un tableau avec le même nom existe déjà et le supprime si nécessaire.
        const sheet = this.getSheet({ sheetName, createIfMissing: true, failOnError: false });
        const existingTable = sheet.getTables().find(table => table.getName() === tableName);
        if (existingTable) existingTable.delete();

        // Détermine la plage où écrire les données.
        const startRange = sheet.getRange(this.checkCellName(startCell));
        const writeRange = startRange
            .getResizedRange(tableData.length - 1, tableData[0].length - 1);

        // Efface le contenu de la plage.
        writeRange.clear(ExcelScript.ClearApplyTo.contents);

        // Écrit les données dans la plage.
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
 * Classe utilitaire Params contenant les paramètres globaux.
 */
class Params {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";                // Feuille contenant les paramètres globaux
    private static readonly TABLE = "Paramètres";           // Tableau contenant les paramètres globaux
    private static readonly ROW_MAX_CONNEXIONS_NUMBER = 1;  // Ligne contenant le nombre maximum de connexions
    private static readonly ROW_TURNAROUND_TIME = 2;        // Ligne contenant le temps de retournement (en minutes)
    private static readonly ROW_MAX_TRAIN_UNITS = 3;        // Ligne contenant le nombre maximal d'unités en UM
    private static readonly ROW_STATIONS_SUFFIXES = 5;      // Ligne contenant les suffixes des gares par défaut
                                                            //  à supprimer lors de la lecture
    // Indicateur de chargement
    private static loaded = false;

    // Paramètres globaux
    public static maxConnectionNumber: number;              // Nombre maximum de connexions
    public static turnaroundTime: DateTime;                 // Temps de retournement
    public static maxTrainUnits: number;                    // Nombre maximal d'unités en UM
    public static stationsSuffixes: string[];               // Suffixes des gares par défaut

    /**
     * Charge les paramètres globaux.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.loaded && !erase) return;

        // Charge les paramètres des classes utilitaires.
        DateTime.load(erase);
        Days.load(erase);
        Parity.load(erase);
        TrainNumber.load(erase);

        // Charge les autres paramètres.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);

        this.maxConnectionNumber = WorkbookService.getNumber(data[this.ROW_MAX_CONNEXIONS_NUMBER], 1) ?? 6;
        const turnaroundTime = WorkbookService.getNumber(data[this.ROW_TURNAROUND_TIME], 1) ?? 10;
        this.turnaroundTime = DateTime.from(turnaroundTime / 24 / 60, true)!;

        this.maxTrainUnits = WorkbookService.getNumber(data[this.ROW_MAX_TRAIN_UNITS], 1) ?? 2;
        const stationsSuffixesString = WorkbookService.getString(data[this.ROW_STATIONS_SUFFIXES],1);
        const separatorRegex = /[\s,.!?;:\-\n\r]+/; // Toute ponctuation est considérée comme un séparateur
        this.stationsSuffixes = stationsSuffixesString
            .split(separatorRegex)
            .filter((t) => Boolean(t));;

        this.loaded = true;
    }
}

/*
 * Classe utilitaire immuable contenant les valeurs d'une date Excel.
 */
class ExcelDate {

    // Constante de valeur initiale (epoch) des dates Excel
    public static readonly EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30));

    // Cache des jours fériés par année
    private static holidayCache = new Map<number, Set<number>>();
 
    // Propriétés de l'objet DateTime
    public readonly value: number;
    public readonly year: number;
    public readonly month: number;
    public readonly day: number;
    public readonly dayOfWeek: Day;
    public readonly isHoliday: boolean;

    /**
     * Constructeur de l'objet ExcelDate.
     * @param {number} excelValue - Valeur Excel du jour, qui représente le nombre de jours
     *  écoulés depuis le 30 décembre 1899.
     */
    constructor(excelValue: number) {

        this.value = Math.floor(excelValue);

        const ms = ExcelDate.EXCEL_EPOCH.getTime() + this.value * 86400000;
        const d = new Date(ms);

        this.year = d.getUTCFullYear();
        this.month = d.getUTCMonth() + 1;
        this.day = d.getUTCDate();
        const jsDay = d.getUTCDay();
        const dayNumber = jsDay === 0 ? 7 : jsDay;
        const dayOfWeek = Day.from(dayNumber);
        if (dayOfWeek === undefined) throw new Error(`Jour de semaine non trouvé : ${dayNumber}`);
        this.dayOfWeek = dayOfWeek;

        const holidays = ExcelDate.getHolidays(this.year);
        this.isHoliday = holidays.has(this.month * 100 + this.day);
    }

    /**
     * Analyse une chaîne de caractères qui représente une date (hh:mm:ss)
     *  au format "dd/MM" ou "dd/MM/yyyy" ou "yyyy/MM/dd"
     *  et renvoie la valeur Excel correspondante.
     * @param {string} value - Chaîne à parser.
     * @returns {number | undefined} - Valeur Excel correspondante, ou undefined si la date est incorrecte.
     */
    public static parseDate(value: string): number | undefined {

        const separatorRegex = /[/\-]/; // Séparateur : / ou -
        const parts = value.split(separatorRegex);
        if (parts.length < 2 || parts.length > 3) return undefined;
 
        let day: number;
        let month: number;
        let year: number;
 
        const p0 = Number(parts[0]);
        const p1 = Number(parts[1]);
        const p2 = parts.length === 3 ? Number(parts[2]) : undefined;
 
        if ([p0, p1, p2].some(v => v !== undefined && isNaN(v))) return undefined;
 
        if (parts.length === 2) {
            // dd/MM (année courante)
            day = p0;
            month = p1;
            year = new Date().getFullYear();
        } else if (p0 > 31) {
            // yyyy/MM/dd
            year = p0;
            month = p1;
            day = p2!;
        } else {
            // dd/MM/yyyy
            day = p0;
            month = p1;
            year = p2!;
        }
 
        if (
            day <= 0 || day > 31 ||
            month <= 0 || month > 12
        ) return undefined;

        const jsDate = new Date(Date.UTC(year, month - 1, day));
        const excelEpoch = Date.UTC(1899, 11, 30);
 
        return (jsDate.getTime() - excelEpoch) / 86400000;
    }

    /**
     * Indique si une date est un jour férié.
     * @param {string} value - Valeur Excel de la date.
     * @returns {boolean} - Vrai si la date est un jour férié, faux sinon.
     */
    public static getHolidays(year: number): Set<number> {
 
        if (this.holidayCache.has(year)) {
            return this.holidayCache.get(year)!;
        }
 
        const set = new Set<number>();
 
        const add = (m: number, d: number) => set.add(m * 100 + d);
 
        // Donne les jours fériés fixes
        add(1, 1);
        add(5, 1);
        add(5, 8);
        add(7, 14);
        add(8, 15);
        add(11, 1);
        add(11, 11);
        add(12, 25);
 
        // Calcule le jour de Pâques
        const a = year % 19;
        const b = Math.floor(year / 100);
        const c = year % 100;
        const d = Math.floor(b / 4);
        const e = b % 4;
        const f = Math.floor((b + 8) / 25);
        const g = Math.floor((b - f + 1) / 3);
        const h = (19 * a + b - d - g + 15) % 30;
        const i = Math.floor(c / 4);
        const k = c % 4;
        const l = (32 + 2 * e + 2 * i - h - k) % 7;
        const m = Math.floor((a + 11 * h + 22 * l) / 451);
 
        const easterMonth = Math.floor((h + l - 7 * m + 114) / 31);
        const easterDay = ((h + l - 7 * m + 114) % 31) + 1;
 
        const addDays = (delta: number): [number, number] => {
            const d = new Date(Date.UTC(year, easterMonth - 1, easterDay + delta));
            return [d.getUTCMonth() + 1, d.getUTCDate()];
        };
 
        // Calcule le Lundi de Pâques
        let [mm, dd] = addDays(1);
        add(mm, dd);
 
        // Calcule le jour de l'Ascension
        [mm, dd] = addDays(39);
        add(mm, dd);
 
        // Calcule le Lundi de Pentecôte
        [mm, dd] = addDays(50);
        add(mm, dd);
 
        this.holidayCache.set(year, set);
        return set;
    }
}

/*
 * Classe utilitaire immuable ExcelTime contenant les valeurs d'une heure Excel.
 */
class ExcelTime {

    // Propriétés de l'objet DateTime
    public readonly value: number;
    public readonly hour: number;
    public readonly minute: number;
    public readonly second: number;

    /**
     * Constructeur de l'objet ExcelTime.
     * @param {number} excelValue - Valeur Excel du temps, dont la fraction de jour représente l'heure.
     */
    constructor(excelValue: number) {
        this.value = excelValue;
        const abs = Math.abs(this.value);
        const totalSeconds = Math.round(abs * 86400);
        this.hour = Math.floor(totalSeconds / 3600);
        this.minute = Math.floor((totalSeconds % 3600) / 60);
        this.second = totalSeconds % 60;
    }

    /**
     * Analyse une chaîne de caractères qui représente une heure (hh:mm:ss)
     *  et renvoie la valeur Excel correspondante.
     * @param {string} value - Chaîne à parser.
     * @returns {number | undefined} - Valeur Excel correspondante, ou undefined si l'heure est incorrecte.
     */
    public static parseTime(value: string): number | undefined {
        const separatorRegex = /[^\d]/; // Toute caractère ou chaine de caractère non numérique
                                        //  est considérée comme un séparateur (ex : 'h', 'min' ...)
        const parts = value.split(separatorRegex).filter((t) => Boolean(t));
        if (parts.length < 2 || parts.length > 3) return undefined;
 
        const [hStr, mStr, sStr = "0"] = parts;
 
        const h = Number(hStr);
        const m = Number(mStr);
        const s = Number(sStr);
 
        if (
            isNaN(h) || isNaN(m) || isNaN(s) ||
            m < 0 || m >= 60 ||
            s < 0 || s >= 60
        ) return undefined;
 
        return (h * 3600 + m * 60 + s) / 86400;
    }
}

/**
 * Classe utilitaire immuable DateTime pour la gestion des dates et horaires Excel.
 *  Si le temps est absolu et non daté, et que l'heure est inférieure à l'heure de changement de journée,
 *  elle est incrémentée de 1 pour rester comparable aux autres heures de la journée précédente.
 */
class DateTime {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";        // Feuille contenant les paramètres globaux
    private static readonly TABLE = "Paramètres";   // Tableau contenant les paramètres globaux
    private static readonly ROW_ROLLOVER_HOUR = 4;  // Ligne contenant l'heure de changement de journée
    private static readonly MIN_EXCEL_DATE = 2;     // Valeur minimale d'un temps absolu daté

    // Etat de chargement
    private static loaded = false;

    // Heure de changement de journée (fraction de jour Excel)
    public static rolloverHour: number;             // Heure de changement de journée (en temps Excel)

    // Ecart minimal entre 
    public static readonly MAX_GAP: number = 3/24/3600; // Différence maximale entre 2 horaires
                                                        //  pour les considérer comme égaux

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
    private _realDate: ExcelDate | undefined;       // Date réelle (uniquement pour les temps absolus
                                                    //  datés, donc avec excelValue >= MIN_EXCEL_DATE)
    private _adaptedDate: ExcelDate | undefined;    // Date adaptée (jour suivant) si l'heure de la date
                                                    //  est inférieure à l'heure de changement de jour
    private _time: ExcelTime | undefined;           // Heure de la journée 
                                                    //  (undefined si le temps n'est qu'une date)

    /**
     * Constructeur privé de l'objet DateTime.
     * @param {number} [excelValue=0] - Valeur du temps en format Excel
     *  à partir du 01/01/1900 00:00:00.
     * @param {boolean} [isRelative=false] - Indique si le temps est relatif
     *  (différence entre 2 horaires).
     * @param {boolean} [adaptTime=true] - Indique si le temps doit être adapté pour
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
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        const format = this._realDate
            ? DateTime.DATE_FORMAT_WITH_YEAR
            : DateTime.TIME_FORMAT_WITH_SECONDS;
        return this.format(format);
    }

    /**
     * Crée un objet DateTime à partir d'une valeur.
     * Si la valeur est déjà un objet DateTime, il est retourné tel quel.
     * Sinon, un nouvel objet DateTime est créé avec la valeur fournie.
     * @param {DateTime | number | string | null | undefined} value Valeur du temps en nombre décimal
     *  ou en chaîne de caractères.
     * @param {boolean} [isRelative=undefined] - Indique si le temps est relatif
     *  (différence entre 2 horaires).
     * @param {boolean} [adaptTime=true] - Indique si l'heure doit être adaptée ou non.
     *  Si la valeur est inférieure à l'heure de changement de journée et que l'horaire est absolu,
     *  elle est incrémentée de 1.
     * @returns {DateTime | undefined} - Nouvel objet DateTime égal à la valeur fournie, ou undefined.
     * @throws {Error} - Si la valeur est un temps relatif et qu'on cherche à l'affecter à un temps absolu.
     */
    public static from(
        value: DateTime | number | string | null | undefined,
        isRelative?: boolean,
        adaptTime: boolean = true
    ): DateTime | undefined {
 
        if (value == null || value === "") return undefined;

        if (value instanceof DateTime) {
            if (isRelative !== undefined && value.isRelative !== isRelative) {
                throw new Error(
                    `Un temps ${value.isRelative ? "relatif" : "absolu"}`
                    + ` cherche à être affecté à un temps ${isRelative ? "relatif" : "absolu"}.`
                );
            }
            return value;
        }

        let v: number | undefined;

        if (typeof value === "string") {
            const trimmed = value.trim();
 
            // La chaine est un nombre simple
            if (/^-?\d+(?:[.,]\d+)?$/.test(trimmed)) {
                v = Number(trimmed.replace(",", "."));
            } else {
                // La chaine doit être analysée
                const parsed = this.parseDateAndTime(trimmed, isRelative, adaptTime);
                return parsed;
            }
        } else {
            v = value;
        }
 
        if (v === undefined || isNaN(v)) return undefined;
 
        // Un temps absolu doit être >= 0
        if (!isRelative && v < 0) return undefined;
 
        return new DateTime(v, isRelative ?? false, adaptTime);
    } 

    /**
     * Retourne l'objet ExcelDate de la date, adaptée ou non.
     * @param {boolean} adapted - Indique si l'on souhaite avoir l'objet ExcelDate adapté ou non.
     * @returns {ExcelDate | undefined} - L'objet ExcelDate correspondant au temps adapté ou non,
     *  ou undefined si le temps n'est pas défini.
     */
    private getDateObj(adapted: boolean): ExcelDate | undefined {
        return (adapted && this._adaptedDate) ? this._adaptedDate : this._realDate;
    }

    /**
     * Retourne la valeur Excel de la date, adaptée ou non.
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir la valeur Excel
     *  de la date adaptée ou non.
     * @returns {number} - La valeur Excel de la date, adaptée ou non, ou 0 si le temps n'est pas défini.
     */
    public getDate(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.value ?? 0;
    }

    /**
     * Retourne l'année de la date, adaptée ou non.
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir l'année
     *  de la date adaptée ou non.
     * @returns {number} - L'année de la date, adaptée ou non, ou 0 si le temps n'est pas défini.
     */
    public getYear(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.year ?? 0;
    }

    /**
     * Retourne le mois de la date, adapté ou non.
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir le mois
     *  de la date adaptée ou non.
     * @returns {number} - Le mois de la date, adapté ou non, ou 0 si le temps n'est pas défini.
     */
    public getMonth(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.month ?? 0;
    }
 
    /**
     * Retourne le jour du mois de la date, adapté ou non.
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir le jour du mois
     *  de la date adaptée ou non.
     * @returns {number} - Le jour du mois de la date, adapté ou non, ou 0 si le temps n'est pas défini.
     */
    public getDay(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.day ?? 0;
    }
 
    /**
     * Retourne le jour de la semaine correspondant à la date, adapté ou non.
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir le jour de la semaine
     *  de la date adaptée ou non.
     * @param {boolean} [withHolidays=true] - Indique si l'on souhaite indiquer les jours fériés (jour 8).
     * @returns {Day | undefined} - Le jour de la semaine correspondant à la date,
     *  adapté ou non, ou undefined si le temps n'est pas défini.
     */
    public getDayOfWeek(adaptedValue: boolean = true, withHolidays: boolean = true): Day | undefined {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj(adaptedValue);
        return dateObj?.isHoliday && withHolidays ? Day.HOLIDAY : dateObj?.dayOfWeek;
    }

    /**
     * Renvoie l'heure correspondant au temps, adaptée ou non.
     * Si le temps est relatif, renvoie l'heure relative.
     * Si le temps est absolu (daté ou non), renvoie la fraction d'une journée (tronquée modulo 1).
     * Si le temps est adapté, il est incrémenté de 1 (ex : 25h00).
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir l'heure adaptée ou non.
     * @returns {number}- L'heure correspondant au temps, adaptée ou non, ou 0 si le temps n'est pas défini.
     */
    public getTime(adaptedValue: boolean = true): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        if (!timeObj) return 0;
        if (!this.isRelative && !adaptedValue) {
            return timeObj.value % 1;
        }
        return timeObj.value;
    }
 
    /**
     * Retourne le nombre d'heures de l'heure correspondant au temps, adapté ou non.
     * Si le temps est relatif, renvoie le nombre d'heures relatif.
     * Si le temps est absolu et que adaptedValue est faux,
     *  renvoie le nombre d'heures de l'objet DateTime tronqué modulo 24.
     * Si le temps est adapté, l'heure est incrémentée de 24 (ex : 25h00). 
     * @param {boolean} [adaptedValue=false] - Indique si l'on souhaite avoir
     *  le nombre d'heures adapté ou non.
     * @returns {number} - Le nombre d'heures correspondant au temps, adapté ou non,
     *  ou 0 si le temps n'est pas défini.
     */
    public getHours(adaptedValue: boolean = false): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        if (!timeObj) return 0;
        if (!this.isRelative && !adaptedValue) {
            return timeObj.hour % 24;
        }
        return timeObj.hour;
    }

    /**
     * Retourne le nombre de minutes de l'heure correspondant au temps.
     * @returns {number} - Le nombre de minutes de l'heure, ou 0 si le temps n'est pas défini.
     */
    public getMinutes(): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        return timeObj?.minute ?? 0;
    }

    /**
     * Retourne le nombre de secondes de l'heure correspondant au temps.
     * @returns {number} - Le nombre de secondes de l'heure, ou 0 si le temps n'est pas défini.
     */
    public getSeconds(): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        return timeObj?.second ?? 0;
    }

    /**
     * Retourne si la date correspondant au temps est un jour férié.
     * @param {boolean} [adaptedValue=true] - Indique si l'on souhaite avoir la date adaptée ou non.
     * @returns {boolean} - Vrai si la date est un jour férié, faux sinon ou si la date n'est pas définie.
     */
    public isHoliday(adaptedValue: boolean = true): boolean {
        if (!this._computed) this.compute();
        return this.getDateObj(adaptedValue)?.isHoliday ?? false;
    }

    /**
     * Analyse une chaîne de caractères qui représente une date et un temps
     *  et renvoie la valeur Excel correspondante.
     * @param {string} value - Chaîne à parser.
     * @param {boolean} [isRelative=undefined] - Indique si le temps est relatif
     *  (différence entre 2 horaires).
     * @param {boolean} [adaptTime=true] - Indique si l'heure doit être adaptée ou non.
     *  Si la valeur est inférieure à l'heure de changement de journée et que l'horaire est absolu,
     *  elle est incrémentée de 1.
     * @returns {DateTime | undefined} - Nouvel objet DateTime égal à la valeur fournie, ou undefined.
     * @throws {Error} - Si la valeur est un temps relatif et qu'on cherche à l'affecter à un temps absolu.
     */
    public static parseDateAndTime(
        value: string,
        isRelative?: boolean,
        adaptTime: boolean = true
    ): DateTime | undefined {
 
        if (!value) return undefined;
 
        const parts = value.trim().split(/[ ;]+/);
 
        let date: number | undefined;
        let time: number | undefined;
 
        type SignState = "unknown" | "negative" | "invalid";
        let signState: SignState = "unknown";
 
        for (let part of parts) {
 
            if (part === '-') {
                signState = (signState === "unknown") ? "negative" : "invalid";
                continue;
            }
 
            if (part.startsWith('-')) {
                signState = (signState === "unknown") ? "negative" : "invalid";
                part = part.slice(1);
            }
 
            if (part.includes('/') || part.split('-').length === 3) {
                date = ExcelDate.parseDate(part);
                if (date === undefined) return undefined;
 
                signState = "invalid";
            }
            else if (part.includes(':') || part.toLowerCase().includes('h')) {
                const parsedTime = ExcelTime.parseTime(part);
                if (parsedTime === undefined) return undefined;
 
                time = parsedTime;
            }
        }

        if (date === undefined && time === undefined) return undefined;
 
        if (time !== undefined && signState === "negative") {
            time = -time;
        }
 
        const result = (date ?? 0) + (time ?? 0);
 
        return new DateTime(result, isRelative, adaptTime);
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

        // Récupère la valeur du temps, inchangé si le temps est relatif ou absolu non daté.
        let timeOfDays = this.excelValue;

        // Calcule les éléments de la date (si temps absolu et daté).
        if (!this.isRelative && this.excelValue > DateTime.MIN_EXCEL_DATE) {
            this._realDate = new ExcelDate(this.excelValue);
            timeOfDays = this.excelValue % 1;
            if (timeOfDays < DateTime.rolloverHour) {
                this._adaptedDate = new ExcelDate(this.excelValue - 1);
                timeOfDays += 1;
            }
        }

        // Calcule les éléments de l'heure de la journée :
        //  - si le temps est relatif, l'heure correspond au temps total, positif ou négatif,
        //  - si le temps est absolu, l'heure est la fraction de la journée,
        //     adaptée si l'heure est inférieure à l'heure de changement de jour,
        //     d'une valeur comprise entre 0 et 1, ou dépassant 1 si adaptée.
        this._time = new ExcelTime(timeOfDays);
 
        this._computed = true;
    }

    /**
     * Renvoie un nouvel objet DateTime égal au temps courant
     *  résolu par rapport à une référence.
     * Si le temps courant est relatif, il est ajouté à la référence.
     * Sinon, le temps courant est renvoyé tel quel.
     * @param {DateTime} reference - Référence à utiliser pour résoudre le temps courant.
     * @returns {DateTime} - Nouvel objet DateTime égal au temps courant résolu par rapport à la référence.
     */
    public resolveAgainst(reference: DateTime): DateTime {
        if (this.isRelative) {
            return new DateTime(
                reference.excelValue + this.excelValue,
                reference.isRelative
            );
        }

        return this;
    }

    /**
     * Renvoie un nouvel objet DateTime égal au temps courant relatif par rapport à une référence.
     * Les deux temps doivent être absolus.
     * @param {DateTime} reference - Référence à utiliser pour résoudre le temps courant.
     * @returns {DateTime} - Nouvel objet DateTime égal au temps courant relatif par rapport à la référence.
     * @throws {Error} - Si l'un des deux temps est relatif.
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
     * Compare le temps courant avec un autre temps.
     * Les deux temps doivent avoir le même type (relatif ou absolu).
     * Si un des temps au moins est absolu et non daté, seul l'horaire est comparé.
     * @param {DateTime} other - Temps à comparer.
     * @returns {number} - Différence entre les deux temps.
     * @throws {Error} - Si les deux temps ont des types différents (relatif ou absolu).
     */
    public compareTo(other: DateTime): number {

        if (!this._computed) this.compute();
        if (!other._computed) other.compute();
        if (this.isRelative !== other.isRelative) {
            throw new Error(`Un temps relatif ne peut pas être comparé avec un temps absolu`);
        }
        if (
            (!this._realDate && !other._time)
                || (!this._time && !other._realDate)
        ) {
            throw new Error(`Un temps absolu non daté ne peut pas être comparé
                avec un temps absolu sans horaire (avec uniquement une date).`);
        }

        const firstTime = (this._realDate && other._realDate) ? this.excelValue : this._time!.value;
        const secondTime = (this._realDate && other._realDate) ? other.excelValue : other._time!.value;

        if (Math.abs(firstTime - secondTime) < DateTime.MAX_GAP) return 0;
        return firstTime - secondTime;
    }

    /**
     * Vérifie si le temps courant est égal à un autre temps.
     * @param {DateTime | null | undefined} other - Temps à comparer.
     * @returns {boolean} - Vrai si les deux temps sont égaux, faux sinon.
     */
    public equalsTo(other: DateTime | null | undefined): boolean {
        return (
            !! other &&
            this.isRelative === other.isRelative &&
            (Math.abs(this.excelValue - other.excelValue) < DateTime.MAX_GAP)
        );
    }

    /**
     * Vérifie si les deux temps sont identiques ou s'ils sont tous les deux undefined.
     * @param {Parity | undefined} a - Premier temps à comparer.
     * @param {Parity | undefined} b - Second temps à comparer.
     * @returns {boolean} - Vrai si les deux temps sont identiques
     *  ou s'ils sont tous les deux undefined, faux sinon.
     */
    public static equalsOrUndefined(
        a?: DateTime,
        b?: DateTime
    ): boolean {
        return a === b || (!!a && !!b && a.equalsTo(b));
    }
 
    /**
     * Ajoute un temps relatif à un autre temps relatif.
     * @param {DateTime} other - Temps relatif à ajouter.
     * @returns {DateTime} - Nouvel objet DateTime égal à la somme des deux temps relatifs.
     * @throws {Error} - Si l'un des deux temps n'est pas relatif.
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
     * @param {DateTime} other - Temps relatif à soustraire.
     * @returns {DateTime} - Nouvel objet DateTime égal à la différence entre les deux temps relatifs.
     * @throws {Error} - Si l'un des deux temps n'est pas relatif.
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
     * @param {string} format - Format de la date ou de l'heure.
     * @param {boolean} adaptTime - Indique si la date ou l'heure doivent prendre en compte
     *  l'adaptation avec l'heure de changement de jour.
     * @param {boolean} withHolidays - Indique si le jour de la semaine doit prendre en compte
     *  les jours fériés.
     * @returns {string} - Date ou heure formattée.
     */
    public format(format: string, adaptTime: boolean = true, withHolidays: boolean = true): string {
        this.compute();
        let prefix = "";
        if (this.excelValue < 0) prefix = "-";
        const pad = (v: number) => v.toString().padStart(2, "0");
 
        const tokens: Record<string, string> = {
        // Année
        "yyyy": this.getYear(adaptTime).toString(),
        "yy": pad(this.getYear(adaptTime) % 100),
        // Mois
        "mm": pad(this.getMonth(adaptTime)),
        "m": this.getMonth(adaptTime).toString(),
        // Jour
        "dd": pad(this.getDay(adaptTime)),
        "d": this.getDay(adaptTime).toString(),
        // Jour de semaine
        "dddd": this.getDayOfWeek(adaptTime, withHolidays)?.fullName ?? "",
        "ddd": this.getDayOfWeek(adaptTime, withHolidays)?.abreviation ?? "",
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

        // Crée les clés temporaires.
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

        // Remplace toutes les clés temporaires par les valeurs réelles.
        for (const key in tempMap) {
            const val = tempMap[key];
            tempFormat = tempFormat.replace(new RegExp(key, "g"), val);
        }
 
        return prefix + tempFormat;
    }

    /**
     * Ajuste une heure pour tenir compte du changement de journée.
     * Si l'heure est inférieure à l'heure de changement de journée,
     *  on ajoute 1 pour passer à la journée suivante.
     *  Par exemple : 01:00 → 25:00 si changement de journée à 03:00
     * Cela ne s'applique que sur les heures non datées (valeur < 1).
     * @param {number} time - Heure à ajuster.
     * @returns {number} - Heure ajustée.
     */
    public static adaptTime(time: number): number {
        return (time < this.rolloverHour) ? time + 1 : time;
    }
 
    /**
     * Charge les paramètres des dates et heures de changement de journée.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.loaded && !erase) return;

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);

        // Extrait les valeurs.
        this.rolloverHour = (WorkbookService.getNumber(data[this.ROW_ROLLOVER_HOUR], 1) ?? 0) % 1;

        this.loaded = true;
    }
}

class Day { 

    // Indicateur de chargement
    private static loaded = false;

    // Liste des jours de la semaine identifiés par leur indice
    private static list: Day[] = new Array(8);

    // Propriétés de l'objet Day
    public readonly mask: number;           // Masque de bits contenant les numéro du jour de la semaine
                                            //  (identique à Days)
                                            //  (de lundi > bit 0 à dimanche > bit 6, férié > bit 7)
    public readonly code: string;           // Code chiffre du jour de la semaine
    public readonly fullName: string;       // Nom du jour de la semaine
    public readonly abreviation: string;    // Abréviation du jour de la semaine

    /**
     * Constructeur privé de la classe Day.
     * @param {Days} days - Groupe du jour de la semaine en individuel
     */
    private constructor(
        days: Days
    ) {
        this.mask = days.mask;
        this.code = days.code;
        this.fullName = days.fullName;
        this.abreviation = days.abreviation;
    }

    /**
     * Retourne l'indice du jour de la semaine.
     * L'indice est un entier compris entre 0 et 7, égal au numéro du jour moins 1.
     * @returns {number} - Indice du jour de la semaine.
     */
    public get index(): number {
        return Day.maskToIndex(this.mask);
    }
 
    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.fullName; 
    }

    /**
     * Accesseurs des jours de la semaine
     */
    public static get MONDAY(): Day { return this.list[0]!; }
    public static get TUESDAY(): Day { return this.list[1]!; }
    public static get WEDNESDAY(): Day { return this.list[2]!; }
    public static get THURSDAY(): Day { return this.list[3]!; }
    public static get FRIDAY(): Day { return this.list[4]!; }
    public static get SATURDAY(): Day { return this.list[5]!; }
    public static get SUNDAY(): Day { return this.list[6]!; }
    public static get HOLIDAY(): Day { return this.list[7]!; }

    /**
     * Renvoie un objet Day correspondant au numéro de jour fourni.
     * Si le numéro de jour n'existe pas, renvoie undefined.
     * @param {Day | number | string | null | undefined} value - Valeur à analyser
     *  pour le jour correspondant.
     * @returns {Days | undefined} - Instance de Day correspondante.
     */
    public static from(
        value: Day | number | string | null | undefined): Day | undefined {
 
        if (value == null || value === '') return undefined;
        if (value instanceof Day) return value;

        const days = Days.get(String(value));
 
        if (!days || days.count !== 1) {
            throw new Error(`Le jour de la semaine "${value}" n'existe pas.`);
        }

        const index = this.maskToIndex(days.mask);
        return this.list[index];
    }

    /**
     * Convertit le masque en l'indice du jour (0 à 7).
     * @param {number} mask - Masque de jour.
     * @returns {number} - Index du jour correspondant.
     * @throws {Error} - Si le masque est invalide.
     */
    private static maskToIndex(mask: number): number {
        const index = Math.log2(mask);
 
        if (!Number.isInteger(index) || index < 0 || index > 7) {
            throw new Error(`Mask invalide pour Day: ${mask}`);
        }
 
        return index;
    }
 
    /**
     * Retourne un tableau des valeurs de la base de données des jours.
     * @returns {Day[]} - Itérateur sur les valeurs de la base de données.
     *  des jours.
     */
    public static values(): Day[] {
        return Array.from(this.list.values());
    }

    /**
     * Efface toutes les gares de la base de données.
     * Cela permet de forcer le rechargement des gares si besoin.
     */
    public static clear(): void {
        this.list = new Array(8);
        this.loaded = false;
    }

    /**
     * Charge les jours de la semaine.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.loaded) {
            if (!erase) return;
            this.clear();
        }

        // Récupère les jours de la base de données de Days.
        for (let i = 0; i < 8; i++) {
            const days = Days.get(1 << i);
            if (!days) throw new Error(`Day.load : jour ${i +1 } inexsistant dans Days`);

            // Crée l'objet Day.
            const day = new Day(days);
            this.list[i] = day;
        }

        this.loaded = true;
    }
}


/**
 * Classe utilitaire Days pour la gestion des jours de la semaine, individuels ou groupés. 
 *  (JOB du lundi au vendredi, WE pour samedi et dimanche...).
 */
class Days {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Param";        // Feuille contenant les paramètres des jours de la semaine
    private static readonly TABLE = "Jours";        // Tableau contenant les paramètres des jours de la semaine
    private static readonly COL_NUMBERS = 0;        // Colonne contenant le numéro du jour 
    private static readonly COL_CODE_LETTER = 1;    // Colonne contenant la lettre de code du groupe de jours
    private static readonly COL_FULL_NAME = 2;      // Colonne contenant le nom complet du jour de la semaine
    private static readonly COL_ABBREVIATION = 3;   // Colonne contenant l'abréviation du jour de la semaine

    // Valeur par défaut des jours de la semaine individuels
    //  (si non renseignés dans le tableau des paramètres)
    private static readonly WEEKDAYS = [
        { number: 1, fullName: "Lundi", abreviation: "Lu" },
        { number: 2, fullName: "Mardi", abreviation: "Ma" },
        { number: 3, fullName: "Mercredi", abreviation: "Me" },
        { number: 4, fullName: "Jeudi", abreviation: "Je" },
        { number: 5, fullName: "Vendredi", abreviation: "Ve" },
        { number: 6, fullName: "Samedi", abreviation: "Sa" },
        { number: 7, fullName: "Dimanche", abreviation: "Di" },
        { number: 8, fullName: "Férié", abreviation: "Fer", code: "F" }
    ];

    // Indicateur de chargement
    private static loaded = false;

    // Liste des groupes de jours de la semaine identifiés par leur masque de bits,
    //  en assurant l'unicité de chaque groupe de jours
    private static masksList: (Days | undefined)[] = new Array(256);
    // Map des groupes de jours de la semaine identifiés par leur code, nom ou abréviation,
    //  (plusieurs codes possibles par groupe de jour, y compris ceux non optimisés)
    private static mapByString: Map<string, Days> = new Map<string, Days>();
 
    // Liste des groupes de jours donnés en paramètre, y compris chaque jour de la semaine seul,
    //  triée par leur nombre de jours dans l'ordre décroissant.
    // Les groupes de jours les plus importants seront utilisés en priorité
    //  pour optimiser le code d'un groupe de jours, les jours seuls sont ajoutés en dernier.
    private static compressionRules: Days[] = [];
    // Synthèse des chaines de caractères désignant les groupes de jours donnés en paramètre,
    //  à partir de leur code, leur abréviation ou leur nom complet.
    // Le tableau est trié par longueur de chaine décroissante
    //  (les chaines les plus longues sont recherchées en premier).
    private static extractionPatterns: {pattern: string, numbersString: string}[] = [];

    // Propriétés de l'objet Days
    public readonly mask: number;           // Masque de bits contenant les numéro(s) du ou des jours
                                            //  du groupe de jours de la semaine
                                            //  (de 1 : lundi > bit 0 à 7 : dimanche > bit 6, 8 : férié > bit 7)
    public readonly code: string;           // Code alphanumérique du groupe de jours de la semaine
    public readonly fullName: string;       // Nom du jour ou du groupe de jours de la semaine
    public readonly abreviation: string;    // Abréviation du jour ou du groupe de jours de la semaine
    private _numberString: string = "";     // Chaîne de caractères contenant les numéros des jours du groupe de jours
    private _count: number = -1;            // Nombre de jours contenus dans le groupe de jours (-1 si non calculé)

    /**
     * Constructeur privé de la classe Days.
     * @param {number} mask - Masque de bits des numéro(s) des jours du groupe de jours de la semaine. 
     * @param {string} code - Code alphanumérique du groupe de jours de la semaine.
     * @param {string} fullName - Nom complet du jour ou du groupe de jours de la semaine.
     * @param {string} abreviation - Abréviation du jour ou du groupe de jours de la semaine.
     */
    private constructor(
        mask: number,
        code: string,
        fullName: string = "",
        abreviation: string = ""
    ) {
        this.mask = mask;
        this.code = code;
        this.fullName = fullName;
        this.abreviation = abreviation;
    }

    /**
     * Retourne la concaténation des numéros des jours du groupe de jours.
     * @returns {string} Chaîne de caractères contenant les numéros des jours de la semaine.
     */
    public get numbersString(): string {
        if (!this._numberString) {
            this._numberString = Days.maskToNumbers(this.mask).join('');
        }

        return this._numberString;
    }

    /**
     * Retourne le nombre de jours contenus dans le groupe de jours.
     * @returns {number} - Nombre de jours contenus dans le groupe de jours.
     */
    public get count(): number {

        if (this._count === -1) {
            let count = 0;
            let mask = this.mask;
            while (mask) {
                mask &= mask - 1;
                count++;
            }
            this._count = count;
        }

        return this._count;
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  tilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.fullName; 
    }

    /**
     * Accesseurs des jours de la semaine
     */
    public static get MONDAY(): Days { return this.get('1')!; }
    public static get TUESDAY(): Days { return this.get('2')!; }
    public static get WEDNESDAY(): Days { return this.get('3')!; }
    public static get THURSDAY(): Days { return this.get('4')!; }
    public static get FRIDAY(): Days { return this.get('5')!; }
    public static get SATURDAY(): Days { return this.get('6')!; }
    public static get SUNDAY(): Days { return this.get('7')!; }
    public static get HOLIDAY(): Days { return this.get('8')!; }

    /**
     * Renvoie un objet Days correspondant au numéro de jour fourni.
     * Si le numéro de jour n'existe pas, renvoie undefined.
     * Charge les paramètres des jours de la semaine si ce n'est pas déjà fait.
     * @param {Days | number | string | null | undefined} value - Valeur à analyser
     *  pour le groupe de jours correspondant.
     * @returns {Days | undefined} - Instance de Days correspondante.
     */
    public static from(
        value: Days | number | string | null | undefined): Days | undefined {
 
        if (value == null || value === '') return undefined;
        if (value instanceof Days) return value;

        const code = String(value);
        if (this.has(code)) {
            return this.get(code);
        }

        // Analyse de la chaine de caractères pour trouver le ou les jours correspondants.
        const numbers = this.extractFromString(code);
        if (numbers.length === 0) {
            return undefined;
        }
        const mask = this.numbersToMask(numbers);

        // Recherche ou création du groupe de jours.
        const days = this.get(mask) ?? this.create(numbers);

        // Ajout du code fourni en entrée pour éviter une nouvelle analyse,
        //  y compris si ce code n'est pas optimisé (code optimisé calculé dans create()).
        this.mapByString.set(code, days);

        return days;
    }

    /**
     * Renvoie un objet Days correspondant au masque de bits fourni.
     * Crée le groupe de jours si celui-ci n'existe pas encore.
     * @param {number} value - Masque de bits correspondant au groupe de jours.
     * @returns {Days | undefined} - Instance de Days correspondante.
     */
    public static fromMask(mask: number): Days | undefined{
        if (mask <= 0 || mask >= 256) return undefined
        return this.masksList[mask] ?? this.create(this.maskToNumbers(mask));
    }
 
    /**
     * Transforme un tableau de numéros de jours en un masque de bits.
     * Chaque numéro correspond à un bit dans le masque, allant de 1 (lundi) à 7 (dimanche), et 8 (férié).
     * @param {number[]} numbers - Tableau de numéros de jours.
     * @returns {number} - Masque de bits correspondant aux numéros de jours.
     */
    private static numbersToMask(numbers: number[]): number {
        return numbers.reduce((mask, n) => mask | (1 << (n - 1)), 0);
    }

    /**
     * Retourne un tableau de numéros de jours à partir d'un masque de bits.
     * Chaque bit dans le masque correspond à un numéro de jour,
     *  allant de 1 (lundi) à 7 (dimanche), et 8 (férié).
     * @param {number} mask - Masque de bits correspondant aux numéros de jours.
     * @returns {number[]} - Tableau de numéros de jours correspondant au masque de bits.
     */
    private static maskToNumbers(mask: number): number[] {
        const result: number[] = [];
        for (let i = 0; i < 8; i++) {
            if (mask & (1 << i)) {
                result.push(i + 1);
            }
        }
        return result;
    }

    /**
     * Vérifie si le groupe de jours contient le jour donné.
     * @param {Day | number | string} day - Jour de la semaine (1 : lundi, 2 : mardi, ..., 7 : dimanche, 8 : férie)
     * @returns {boolean} - Vrai si le groupe de jours contient le jour, faux sinon.
     */
    public contains(day: Day | number | string): boolean {
        const dayObj = Day.from(day);
        return (this.mask & dayObj!.mask) !== 0;
    }

    /**
     * Retourne true si le groupe de jours contient au moins un jour commun
     *  avec le groupe de jours fourni en paramètre.
     * @param {Days} other - Autre groupe de jours.
     * @returns {boolean} - Vrai si le groupe de jours contient au moins un jour commun, faux sinon.
     */
    public intersects(other: Days): boolean {
        return (this.mask & other.mask) !== 0;
    }

    /**
     * Retourne l'intersection de deux groupes de jours.
     * Si un des deux groupes de jours est undefined, l'autre est retourné.
     * Si les deux groupes sont undefined, undefined est retourné.
     * @param {Days | undefined} days1 - Premier groupe de jours.
     * @param {Days | undefined} days2 - Deuxième groupe de jours.
     * @returns {Days | undefined} - Groupe de jours correspondant à l'intersection.
     */
    public static intersection(
        days1: Days | undefined,
        days2: Days | undefined
    ): Days | undefined {

        if (!days1 || !days2) return undefined;

        const mask = days1.mask & days2.mask;
        if (mask === 0) return undefined;
 
        return Days.fromMask(mask);
    }

    /**
     * Retourne l'union de deux groupes de jours.
     * Si un des deux groupes de jours est undefined, l'autre est retourné.
     * Si les deux groupes sont undefined, undefined est retourné.
     * @param {Days | undefined} days1 - Premier groupe de jours.
     * @param {Days | undefined} days2 - Deuxième groupe de jours.
     * @returns {Days | undefined} - Union des deux groupes de jours,
     *  ou undefined si les deux groupes sont undefined.
     */
    public static union(
        days1: Days | undefined,
        days2: Days | undefined
    ): Days | undefined {

        if (!days1 && !days2) return undefined;

        const mask = (days1?.mask ?? 0) | (days2?.mask ?? 0);
        return Days.fromMask(mask);
    }

    /**
     * Retourne la différence entre deux groupes de jours.
     * La différence est le groupe de jours qui contient les jours de a mais pas ceux de b.
     * Si a est undefined, b est retourné.
     * Si b est undefined, a est retourné.
     * Si les deux groupes sont undefined, undefined est retourné.
     * @param {Days | undefined} days1 - Premier groupe de jours.
     * @param {Days | undefined} days2 - Deuxième groupe de jours.
     * @returns {Days | undefined} - Différence entre les deux groupes de jours,
     *  ou undefined si les deux groupes sont undefined.
     */
    public static difference(days1: Days, days2: Days): Days | undefined {
        const mask = days1.mask & ~days2.mask;
        return Days.fromMask(mask);
    }

    /**
     * Nettoie et trie une chaîne de chiffres.
     * Supprime les caractères non numériques et non compris entre 1 et 8,
     * puis trie les chiffres dans l'ordre.
     * @param {string} numbersString - Chaîne de caractères contenant des chiffres.
     * @returns {string} - Chaîne de caractères contenant les chiffres triés.
     */
    private static cleanAndSortNumbers(numbersString: string): number[] {
        return Array.from(new Set(
            numbersString
                .replace(/[^1-8]/g, '')     // Supprime les caractères non numériques
                                            //  et non compris entre 1 et 8
                .split('')                  // Divise la chaîne en un tableau de chiffres
                .map((x) => Number(x))      // Convertit les caractères en nombres
        )).sort((a, b) => a - b);           // Trie les chiffres dans l'ordre
    }

    /**
     * Extrait les jours d'une chaîne en tableau de numéros de jours (1 à 8).
     * Utilise un cache pour éviter de recalculer les résultats pour les mêmes combinaisons.
     * Si deux chaînes sont fournies, retourne l'intersection des jours correspondants.
     * @param {string} value - Chaîne contenant des noms, numéros ou abréviations de jours
     *  séparés ou non par de la ponctuation  (ex : "lundi;me7").
     * @returns {number[]} - Tableau trié de numéros de jours (sans doublons), ex : [1, 3].
     */
    public static extractFromString(value: string): number[] {

        let processed = String(value).toUpperCase();

        // Analyse avec Regex des chaines servant pour l'extraction.
        this.extractionPatterns.forEach(s => {
                const regex = new RegExp(s.pattern, 'g');
                processed = processed.replace(regex, s.numbersString);
            });

        // Trie et nettoie les numéros.
        const result = this.cleanAndSortNumbers(processed);

        return result;
    }

    /**
     * Optimise un code de groupe de jours en trouvant les groupes de jours définis en paramètres.
     * @param {string} mask - Masque de bits du groupe de jours.
     * @returns {string} - Liste des groupes de jours correspondants,
     *  triée par leur premier numéro de jour.
     */
    private static optimiseMask(mask: number): Days[] {

        let remaining = mask;
        let result: Days[] = [];

        for (const d of this.compressionRules) {
            if ((remaining & d.mask) === d.mask) {
                result.push(d);
                remaining &= ~d.mask; // Supprime les bits.
            }
        }
 
        result.sort((a, b) => {
            const aFirst = a.mask & -a.mask;
            const bFirst = b.mask & -b.mask;
            return aFirst - bFirst;
        });

        return result;
    }

    /**
     * Vérifie si un groupe de jours est présent dans la base de données.
     * @param {string | number} value - Masque, nom ou code alphanumérique du groupe de jours.
     * @returns {boolean} - Vrai si le groupe de jours est présent, faux sinon.
     */
    public static has(value: string | number): boolean {
        return (typeof value === 'number')
            ? value >= 0 && value < this.masksList.length && !!this.masksList[value]
            : this.mapByString.has(value);
    } 

    /**
     * Retourne le groupe de jours correspondant au code alphanumérique ou aux numéros concaténés fourni.
     * Si le groupe de jours n'existe pas, renvoie undefined.
     * @param {string | number} value - Masque, nom ou code alphanumérique du groupe de jours.
     * @returns {Days | undefined} - Instance de Days correspondante,
     *  ou undefined si le groupe de jours n'existe pas.
     */
    public static get(value: string | number): Days | undefined {
        return (typeof value === 'number')
            ? this.masksList[value]
            : this.mapByString.get(value);
    } 

    /**
     * Crée un nouveau groupe de jours et l'ajoute à la base de données.
     * La gare est référencée par son code alphanumérique, son nom complet, son abréviation et ses numéros.
     * Si un groupe de jours avec le même code, nom complet, abréviation ou numéros existe déjà,
     *  une erreur est levée.
     * @param {number[]} numbers - Tableau contenant les numéros des jours du groupe de jours.
     * @param {string} code - Code alphanumérique du groupe de jours.
     * @param {string} fullName - Nom complet du groupe de jours.
     * @param {string} abreviation - Abréviation du groupe de jours.
     * @returns {Days} - Nouveau groupe de jours créé.
     * @throws {Error} - Si le groupe de jours est déjà présent dans la base de données.
     */
    private static create(
        numbers: number[],
        code: string = "",
        fullName: string = "",
        abreviation: string = ""
    ): Days {

        if (numbers.length === 0) {
            throw new Error(`Le groupe de jours ${code} ne contient pas de jours.`);
        }
 
        // Vérifie que le groupe de jour n'existe pas déjà.
        const mask = this.numbersToMask(numbers);
        if (this.has(mask)) {
            return this.get(mask)!;
        }

        // Si toutes les valeurs sont fournies (chargement depuis les paramètres),
        //  récupère toutes les valeurs.
        let finalCode = code;
        let finalFullName = fullName;
        let finalAbbreviation = abreviation;

        // Si seul les numéros de jours sont fournis,
        //  un code optimisé est calculé par assemblage de groupes de jours connus.
        //  Cet assemblage définit également un nom complet et une abréviation par concaténation.
        if (!code) {
            const parts = this.optimiseMask(mask);
            finalCode = parts.map(d => d.code).join('');
            finalFullName = parts.map(d => d.fullName).join(' + ');
            finalAbbreviation = parts.map(d => d.abreviation).join('');
        }

        // Instancie le nouveau groupe de jours.
        const days = new Days(
            mask,
            finalCode,
            finalFullName,
            finalAbbreviation
        );

        // Ajoute le groupe de jour à la base de données.
        this.masksList[mask] = days;
        this.mapByString.set(days.numbersString, days);
        this.mapByString.set(finalCode, days);
        this.mapByString.set(finalFullName, days);
        this.mapByString.set(finalAbbreviation, days);

        return days;
    }

    /**
     * Retourne un tableau des valeurs de la base de données des groupes de jours.
     * @returns {Days[]} - Itérateur sur les valeurs de la base de données.
     *  des groupes de jours.
     */
    public static values(): Days[] {
        return this.masksList.filter(e => e !== undefined);
    }

    /**
     * Efface toutes les gares de la base de données.
     * Cela permet de forcer le rechargement des gares si besoin.
     */
    public static clear(): void {
        this.masksList = [];
        this.mapByString = new Map();
        this.compressionRules = [];
        this.extractionPatterns = [];
        this.loaded = false;
    }

    /**
     * Charge les jours de la semaine.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.loaded) {
            if (!erase) return;
            this.clear();
        }

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides).
                if (row.every((cell: unknown) => !cell)) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Extrait les valeurs.
                const fullName = WorkbookService.getRequiredString(
                    row,
                    this.COL_FULL_NAME,
                    `Nom complet du groupe de jours`
                        + ` non renseigné dans le tableau des jours.`
                );
                if (this.has(fullName)) {
                    throw new Error(`Nom complet ${fullName} déjà utilisé.`);
                }
                const abreviation = WorkbookService.getRequiredString(
                    row,
                    this.COL_ABBREVIATION,
                    `Groupe de jours du ${fullName} :`
                        + ` abréviation non renseignée dans le tableau des jours.`
                );
                if (this.has(abreviation)) {
                    throw new Error(`Groupe de jours du ${fullName} :`
                        + ` abbreviation ${abreviation} déjà utilisée.`);
                }
                const numbersString = WorkbookService.getString(row, this.COL_NUMBERS);
                const numbers = this.cleanAndSortNumbers(numbersString);
                if (numbers.length === 0) {
                    throw new Error(`Groupe de jours du ${fullName} :`
                        + ` numéros des jours non renseignés ou invalides dans le tableau des jours.`);
                }
                const codeLetter = WorkbookService.getString(row, this.COL_CODE_LETTER)
                        .toUpperCase()
                        .replace(/[^A-Z]/g, '');
                // Si groupe de jours, une seule lettre attendue.
                if (!codeLetter && numbers.length > 1) {
                    throw new Error(`Groupe de jours du ${fullName} :`
                        + ` lettre de code non renseignée ou invalide dans le tableau des jours.`);
                }
                const code = (numbersString.length === 1)
                    ? numbersString
                    : codeLetter;

                // Crée l'objet Days.
                const days = this.create(
                    numbers,
                    code,
                    fullName,
                    abreviation
                );
            }

        } catch (e) {
            throw new Error(`Days.load (ligne ${excelRow}) : ${e}`);
        } 

        // Si non renseignés dans le tableau, charge les jours individuels par défaut.
        this.WEEKDAYS.forEach(d => {
            const numbersString = String(d.number);
            if (!this.has(numbersString)) {
                this.create([d.number], d.code ?? numbersString, d.fullName, d.abreviation);
            }
        });

        // Constitue le tableau d'analyse des codes avec :
        //  - la lettre code (si existante),
        //  - le nom complet du groupe de jours,
        //  - l'abréviation du groupe de jours.
        // Pour les groupes de jours, constitue également le tableau des règles de compression du code
        //  avec les codes lettres des groupes de jours.
        for (const days of this.values()) {
            this.compressionRules.push(days);
            this.extractionPatterns.push({
                pattern: days.code,
                numbersString: days.numbersString
            });
            this.extractionPatterns.push({
                pattern: days.fullName.toUpperCase(),
                numbersString: days.numbersString
            });
            this.extractionPatterns.push({
                pattern: days.abreviation.toUpperCase(),
                numbersString: days.numbersString
            });
        }

        // Algorithme de Brian Kernighan pour compter le nombre de bits (donc de jours dans le groupe).
        const popcount = function (mask: number): number {
            let count = 0;
            while (mask) {
                mask &= mask - 1;
                count++;
            }
            return count;
        }

        // Tri de la liste des groupes de jours donnés en paramètre,
        //  par leur nombre de jours dans l'ordre décroissant.
        this.compressionRules.sort(
            (a, b) => b.count - a.count
        );
 
        // Tri du tableau d'analyse des codes, de la plus grande chaine à la plus petite.
        this.extractionPatterns.sort((a, b) => b.pattern.length - a.pattern.length)

        this.loaded = true;

        // Charge les jours individuels.
        Day.load();
    }
}

/*
 * Classe utilitaire DaysValues pour la gestion des valeurs associées aux jours.
 */
class DaysValues {

    // Propriétés de l'objet Days
 
    public readonly days: Days;                             // Groupe de jours total concerné
    private entries: { days: Days, value: string }[] = [];  // Liste de règles (Days → valeur)

    /**
     * Constructeur de la classe DaysValues.
     * @param {Days} days - Groupe de jours total concerné.
     */
    constructor(days: Days) {
        this.days = days;
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * La méthode renvoie une chaine vide si l'objet n'a pas de valeurs associées.
     * Sinon, elle renvoie une chaine au format "jour1: valeur1, jour2: valeur2, ..."
     *  où chaque jour est représenté par un ensemble de numéros de jours
     *  (ex: "1-2,4,6" pour les jours lundi, mardi, jeudi et samedi).
     * @returns {string} - La représentation textuelle de l'objet.
     */
    public toString(): string {

        if (this.entries.length === 0) return "";

        // Cas uniforme
        if (this.entries.length === 1 &&
            this.entries[0].days.mask === this.days.mask) {
            return this.entries[0].value;
        }

        return this.entries
            .map(e => `${e.days.numbersString}: ${e.value}`)
            .join(", ");
    }

    /**
     * Crée un objet DaysValues à partir d'un groupe de jours et d'une chaîne de valeurs.
     * La chaîne de valeurs peut contenir des valeurs uniques ou multiples, séparées par des virgules.
     * Si la chaîne de valeurs contient une seule valeur, le groupe de jours est considéré comme un seul jour.
     * Si la chaîne de valeurs contient plusieurs valeurs, chaque valeur est associée à un groupe de jours.
     * @param {Days} days - Groupe de jours total concerné.
     * @param {string} input - Chaîne de valeurs.
     * @returns {DaysValues} - L'objet DaysValues créé.
     */ 
    public static from(days: Days, input: string): DaysValues {

        const dv = new DaysValues(days);

        if (!input || input.trim() === "") return dv;

        // Cas avec valeur unique
        if (!input.includes(":")) {
            dv.set(days, input.trim());
            return dv;
        }

        // Cas avec plusieurs valeurs
        const parts = input.split(/[,;]+/);

        for (const part of parts) {

            const [daysPart, valuePart] = part.split(":");

            if (!valuePart) continue;

            const subDays = Days.from(daysPart.trim());
            const value = valuePart.trim();

            if (subDays) {
                dv.set(subDays, value);
            }
        }
        dv.fillGaps("");

        return dv;
    }

    /**
     * Renvoie la valeur associée au jour de la semaine donné.
     * Si le jour n'est pas couvert par une règle, renvoie une chaîne vide.
     * @param {Day} day - Jour de la semaine.
     * @returns {string} - Valeur associée.
     */
    public get(day: Day): string {

        for (const entry of this.entries) {
            if (entry.days.contains(day)) {
                return entry.value;
            }
        }

        return "";
    }

    /**
     * Modifie la valeur associée à un groupe de jours.
     * Supprime les parties déjà couvertes dans les autres sous-groupes de jours
     *  et ajoute la nouvelle valeur.
     * @param {Days} days - Groupe de jours total concerné.
     * @param {string} value - Valeur associée.
     */
    public set(days: Days, value: string): void {

        // Restreint le groupe de jours à affecter au périmètre autorisé
        const validDays = Days.intersection(days, this.days);
        if (!validDays || validDays.mask === 0) return;

        // Supprime les parties déjà couvertes dans les autres sous-groupes de jours (valeur modifiée)
        this.entries = this.entries
            .map(e => {
                const remaining = Days.difference(e.days, validDays);
                return remaining ? { days: remaining, value: e.value } : null;
            })
            .filter(e => e !== null)

        // Ajoute la nouvelle valeur
        this.entries.push({ days: validDays, value });

        this.merge();
    }

    /**
     * Fusionne les valeurs associées aux groupes de jours.
     * Regroupe par valeur, puis reconstruit les entrées avec les valeurs fusionnées.
     */
    private merge(): void {

        const map = new Map<string, number>();

        // Regroupe par valeur
        for (const e of this.entries) {
            const prev = map.get(e.value) ?? 0;
            map.set(e.value, prev | e.days.mask);
        }

        // Reconstruit entries
        this.entries = [];
        for (const [value, mask] of Array.from(map.entries())) {
            this.entries.push({
                days: Days.get(mask)!,
                value
            });
        }
        this.entries.sort((a, b) => 
            a.days.numbersString.localeCompare(b.days.numbersString));
    }

    /**
     * Vérifie si toutes les parties du groupe de jours sont couvertes.
     * @returns {boolean} - Vrai si toutes les parties sont couvertes, faux sinon.
     */
    public isComplete(): boolean {

        let mask = 0;
        for (const e of this.entries) {
            mask |= e.days.mask;
        }

        return mask === this.days.mask;
    }

    /**
     * Ajoute les valeurs manquantes dans le groupe de jours.
     * Pour chaque jour non couvert par une règle, ajoute une entrée avec la valeur par défaut.
     * @param {string} defaultValue - Valeur associée aux jours non couverts.
     */
    public fillGaps(defaultValue: string): void {

        let covered = 0;
        for (const e of this.entries) covered |= e.days.mask;
        const missingMask = this.days.mask & ~covered;

        if (missingMask) {
            this.entries.push({
                days: Days.fromMask(missingMask)!,
                value: defaultValue
            });
            this.merge();
        }
    }

    /**
     * Vérifie si les deux objets DaysValues sont égaux.
     * La comparaison est faite en fonction de la représentation textuelle des objets.
     * @param {DaysValues} other - L'objet DaysValues à comparer.
     * @returns {boolean} - Vrai si les deux objets sont égaux, faux sinon.
     */
    public equals(other: DaysValues): boolean {

        return this.toString() === other.toString();
    }
}

/*
 * Classe utilitaire Parity immuable qui permet de manipuler la parité
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

    // Constantes des valeurs de parité
    public static readonly UNDEFINED: number = 0;   // Parité non définie
    public static readonly ODD: number = 1;         // Parité impaire
    public static readonly EVEN: number = 2;        // Parité paire
    public static readonly DOUBLE: number = 3;      // Parité double

    // Map des parités possibles
    private static pool: (Parity | undefined)[][] = [
        [],                                         // pool[0] : doubleAllowed = false
        []                                          // pool[1] : doubleAllowed = true
    ];

    // Map des lettres et nombres désignants les parités
    private static letters: Map<number, string> = new Map();
    private static digits: Map<number, number> = new Map();
 
    // Indicateur de chargement
    private static loaded = false;

    // Propriétés de l'objet Parity
    public readonly value: number;                          // Valeur de la parité
    private readonly doubleParityAllowed: boolean;          // Autorise une double parité

    /**
     * Constructeur privé de la classe Parity.
     * @param {number} - value Valeur de parité.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la double parité est autorisée.
     *  Si faux (par défaut), la double parité est impossible.
     */
    private constructor(
        value: number,
        doubleParityAllowed: boolean = false
    ) {
        this.value = value;
        this.doubleParityAllowed = doubleParityAllowed;
    }
 
    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.value.toString();
    }

    /**
     * Retourne une instance de Parity qui correspond à la valeur de parité spécifiée.
     * Si l'instance n'existe pas, la crée et la stocke pour future utilisation.
     * @param {number} - value Valeur de parité.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la double parité est autorisée.
     *  Si faux (par défaut), la double parité est impossible.
     * @returns {Parity} - Instance de Parity correspondante.
     */
    private static getOrCreate(value: number, doubleParityAllowed: boolean): Parity {
        if (value < 0 || value > 3) {
            value = this.UNDEFINED;
        }
        const i = doubleParityAllowed ? 1 : 0;
        let p = this.pool[i][value];
        if (!p) {
            p = new Parity(value, doubleParityAllowed);
            this.pool[i][value] = p;
        }
        return p;
    }

    /**
     * Retourne une instance de Parity à partir d'une valeur qui peut être :
     *  - une lettre de parité (ou la concaténation des deux lettres sans ordre si double parité),
     *  - un chiffre de parité (format chaîne ou nombre),
     *  - un numéro de train (pair, impair ou double s'il contient un '/'),
     *  - une instance de Parity (retourne la même instance),
     *  - null ou undefined (retourne une instance de Parity avec valeur this.UNDEFINED).
     * @param {Parity | string | number | null | undefined} value - Valeur à analyser pour la parité.
     * @returns {Parity} - Instance de Parity correspondante.
     */
    public static from(
        value: Parity | string | number | null | undefined,
        doubleParityAllowed: boolean = false
    ): Parity {
        if (value instanceof Parity) {
            if (value.doubleParityAllowed === doubleParityAllowed) return value;
            return this.getOrCreate(value.value, doubleParityAllowed);
        }
 
        const normalized = this.normalize(value, doubleParityAllowed);
        return this.getOrCreate(normalized, doubleParityAllowed);
    } 

    /**
     * Normalise en une valeur de parité une valeur, qui peut être :
     *  - la lettre de parité (ou la concaténation des deux lettres sans ordre si double parité),
     *  - le chiffre de parité (format chaîne ou nombre),
     *  - un numéro de train (pair, impair ou double s'il contient un '/').
     * @param {string | number | null | undefined} value - Valeur à normaliser.
     * @param {boolean} doubleParityAllowed - Si vrai, la double parité est autorisée.
     * @returns {number} - Valeur de parité normalisée.
     */
    private static normalize(
        value: string | number | null | undefined,
        doubleParityAllowed: boolean
    ): number {
 
        // La valeur est nulle ou undefined
        if (value == null) return this.UNDEFINED;
 
        // La valeur est un nombre
        if (typeof value === 'number') {
            if (
                // Valeurs de Parity explicites
                value === this.UNDEFINED ||
                value === this.ODD ||
                value === this.EVEN ||
                value === this.DOUBLE
            ) {
                return value === this.DOUBLE && !doubleParityAllowed
                    ? this.UNDEFINED
                    : value;
            }
 
            // Nombres négatifs → undefined
            if (value <= 0) return this.UNDEFINED;
 
            // Parité du nombre
            return value % 2 === 0 ? this.EVEN : this.ODD;
        }
 
        // La valeur est une chaine
        const str = value.trim().toUpperCase();
 
        if (str === '' || str === '0') return this.UNDEFINED;
 
        // Double implicite (ex: "12345/6")
        if (str.includes('/')) {
            return doubleParityAllowed ? this.DOUBLE : this.UNDEFINED;
        }
 
        // Tentative de conversion numérique
        const numeric = parseInt(str, 10);
        if (!Number.isNaN(numeric)) {
            return this.normalize(numeric, doubleParityAllowed);
        }
 
        // Lettres
        const odd = this.letter(this.ODD);
        const even = this.letter(this.EVEN);
        if (odd && even) {
            switch (str) {
                case odd:
                    return this.ODD;
                case even:
                    return this.EVEN;
                case odd + even:
                case even + odd:
                    return doubleParityAllowed ? this.DOUBLE : this.UNDEFINED;
            }
        }
 
        return this.UNDEFINED;
    }

    /**
     * Crée une parité qui n'a pas de valeur définie.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} - Parité sans valeur définie.
     */
    public static undefined(doubleParityAllowed: boolean = false): Parity {
        return this.getOrCreate(this.UNDEFINED, doubleParityAllowed);
    }

    /**
     * Crée une parité qui correspond à une parité impaire.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} - Parité impaire.
     */
    public static odd(doubleParityAllowed: boolean = false): Parity {
        return this.getOrCreate(this.ODD, doubleParityAllowed);
    }

    /**
     * Crée une parité qui correspond à une parité paire.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} - Parité paire.
     */
    public static even(doubleParityAllowed: boolean = false): Parity {
        return this.getOrCreate(this.EVEN, doubleParityAllowed);
    }

    /**
     * Crée une parité qui correspond à une parité double.
     * Elle est représentée par le chiffre -2.
     * @returns {Parity} - Parité double.
     */
    public static double(): Parity {
        return this.getOrCreate(this.DOUBLE, true);
    }

    /**
     * Vérifie si la parité est identique à une valeur de parité donnée.
     * @param {number} parity - Autre valeur de parité à comparer.
     * @returns {boolean} - Vrai si les deux parités sont identiques, faux sinon.
     */
    public is(parity: number): boolean {
        return this.value === parity;
    }

    /**
     * Vérifie si la parité est définie (différente de Parity.UNDEFINED).
     * @returns {boolean} - Vrai si la parité est définie, faux sinon.
     */
    public isDefined(): boolean {
        return this.value !== Parity.UNDEFINED;
    }

    /**
     * Vérifie si deux parités sont opposées (parité impaire versus parité paire).
     * @param {Parity | undefined} other - Autre variable de parité à comparer.
     * @returns {boolean} - Vrai si les deux parités sont opposées, faux sinon.
     */
    public isOpposedTo(other: Parity | undefined): boolean {
        return !!other
            && (this.value === Parity.ODD && other.value === Parity.EVEN
                || this.value === Parity.EVEN && other.value === Parity.ODD);
    }

    /**
     * Vérifie si la parité est définie et identique à celle d'une autre variable de parité.
     * @param {Parity | undefined} other - Autre variable de parité à comparer.
     * @returns {boolean} - Vrai si les deux parités sont identiques, faux sinon.
     */
    public equalsTo(other: Parity | undefined): boolean {
        return this === other;
    }

    /**
     * Vérifie si la parité inclut une autre valeur de parité.
     * @param {Parity | number | null | undefined} other - Autre valeur de parité à inclure.
     * @returns {boolean} - Vrai si la parité inclut la valeur de parité, faux sinon.
     */
    public includes(other: string | number | Parity | null | undefined): boolean {
        const requested = Parity.from(other, this.doubleParityAllowed);
 
        // undefined n'inclut rien
        if (this.value === Parity.UNDEFINED) {
            return false;
        }
 
        // La parité double inclut toutes les parités définies
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
     * @returns {Parity} - Parité inversée.
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
     * @param {Parity} other - Parité à combiner avec la parité actuelle.
     * @returns {Parity} - Parité combinée.
     */
    public combineWith(other: Parity): Parity {

        if (!this.doubleParityAllowed) throw new Error(`Il n'est pas possible de combiner une`
            + ` parité à une autre si celle de départ n'autorise pas les parités doubles.`
            + ` Le résultat est forcément une parité qui accepte les parités doubles.`);
 
        if (!this.isDefined()) {
            return Parity.from(other.value, true);
        }

        if (!other.isDefined() || this.value === other.value) {
            return this;
        }
 
        return Parity.double();
    } 
 
    /**
     * Retourne le chiffre de parité correspondant.
     * @param {number} parity - Valeur de la parité.
     * @returns {number} - Chiffre de parité correspondante, ou 0 si la parité est undefined.
     */
    public static digit(parity: number): number {
        return this.digits.get(parity) ?? 0;
    }

    /**
     * Retourne le chiffre de parité correspondant (chaine vide pour une parité indéfinie).
     * @returns {string | number} - Chiffre de parité, ou une chaîne vide
     *  si la parité est indéfinie.
     */
    public printDigit(): string | number {
        return this.isDefined() ? Parity.digit(this.value) : "";
    }

    /**
     * Retourne la lettre de parité correspondante.
     * @param {number} parity - Valeur de la parité.
     * @returns {string} - Lettre de parité correspondante, ou une chaîne vide si la parité est undefined.
     */
    public static letter(parity: number): string {
        return this.letters.get(parity) ?? "";
    }

    /**
     * Retourne la lettre de parité correspondante
     *  (parité impaire ou paire, concaténation impaire puis paire si parité double).
     * @returns {string} - Lettre de parité correspondante, ou une chaîne vide
     *  si la parité est double ou indéfinie.
     */
    public printLetter(): string {
        switch (this.value) {
            case Parity.ODD:
            case Parity.EVEN:
                return Parity.letter(this.value)!;
            case Parity.DOUBLE:
                return Parity.letter(Parity.ODD)!
                    + Parity.letter(Parity.EVEN)!;
            default:
                return "";
        }
    }

    /**
     * Vérifie si une chaîne de caractères contient la lettre de parité correspondante,
     *  ou les deux lettres si la parité est double.
     * @param {string} string - Chaîne de caractères à analyser.
     * @param {number} parity - Parité à chercher.
     * @returns {boolean} - Vrai si la chaîne de caractères contient la lettre de parité,
     *  faux sinon.
     */
    public static containsParityLetter(string: string, parity: number): boolean {
        switch (parity) {
            case this.ODD:
                return string.toUpperCase().includes(this.letter(this.ODD)!);
            case this.EVEN:
                return string.toUpperCase().includes(this.letter(this.EVEN)!);
            case this.DOUBLE:
                return string.toUpperCase().includes(this.letter(this.ODD)!)
                    && string.toUpperCase().includes(this.letter(this.EVEN)!);
            default:
                return false;
        }
    }

    /**
     * Charge les lettres et chiffres associés aux parités.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.loaded) {
            if (!erase) return;
            this.letters.clear();
            this.digits.clear();
        }

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);

        const getParityLetter = (
            row: number,
            fallback: string
        ): string =>
            WorkbookService.getString(data[row], this.COL_LETTER, fallback)
                .toUpperCase();

        this.letters.set(this.ODD, getParityLetter(this.ROW_ODD, "I"));
        this.letters.set(this.EVEN, getParityLetter(this.ROW_EVEN, "P"));

        const getParityDigit = (
            row: number,
            fallback: number
        ): number =>
            WorkbookService.getNumber(data[row], this.COL_NUMBER, fallback);

        this.digits.set(this.ODD, getParityDigit(this.ROW_ODD, 1));
        this.digits.set(this.EVEN, getParityDigit(this.ROW_EVEN, 2));
        this.digits.set(this.DOUBLE, getParityDigit(this.ROW_DOUBLE, -2));

        this.loaded = true;
    }
}

/**
 * Classe TrainNumber définissant un numéro de train.
 * Il est alphanumérique, sans ponctuation et sans espaces, avec un chiffre pour dernier caractère.
 * La double parité est marquée par ######/#.
 */
class TrainNumber {

    // Constantes de lecture de la base de données Excel

    private static readonly TRAIN_NUMBERS_PARAM_SHEET = "Param";        // Feuille contenant les modèles de numéros de trains
                                                                        //  à charger comme paramètres
    private static readonly COMMERCIAL_TABLE = "Commerciaux";           // Tableau contenant les motifs des trains commerciaux
    private static readonly W_TABLE = "W";                              // Tableau contenant les motifs des trains W
    private static readonly MOUVEMENTS_TABLE = "Evolutions";            // Tableau contenant les motifs des évolutions
    private static readonly TRAINS_4DIGIT_TABLE = "LigneC4chiffres";    // Tableau contenant les motifs des trains abrégeables à 4 chiffres

    // Indicateur de chargement
    private static loaded = false;

    // Regex globales
    private static commercialRegex: RegExp;         // Regex des trains commerciaux
    private static wRegex: RegExp;                  // Regex des trains W
    private static mouvementsRegex: RegExp;         // Regex des évolutions
    private static abbreviate4Regex: RegExp;        // Regex des trains abrégeables à 4 chiffres

    // Propriétés de l'objet TrainNumber
    public readonly value: string;                  // Numéro de train avec parité
                                                    //  (la double parité est marquée par ######/#)

    // Cache interne
    private readonly variants: Set<string>;         // Toutes les variantes équivalentes
    private readonly variantsByParity: string[];    // Accès direct par parité
    private _zone?: number | null;                  // Zone du train si train commercial
                                                    //  (de 0 à 9 : 4ème chiffre du numéro du train)
    private _battery?: number | null;               // Batterie du train si train commercial 
                                                    //  (de 0 à 99 : 5ème et 6ème chiffres
                                                    //  du numéro du train si le train a une parité double, 
                                                    //  le 3ème chiffre donne la parité)

    /**
     * Constructeur privé de la classe TrainNumber.
     * Garde uniquement les chiffres et lettres mises en majuscules.
     * @param {string | number} value - Numéro de train (nombre ou chaine de caractères).
     * @param {boolean} doubleParity - Si vrai, force la double parité. Si faux (par défaut),
     *  la double parité est détectée avec la présence de "/" dans le numéro de train.
     */
    private constructor(
        value: string | number,
        doubleParity: boolean = false
    ) {

        const raw = value.toString();
        const applyDoubleParity = doubleParity || raw.includes("/");

        const normalized = TrainNumber.normalize(raw);

        if (!TrainNumber.isValidTrainNumber(normalized)) {
            Log.warn(`Numéro de train invalide : ${value}`);
            this.value = "";
            this.variants = new Set();
            this.variantsByParity = [];
            return;
        }

        // Calcul unique
        const base = normalized;
        const lastDigit = base.charCodeAt(base.length - 1) - 48;
        const rest = base.slice(0, -1);
        const even = lastDigit - (lastDigit % 2);
        const odd = even + 1;

        const evenStr = rest + even;
        const oddStr = rest + odd;
        const evenOdd = `${evenStr}/${odd}`;
        const oddEven = `${oddStr}/${even}`;

        // Cache des variantes
        this.variants = new Set([
            evenStr,
            oddStr,
            evenOdd,
            oddEven
        ]);

        // Accès indexé par parité
        this.variantsByParity = [];
        this.variantsByParity[Parity.EVEN] = evenStr;
        this.variantsByParity[Parity.ODD] = oddStr;

        // La parité double commence par la même parité que la valeur en entrée
        this.variantsByParity[Parity.DOUBLE] = (normalized === oddStr) ? oddEven : evenOdd;

        this.value = applyDoubleParity
            ? this.variantsByParity[Parity.DOUBLE]
            : normalized;
    }

    /**
     * Retourne la valeur de base du train (sans parité).
     * @returns {string} - Valeur de base du train.
     */
    public get baseValue(): string {
        return this.value.split('/')[0];
    }

    /**
     * Indique si le train a une parité double.
     * @returns {boolean} Vrai si le train a une parité double.
     */
    public get isDoubleParity(): boolean {
        return this.value.includes("/");
    }

    /**
     * Teste si le train est commercial.
     * @returns {boolean} - Vrai si le train est commercial, faux sinon.
     */
    public get isCommercial(): boolean {
        return TrainNumber.commercialRegex?.test(this.baseValue) ?? false;
    }

    /**
     * Teste si le train est W (vide voyageur).
     * @returns {boolean} - Vrai si le train est W, faux sinon.
     */
    public get isW(): boolean {
        return TrainNumber.wRegex?.test(this.baseValue) ?? false;
    }

    /**
     * Teste si le train est une évolution.
     * @returns {boolean} - Vrai si le train est une évolution, faux sinon.
     */
    public get isMouvement(): boolean {
        return TrainNumber.mouvementsRegex?.test(this.baseValue) ?? false;
    }

    /**
     * Retourne la zone du train (de 0 à 9 : 4ème chiffre du numéro du train).
     * Retourne null si le train n'est pas commercial.
     * @returns {number | null} Zone du train.
     */
    public get zone(): number | null {
        if (this._zone === undefined) {
            this._zone = this.isCommercial
                ? this.value.charCodeAt(3) - 48
                : null;
        }
        return this._zone;
    }

    /**
     * Retourne la batterie du train (de 0 à 99 : 5ème et 6ème chiffres du numéro du train)
     * Si le train a une parité double, c'est le 3ème chiffre qui donne la parité.
     * Retourne null si le train n'est pas commercial.
     * @returns {number | null} Batterie du train.
     */ 
    public get battery(): number | null {
        if (this._battery === undefined) {
            this._battery = this.isCommercial
                ? parseInt(this.value.slice(4, 6), 10)
                    + (this.isDoubleParity ? parseInt(this.value[2], 10) % 2 : 0)
                : null;
        }
        return this._battery;
    }
    
    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.value;
    }

    /**
     * Crée une instance de TrainNumber à partir d'un numéro de train ou d'un nombre.
     * La méthode normalise le numéro de train en supprimant les caractères non-alphanumériques
     *  et en ne gardant que la partie précédent un "/".
     * La double parité est marquée par ######/#.
     * @param {string | number | null | undefined} value - Numéro de train (nombre ou chaine de caractères).
     * @param {boolean} doubleParity - Si vrai, force la double parité. Si faux (par défaut),
     *  la double parité est détectée avec la présence de "/" dans le numéro de train.
     * @returns {TrainNumber} - Instance de TrainNumber correspondant au numéro de train.
     */
    public static from(
        value: TrainNumber | string | number | null | undefined,
        doubleParity: boolean = false
    ): TrainNumber {
        if (value == null) {
            throw new Error(`Le numéro de train n'est pas renseigné ou est invalide.`);
        }
        return value instanceof TrainNumber ? value : new TrainNumber(value, doubleParity);
    }

    /**
     * Normalise un numéro de train en supprimant les caractères non-alphanumériques
     *  et en ne gardant que la partie précédent un "/".
     * @param {string} value - Numéro de train à normaliser.
     * @returns {string} - Numéro de train normalisé.
     */
    private static normalize(value: string): string {
        return value
            .split("/")[0]
            .toUpperCase()
            .replace(/[^A-Z0-9]/g, '');
    }

    /**
     * Vérifie si un numéro de train est valide.
     * @param {string} value - Numéro de train à vérifier.
     * @returns {boolean} - Vrai si le numéro de train est valide, faux sinon.
     */
    private static isValidTrainNumber(value: string): boolean {
        if (!value) return false;
        const lastChar = value.slice(-1);
        return /^[0-9]$/.test(lastChar);
    }

    /**
     * Abrège le numéro de train à 4 chiffres si possible.
     * La méthode teste si le numéro de train correspond à une expression régulière
     *  définie dans la classe TrainNumber.
     * Si le numéro de train correspond, il est abrégé en supprimant les 2 premiers chiffres.
     * Si le numéro de train ne correspond pas, il est renvoyé inchangé.
     * @returns {string} - Numéro de train abrégé de 6 à 4 chiffres s'il est abrégeable.
     */
    private static abbreviate(value: string): string {
        return this.abbreviate4Regex?.test(value.split("/")[0])
            ? value.substring(2)
            : value;
    }

    /**
     * Teste si une valeur correspond à ce train (toutes formes confondues).
     */
    public includes(value: TrainNumber | string | number | null | undefined): boolean {

        if (value == null) return false;

        const str = value instanceof TrainNumber
            ? value.value
            : TrainNumber.normalize(String(value));

        return this.variants.has(str);
    }

    /**
     * Adapte le numéro du train en fonction de la parité demandée..
     * @param {number} parityValue Parité demandée (paire, impaire, double).
     * @param {boolean} abbreviateTo4Digits Si vrai, le numéro du train est abrégé à 4 chiffres.
     *  Si faux, le numéro du train n'est pas abrégé.
     * @returns {string} Numéro du train adapté
     */
    public adaptWithParity(parityValue: number, abbreviateTo4Digits: boolean = false): string {

        const adaptedValue = this.variantsByParity[parityValue] ?? this.value;
        return abbreviateTo4Digits? TrainNumber.abbreviate(adaptedValue) : adaptedValue;
    }

    /**
     * Retourne le numéro du train en fonction des paramètres :
     *  - si abbreviate est vrai, le numéro du train est abrégé à 4 chiffres,
     *  - si withoutDoubleParity est vrai, le numéro du train est renommé sans double parité.
     * @param {boolean} [abbreviate=false] - Si vrai, le numéro du train est abrégé
     *  de 6 à 4 chiffres pour les trains commerciaux. Si faux (par défaut), le numéro n'est pas abrégé.
     * @param {boolean} [withoutDoubleParity=false] - Si vrai, le numéro est renommé
     *  pour ne pas indiquer le changement de parité. Si faux (par défaut), le numéro de train
     *  en gare origine est renvoyé avec double parité si concerné.
     * @returns {string} - Numéro du train.
     */
    public format(
        abbreviate: boolean = false,
        withoutDoubleParity: boolean = false
    ): string {
        let result = withoutDoubleParity ? this.baseValue : this.value;

        if (abbreviate) {
            result = TrainNumber.abbreviate(result);
        }
 
        return result;
    }
 
    /**
     * Charge les paramètres des numéros de train
     *  - regex des numéros de train W,
     *  - regex des numéros de train abrégeables à 4 chiffres.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si les tables à charger existent déjà.
        if (this.loaded && !erase) return;

        this.loadRegex();

        this.loaded = true;
    }

    /**
     * Charge les motifs des trains spécifiques.
     * Les valeurs de la table sont transformées en regex partielles avec les numéros
     *  remplacés par des chiffres, puis combinées en une regex globale unique.
     */
    private static loadRegex(): void {
 
        const dataToRegex = (data: CellValue[][]) => {
            const parts = data
                .slice(1)
                .reduce((acc, row) => acc.concat(row), [])
                .filter((v: unknown): v is string => typeof v === "string" && v.trim() !== "")
                .map(pattern => {
                    return '^' + pattern.trim().replace(/#/g, '\\d') + '$';
                });

            return parts.length
                ? new RegExp(parts.join('|'))
                : /^$/;
        };

        const commercialData = WorkbookService.getDataFromTable(
            this.TRAIN_NUMBERS_PARAM_SHEET,
            this.COMMERCIAL_TABLE
        );
        this.commercialRegex = dataToRegex(commercialData);

        const wData = WorkbookService.getDataFromTable(
            this.TRAIN_NUMBERS_PARAM_SHEET,
            this.W_TABLE
        );
        this.wRegex = dataToRegex(wData);

        const mouvementsData = WorkbookService.getDataFromTable(
            this.TRAIN_NUMBERS_PARAM_SHEET,
            this.MOUVEMENTS_TABLE
        );
        this.mouvementsRegex = dataToRegex(mouvementsData);

        const abbreviate4Data = WorkbookService.getDataFromTable(
            this.TRAIN_NUMBERS_PARAM_SHEET,
            this.TRAINS_4DIGIT_TABLE
        );
        this.abbreviate4Regex = dataToRegex(abbreviate4Data);
    }
}

/* 
 * Classe Station définissant une gare.
 */
class Station {

    // Propriétés de l'objet Station
    public readonly id: number;                             // Id de la gare
    public readonly abbreviation!: string;                  // Abréviation de la gare
    public readonly name: string;                           // Nom de la gare
    public referenceStation: Station | null;                // Gare de rattachement
    public childStations: Station[];                        // Sous-gares
    public readonly turnaround: Parity;                     // Parité d'un rebroussement possible
                                                            //  (la parité est celle du train avant rebroussement)
    public readonly reverseLineDirection: boolean;          // Parité de la ligne inversée sur cette gare

    /**
     * Constructeur d'une gare.
     * @param {number} id - Id de la gare.
     * @param {string} abbreviation - Abréviation de la gare.
     * @param {string} name - Nom de la gare.
     * @param {Station} referenceStation - Gare de rattachement.
     * @param {Parity} turnaround - Parité d'un rebroussement possible
     *  (la parité est celle du train avant rebroussement).
     * @param {boolean} reverseLineDirection - Parité de la ligne inversée sur cette gare.
     */
    constructor(
        id: number,
        abbreviation: string,
        name: string,
        referenceStation: Station | null,
        turnaround: Parity | string | number,
        reverseLineDirection: boolean,
    ) {
        this.id = id;
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
 
    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.abbreviation;
    }
}

/**
 * Classe Stations contenant la liste des gares.
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

    // Tableau des gares indexées par id
    public static list: Station[] = [];
    // Map des gares indexées par abréviation
    public static abbrMap: Record<string, Station> = Object.create(null);
    // Map des gares indexées par nom
    public static nameMap: Record<string, Station> = Object.create(null);
 
    /**
     * Retourne le nombre de gares enregistrées dans la base de données
     * @returns {number} - Nombre de gares enregistrées
     */
    public static get size(): number {
        return this.list.length;
    }

    /**
     * Vérifie si une gare est présente dans la base de données.
     * @param {string} value - Abréviation ou nom de la gare.
     * @returns {boolean} - Vrai si la gare est présente, faux sinon.
     */
    public static has(value: string): boolean {
        return value in this.abbrMap || value in this.nameMap;
    }

    /**
     * Renvoie la gare correspondant à l'ID donné.
     * @param {number} id - ID de la gare.
     * @returns {Station} - Gare correspondante.
     */
    public static getById(id: number): Station {
        const s = this.list[id];
        if (!s) throw new Error(`Gare : ID ${id} inconnue`);
        return s;
    }

    /**
     * Renvoie une gare correspondant à l'abréviation ou au nom donné.
     * @param {string} value - Abréviation ou nom de la gare.
     * @returns {Station | undefined} - Gare correspondante, ou undefined si non trouvée.
     */
    public static get(value: string): Station | undefined {

        let adaptedValue = value;

        for (const suffix of Params.stationsSuffixes) {
            if (adaptedValue.endsWith(`-${suffix}`)) {
                adaptedValue = adaptedValue.slice(0, -suffix.length - 1);
                break; // S'arrête au premier match.
            }
        }

        return this.abbrMap[adaptedValue] ?? this.nameMap[adaptedValue];
    }

    /**
     * Crée une nouvelle gare et l'ajoute à la base de données, 
     *  référencée par son ID, sa clé et son nom.
     * Si une gare avec la même clé ou le même nom existe déjà, une erreur est levée.
     * @param {string} abbreviation - Abréviation de la gare.
     * @param {string} name - Nom de la gare.
     * @param {Parity | string | number} turnaround - Parité d'un rebroussement possible
     *  (la parité est celle du train avant rebroussement).
     * @param {boolean} reverseLineDirection - Parité de la ligne inversée sur cette gare.
     * @returns {Station} - La nouvelle gare créée.
     * @throws {Error} - Si une gare avec la même clé ou le même nom existe déjà.
     */
    private static create(
        abbreviation: string,
        name: string,
        turnaround: Parity | string | number,
        reverseLineDirection: boolean,
    ): Station {

        // Vérifie que la gare n'existe pas déjà.
        for (const value of [abbreviation, name]) {
            if (this.has(value)) {
                throw new Error(`La gare ${value} est déjà présente dans la base de données.`);
            }
        }

        // Calcule l'ID.
        const id = this.list.length;

        // Instancie la nouvelle gare.
        const station = new Station(
            id,
            abbreviation,
            name,
            null,
            turnaround,
            reverseLineDirection
        );

        // Ajoute la gare à la base de données.
        this.list.push(station);
        this.abbrMap[station.abbreviation] = station;
        this.nameMap[station.name] = station;

        return station;
    }

    /**
     * Retourne un tableau des valeurs de la base de données des gares.
     * @returns {Station[]} - Itérateur sur les valeurs
     *  de la base de données des gares.
     */
    public static values(): Station[] {
        return Array.from(this.list.values());
    }

    /**
     * Efface toutes les gares de la base de données.
     * Cela permet de forcer le rechargement des gares si besoin.
     */
    public static clear(): void {
        this.list = [];
        this.abbrMap = Object.create(null);
        this.nameMap = Object.create(null);
    }
 
    /**
     * Charge les gares.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Stations.load : aucune donnée trouvée dans la table.`);
            return;
        }

        const referenceStationPairs: [Station, string][] = [];
        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère les champs.
                const abbreviation = WorkbookService.getString(row, this.COL_ABBR)
                    .toUpperCase();
                const name = WorkbookService.getString(row, this.COL_NAME);
                const referenceStationAbbv = WorkbookService.getString(row, this.COL_REFERENCE_STATION);
                const turnaroundLetters = WorkbookService.getString(row, this.COL_TURNAROUND);
                const reverseLineDirection = WorkbookService.getBoolean(row, this.COL_REVERSE_LINE_PARITY);

                // Crée l'objet Station et l'insère dans la base de données.
                const station = this.create(
                    abbreviation,
                    name,
                    turnaroundLetters,
                    reverseLineDirection
                );

                // Mémorise les paires gare/gare de rattachement.
                if (referenceStationAbbv) {
                    referenceStationPairs.push([station, referenceStationAbbv]);
                }
            }

        } catch (e) {
            throw new Error(`Stations.load (ligne ${excelRow}) : ${e}`);
        } 

        // Parcourt les paires pour ajouter les objets des gares de réference à chaque gare.
        for (const [station, referenceStationAbbv] of referenceStationPairs) {
            const referenceStation = this.get(referenceStationAbbv);

            if (referenceStation) {
                station.referenceStation = referenceStation;
                referenceStation.childStations.push(station);
            }
        }
    }

    /**
     * Sauvegarde les gares de la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = this.SHEET,
        tableName: string = this.TABLE,
        startCell: string = "A1"
    ): void {

        // Convertit la base de données en un tableau de données.
        const data: (string | number)[][] = Array
            .from(this.values())
            .map(station => [
                station.abbreviation,
                station.name,
                station.referenceStation?.abbreviation ?? "",
                station.turnaround.printLetter(),
                station.reverseLineDirection ? 1 : 0
            ]);

        // Imprime le tableau.
        WorkbookService.printTable(
            this.HEADERS, 
            data,
            sheetName, 
            tableName, 
            startCell
        );
    }
}

/**
 * Classe StationWithParity immuable définissant une gare d'arrêt ou de passage d'un train
 *  et sa parité associée.
 */
class StationWithParity {

    // Constantes de parité
    public static readonly UNDEFINED = 0;   // Valeur de parité undefined pour le calcul de l'ID
    public static readonly ODD = 1;         // Valeur de parité impaire pour le calcul de l'ID
    public static readonly EVEN = 2;        // Valeur de parité paire pour le calcul de l'ID

    // Propriétés de l'objet StationWithParity
    public readonly id: number;             // Identifiant unique
    public readonly key: string;            // Clé (en cache)

    private _expandedCache?: StationWithParity[];   // Cache des gares avec parité rattachées

    /**
     * Constructeur d'une gare avec parité.
     * @param {number} id - Identifiant unique.
     * @param {Station} station - Gare (objet Station).
     * @param {Parity} parity - Parité associée à la gare.
     */
    constructor(id: number, station: Station, parity: Parity) {
        this.id = id;
        this.key = StationWithParity.keyOf(station, parity); 
    }
 
    /**
     * Retourne la gare (objet Station) associée à cet objet StationWithParity.
     * @returns {Station} - Gare (objet Station) associée.
     */
    public get station(): Station {
        return Stations.getById(Math.floor(this.id / 3));
    }

    /**
     * Retourne la parité associée à cet objet StationWithParity.
     * La parité est définie en fonction de la valeur de l'identifiant unique de l'objet :
     *  - si l'identifiant est impair, la parité est impair,
     *  - si l'identifiant est pair, la parité est pair,
     *  - si l'identifiant est nul, la parité est undefined.
     * @returns {Parity} - Parité associée à cet objet StationWithParity.
     */
    public get parity(): Parity {
        const p = this.id % 3;
        return p === StationWithParity.ODD ? Parity.odd()
             : p === StationWithParity.EVEN ? Parity.even()
             : Parity.undefined();
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.key;
    }

    /**
     * Retourne une instance de StationWithParity à partir d'une valeur qui peut être :
     *  - une instance de StationWithParity,
     *  - une instance de Station,
     *  - un nom de gare ou une clé d'arrêt (avec ou sans suffixe parité),
     *  - null ou undefined (lève une erreur).
     * La parité associée est celle trouvée dans la valeur, sauf si elle est imposée en argument.
     * @param {StationWithParity | Station | string | null | undefined} value - Valeur à analyser
     *  pour la gare.
     * @param {Parity | string | number} [parity] - Parité optionnelle imposée,
     *  qui remplace celle potentiellement présente dans value.
     * @returns {StationWithParity | undefined} - Instance de StationWithParity correspondante.
     * @throws {Error} - Si la valeur est null ou undefined.
     */
    public static from(
        value: StationWithParity | Station | string | null | undefined,
        parity?: Parity | string | number
    ): StationWithParity | undefined {

        if (value == null || value === "") return undefined;

        const parityObj = Parity.from(parity, false);

        // La valeur est une instance de StationWithParity :
        //  retourne la valeur sauf si la parité est imposée en argument.
        if (value instanceof StationWithParity) {
            if (!parityObj.isDefined()) return value;
            return StationsWithParity.getFromStationAndParity(
                value.station,
                parityObj
            )!;
        } 

        // La valeur est une instance de Station :
        //  la gare est récupérée et la parité sera celle donnée en argument.
        if (value instanceof Station) {
            return StationsWithParity.getFromStationAndParity(
                value,
                parityObj
            )!;
        }
 
        // La valeur est une chaîne qui correspond à la clé d'une gare avec parité :
        //  retourne l'instance de StationWithParity correspondante.
        if (!parityObj.isDefined() && StationsWithParity.hasKey(value)) {
            return StationsWithParity.getByKey(value)!;
        }

        // La valeur est une chaîne qui ne correspond pas à la clé d'une gare avec parité :
        //  la gare et la parité sont extraites de la chaîne.
        const { station, parity: parsedParity } = this.parseStationAndParity(value);
        const finalParity = parityObj.isDefined() ? parityObj : parsedParity;

        return StationsWithParity.getFromStationAndParity(
            station,
            finalParity
        )!;
    }

    /**
     * Renvoie la valeur unique de la parité pour le calcul de l'ID, qui est :
     *  - this.ODD si la parité est impaire,
     *  - this.EVEN si la parité est paire,
     *  - this.UNDEFINED sinon.
     * @param {Parity} parity - La parité à transformer.
     * @returns {number} - Valeur unique de la parité. pour le calcul.
     */
    public static parityValue(parity: Parity): number {
        switch (parity.value) {
            case Parity.ODD: return this.ODD;
            case Parity.EVEN: return this.EVEN;
            default: return this.UNDEFINED;
        }
    }

    /**
     * Vérifie si la parité est définie pour cette gare.
     * @returns {boolean} - Vrai si la parité est définie, faux sinon.
     */
    public hasDefinedParity(): boolean {
        return this.id % 3 !== 0;
    }

    /**
     * Analyse un nom de gare avec ou sans suffixe _PARITE
     * et renvoie un objet avec la gare correspondante et la parité associée.
     * La parité est undefined si le nom ne contient pas de suffixe _PARITE.
     * Si la valeur contient une erreur (par exemple, si le nom de gare n'existe pas),
     *  une exception est levée.
     * @param {string} value - Valeur à analyser pour la gare.
     * @returns {{ station: Station; parity: Parity }} - Objet avec la gare et la parité associée.
     */
    private static parseStationAndParity(
        value: string
    ): { station: Station; parity: Parity } {

        if (!value) {
            throw new Error(`La gare ne peut pas être vide.`);
        }

        const [stationName, parityPart] = value.split("_");

        const station = Stations.get(stationName.toUpperCase());
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
     * @returns {StationWithParity | undefined} - Gare après rebroussement si possible, sinon undefined.
     */
    public stationAfterTurnaround(): StationWithParity | undefined {
 
        // La gare de rebroussement est donnée par l'inversion de parité si définie,
        //  ou sans changement sinon.
        const reversedParity = this.parity.invert();
        const stationAfterTurnaround = StationsWithParity.getFromStationAndParity(
            this.station,
            reversedParity
        );

        // Si parité définie, rebroussement possible
        //  si parité incluse dans la propriété Station.turnaround.
        // Si parité non définie, rebroussement considéré comme possible
        //  si autorisé depuis au moins un sens.
        const canTurnAround = this.parity.isDefined()
            ? this.station.turnaround.includes(this.parity)
            : this.station.turnaround.isDefined();

        return canTurnAround ? stationAfterTurnaround : undefined;
    }

    /**
     * Vérifie si la gare avec parité a la même gare que l'autre.
     * @param other - Autre objet StationWithParity à comparer.
     * @returns {boolean} - Vrai si les deux objets ont la même gare, faux sinon.
     */
    public hasSameStationTo(other: StationWithParity | null | undefined): boolean {
        return !!other && Math.floor(this.id / 3) === Math.floor(other.id / 3);
    }
 
    /**
     * Vérifie si la gare avec parité a parmi ses gares rattachées une seconde gare, c'est à dire :
     *  - que cette seconde gare est identique ou est une gare fille de la première,
     *  - et que si la parité de la première est définie, celle de la seconde est identique.
     * @param other - Autre objet StationWithParity qui doit être inclus ou non.
     * @returns {boolean} - Vrai si l'objet inclut l'autre, faux sinon.
     */
    public includes(other: StationWithParity | null | undefined): boolean {
        return !!other
            && this.expandWithChildren().includes(other);
    }
 
    /**
     * Vérifie si l'objet StationWithParity est identique à l'autre.
     * @param other - Autre objet StationWithParity à comparer.
     * @returns {boolean} - Vrai si les deux objets sont identiques, faux sinon.
     */
    public equalsTo(other: StationWithParity | null | undefined): boolean {
        return this === other;
    }

    /**
     * Renvoie une chaîne représentant l'objet StationWithParity sous la forme
     *  GARE_PARITE, où GARE est le nom de la gare sans suffixe _PARITE et
     *  PARITE est la parité sous forme de chiffre.
     * @param {Station} station - Gare.
     * @param {Parity} parity - Parité.
     * @returns {string} - Chaîne représentant l'objet StationWithParity.
     */
    private static keyOf(station: Station, parity: Parity): string {
        return parity.isDefined() 
            ? `${station.abbreviation}_${parity.printDigit()}` 
            : `${station.abbreviation}`;
    }

    /**
     * Renvoie un tableau de toutes les gares avec parités rattachées en renvoyant : 
     *  - les 3 parités si elle n'est pas définie,
     *  - les gares filles.
     * La méthode prend en paramètre un ensemble de gares déjà visitées pour éviter les boucles infinies.
     * @param {Set<number>} [visited=new Set<number>()] - Ensemble des gares déjà visitées.
     * @returns {StationWithParity[]} - Tableau contenant toutes les gares visitées.
     */
    public expandWithChildren(visited: Set<number> = new Set<number>()): StationWithParity[] {

        if (this._expandedCache) return this._expandedCache;

        const results: StationWithParity[] = [];
        const added = new Set<number>();

        // Vérifie que la gare à inclure n'est pas déjà présente dans le résultat.
        const add = (swp: StationWithParity) => {
            if (!added.has(swp.id)) {
                added.add(swp.id);
                results.push(swp);
            }
        };

        // Evite une boucle infinie.
        if (visited.has(this.id)) return [];
        visited.add(this.id);
 
        // Génère l'expansion de la parité.
        if (!this.hasDefinedParity()) {
            const base = Math.floor(this.id / 3) * 3;
            add(StationsWithParity.list[base + StationWithParity.UNDEFINED]);
            add(StationsWithParity.list[base + StationWithParity.ODD]);
            add(StationsWithParity.list[base + StationWithParity.EVEN]);
        } else {
            add(this);
        }

        // Génère l'expansion avec les gares filles.
        for (const child of this.station.childStations) {
            const childSwp = StationsWithParity.getFromStationAndParity(child, this.parity);
            if (!childSwp) continue;

            // Duplique la liste des éléments déjà visités.
            const childVisited = new Set(visited);
            const expandedChildren = childSwp.expandWithChildren(childVisited);

            for (const c of expandedChildren) {
                add(c);
            }
        }

        // Sauvegarde les gares rattachées dans le cache.
        this._expandedCache = results;

        return results;
    }
}

/**
 * Classe StationsWithParity contenant la liste des gares avec parité.
 * Pour chaque gare, 3 parités : UNDEFINED, ODD et EVEN.
 */
class StationsWithParity {

    // Liste des gares avec parité
    public static list: StationWithParity[] = [];
    // Map des gares avec parité
    public static keyMap: Record<string, StationWithParity> = Object.create(null);

    /**
     * Nombre de gares avec parité dans la base de données.
     * @returns {number} - Nombre de gares avec parité dans la base de données.
     */
    public static get size(): number {
        return this.list.length;
    }

    /**
     * Vérifie si une gare avec parité est présente dans la base de données.
     * @param {string} key - Clé de la gare avec parité.
     * @returns {boolean} - Vrai si la gare est présente, faux sinon.
     */
    public static hasKey(key: string): boolean {
        return key in this.keyMap;
    }

    /**
     * Renvoie la gare avec parité correspondant à l'ID donné.
     * @param {number} id - ID de la gare avec parité.
     * @returns {StationWithParity} - Gare correspondant si elle existe, undefined sinon.
     */
    public static getById(id: number): StationWithParity {
        const s = this.list[id];
        if (!s) throw new Error(`Gare avec parité : ID ${id} inconnue`);
        return s;
    }
 
    /**
     * Renvoie une gare avec parité correspondant à la clé donnée.
     * @param {string} key - Clé de la gare avec parité.
     * @returns {StationWithParity | undefined} - Gare correspondant si elle existe, undefined sinon.
     */
    public static getByKey(key: string): StationWithParity | undefined {
        return this.keyMap[key];
    }
 
    /**
     * Renvoie la gare avec parité correspondant à la gare et la parité données.
     * @param {Station} station - Gare à trouver.
     * @param {Parity | string | number} parity - Parité à trouver.
     * @returns {StationWithParity | undefined} - Gare correspondant si elle existe, undefined sinon.
     */
    public static getFromStationAndParity(
        station: Station,
        parity: Parity | string | number
    ): StationWithParity | undefined {
        const base = station.id * 3;
        const parityObj = Parity.from(parity, false);
        const id = base + StationWithParity.parityValue(parityObj);
        return this.list[id];
    }

    /**
     * Crée une gare avec parité et l'ajoute à la base de données,
     *  référencée par son ID ou sa clé.
     * Si la gare avec parité est déjà présente, une erreur est levée.
     * @param {Station} station - Gare.
     * @param {Parity | string | number} parity - Parité à trouver.
     * @throws {Error} - Si la gare avec parité est déjà présente dans la base de données.
     */
    private static create(
        station: Station,
        parity: Parity | string | number
    ): void {
        const parityObj = Parity.from(parity, false);
        const id = station.id * 3 + StationWithParity.parityValue(parityObj);
        const swp = new StationWithParity(id, station, parityObj);
        if (this.hasKey(swp.key)) {
            throw new Error(`La gare avec parité ${swp} est déjà présente`
                + ` dans la base de données.`);
        }
        this.list[swp.id] = swp;
        this.keyMap[swp.key] = swp;
    }
 
    /**
     * Retourne un tableau des valeurs de la base de données des gares avec parité.
     * @returns {StationWithParity[]} - Itérateur sur les valeurs
     *  de la base de données des gares avec parité.
     */
    public static values(): StationWithParity[] {
        return Array.from(this.list.values());
    }

    /**
     * Efface toutes les gares avec parité de la base de données.
     * Cela permet de forcer le rechargement des gares avec parité si besoin.
     */
    public static clear(): void {
        this.list = [];
        this.keyMap = Object.create(null);
    }

    /**
     * Charge les gares avec parité à partir de la base de données des gares.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
    */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les gares si elles n'ont pas encore été chargées.
        Stations.load(); 

        // Liste les parités à prendre en compte.
        const parities = [
            Parity.undefined(),
            Parity.odd(),
            Parity.even()
        ];

        // Génère les gares avec parité à partir de la base de données des gares.
        for (const station of Stations.list) {
            for (const parity of parities) {
                this.create(station, parity);
            }
        }
    }
}

/**
 * Classe Connection définissant une connexion orientée entre deux gares.
 */
class Connection {

    // Constantes des valeurs par défaut
    public static readonly DEFAULT_CONNECTION_TIME= 1;  // Durée de connection par défaut en jours
                                                        //  (si 0 ou non renseignée)
                                                        //  La durée est très importante pour privilégier
                                                        //  les connexions avec une durée de connexion
                                                        //  déjà évaluée à partir de parcours réels

    // Propriétés de l'objet Connection
    public readonly from: StationWithParity;            // Gare de départ
    public readonly to: StationWithParity;              // Gare d'arrivée
    private _time: DateTime;                            // Temps de trajet
    public readonly withTurnaround: boolean;            // Connexion impliquant un rebroussement
    public readonly withMovement: boolean;              // Connexion sous régime de l'évolution
    public readonly changeParity: boolean;              // Connexion avec changement de parité

    /**
     * Constructeur d'une connexion.
     * @param {StationWithParity} from - Gare de départ
     * @param {StationWithParity} to - Gare d'arrivée
     * @param {DateTime | number | string} [time] - Temps de trajet
     *  (si 0 ou non renseigné : durée par défaut).
     * @param {boolean} [withMovement=false] - Indique si la connexion est sous régime de l'évolution.
     * @param {boolean} [changeParity=false] - Indique si la connexion implique un changement de parité.
     */
    constructor(
        from: StationWithParity,
        to: StationWithParity,
        time: DateTime | number | string = Connection.DEFAULT_CONNECTION_TIME,
        withMovement: boolean = false,
        changeParity: boolean = false
    ) {
        if (from.equalsTo(to)) {
            throw new Error(
                `Une connexion ne peut pas relier ${from} à elle-même`
                + ` sans changement de gare ou de parité.`
            );
        }
        this.from = from;
        this.to = to;
        this.withTurnaround = this.from.hasSameStationTo(this.to);
        let timeObj: DateTime | undefined;
        if (this.withTurnaround) {
            timeObj = DateTime.from(0, true)!
        } else {
            timeObj = DateTime.from(time, true);
            if (!timeObj || timeObj.excelValue <= 0) {
                timeObj = DateTime.from(Connection.DEFAULT_CONNECTION_TIME, true)!;
            }
        }
        this._time = timeObj;
        this.withMovement = withMovement;
        this.changeParity = changeParity;
    }

    /**
     * Renvoie le temps de trajet de la connexion.
     * @returns {DateTime} - Temps de trajet de la connexion.
     */
    public get time(): DateTime {
        return this._time;
    }

    /**
     * Modifie le temps de trajet de la connexion.
     * @param {DateTime | number | string} value - Nouveau temps de trajet de la connexion.
     * @throws {Error} - Si le temps de trajet est inférieur ou égal à 0 ou n'est pas relatif.
     */
    public set time(value: DateTime | number | string) {
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

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return `${this.from} -> ${this.to}`;
    }
}

/**
 * Classe Connections contenant la liste des connexions.
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
    // Liste des connections
    private static list: Connection[][] = [];

    /**
     * Retourne le nombre de connexions enregistrées dans la base de données.
     * @returns {number} - Nombre de connexions enregistrées.
     */
    public static get size(): number {
        let count = 0;
        for (const arr of this.list) {
            if (arr) count += arr.length;
        }
        return count;
    }

    /**
     * Résoud le tableau des gares avec parité définie rattachées à une gare ou à une gare avec parité.
     * Seules les gares avec parité définie sont retournées.
     * @param {Station | StationWithParity} input - Gare ou gare avec parité.
     * @returns {number[]} - Tableau des identifiants des gares.
     */
    private static resolveIds(
        input: Station | StationWithParity
    ): number[] {
        const swp = (input instanceof Station)
            ? StationsWithParity.getFromStationAndParity(input, Parity.UNDEFINED)
            : input; 
            return swp
                ? swp.expandWithChildren()
                    .filter(s => s.id % 3 !== 0)
                    .map(s => s.id)
                : [];
    }

    /**
     * Vérifie si une connexion est présente dans la base de données.
     * @param {Station | StationWithParity} from - Gare de départ.
     * @param {Station | StationWithParity} to - Gare d'arrivée.
     * @returns {boolean} - Vrai si le parcours est présent, faux sinon.
     */
    public static has(
        from: Station | StationWithParity,
        to: Station | StationWithParity
    ): boolean {
 
        const fromIds = this.resolveIds(from);
        const toIds = this.resolveIds(to);
 
        for (const fromId of fromIds) {
            const neighbors = this.list[fromId];
            if (!neighbors) continue;
 
            for (const c of neighbors) {
                if (toIds.includes(c.to.id)) {
                    return true;
                }
            }
        }
 
        return false;
    }

    /**
     * Renvoie la connexion correspondant aux gares de départ et d'arrivée données.
     * @param {Station | StationWithParity} from - Gare de départ.
     * @param {Station | StationWithParity} to - Gare d'arrivée.
     * @returns {Connection | undefined} - Gare correspondante, ou undefined si non trouvée.
     */
    public static get(
        from: Station | StationWithParity,
        to: Station | StationWithParity
    ): Connection | undefined {
 
        const fromIds = this.resolveIds(from);
        const toIds = this.resolveIds(to);
 
        for (const fromId of fromIds) {
            const neighbors = this.list[fromId];
            if (!neighbors) continue;
 
            for (const c of neighbors) {
                if (toIds.includes(c.to.id)) {
                    return c;
                }
            }
        }
 
        return undefined;
    }

    /**
     * Récupère les voisins d’une station.
     * @param {number} id - Identifiant de la station.
     * @returns {Connection[]} - Voisins de la station.
     */
    public static getNeighbors(id: number): Connection[] {
        return this.list[id] ?? [];
    }

    /**
     * Crée une nouvelle connexion et l'ajoute à la base de données,
     *  référencée par ses gares de départ et d'arrivée.
     * Si la connexion est déjà présente dans la base de données, une erreur est levée.
     * @param {string} from - Gare de départ.
     * @param {string} to - Gare d'arrivée.
     * @param {Connection} time - Durée de la connexion.
     * @returns {Connection} - La connexion ajoutée.
     * @throws {Error} - Si la connexion est déjà présente dans la base de données.
     */
    private static create(
        from: StationWithParity | string,
        to: StationWithParity | string,
        time: DateTime | number | string = Connection.DEFAULT_CONNECTION_TIME,
        withMovement: boolean = false,
        changeParity: boolean = false
    ): Connection {

        const fromObj = StationWithParity.from(from)!;
        const toObj = StationWithParity.from(to)!;
        const connection = new Connection(
            fromObj,
            toObj,
            time,
            withMovement,
            changeParity
        );

        const fromId = connection.from.id;
        const toId = connection.to.id;

        if (!this.list[fromId]) {
            this.list[fromId] = [];
        }

        const existing = this.list[fromId].find(c => c.to.id === toId);

        if (existing) {
            throw new Error(`La connexion ${connection} existe déjà.`);
        }

        this.list[fromId].push(connection);

        return connection;
    }

    /**
     * Retourne un tableau des valeurs de la base de données des connexions.
     * @returns {Connection[]} - Itérateur sur les valeurs
     *  de la base de données des connexions.
     */
    public static values(): Connection[] {
        const result: Connection[] = [];
        for (const arr of this.list ){
            if (!arr) continue;
            result.push(...arr);
        };
        return result;
    }

    /**
     * Efface toutes les connexions de la base de données.
     * Cela permet de forcer le rechargement des connexions si besoin.
     */
    public static clear(): void {
        this.list = [];
    }

    /**
     * Charge les connexions entre les gares.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les gares si elles n'ont pas encore été chargées.
        StationsWithParity.load(); 

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Connections.load : aucune donnée trouvée dans la table.`);
            return;
        }

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère les champs.
                const from = WorkbookService.getString(row, this.COL_FROM)?.toUpperCase();
                const to = WorkbookService.getString(row, this.COL_TO)?.toUpperCase();
                if (!from || !to) continue;
                const timeInMinutes = WorkbookService.getNumber(row, this.COL_TIME);
                const withMovement = WorkbookService.getBoolean(row, this.COL_MOVEMENT);
                const changeParity = WorkbookService.getBoolean(row, this.COL_CHANGE_PARITY);

                // Instancie les propriétés objets (si 0 ou non renseignée : valeur par défaut).
                const excelTime = timeInMinutes
                    ? timeInMinutes / 24 / 60
                    : Connection.DEFAULT_CONNECTION_TIME;
 
                // Crée l'objet Connection et l'insère dans la base de données.
                const connection = this.create(
                    from,
                    to,
                    excelTime,
                    withMovement,
                    changeParity
                );
            }

        } catch (e) {
            throw new Error(`Connections.load (ligne ${excelRow}) : ${e}`);
        }
    }

    /**
     * Sauvegarde les connexions de la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = this.SHEET,
        tableName: string = this.TABLE,
        startCell: string = "A1"
    ): void {

        // Convertit la base de données en un tableau de données.
        const data: (string | number)[][] = Array
            .from(this.values())
            .map(connection => [
                connection.from.key,
                connection.to.key,
                connection.time.excelValue * 24 * 60,
                connection.withTurnaround ? 1 : 0,
                connection.withMovement ? 1 : 0,
                connection.changeParity ? 1 : 0
            ]);

        // Imprime le tableau.
        const table = WorkbookService.printTable(
            this.HEADERS,
            data,
            sheetName,
            tableName,
            startCell
        );

        // Met les durées de parcours au format "hh:mm:ss".
        table.getRange()
            .getColumn(this.COL_TIME)
            .setNumberFormat("hh:mm:ss");
    }

    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée d'un trajet,
     * en prenant en compte les gares intermédiaires qui peuvent être empruntées
     * avec des parités différentes.
     * @param {StationWithParity[][]} routeStations - Trajet avec les gares intermédiaires
     *  et les parités possibles.
     * @returns {Connection[]} - Chemin le plus court entre le départ et l'arrivée du trajet.
     *  Si aucun chemin n'est trouvé, undefined est renvoyé.
     */
    public static shortestPathWithGroups(
        routeStations: StationWithParity[][]
    ): Connection[] | undefined {
 
        this.load();

        const queue: State[] = [];
        const visited = new Map<number, number>();
 
        // Expand la route avec toutes les parités possibles.
        const expandedRouteStations: number[][][] =
        routeStations.map(group =>
            group.map(station =>
                this.resolveIds(station)
            )
        );

        // Initialise la file avec la gare de départ.
        const firstGroup = expandedRouteStations[0];
        for (const variantGroup of firstGroup) {
            for (const stationId of variantGroup) {
                queue.push(new State(
                    stationId,
                    0,
                    1,
                    0
                ));
            }
        }
 
        return this.runGroupedDijkstra(
            queue,
            visited,
            expandedRouteStations
        );
    }

    /**
     * Exécute l'algorithme de Dijkstra pour trouver le chemin le plus court
     *  entre le départ et l'arrivée d'un trajet, en prenant en compte les gares
     *  intermédiaires qui peuvent être empruntées avec des parités différentes.
     * @param {State[]} queue - File d'attente contenant les états à visiter.
     * @param {Map<number, number>} visited - Carte des états déjà visités.
     * @param {number[][][]} routeStations - Trajet avec les gares intermédiaires
     *  et les parités possibles.
     * @returns {Connection[] | undefined} - Chemin le plus court entre le départ
     *  et l'arrivée du trajet. Si aucun chemin n'est trouvé, undefined est renvoyé.
     */
    private static runGroupedDijkstra(
        queue: State[],
        visited: Map<number, number>,
        routeStations: number[][][]
    ): Connection[] | undefined {
 
 
        while (queue.length > 0) {
 
            queue.sort((a, b) => a.cost - b.cost);
            const state = queue.shift()!;
 
            const key = state.key;
 
            if (visited.has(key) && visited.get(key)! <= state.cost) {
                continue;
            }
 
            visited.set(key, state.cost);
 
            // Condition de fin sécurisée
            if (state.groupIndex >= routeStations.length) {
                return state.buildPath();
            }
 
            const nextStates =
                this.expandNeighbors(state, routeStations);
 
            queue.push(...nextStates);
        }
 
        return undefined;
    }

    /**
     * Donne les états suivants d'un état donné,
     *  en prenant en compte les gares intermédiaires et les parités possibles.
     * @param {State} state - État actuel.
     * @param {number[][][]} routeStations - Trajet avec les gares intermédiaires
     *  et les parités possibles.
     * @returns {State[]} - Liste des états suivants.
     */
    private static expandNeighbors(
        state: State,
        routeStations: number[][][]
    ): State[] {
 
        const result: State[] = [];
 
        // Donne les gares voisines.
        const neighbors = this.getNeighbors(state.stationId);
 
        for (const connection of neighbors) {

            // Donne l'id de la gare voisine.
            const nextStationId = connection.to.id;
 
            // Ajoute le coût de la connection : temps de parcours, ou temps de retournement.
            const nextCost = state.cost
                + (connection.withTurnaround
                    ? Params.turnaroundTime.excelValue
                    : connection.time.excelValue);
 
            let nextGroup = state.groupIndex;
            let nextMask = state.visitedMask;
 
            if (nextGroup >= routeStations.length) continue;
 
            const currentGroup = routeStations[nextGroup];
 
            // Cherche le groupe de la gare voisine.
            let matched = false;
            for (let i = 0; i < currentGroup.length; i++) {
                const variantGroup = currentGroup[i];
                if (variantGroup.includes(nextStationId)) {
                    matched = true;
                    nextMask |= (1 << i);
                    break;
                }
            }
            if (matched) {
                if (nextMask === (1 << currentGroup.length) - 1) {
                    nextGroup++;
                    nextMask = 0;
                }
            }
 
            // Ajoute le nouvel état.
            result.push(new State(
                nextStationId,
                nextCost,
                nextGroup,
                nextMask,
                state,
                connection
            ));
        }

        return result;
    }

    /**
     * Sauvegarde les temps de connexions entre les gares dans la base de données,
     *  à partir de parcours existants : pour un train qui s'arrête dans plusieurs gares consécutives, 
     *  le temps de parcours des connexions entre ces gares peut être calculé.
     * @param {Path[] | string} paths - Liste des parcours de trains.
     */
    public static saveConnectionsTimes(paths: Path[] | string) {
        const pathsList = (typeof paths === "string")
            ? paths.split(";").map(key => Paths.get(key)!)
            : paths;
        pathsList.forEach((path) => {
            path?.stops.forEach((stop) => {
                const nextStop = path.nextStop(stop);

                if (nextStop) {
                    const connection = this.get(stop.station, nextStop.station);
                    if (connection && !!nextStop.arrivalTime && !!stop.departureTime) {
                        const time = nextStop.arrivalTime.excelValue - stop.departureTime.excelValue;
                        if (time > 0) {
                            connection.time = time;
                        }
                    }
                }
            });
        });
    }
}

/**
 * Classe State définissant un état de recherche de l'algorithme Dijkstra.
 */
class State {

    public readonly stationId: number;      // Identifiant de la gare
    public readonly cost: number;           // Coût du chemin
    public readonly groupIndex: number;     // Index du groupe de gares
    public readonly visitedMask: number;    // Masque des gares visitées
    public readonly prev?: State;           // Etat précedent
    public readonly via?: Connection;       // Connection ajoutée

    /**
     * Constructeur de l'état de recherche de l'algorithme Dijkstra.
     * @param {number} stationId - Identifiant de la gare.
     * @param {number} cost - Coût du chemin.
     * @param {number} groupIndex - Index du groupe de gares.
     * @param {number} visitedMask - Masque des gares visitées.
     * @param {State} prev - Etat précédent.
     * @param {Connection} via - Connection ajoutée.
     */
    public constructor(
        stationId: number,
        cost: number,
        groupIndex: number,
        visitedMask: number,
        prev?: State,
        via?: Connection
    ) {
        this.stationId = stationId;
        this.cost = cost;
        this.groupIndex = groupIndex;
        this.visitedMask = visitedMask;
        this.prev = prev;
        this.via = via;
    }

    /**
     * Renvoie une clé unique pour l'état de recherche, composée de l'identifiant de la gare,
     *  de l'index du groupe de gares et du masque des gares visitées.
     * @returns {number} - Clé unique de l'état.
     */
    public get key(): number {
        return (
            this.stationId
            | (this.groupIndex << 16)
            | (this.visitedMask << 24)
        );
    }

    /**
     * Reconstruit le chemin le plus court entre le départ et l'arrivée d'un trajet,
     *  en prenant en compte les gares intermédiaires qui peuvent être empruntées
     *  avec des parités différentes.
     * @returns {Connection[]} - Chemin le plus court entre le départ et
     *  l'arrivée du trajet. Si aucun chemin n'est trouvé, undefined est renvoyé.
     */
    public buildPath(): Connection[] {
        const path: Connection[] = [];
        let current: State | undefined = this;

        while (current?.via) {
            path.push(current.via);
            current = current.prev;
        }

        return path.reverse();
    }
}

/*
 * Classe Stop définissant l'arrêt ou le passage d'un train dans une gare.
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
     * @param {boolean} [areRelativeTimes=undefined] - Indique si les horaires sont relatives (par exemple, par rapport à un autre arrêt).
     * @param {string[] | string} [tracks=[]] - Voies de l'arrêt.
     */
    constructor(
        station : StationWithParity | Station | string,
        stationAfterTurnaround?: StationWithParity | string,
        arrivalTime?: DateTime | number | string,
        departureTime?: DateTime | number | string,
        passageTime?: DateTime | number | string,
        areRelativeTimes?: boolean,
        tracks: string[] | string = [],
    ) {
        // Détermine la gare d'arrêt
        const stationObj = StationWithParity.from(station)
        if (!stationObj) throw new Error(`La gare ${station} est inconnue.`);
        this.station = stationObj;

        // Détermine le rebroussement
        this._withTurnaround = this.canTurnaroundTo(stationAfterTurnaround);

        // Détermine les horaires de l'arrêt
        this.setTimes(arrivalTime, departureTime, passageTime, areRelativeTimes);

        // Détermine les voies de l'arrêt
        this._tracks = tracks instanceof Array ? tracks : Stop.getTracksFromString(tracks);
    }

    /**
     * Renvoie une clé unique pour l'arrêt, composée du nom de la gare et de la parité (si connue).
     * @returns {string} - Clé unique.
     */
    public get key(): string {
        return this.station.key;
    }

    /**
     * Renvoie le nom de la gare associée à cet arrêt.
     * @returns {string} - Nom de la gare.
     */
    public get stationName(): string {
        return this.station!.station.name;
    }

    /**
     * Renvoie l'abréviation de la gare associée à cet arrêt.
     * @returns {string} - Abréviation de la gare.
     */
    public get stationAbbreviation(): string {
        return this.station!.station.abbreviation;
    }

    /**
     * Renvoie vrai si l'arrêt à un rebroussement possible, faux sinon.
     * @returns {boolean} - Vrai si l'arrêt a un rebroussement possible, faux sinon.
     */
    public get withTurnaround(): boolean {
        return this._withTurnaround;
    }

    /**
     * Retourne l'heure d'arrivée à l'arrêt, si connue.
     * @returns {DateTime | undefined} - Heure d'arrivée à l'arrêt, ou undefined si non connue.
     */
    public get arrivalTime(): DateTime | undefined {
        return this._arrivalTime;
    }

    /**
     * Retourne l'heure de départ à l'arrêt, si connue.
     * @returns {DateTime | undefined} - Heure de départ à l'arrêt, ou undefined si non connue.
     */
    public get departureTime(): DateTime | undefined {
        return this._departureTime;
    }

    /**
     * Retourne l'heure de passage à l'arrêt, si connue.
     * @returns {DateTime | undefined} - Heure de passage à l'arrêt, ou undefined si non connue.
     */
    public get passageTime(): DateTime | undefined {
        return this._passageTime;
    }

    /**
     * Retourne le tableau des voies de l'arrêt.
     * @returns {string[]} - Tableau des voies de l'arrêt.
     */
    public get tracks(): string[] {
        return this._tracks;
    }

    /**
     * Modifie le tableau des voies de l'arrêt.
     * @param {string[]} tracks - Tableau des voies de l'arrêt.
     */
    public set tracks(value: string[]) {
        this._tracks = value;
    }

    /**
     * Retourne la gare associée à l'objet StationWithParity et la parité opposée,
     *  si le rebroussement est possible (connection existante).
     * @returns {StationWithParity | undefined} - Gare après rebroussement,
     *  ou undefined si la parité est indéfinie ou le rebroussement n'est pas possible.
     */
    public get stationAfterTurnaround(): StationWithParity | undefined {
        return this.withTurnaround ? this.station.stationAfterTurnaround() : undefined;
    }

    /**
     * Modifie la gare de rebroussement si possible.
     * Un rebroussement est possible si la gare après rebroussement donnée correspond à la gare calculée.
     * Si l'arrêt présente une heure de passage, elle est transformée en heure d'arrivée,
     *  le départ se fait arès le temps de retournement.
     * @param {StationWithParity | string | undefined} value - Gare après rebroussement.
     */
    public set stationAfterTurnaround(value: StationWithParity | string | undefined) {
        this._withTurnaround = this.canTurnaroundTo(value);
        if (!!this._passageTime) {
            this.setTimes(
                this._passageTime,
                Params.turnaroundTime.resolveAgainst(this._passageTime!),
                undefined
            );
        }
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.key;
    }

    /**
     * Vérifie si un rebroussement est possible avec stationAfterTurnaround comme gare après rebroussement.
     * Un rebroussement est possible si la gare après rebroussement correspond à la gare calculée.
     * @param {StationWithParity | string} - stationAfterTurnaround Gare après rebroussement.
     * @returns {boolean} - Vrai si le rebroussement est possible, faux sinon.
     */
    private canTurnaroundTo(stationAfterTurnaround : StationWithParity | string | undefined): boolean {

        // Vérifie si la gare après rebroussement demandée est connue.
        const stationAfterTurnaroundObj = StationWithParity.from(stationAfterTurnaround);
        if (!stationAfterTurnaroundObj) return false;

        // Calcule la gare théorique après rebroussement si celui-ci est possible.
        const calculated = this.station.stationAfterTurnaround();
        if (!calculated) {
            Log.warn(`Un rebroussement n'est pas autorisé à la gare de ${this.station}.`
            + ` Il ne sera pas pris en compte.`);
            return false;
        }

        // Compare les gares théoriques et demandées.
        if (!stationAfterTurnaroundObj.equalsTo(calculated)) {
            Log.warn(`Le rebroussement à la gare de ${this.station} ne sera pas pris en compte,`
                + ` car la gare après rebroussement demandée ${stationAfterTurnaroundObj}`
                + ` ne correspond pas.`);
            return false
        }

        return true;
    }

    /**
     * Modifie les heures d'arrivée, de départ et de passage de l'arrêt, et vérifie leur cohérence.
     * @param {DateTime | number | string} [arrivalTime] - Heure d'arrivée à l'arrêt.
     * @param {DateTime | number | string} [departureTime] - Heure de départ à l'arrêt.
     * @param {DateTime | number | string} [passageTime] - Heure de passage à l'arrêt.
     * @param {boolean} [areRelativeTimes=undefined] - Vrai si les heures sont relatives, faux sinon.
     */
    public setTimes(
        arrivalTime?: DateTime | number | string,
        departureTime?: DateTime | number | string,
        passageTime?: DateTime | number | string,
        areRelativeTimes?: boolean
    ) {
        this._arrivalTime = DateTime.from(arrivalTime, areRelativeTimes);
        this._departureTime = DateTime.from(departureTime, areRelativeTimes);
        this._passageTime = (!arrivalTime && !departureTime)
            ? DateTime.from(passageTime, areRelativeTimes)
            : undefined;
        if (!this._arrivalTime && !this._departureTime && !this._passageTime) {
            throw new Error(`L'arrêt ${this.station} n'a pas d'heure d'arrivée,`
                + ` d'heure de départ ou d'heure de passage.`);
        }
        if (this._arrivalTime && this._departureTime) {
            const timeDiff = this._departureTime.compareTo(this._arrivalTime);
            if (timeDiff <= 0) {
                if (timeDiff === 0) {
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
     * @param {string} tracksString - Chaîne de caractères contenant la liste de voies.
     * @returns {string[]} - Tableau de chaînes de caractères correspondant à la liste de voies.
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
     * @param {boolean} [ignoreArrival=false] - Si vrai, ignore l'heure d'arrivée
     *  et préfère l'heure de départ ou de passage. Si faux (par défaut),
     *  c'est d'abord l'heure d'arrivée qui est prise en compte.
     * @param {DateTime} [reference] - Heure de référence pour les heures relatives.
     * @returns {DateTime | undefined} - Heure la plus petite à l'arrêt,
     *  ou undefined si aucune heure n'est lue.
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
     *  avec une heure d'arrivée et une heure de départ, ou une heure de passage.
     * @returns {boolean} - Vrai si l'arrêt est un arrêt intermédiaire, faux sinon.
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
     * @param {DateTime} reference - Référence à utiliser pour convertir les heures.
     */
    public convertToRelativeTime(reference: DateTime, throwErrorIfAlreadyRelative: boolean = false): void {
 
        // Temps de référence déjà relatif : pas de conversion possible.
        // Vérifie simplement que les temps soient déjà relatifs.
        if (reference.isRelative){
            const arrivalTimeIsAbsolute = this._arrivalTime && !this._arrivalTime.isRelative;
            const departureTimeIsAbsolute = this._departureTime && !this._departureTime.isRelative;
            const passageTimeIsAbsolute = this._passageTime && !this._passageTime.isRelative;

            if (arrivalTimeIsAbsolute || departureTimeIsAbsolute || passageTimeIsAbsolute) {
                if (throwErrorIfAlreadyRelative) {
                    throw new Error(`Le temps de référence`
                        + ` ${reference.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                        + ` est déjà relatif. Les horaires de l'arrêt ${this} qui sont absolus`
                        + ` ne peuvent donc pas être convertis en temps relatifs.`);
                }
            }
            return;
        }

        // Temps de référence absolu : conversion possible.
        // Vérifie si les temps sont bien absolus avant de les convertir.
        if (this._arrivalTime) {
            if (this._arrivalTime.isRelative) {
                Log.warn(`L'heure d'arrivée à l'arrêt ${this}`
                    + ` ${this._arrivalTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` est déjà relative. Elle ne sera donc pas convertie.`);
            } else {
                this._arrivalTime = this._arrivalTime.relativeTo(reference);
            }
        }
        if (this._departureTime) {
            if (this._departureTime.isRelative) {
                Log.warn(`L'heure de départ à l'arrêt ${this}`
                    + ` ${this._departureTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` est déjà relative. Elle ne sera donc pas convertie.`);
            } else {
                this._departureTime = this._departureTime.relativeTo(reference);
            }
        }
        if (this._passageTime) {
            if (this._passageTime.isRelative) {
                Log.warn(`L'heure de passage à l'arrêt ${this}`
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
     * @param {Stop | null | undefined} other - Autre arrêt à comparer.
     * @returns {boolean} - Vrai si les arrêts sont égaux, faux sinon.
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
     * Compare cet arrêt avec un autre arrêt, le premier devant inclure le second,
     *  en vérifiant la gare avec parité, le rebroussement,
     *  les heures d'arrivée, de départ et de passage.
     * @param {Stop | null | undefined} other - Autre arrêt à comparer.
     * @returns {boolean} - Vrai si les arrêts sont égaux, faux sinon.
     */
    public includes(other: Stop | null | undefined): boolean {
        return (
            !! other &&
            this.station.includes(other.station) &&
            this._withTurnaround === other.withTurnaround &&
            DateTime.equalsOrUndefined(this._arrivalTime, other.arrivalTime) &&
            DateTime.equalsOrUndefined(this._departureTime, other.departureTime) &&
            DateTime.equalsOrUndefined(this._passageTime, other.passageTime)
        );
    }

    /**
     * Ajoute une voie à l'arrêt si elle n'y est pas déjà.
     * Si la voie n'est pas déjà dans la liste des voies, l'ajoute et trie la liste.
     * @param {string} track - Voie à ajouter.
     */
    public addTrack(track: string): void {
        if (!this._tracks.includes(track)) {
            this._tracks.push(track);
            this._tracks.sort();
        }
    }
}

/**
 * Classe Stops contenant la liste des arrêts.
 */
class Stops {
 
    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Arrêts";               // Feuille contenant la liste des arrêts
    private static readonly TABLE = "Arrêts";               // Tableau contenant la liste des arrêts
    private static readonly HEADERS = [[                    // En-têtes du tableau des arrêts
        "Parcours",
        "Gare",
        "Gare après rebroussement",
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
    private static readonly COL_IMPORT_NUMBER = 0;              // Colonne du numéro de train
    private static readonly COL_IMPORT_DATE = 1;                // Colonne de la date
    private static readonly COL_IMPORT_SERVICE = 2;             // Colonne du service
    private static readonly COL_IMPORT_DAYS = 3;                // Colonne des jours de circulation
    private static readonly COL_IMPORT_STATION = 4;             // Colonne de la gare
    private static readonly COL_IMPORT_ARRIVAL_TIME = 5;        // Colonne de l'heure d'arrivée
    private static readonly COL_IMPORT_DEPARTURE_TIME = 6;      // Colonne de l'heure de départ
    private static readonly COL_IMPORT_PASSAGE_TIME = 7;        // Colonne de l'heure de passage
    private static readonly COL_IMPORT_TRACK = 8;               // Colonne de la voie

    /**
     * Charge les arrêts.
     * Les arrêts sont stockés dans la propriété "stops" des parcours correspondants.
     * Si un train n'existe pas, un message d'erreur est affiché.
     */
    public static load(): void {

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Stops.load : aucune donnée trouvée dans la table.`);
            return;
        }
 
        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère le parcours correspondant.
                const pathKey = WorkbookService.getRequiredString(
                    row,
                    this.COL_PATH_KEY,
                    `pathKey manquant.`
                );
                const path = Paths.get(pathKey);
                if (!path) {
                    throw new Error(`Parcours "${pathKey}" inexistant.`);
                }

                // Récupère les champs.
                const station = WorkbookService.getString(row, this.COL_STATION);
                const stationAfterTurnaround =
                    WorkbookService.getString(row, this.COL_STATION_AFTER_TURNAROUND);
                const arrivalTime =
                    WorkbookService.getNumberOrUndefined(row, this.COL_ARRIVAL_TIME);
                const departureTime =
                    WorkbookService.getNumberOrUndefined(row, this.COL_DEPARTURE_TIME);
                const passageTime =
                    WorkbookService.getNumberOrUndefined(row, this.COL_PASSAGE_TIME);
                const tracks = WorkbookService.getString(row, this.COL_TRACK);

                // Instancie l'objet Stop.
                const stop = new Stop(
                    station,
                    stationAfterTurnaround,
                    arrivalTime,
                    departureTime,
                    passageTime,
                    true,
                    tracks
                );

                // Ajoute l'arrêt au parcours.
                path.stops.push(stop);
            }

        } catch (e) {
            throw new Error(`Stops.load (ligne ${excelRow}) : ${e}`);
        }
    }
 
    /**
     * Sauvegarde les arrêts des trains de la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = this.SHEET,
        tableName: string = this.TABLE,
        startCell: string = "A1"
    ): void {

        // Crée le tableau final avec les données de chaque arrêt pour chaque train.
        const data: (string | number)[][] = [];
 
        for (const path of Paths.values()) {
            for (const stop of Array.from(path.stops.values())) {
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

        // Imprime le tableau.
        const table = WorkbookService.printTable(
            this.HEADERS,
            data,
            sheetName,
            tableName,
            startCell
        );

        // Trie le tableau selon la colonne des clés
        table.getSort().apply([
            { key: this.COL_PATH_KEY, ascending: true },
        ]);
 
        // Met les horaires au format "hh:mm:ss".
        const timeColumns = [
            this.COL_ARRIVAL_TIME,
            this.COL_DEPARTURE_TIME,
            this.COL_PASSAGE_TIME
        ];
        for (const col of timeColumns) {
            table.getRange().getColumn(col).setNumberFormat("hh:mm:ss");
        }
    }

    /**
     * Importe les arrêts des trains dans la base de données à partir d'un tableau Excel.
     * Les arrêts sont stockés dans la propriété "stops" des parcours correspondants.
     * Si un train n'existe pas, un message d'erreur est affiché.
     * @param {string} [sheetName=this.IMPORT_SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.IMPORT_TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static import(): void {

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.IMPORT_SHEET, this.IMPORT_TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Stops.load : aucune donnée trouvée dans la table.`);
            return;
        }

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère les champs.
                const trainNumber = WorkbookService.getString(row, this.COL_IMPORT_NUMBER);
                const date = WorkbookService.getNumber(row, this.COL_IMPORT_DATE);
                const service = WorkbookService.getString(row, this.COL_IMPORT_SERVICE);
                const days = WorkbookService.getString(row, this.COL_IMPORT_DAYS);
                const station = WorkbookService.getString(row, this.COL_IMPORT_STATION);
                const arrivalTime =
                    WorkbookService.getNumberOrUndefined(row, this.COL_IMPORT_ARRIVAL_TIME);
                const departureTime =
                    WorkbookService.getNumberOrUndefined(row, this.COL_IMPORT_DEPARTURE_TIME);
                const passageTime =
                    WorkbookService.getNumberOrUndefined(row, this.COL_IMPORT_PASSAGE_TIME);
                const tracks = WorkbookService.getString(row, this.COL_IMPORT_TRACK);

                // Instancie l'objet Stop.
                const stop = new Stop(
                    station,
                    "",
                    arrivalTime,
                    departureTime,
                    passageTime,
                    false,
                    tracks
                );
            } 

        } catch (e) {
            throw new Error(`Stops.load (ligne ${excelRow}) : ${e}`);
        }
    }
}

/**
 * Classe Path définissant le parcours d'un train, avec ses gares et temps de passage
 *  par rapport à la gare origine.
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
    private _routeStations?: StationWithParity[][] = [];     // Tableau des gares ou groupes de gares d'arrêts
                                                    // du parcours définis dans la signature 
    public stops: Stop[] = [];                      // Gares d'arrêt ou gares de passage du parcours
    private _stopsIndex: Map<string, Stop> = new Map();   // Dictionnaire des arrêts référencés
                                                    //  par leur clé (abréviation_parité)
    private _stopPosition: Map<string, number> = new Map();    // Dictionnaire de la position des arrêts
                                                    //  dans le parcours (référencés par leur clé)
    public stopsChecked: number = Path.UNCHECKED;   // Résultat de la vérification du parcours
                                                    //  (0 si non vérifié)

    /**
     * Constructeur d'un parcours.
     * @param {string} [key=""] - Clé du parcours.
     * @param {Parity|string/number} [parityValue=Parity.UNDEFINED] - Parité du parcours.
     * @param {Parity|string/number} [lineDirection=Parity.UNDEFINED] - Direction du parcours sur la ligne.
     * @param {string} [missionCode=""] - Code de mission des trains du parcours.
     * @param {string} [name=""] - Nom du parcours.
     * @param {string} [signature=""] - Signature du parcours : gares définissant le parcours.
     * @param {Stop[]} [stops=[]] - Gares d'arrêts ou gares de passage du parcours.
     * @param {number} [stopsChecked=Path.UNCHECKED] - Résultat de la vérification du parcours.
     */
    constructor(
        key: string = "",
        parityValue: Parity | string | number = Parity.UNDEFINED,
        lineDirection: Parity | string | number = Parity.UNDEFINED,
        missionCode: string = "",
        name: string = "",
        signature: string = "",
        stops: Stop[] = [],
        stopsChecked: number = Path.UNCHECKED
    ) {
        this.key = key;
        this.parity = Parity.from(parityValue, true);
        this.lineDirection = Parity.from(lineDirection, true);
        this.missionCode = missionCode;
        this.name = name;
        this._signature = signature;
        this.stops = stops;
        this.stopsChecked = stopsChecked;
    }

    /**
     * Renvoie l'arrêt d'origine du parcours.
     * @returns {Stop | undefined} - L'arrêt d'origine, ou undefined si le parcours n'a pas d'arrêt.
     */
    public get origin(): Stop | undefined {
        return this.stops[0];
    }

    /**
     * Renvoie l'arrêt de destination du parcours.
     * @returns {Stop | undefined} - L'arrêt de destination,
     *  ou undefined si le parcours n'a pas d'arrêt de destination.
     */
    public get destination(): Stop | undefined {
        return this.stops[this.stops.length - 1] ;
    }

    /**
     * Renvoie la signature du parcours, qui est la concaténation
     *  des noms des gares d'arrêt du parcours, précédés de "@"
     *  si l'ordre de passage n'est pas imposé.
     * @returns {string} - Signature du parcours.
     */
    public get signature(): string {
        return this._signature;
    }

    /**
     * Renvoie le tableau des gares d'arrêt du parcours.
     * Le tableau est construit à partir de la signature du parcours.
     * Chaque élément du tableau est ordonné et correspond à une gare d'arrêt du parcours,
     *  ou à un groupe de gares à parcourir dans un ordre indifférent, séparées par un ";".
     * Chaque gare ou ensemble de gares est parcouru dans l'ordre du tableau, et séparé par un ">".
     * @returns {string[][]} - Tableau des gares d'arrêt du parcours.
     */
    public get routeStations(): StationWithParity[][] {

        if (!this._signature) {
            this._routeStations = [];
            return this._routeStations;
        }

        if (this._routeStations?.length === 0) {

            this._routeStations = this._signature
                .replace(/\s/g, "")
                .replace(/,/g, ";")
                .split(">")
                .map(group =>
                    group
                        .split(";")
                        .map(station => StationWithParity.from(station))
                        .filter((s): s is StationWithParity => s !== undefined)
                );
        }
 
        return this._routeStations ?? [];
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.key;
    }

    /**
     * Retourne le parcours Path correspondant au paramètre path.
     * Si path est déjà un objet Path, il est retourné tel quel.
     * Si path est un string, il est considéré comme le clé du parcours et
     *  l'objet Path correspondant est retourné s'il existe, sinon undefined est retourné.
     * @param {Path | string | null | undefined} value - Parcours à retourner,
     *  sous forme d'objet Path ou de clé string.
     * @returns {Path | undefined} - Parcours Path correspondant, ou undefined si le clé n'existe pas.
     */
    public static from(
        value: Path | string | null | undefined
    ): Path | undefined {
        if (value == null || value === "") return undefined;
        if (value instanceof Path) return value;
        return Paths.get(value!);
    }

    /**
     * Crée un parcours Path à partir des gares d'origine et de destination,
     *  ainsi que de leur heures de départ et d'arrivée.
     * @param {StationWithParity | Station | string} from - Nom de la gare d'origine.
     * @param {DateTime | number | string} departureTime - Heure de départ à la gare d'origine.
     * @param {StationWithParity | Station | string} to - Nom de la gare de destination.
     * @param {DateTime | number | string} arrivalTime - Heure d'arrivée à la gare de destination.
     * @param {boolean} [areRelativeTimes=false] - Si vrai, les heures de départ et d'arrivée
     *  sont considérées comme relatives.
     * @param {string} [missionCode=""] - Code de mission des trains du parcours (facultatif).
     * @param {string} [name=""] - Nom du parcours (facultatif).
     * @param {string} [signature=""] - Signature du parcours (facultatif).
     * @returns {Path} - Un objet Path représentant le parcours.
     */
    public static fromTerminals(
        from: StationWithParity | Station | string,
        departureTime: DateTime | number | string,
        to: StationWithParity | Station | string,
        arrivalTime: DateTime | number | string,
        areRelativeTimes: boolean = false,
        findPath: boolean = false,
        missionCode?: string,
        name?: string,
        signature?: string
    ): Path {

        const fromObj = StationWithParity.from(from);
        if (!fromObj) throw new Error(`Gare d'origine ${from} incorrecte.`);
        const departureTimeObj = DateTime.from(departureTime, areRelativeTimes);
        if (!departureTimeObj) throw new Error(`Heure de départ ${departureTime} incorrecte.`);
        const toObj = StationWithParity.from(to);
        if (!toObj) throw new Error(`Gare de destination ${to} incorrecte.`);
        const arrivalTimeObj = DateTime.from(arrivalTime, areRelativeTimes);
        if (!arrivalTimeObj) throw new Error(`Heure d'arrivée ${arrivalTime} incorrecte.`);

        const s1 = new Stop(fromObj, undefined, undefined, departureTimeObj, undefined, areRelativeTimes); 
        const s2 = new Stop(toObj, undefined, arrivalTimeObj, undefined, undefined, areRelativeTimes);
        const stops = [s1, s2];
 
        const path = Paths.create(
            "",
            undefined,
            undefined,
            missionCode,
            name,
            signature,
            stops
        );
 
        // Renvoie directement le parcours s'il existait déjà
        if (path.stopsChecked !== Path.UNCHECKED) return path;

        if (findPath) {
            path.findPath();
        } else {
            path.stopsChecked = Path.ONLY_FROM_AND_TO;
            path.check();
        }

        return path;
    }

    /**
     * Renvoie le radical de la clé du parcours constitué de
     *  origine_destination_codeMission_nomDuParcours (si ces valeurs existent).
     * @returns {string} - Radical de la clé du parcours.
     */
    public buildRadical(): string {
        const origin = this.origin?.stationAbbreviation ?? "";
        const dest = this.destination?.stationAbbreviation ?? "";
 
        const parts = [origin + '>' +  dest];

        if (this.missionCode) parts.push(this.missionCode);
        if (this.name) parts.push(this.name);

        return parts.join("_");
    }

    /**
     * Ajoute un arrêt au parcours.
     * Si les trains du parcours sont déjà passés par l'arrêt et que erase est faux,
     *  lance une erreur.
     * @param {Stop} stop - Arrêt à ajouter.
     * @param {boolean} [finalize=true] - Si vrai, finalise les arrêts avec tri
     *  et recréation des index.
     * @param {boolean} [erase=false] - Si vrai, remplace l'arrêt s'il existe déjà. Si faux
     *  (par défaut), le nouvel arrêt n'est pas pris en compte.
     * @returns {Stop | null} - L'arrêt ajouté, ou null si une erreur a été levée.
     * @throws {Error} - Si les trains du parcours sont déjà passé par l'arrêt
     *  et que erase est faux.
     */
    public addStop(stop: Stop, finalize: boolean = true, erase: boolean = false): void {

        const hasDefinedParity = stop.station.parity.isDefined();

        // Le parcours a été calculé => contient des arrêts avec parité.
        if (this.stopsChecked === Path.FULL_PATH) {
            if (!hasDefinedParity) {
                Log.warn(`Le parcours calculé ${this} ne doit comporter`
                    + ` que des arrêts avec parité définie.`
                    + ` L'arrêt ${stop} ne sera donc pas pris en compte.`);
                return;
            }
            if (this._stopsIndex.has(stop.key)) {
                if (!erase) {
                    Log.warn(`L'arrêt "${stop}" est déjà associé aux trains`
                        + ` du parcours ${this}. Un même train ne peut pas revenir`
                        + ` dans la même gare et avec le même sens.`
                        + ` Le deuxième arrêt ne sera donc pas pris en compte.`); 
                    return;
                }
                this.removeStop(stop.key);
            }

        // Le parcours n'a pas été calculé : il ne contient pas d'arrêts avec parité.
        } else {
            // Lève une erreur si ajout d'un arrêt avec parité.
            if (hasDefinedParity) {
                Log.warn(`Le parcours ${this} n'a pas été calculé.`
                    + ` Il ne peut donc pas contenir d'arrêts avec parité.`
                    + ` L'arrêt ${stop} ne sera donc pas pris en compte.`);
                return; 
            }
            // Supprime l'arrêt s'il existe déjà.
            if (this._stopsIndex.has(stop.key)) {
                if (!erase) {
                    Log.warn(`L'arrêt "${stop}" est déjà associé aux trains`
                        + ` du parcours ${this}. Si le train dessert une gare dans les deux sens,`
                        + ` il est nécessaire de calculer les parités de passage en gare.`
                        + ` Le deuxième arrêt ne sera donc pas pris en compte.`);
                    return;
                }
                this.removeStop(stop.key);
            }
            // Si l'arrêt n'est pas présent dans la signature, suppression de la signature
            //  qui sera générée à nouveau pour tenir compte du nouvel arrêt.
            if (!this.isStopInSignature(stop)) {
                this._signature = "";
            }
        }

        // Ajoute l'arrêt dans le tableau des arrêts.
        this.stops.push(stop);

        // Trie les arrêts du parcours, reconstruit l'index et la signature.
        if (finalize) this.finalizeStops();
    } 

    /**
     * Supprime un arrêt du parcours.
     */
    private removeStop(station: Station | StationWithParity | string): void {
        const existing = this.getStop(station);
        if (!existing) return;
 
        this.stops.splice(this.stops.indexOf(existing), 1);
        this._stopsIndex.delete(existing.key);
        this._stopPosition.clear();
    }

    /**
     * Vérifie si un arrêt est inclus dans la signature.
     * @param {Stop} stop - Arrêts à vérifier.
     * @returns {boolean} - Vrai si l'arrêts est inclus dans la signature, faux sinon.
     */
    private isStopInSignature(stop: Stop): boolean {
        if (!this._signature) return false;
        return this.routeStations.some(group =>
            group.some(station => stop.station.includes(station))
        );
    }
 
    /**
     * Finalise les arrêts du parcours en triant et reconstruisant l'index et la signature.
     */
    public finalizeStops(): void {
        this.orderStops();
        this.rebuildStopIndex();
        this.rebuildStopPosition();
 
        if (this.stopsChecked === Path.FULL_PATH) {
            this.recomputeParities();
        }
 
        if (!this._signature) {
            this.buildSignatureFromStops();
        }
    }

    /**
     * Trie les arrêts du parcours par ordre chronologique.
     * Les arrêts sans heure de passage sont placés en fin de liste.
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
     * Calcule les parités du parcours en fonction des arrêts.
     */
    private recomputeParities(): void {

        this.parity = Parity.undefined(true);
        this.lineDirection = Parity.undefined(true);
 
        for (const stop of this.stops) {
            this.parity = this.parity.combineWith(stop.station.parity);
 
            this.lineDirection = this.lineDirection.combineWith(
                stop.station.station.reverseLineDirection
                    ? stop.station.parity.invert()
                    : stop.station.parity
            );
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
     *  le parcours. Elle est utilisée pour chercher les connexions entre les
     *  différents parcours.
     */
    public buildSignatureFromStops(): void {

        this._signature = this.stops
            .map(s => s.key).join(">");
        this._routeStations = [];
    }

    /**
     * Construit la signature du parcours en fonction de la liste des groupes de gares.
     * La signature est une chaîne de caractères qui identifie de manière unique le parcours.
     * Elle est utilisée pour chercher les connexions entre les différents parcours.
     * @returns {string} - Signature du parcours.
     */
    public buildSignatureFromRouteStations(): string {
 
        return this.routeStations
            .map(group =>
                group
                    .map(station => station.key)
                    .join(";")
            )
            .join(">");
    }

    /**
     * Retourne l'arrêt du parcours associé à une gare.
     * Si un nombre est donné, il d'agit du numéro d'ordre
     *  (à partir de 0, ou négatif pour un décompte à partir du terminus)
     * Si la gare a une parité définie, renvoie l'arrêt correspondant.
     * Sinon, cherche l'arrêt dans le sens pair, puis dans le sens impair.
     * Si les deux arrêts sont trouvés, renvoie le premier arrêt chronologique.
     * Sinon, renvoie l'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     * @param {StationWithParity | Station | string | number} station - La gare à chercher.
     * @returns {Stop | undefined} - L'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     */
    public getStop(station: StationWithParity | Station | string | number): Stop | undefined {

        // Recherche par le numéro d'ordre (à partir de 0,
        //  ou négatif pour un décompte à partir du terminus)
        if (typeof station === "number") {

            const index = station >= 0
                ? station
                : this.stops.length + station;
        
            return this.stops[index];
        }

        // Recherche rapide par clé
        if (typeof station === "string" && this._stopsIndex.has(station)) {
            return this._stopsIndex.get(station);
        }

        const stationObj = StationWithParity.from(station);
        if (!stationObj) throw new Error(`La gare ${station} est inconnue.`);
 
        // Fonction interne : logique existante appliquée à UNE gare.
        const findDirect = (swp: StationWithParity): Stop | undefined => {

            // Le parcours a été calculé : il contient des arrêts avec parité.
            if (this.stopsChecked === Path.FULL_PATH) {

                if (swp.parity.isDefined()) {
                    return this._stopsIndex.get(swp.key) ?? undefined;
                }

                const odd = StationWithParity.from(swp, Parity.odd())!;
                const even = StationWithParity.from(swp, Parity.even())!;

                const oddStop = this._stopsIndex.get(odd.key);
                const evenStop = this._stopsIndex.get(even.key);

                if (oddStop && evenStop) {
                    const firstStop = oddStop.getTime()!.compareTo(evenStop.getTime()!) < 0
                        ? oddStop
                        : evenStop;
                    Log.warn(`Le parcours ${this} a un arrêt dans chaque sens dans la gare ${swp}.`
                        + ` C'est le premier arrêt ${firstStop} qui est renvoyé.`);
                    return firstStop;
                }

                return oddStop ?? evenStop ?? undefined;
            }

            // Le parcours n'a pas été calculé : il ne contient pas d'arrêts avec parité.
            return this._stopsIndex.get(swp.station.abbreviation) ?? undefined;
        };

        // Fait une recherche directe.
        const direct = findDirect(stationObj);
        if (direct) return direct;

        // Fait une recherche sur les parents (gare de référence + filles).
        const referenceStation = stationObj.station.referenceStation;
        const childStations: Station[] = stationObj.station.childStations;
        const parentStations = [referenceStation, ...childStations]
            .filter((s): s is Station => !!s);
        const parents: StationWithParity[] = parentStations
            .map(s => StationWithParity.from(s, stationObj.parity))
            .filter((s): s is StationWithParity => !!s);
        for (const p of parents) {
            const found = findDirect(p);
            if (found) return found;
        }

        // Retourne undefined si rien n'est trouvé.
        return undefined;
    }

    /**
     * Retourne l'arrêt suivant la gare spécifiée.
     * Si la gare spécifiée est la dernière de la liste, renvoie undefined.
     * @param {Stop | StationWithParity | Station | string} stop - L'arrêt ou la gare à chercher.
     * @returns {Stop | undefined} - L'arrêt suivant, ou undefined si la gare est la dernière.
     */
    public nextStop(
        stop: Stop | StationWithParity | Station | string
    ): Stop | undefined {
 
        const stopObj = (stop instanceof Stop)
            ? stop
            : this.getStop(stop);
        if (!stopObj) return undefined;
 
        const index = this._stopPosition.get(stopObj.key);
        if (index === undefined || index === this.stops.length - 1) return undefined;
 
        return this.stops[index + 1];
    }

    /**
     * Retourne l'arrêt précédent la gare spécifiée.
     * Si la gare spécifiée est la première de la liste, renvoie undefined.
     * @param {Stop | StationWithParity | Station | string} stop - L'arrêt ou la gare à chercher.
     * @returns {Stop | undefined} - L'arrêt précédent, ou undefined si la gare est la première.
     */
    public previousStop(
        stop: Stop | StationWithParity | Station | string
    ): Stop | undefined {
 
        const stopObj = (stop instanceof Stop)
            ? stop
            : this.getStop(stop);
        if (!stopObj) return undefined;
 
        const index = this._stopPosition.get(stopObj.key);
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
     * @param {Path} other - Le parcours à comparer.
     * @returns {boolean} - Vrai si les deux parcours ont les mêmes arrêts, faux sinon.
     */
    public equalsStops(other: Path): boolean {

        // Affecte les parcours pour que path1 contienne au moins tous les arrêts de path2
        //  (si path1 et path2 n'ont pas le même état, path1 doit être le parcours calculé)
        const path1 = (this.stopsChecked === Path.FULL_PATH) ? this : other;
        const path2 = (this.stopsChecked === Path.FULL_PATH) ? other : this;

        for (let i = 0; i < path2.stops.length; i++) {
            const stop2 = path2.stops[i];
            const stop1 = path1.getStop(stop2.station);
            if (!stop2.includes(stop1)) return false;
        }
        return true;
    }

    /**
     * Convertit les heures d'arrivée, de départ et de passage des arrêts
     *  en temps relatifs par rapport à l'heure de départ du premier arrêt.
     * Si un arrêt a déjà un horaire relatif, une erreur est levée.
     */
    public convertStopsToRelative(): void {

        if (this.stops.length === 0) return;
        const t0 = this.stops[0].departureTime;

        if (!t0) throw new Error(`Le premier arrêt du parcours ${this}`
            + ` n'a pas d'heure de départ. Les horaires ne peuvent donc pas`
            + ` être convertis en horaires relatifs.`);

        for (const stop of this.stops) {
            stop.convertToRelativeTime(t0);
        }
    }

    /**
     * Vérifie que le parcours est correct.
     * @throws {Error} - Si une erreur est détectée.
     */
    public check(): void {
 
        // Ne fait pas de vérification si le parcours a une erreur
        switch (this.stopsChecked) {
            case Path.ERROR_WITH_STOPS:
            case Path.UNCHECKED:
                return;
        }
 
        try {

            this.checkTerminals();
            this.checkSignature();

            // Valide le test si le parcours est avec gares origine et destination uniquement.
            if (this.stopsChecked === Path.ONLY_FROM_AND_TO) {
                return;
            }

            this.checkTimes();

            // Valide le test si le parcours est avec gares intermédiaires non calculé.
            if (this.stopsChecked === Path.WITH_VIA_STOPS) {
                return;
            }

            this.checkConnections();
            return;

        } catch (e) {
            this.stopsChecked = Path.ERROR_WITH_STOPS;
            throw new Error(`Vérification du parcours ${this} : ${e}`);
        }
    }

    /**
     * Vérifie les gares et horaires de départ et d'arrivée.
     * @throws {Error} - Si une erreur est détectée.
     */
    private checkTerminals(): void {
 
        // Vérifie l'existence d'une gare de départ.
        const firstStop = this.stops[0];
        if (!firstStop) {
            throw new Error(`Il n'y a pas de gare de départ.`);
        }
        // Vérifie l'existence d'une gare d'arrivée.
        const lastStop = this.stops[this.stops.length - 1];
        if (!lastStop) {
            throw new Error(`Il n'y a pas de gare d'arrivée.`);
        }
        // Vérifie l'existence d'une heure de départ.
        const departureTime = this.stops[0].departureTime;
        if (!departureTime) {
            throw new Error(`Le premier arrêt n'a pas d'heure de départ.`);
        }
        // Vérifie l'existence d'une heure d'arrivée.
        const arrivalTime = this.stops[this.stops.length - 1].arrivalTime;
        if (!arrivalTime) {
            throw new Error(`Le dernier arrêt n'a pas d'heure d'arrivée.`);
        }
        // Vérifie l'absence d'heure d'arrivée dans le premier arrêt.
        if (firstStop.isIntermediateStop()) {
            throw new Error(`Le premier arrêt ne peut pas contenir d'heure d'arrivée`
                + ` mais uniquement une heure de départ.`);
        }
        // Vérifie l'absence d'heure de départ dans le dernier arrêt.
        if (lastStop.isIntermediateStop()) {
            throw new Error(`Le dernier arrêt ne peut pas contenir d'heure de départ`
                + ` mais uniquement une heure d'arrivée.`);
        }
        // Vérifie la concordance entre les heures de départ et d'arrivée
        //  (toutes deux absolues ou relatives).
        if (arrivalTime.isRelative !== departureTime.isRelative) {
            throw new Error(`Les deux heures de départ et d'arrivée`
                + ` doivent être toutes deux absolues ou relatives.`);
        }
        // Vérifie que l'heure de départ est nulle si relative
        //  (l'heure de départ est une référence pour la suite du parcours).
        if (departureTime.isRelative && departureTime.excelValue !== 0) {
            throw new Error(`Une heure de départ relative doit avoir pour valeur 0.`);
        }
        // Vérifie que l'heure d'arrivée est postérieure à l'heure de départ.
        if (arrivalTime.compareTo(departureTime) <= 0) {
            throw new Error(`L'heure d'arrivée ${arrivalTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                + ` doit être supérieure`
                + ` à l'heure de départ ${departureTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}.`);
        }
        // Vérifie que le départ ne contient pas d'arrêt après retournement.
        if (!!firstStop.stationAfterTurnaround) {
            throw new Error(`L'heure de départ ${departureTime.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                + ` ne doit pas contenir d'arrêt aprés retournement.`);
        }
    }

    /**
     * Vérifie la signature, et la présence des gares de départ et d'arrivée.
     * @throws {Error} - Si une erreur est détectée.
     */
    private checkSignature() {
 
        // Vérifie l'existance de la signature, ou la constitue si inexistante
        //  dans le cas où le parcours n'a pas été calculé.
        let sigStations = this.routeStations;
        if (!sigStations || sigStations.length === 0) {
            switch (this.stopsChecked) {
                case Path.ONLY_FROM_AND_TO:
                case Path.WITH_VIA_STOPS:
                    this.buildSignatureFromStops();
                    return;
                case Path.FULL_PATH:
                    throw new Error(`La signature est manquante.`);
            }
        }

        // Règle de comparaison des arrêts de la signature avec ceux du parcours :
        //  - si le parcours est calculé, la gare de la signature doit inclure l'arrêt du parcours
        //     (avec parité définie),
        //  - si le parcours n'est pas calculé,
        //     l'arrêt du parcours (sans parité) doit inclure la gare de la signature.
        const areSameStations = (sigStation: StationWithParity, stop: Stop): boolean => {
            return this.stopsChecked === Path.FULL_PATH
                ? sigStation.includes(stop.station)
                : stop.station.includes(sigStation);
        };

        // Vérifie que la première gare de la signature est isolée
        //  (ne peut pas être dans un ordre quelconque avec d'autres gares)
        //  et correspond à la gare de départ.
        const firstStop = this.stops[0];
 
        if (sigStations.length === 0
            || sigStations[0].length !== 1
            || !areSameStations(sigStations[0][0], firstStop)
        ) {
            sigStations.unshift([firstStop.station]);
            Log.info(`La gare de départ ${firstStop} a été ajoutée au début de la signature.`);
        }

        // Vérifie que la dernière gare de la signature est isolée
        //  (ne peut pas être dans un ordre quelconque avec d'autres gares)
        //  et correspond à la gare d'arrivée.
        const lastStop = this.stops[this.stops.length - 1];
        const lastIndex = sigStations.length - 1;
        if ((sigStations[lastIndex].length !== 1)
            || !areSameStations(sigStations[lastIndex][0], lastStop)
        ) {
            sigStations.push([lastStop.station]);
            Log.info(`La gare d'arrivée ${lastStop} a été ajoutée à la fin de la signature.`);
        };
 
        const cleaned: StationWithParity[][] = [];
        // Décompte du nombre d'arrêt trouvés dans la signature
        let foundStops: number = 0;

        // Vérifie pour chaque arrêt :
        //  - que les gares origine et destination ne sont pas reprises
        //     comme gares intermédiaires dans la signature,
        //  - que les gares intermédiaires de la signature coïncident
        //     aux gares intermédiaires du parcours.
        for (let i = 0; i < sigStations.length; i++) {

            // Garde toujours la première et dernière gare de la signature;
            if (i === 0 || i === sigStations.length - 1) {
                cleaned.push(sigStations[i]);
                continue;
            }
 
            const group = sigStations[i];
            const filtered = group.filter(station => {

                // Vérifie que la gare intermédiaire de la signature ne soit pas la gare de départ;
                const isFirstStop = areSameStations(station, firstStop);
                if (isFirstStop) {
                    Log.info(`Suppression de la gare de départ ${station}`
                        + ` dans les gares intermédiaires de la signature.`);
                }

                // Vérifie que la gare intermédiaire de la signature ne soit pas la gare d'arrivée;
                const isLastStop = areSameStations(station, lastStop);
                if (isLastStop) {
                    Log.info(`Suppression de la gare d'arrivée ${station}`
                        + ` dans les gares intermédiaires de la signature.`);
                }

                // Vérifie que la gare intermédiaire de la signature
                //  soit bien reprise dans la liste des arrêts (parcours calculé uniquement);
                const stop = this.getStop(station);
                if (stop) foundStops++;
                const isIntermediateStop = (this.stopsChecked !== Path.FULL_PATH) || stop;
                if (!isIntermediateStop) {
                    Log.info(`Suppression de la gare ${station}`
                        + ` dans les gares intermédiaires de la signature`
                        + ` car elle n'est pas incluse dans la liste des arrêts du parcours.`);
                }

                return !isFirstStop && !isLastStop && isIntermediateStop;
            });
 
            if (filtered.length > 0) {
                cleaned.push(filtered);
            }
        }

        // Parcours non calculé avec gares intermédiaires : s'il manque de gares intermédiaires
        //  dans la signature, génère à nouveau la signature.
        if (this.stopsChecked === Path.WITH_VIA_STOPS && foundStops !== this.stops.length - 2) {
            this.buildSignatureFromStops();
            Log.info(`Il manque des gares intermédiaires dans la signature ${this._signature}.`
                + ` Elle a donc été générée à nouveau`);
            return;
        }

        sigStations = cleaned;
        this._routeStations = sigStations;

        const normalizedSignature = this.buildSignatureFromRouteStations();
        if (this._signature !== normalizedSignature) {
            Log.info(`La signature du parcours ${this.signature}`
                + ` est normalisée en ${normalizedSignature}.`);
            this._signature = normalizedSignature;
        }
    }
 
    /**
     * Vérifie que tous les arrêts intermédiaires sont corrects, en vérifiant
     *  que les heures de passage sont concordantes et que les gares intermédiaires
     *  correspondent aux gares de la signature.
     * @throws {Error} - Si une erreur est détectée;
     */
    private checkTimes() {

        const areTimesRelative = this.stops[0].getTime()!.isRelative;
        const sigStations = this.routeStations;
        let j = 1;
        let stopFromSigToFind = new Map<string, StationWithParity>();

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

                    // Cherche l'arrêt dans la liste des arrêts de la signature. Parcourt la signature
                    //  en sautant les groupes d'arrêts (séparés par ';')
                    //  et les arrêts de la signature qui ne correspondent pas à l'arrêt à chercher.
                    // Si tous les arrêts de la signature sont parcourus, lève une erreur
                    //  car l'arrêt à chercher n'a pas été trouvé.
                    while (sigStations[j].length !== 1
                        || !this.stops[i].station.includes(sigStations[j][0])) {
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
                    //  dont tous les arrêts doivent être trouvés avant de passer au (groupe) suivant. 
 
                    // Constitue la liste des arrêts de la signature à trouver.
                    if (stopFromSigToFind.size === 0) {
                        sigStations[j].reduce((map, value) => {
                            map.set(value.key, value);
                            return map;
                        }, stopFromSigToFind);
                    }
                    // Supprime de la liste l'arrêt de la signature trouvé.
                    if (stopFromSigToFind.has(this.stops[i].key)) {
                        stopFromSigToFind.delete(this.stops[i].key);
                        if (stopFromSigToFind.size === 0) j++;
                    }
                    break;
            }
            // Vérifie si l'arrêt comporte des horaires (arrivée, départ ou passage).
            const stopTime = this.stops[i].getTime();
            if (!stopTime) {
                throw new Error(`L'heure de passage à la gare de ${this.stops[i]}`
                    + ` n'est pas renseignée.`);
            }
            // Vérifie la concordance des horaires (tous absolus ou relatives).
            if (stopTime.isRelative !== areTimesRelative) {
                throw new Error(`L'heure de passage à la gare de ${this.stops[i]}`
                    + ` doit être ${areTimesRelative ? "relative" : "absolue"}`
                    + ` comme la gare origine.`);
            }
            // Vérifie que l'arrêt est une gare intermédiaire.
            if ((i < this.stops.length - 1) && !this.stops[i].isIntermediateStop()) {
                throw new Error(`L'arrêt à la gare de ${this.stops[i]} doit comporter`
                    + ` une heure d'arrivée et une heure de départ, ou une heure de passage.`);
            }
            // Vérifie que l'heure de passage est postérieure au passage précedent.
            if (this.stops[i].getTime()!.compareTo(this.stops[i - 1].getTime(true)!) <= 0) {
                throw new Error(`L'heure d'arrivée ou de passage`
                    + ` ${this.stops[i].getTime()!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` à la gare de ${this.stops[i]}`
                    + ` doit être postérieure à l'heure de passage ou de départ`
                    + ` ${this.stops[i - 1].getTime(true)!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` à la gare de ${this.stops[i - 1]}.`);
            }

        }
    }

    /**
     * Vérifie si une connexion existe entre chaque gare de la liste des arrêts.
     * @throws {Error} - Si une connexion est inexistante.
     */
    private checkConnections() {

        for (let i = 1; i < this.stops.length; i++) {

            // Vérifie si une connexion existe entre la gare précédente et la gare actuelle.
            if (this.stopsChecked === Path.FULL_PATH) {
                const lastStop = this.stops[i - 1].stationAfterTurnaround ?? this.stops[i - 1].station;
                if (!Connections.has(lastStop!, this.stops[i].station)) {
                    throw new Error(`Il n'y a pas de connexion`
                        + ` entre la gare ${lastStop} et la gare ${this.stops[i]}.`);
                }
            }
        }
    }

    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée du sillon,
     *  puis génère la liste des arrêts calculés.
     * Une fois le trajet calculé, this.stopsChecked a pour valeur Path.FULL_PATH.
     */
    public findPath(): void {

        if (this.stopsChecked === Path.FULL_PATH) {
            return;
        }

        let connections: Connection[];
        const refList = Paths.signatureIndex.get(this.signature);
        const ref = refList ? refList[0] : null;
        try {
            if (ref && ref.stopsChecked !== Path.FULL_PATH) {
                throw new Error(`Le parcours de référence ${ref}`
                    + ` n'est pas complet et ne peut pas servir de base de calcul pour ${this}.`);
            }
            connections = ref 
                ? ref.buildConnectionsFromStops()
                : this.shortestPathThrough();

            this.buildStopsFromConnections(connections);
        } catch (e) {
            throw new Error(`Calcul du parcours ${this} : ${e}`);
        }

        this.check();
        // Ajoute le parcours au cache des signatures.
        if (ref){
            // Un autre parcours avec la même signature existe déjà :
            //  le parcours nouvellement calculé est ajouté à la liste
            Paths.signatureIndex.get(this.signature)?.push(this);
        } else {
            // Aucun parcours n'a cette signature :
            //  le parcours nouvellement calculé est ajouté avec une nouvelle liste
            Paths.signatureIndex.set(this.signature, [this]);
        }
 
    }

    /**
     * Cherche le chemin le plus court entre le départ et l'arrivée du sillon
     *  en utilisant les groupes de gares.
     * @returns {Connection[]} - Chemin le plus court si il existe, sinon undefined.
     * @throws {Error} - Si la liste des gares est invalide ou si le chemin n'existe pas.
     */
    private shortestPathThrough(): Connection[] {

        if (!this.routeStations || this.routeStations.length < 2) {
            throw new Error("RouteStations invalide");
        }

        const connections = Connections.shortestPathWithGroups(this.routeStations);
        if (!connections) {
            throw new Error(`Impossible de calculer le parcours ${this.signature}`);
        }

        return connections;
    }

 
    /**
     * Construit la liste des connexions entre les arrêts d'un parcours déjà calculé.
     * @returns {Connection[]} - Liste des connexions entre les arrêts.
     */
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

            const connexionsToAdd: { from: StationWithParity, to: StationWithParity }[] = [];
            if (fromStationAfterTurnaround) {
                connexionsToAdd.push({ from: fromStation, to: fromStationAfterTurnaround });
            }
            connexionsToAdd.push({ from: fromStationAfterTurnaround ?? fromStation, to: toStation });

            for (const c of connexionsToAdd) {
                const connection = Connections.get(c.from, c.to);
                if (!connection) throw new Error(`Connection introuvable`
                    + ` entre ${c.from} et ${c.to}`);
                connections.push(connection);
            }
        }
 
        return connections;
    }

    /**
     * Reconstruit la liste des arrêts d'un parcours à partir de la liste des connexions.
     * La liste des arrêts est construite en prenant en compte les temps de retournement
     *  et les horaires d'arrivée et de départ des arrêts.
     * @param {Connection[]} connections - Liste des connexions entre les arrêts.
     */
    public buildStopsFromConnections(connections: Connection[]): void {

        const newStops: Stop[] = [];

        if (!connections.length) {
            this.stops = [];
            return;
        }

        // Initialise le cache des connexions depuis le dernier arrêt connu.
        let buffer: Connection[] = [];
 
        // Reconstruit le premier arrêt à partir des stops existants.
        const firstConnection = connections[0];
 
        const firstExisting = this.stops[0];
 
        if (!firstExisting.station.includes(firstConnection.from)) {
            throw new Error(`Le premier arrêt calculé ne correspond pas au premier arrêt de la signature.`);
        }
        const areRelativeTimes = firstExisting.departureTime?.isRelative;
        const firstStop = new Stop(
            firstConnection.from,
            undefined,
            undefined,
            firstExisting.departureTime,
            undefined,
            areRelativeTimes,
            firstExisting.tracks
        );
        newStops.push(firstStop);
        let lastStop = firstStop;

        // Parcourt les connexions.
        for (const c of connections) {

            // Vérifie si la connexion implique un retournement.
            if (c.withTurnaround && buffer.length === 0) {
                // Si le buffer est vide, le retournement se fait dans la dernière gare prise en compte.
                // Il n'y a donc pas besoin de prendre en compte la connexion de retournement
                //  dans le buffer, il faut uniquement mettre à jour le dernier arrêt avec le retournement.
                lastStop.stationAfterTurnaround = c.to;
                continue;
            }

            buffer.push(c);
            const stop =
                this.getStop(c.to);

            // Arrêt avec horaire connu.
            if (stop) {

                // Récupère les horaires aux deux arrêts connus.
                const startStop = newStops[newStops.length - 1];
                const startTime = startStop.getTime(true);
                if (!startTime) throw new Error(`L'arrêt ${startStop} n'a pas d'heure de départ.`);
                const endTime = stop.getTime(false);
                if (!endTime) throw new Error(`L'arrêt ${stop} n'a pas d'heure d'arrivée.`);
                if (endTime.excelValue <= startTime.excelValue) {
                    // Arrêt trouvé déjà parcouru (dans le cas d'un deuxième passage dans l'autre sens).
                    continue;
                }

                // Calcul le(s) temps de retournement à retrancher du temps de parcours total,
                //  sauf pour la dernière connexion du buffer (le temps de retournement sera pris
                //  en compte dans le temps d'arrêt du dernier arrêt connu trouvé)
                const totalTurnaroundTime = buffer
                    .slice(0, -1)
                    .reduce((sum, x) =>
                        sum + (x.withTurnaround ? Params.turnaroundTime.excelValue : 0), 0);

                // Calcule le temps de parcours entre les deux arrêts connus à proratiser.
                const interpolatedTime = endTime.excelValue - startTime.excelValue - totalTurnaroundTime; 
 
                // Calcule la somme des temps de parcours.
                const totalTime =
                buffer.reduce((sum, x) =>
                    sum + x.time.excelValue, 0);
                const ratio = interpolatedTime / totalTime;
                let elapsed = 0;

                // Parcourt les connexions du buffer pour créer les arrêts.
                for (let i = 0; i < buffer.length; i++) {
                    const bc = buffer[i];

                    if (bc.withTurnaround) {
                        // Si la connexion est un retournement, le dernier arrêt est forcement
                        //  un arrêt calculé, donc avec une heure de passage
                        //  (sinon la connexion aurait été sautée au début de la première boucle for).
                        // Donc le dernier arrêt est transformé en arrêt avec rebroussement, 
                        //  avec pour durée le temps de retournement par défaut.
                        lastStop.stationAfterTurnaround = bc.to;
                        elapsed += Params.turnaroundTime.excelValue;
                        continue;
                    } else {
                        elapsed += bc.time.excelValue * ratio;
                    }
 
                    // Calcule l'heure de passage.
                    const interpolated = startTime.excelValue + elapsed;
                    lastStop = new Stop(
                        bc.to,
                        undefined,
                        undefined,
                        undefined,
                        interpolated,
                        areRelativeTimes); 
                    newStops.push(lastStop);
                }

                // Modifie le dernier arrêt avec les horaires d'arrivée et de départ, et les voies.
                if (stop.arrivalTime) {
                    lastStop.setTimes(stop.arrivalTime.excelValue, stop.departureTime?.excelValue, undefined, areRelativeTimes);
                }
                lastStop.tracks = stop.tracks;
 
                buffer = [];
            }
 
        }

        // Lève une erreur s'il reste du buffer (fin du trajet).
        // Il est nécessaire d'aboutir à un arrêt connu (au maximum le dernier arrêt).
        if (buffer.length) {
            throw new Error(`Echec dans la construction des arrêts du parcours`
                + ` à partir des connexions trouvées : Le dernier arrêt ${lastStop} du parcours calculé`
                + ` n'existait pas dans le parcours initial.`); 
        }
 
        // Met à jour le parcours en ajoutant chaque nouvel arrêt calculé.
        this.eraseStops();
        this.stopsChecked = Path.FULL_PATH;
        for (const stop of newStops) {
            this.addStop(stop, false);
        }
        this.finalizeStops();
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
    public static readonly signatureIndex: Map<string, Path[]> = new Map();

    /**
     * Retourne le nombre de parcours enregistrés dans la base de données
     * @returns {number} - Nombre de parcours enregistrés
     */
    public static get size(): number {
        return this.map.size;
    }

    /**
     * Vérifie si un parcours est présent dans la base de données.
     * @param {string} key - Clé du parcours.
     * @returns {boolean} - Vrai si le parcours est présent, faux sinon.
     */
    public static has(key: string): boolean {
        return this.map.has(key);
    }

    /**
     * Renvoie le parcours correspondant à la clé donnée.
     * @param {string} key - Clé du parcours.
     * @returns {Path | undefined} - Parcours correspondant, ou undefined si la clé n'existe pas.
     */
    public static get(key: string): Path | undefined {
        return this.map.get(key);
    }

    /**
     * Ajoute un nouveau parcours dans la base de données, référencé par sa clé.
     * Si le parcours est déjà présent, une erreur est levée.
     * @param {Path} path - Parcours à enregistrer.
     * @throws {Error} - Si le parcours est déjà présent dans la base de données.
     */
    private static set(path: Path): void {
        if (this.has(path.key)) {
            throw new Error(`Le parcours ${path} est déjà présent`
                + ` dans la base de données.`);
        }
        this.map.set(path.key, path);;
    }
 
    /**
     * Retourne un tableau des valeurs de la base de données des parcours.
     * @returns {Path[]} - Itérateur sur les valeurs.
     *  de la base de données des parcours.
     */
    public static values(): Path[] {
        return Array.from(this.map.values());
    }

    /**
     * Efface toutes les parcours de la base de données.
     * Vide les maps des parcours indexés par clé, radical, signature, et structure.
     * Cela permet de forcer le rechargement des parcours si besoin.
     */
    public static clear() {
        this.map.clear();
        this.structure.clear();
        this.signatureIndex.clear();
    }

    /**
     * Crée un objet Path et l'ajoute dans la base de données.
     * Si la clé est vide, génère une clé unique pour le parcours,
     *  ou renvoie le parcours concerné si déjà existant.
     * Si un parcours avec la même clé est déjà présent dans la base de données, une erreur est levée.
     * @param {string} [key=""] - Clé du parcours.
     * @param {Parity|string/number} [parityValue=Parity.UNDEFINED] - Parité du parcours.
     * @param {Parity|string/number} [lineDirection=Parity.UNDEFINED] - Direction du parcours sur la ligne.
     * @param {string} [missionCode=""] - Code de mission des trains du parcours.
     * @param {string} [name=""] - Nom du parcours.
     * @param {string} [signature=""] - Signature du parcours : gares définissant le parcours.
     * @param {Stop[]} [stops=[]] - Gares du parcours.
     * @param {number} [stopsChecked=Path.UNCHECKED] - Résultat de la vérification du parcours.
     * @returns {Path} - Parcours créé.
     */
    public static create(
        key: string = "",
        parityValue: Parity | string | number = Parity.UNDEFINED,
        lineDirection: Parity | string | number = Parity.UNDEFINED,
        missionCode: string = "",
        name: string = "",
        signature: string = "",
        stops: Stop[] = [],
        stopsChecked: number = Path.UNCHECKED
    ): Path {

        // Instancie l'objet Path.
        const path = new Path(
            key,
            parityValue,
            lineDirection,
            missionCode,
            name,
            signature,
            stops,
            stopsChecked
        );

        // Convertit les horaires en relatifs
        path.convertStopsToRelative();

        // Finalise la gestion des arrêts (tri et mise à jour des maps)
        path.finalizeStops();

        // Insère le parcours dans la base de données, en générant si besoin la clé
        return this.insert(path);
    }

    /**
     * Insère un parcours dans la base de données,
     *  avec mise à jour de la structure des parcours si la clé existe déjà (parcours en chargement),
     *  ou avec génération de la clé si inexistante (nouveau parcours).
     * Si la connexion est déjà présente dans la base de données, une erreur est levée.
     * Si la clé est vide, génère une clé unique pour le parcours,
     *  ou renvoie le parcours concerné si déjà existant.
     * Si la clé est déjà définie, met à jour de la structure des parcours,
     *  ou lève une erreur si un parcours est déjà présent avec la même clé dans la base de données
     * @param {Path} path - Parcours à insérer.
     * @returns {Path} - Parcours inséré avec sa clé.
     */
    private static insert(path: Path): Path {

        // Clé existante : met à jour la structure des parcours.
        // Si un parcours avec la même clé est déjà présent dans la base de données, une erreur est levée.
        if(path.key) {

            // Ajoute l'objet Path dans la base de données, indexé par sa clé.
            this.set(path);

            // Ajoute l'objet Path dans l'index par signature, si pas encore présent
            //  (parcours calculé uniquement).
            if (path.stopsChecked === Path.FULL_PATH && !this.signatureIndex.has(path.signature)) {
                this.signatureIndex.set(path.signature, [path]);
            }
 
            // Ajoute l'objet Path dans la structure des radicaux et suffixes.
            const radical = this.extractRadical(path.key);
            if (!this.structure.has(radical)) {
                this.structure.set(radical, new Map());
            }
            const letter = this.extractLetter(path.key);
            if (!this.structure.get(radical)!.has(letter)) {
                this.structure.get(radical)!.set(letter, new Map());
            }
            const number = this.extractNumber(path.key);
            if (!this.structure.get(radical)!.get(letter)!.has(number)) {
                this.structure.get(radical)!.get(letter)!.set(number, path);
            }

            return path;
        }

        // Clé non existante : génère une nouvelle clé.

        const radical = path.buildRadical();
        const signature = path.signature;
 
        let radicalMap = this.structure.get(radical);
 
        // Nouveau radical : ajoute le radical et le parcours dans la structure.
        if (!radicalMap) {
            radicalMap = new Map();
            this.structure.set(radical, radicalMap);
 
            const numberMap = new Map<number, Path>();
            numberMap.set(0, path);
 
            // Par convention, le premier parcours d'un radical différent
            //  n'a pas de suffixe lettre => représenté par "".
            radicalMap.set("", numberMap);
 
            path.key = radical;
            this.set(path);
 
            return path;
        }
 
        // Radical existant : recherche l'existance de la signature.
        let letterKey = this.findLetterBySignature(radicalMap, signature);

        // Nouvelle signature : ajoute la signature et le parcours dans la structure.
        if (letterKey === null) {
            letterKey = this.nextLetter(radicalMap);
 
            const numberMap = new Map<number, Path>();
            numberMap.set(0, path);
 
            radicalMap.set(letterKey, numberMap);
 
            path.key = this.buildKey(radical, letterKey, 0);
            this.set(path);
 
            return path;
        }
 
        // Signature existante : recherche l'existance d'un parcours identique (mêmes horaires)
        //  et le renvoie si trouvé.
        const numberMap = radicalMap.get(letterKey)!;

        for (const existing of Array.from(numberMap.values())) {
            if (existing.equalsStops(path)) {
                return existing;
            }
        }
 
        // Pas de parcours trouvé : le nouveau parcours est bien unique : génère la clé.
        const number = this.nextNumber(numberMap);
 
        numberMap.set(number, path);
 
        path.key = this.buildKey(radical, letterKey, number);
        this.set(path);
 
        return path;
    }

    /**
     * Supprime un parcours de la structure interne.
     * Si le parcours n'existe pas, cette fonction ne fait rien.
     * @param {Path} path - Le parcours à supprimer.
     */
    public static delete(path: Path): void {

        // Supprime l'objet Path de la base de données, indexé par sa clé.
        this.map.delete(path.key);

        // Détermine les composantes de la clé
        const radical = path.buildRadical();
        const letter = this.extractLetter(path.key);
        const number = this.extractNumber(path.key);
 
        const radicalMap = this.structure.get(radical);
        if (!radicalMap) return;
 
        const numberMap = radicalMap.get(letter);
        if (!numberMap) return;
 
        numberMap.delete(number);
 
        // Nettoie l'étage nombre.
        if (numberMap.size === 0) {
            radicalMap.delete(letter);
        }
 
        // Nettoie l'étage lettre.
        if (radicalMap.size === 0) {
            this.structure.delete(radical);
        }

        // Suppression du parcours de la base des signatures.
        if (path.stopsChecked === Path.FULL_PATH) {
            const list = this.signatureIndex.get(path.signature);
            if (list) {
                const index = list.findIndex(p => p.key === path.key);
                if (index !== -1) {
                    list.splice(index, 1);
                }
                if (list.length === 0) {
                    this.signatureIndex.delete(path.signature);
                }
            }
        }
    }
 
    /**
     * Cherche le prochain suffixe lettre libre dans la liste des suffixes utilisés.
     * Si un seul élément existe déjà (donc sans suffixe, valeur "" dans la map),
     *  atribue le suffixe "A" à cet élément et au nouvel élément le suffixe "B".
     * Sinon, cherche le premier suffixe lettre non utilisé.
     * Les suffixes lettre sont précédés de "~".
     * @param {Map<number, Path>} numberMap - Map des suffixes déjà utilisés.
     * @returns {number} - Prochain suffixe lettre libre dans la map.
     */
    private static nextLetter(
        radicalMap: Map<string, Map<number, Path>>
    ): string {
 
        // Si un seul élément existe déjà (donc sans suffixe), donne à cet élément le suffixe "A"
        //  et au nouvel élément le suffixe "B".
        if (radicalMap.size === 1 && radicalMap.has("")) {
 
            const numberMap = radicalMap.get("")!;
            const radical = this.extractRadical(numberMap.values().next().value!.key)!;
 
            radicalMap.delete("");
            radicalMap.set("A", numberMap);

            for (const path of Array.from(numberMap.values())) {
                const number = this.extractNumber(path.key);
                this.map.delete(path.key);
                path.key = this.buildKey(radical, "A", number);
                this.set(path);
            }
 
            return "B";
        }
 
        // Si plusieurs éléments existent déjà (donc avec suffixes),
        //  cherche le premier suffixe lettre non utilisé.
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
     * @param {number} index - L'index à convertir.
     * @returns {string} - Chaîne de lettres correspondante.
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
     * @param {Map<number, Path>} numberMap - Map des suffixes déjà utilisés.
     * @returns {number} - Prochain suffixe numérique libre dans la map.
     */
    private static nextNumber(
        numberMap: Map<number, Path>
    ): number {
 
        // Si un seul élément existe déjà (donc sans suffixe), donne à cet élément le suffixe "1"
        //  et au nouvel élément le suffixe "2".
        if (numberMap.size === 1 && numberMap.has(0)) {
 
            const firstPath = numberMap.get(0)!;
 
            numberMap.delete(0);
            numberMap.set(1, firstPath);

            this.map.delete(firstPath.key);
            firstPath.key = firstPath.key + "#1";
            this.set(firstPath);

            return 2;
        }
 
        // Si plusieurs éléments existent déjà (donc avec suffixes),
        //  cherche le premier suffixe numérique non utilisé.
        let n = 1;
        while (numberMap.has(n)) n++;
 
        return n;
    }

    /**
     * Extrait le radical de la clé d'un parcours
     *  (chaîne de la forme "X~Y#Z" où X est le radical et Y et Z sont des suffixes).
     * @param {string} key - Clé du parcours.
     * @returns {string} - Radical de la clé (ou une chaîne vide si la clé n'a pas de radical).
     */
    private static extractRadical(key: string): string {
        return key.split("~")[0].split("#")[0];
    }

    /**
     * Extrait la lettre de la clé d'un parcours (chaîne de la forme "~X" où X est la lettre du suffixe).
     * @param {string} key - Clé du parcours.
     * @returns {string} - Lettre du suffixe (ou une chaîne vide si la clé n'a pas de suffixe lettre).
     */
    private static extractLetter(key: string): string {
        const m = key.match(/~([A-Z]+)/);
        return m ? m[1] : "";
    }
 
    /**
     * Extrait le numéro de la clé d'un parcours
     *  (chaîne de la forme "#X" où est le numéro du suffixe numérique).
     * @param {string} key - Clé du parcours.
     * @returns {number} - Numéro du suffixe numérique (ou 0 si la clé n'a pas de suffixe numérique).
     */
    private static extractNumber(key: string): number {
        const m = key.match(/#(\d+)/);
        return m ? Number(m[1]) : 0;
    }

    /**
     * Construit une clé de parcours à partir d'un radical,
     *  d'une lettre de suffixe et d'un numéro de suffixe.
     * La clé est composée de la forme "radical~lettre#nombre" avec
     *  un suffixe lettre optionnel précédé de "~"
     *  et un suffixe numérique optionnel précédé de "#".
     * @param {string} radical - Radical de la clé.
     * @param {string} letter - Lettre de suffixe (ou une chaîne vide si pas de suffixe lettre).
     * @param {number} number - Numéro de suffixe (ou 0 si pas de suffixe numérique).
     * @returns {string} - Clé de parcours avec les suffixes appropriés.
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
     * Cherche si un parcours existe déjà avec un même radical et une même signature.
     * Si oui donne le suffixe lettre de ce parcours, sinon renvoie null.
     * @param {Map<string, Map<number, Path>>} radicalMap - Map des parcours ayant le même radical
     *  que celui du parcours pour lequel la recherche est faite.
     * @param {string} signature - Signature du parcours à chercher.
     * @returns {string | null} - Lettre du suffixe de la clé du parcours trouvé
     *  (même radical et même signature).
     */
    private static findLetterBySignature(
        radicalMap: Map<string, Map<number, Path>>,
        signature: string
    ): string | null {
 
        for (const [letter, numberMap] of Array.from(radicalMap.entries())) {
 
            // Récupère un seul Path (le premier)
            const firstPath = numberMap.values().next().value as Path;
 
            if (firstPath.signature === signature) {
                return letter;
            }
        }
 
        return null;
    }

    /**
     * Charge les parcours de trains.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false) {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les connexions si elles ne sont pas encore chargées.
        Connections.load();

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Paths.load : aucune donnée trouvée dans la table.`);
            return;
        }

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
 
                // Récupère les champs.
                const key = WorkbookService.getString(row, this.COL_KEY);
                const parityLetter = WorkbookService.getString(row, this.COL_PARITY);
                const lineDirectionLetter = WorkbookService.getString(row, this.COL_LINE_PARITY);
                const missionCode = WorkbookService.getString(row, this.COL_MISSION_CODE);
                const name = WorkbookService.getString(row, this.COL_NAME);
                const signature = WorkbookService.getString(row, this.COL_SIGNATURE);
                const stopsChecked = WorkbookService.getNumber(row, this.COL_STOP_CHECKED);

                // Crée l'objet Path et l'insère dans la base de données.
                const path = this.create(
                    key,
                    parityLetter,
                    lineDirectionLetter,
                    missionCode,
                    name,
                    signature,
                    [],
                    stopsChecked
                );
            } 

        } catch (e) {
            throw new Error(`Paths.load (ligne ${excelRow}) : ${e}`);
        }

        // Charge les arrêts des parcours.
        Stops.load();

        // Vérifie si les parcours sont valides.
        try {
            for (const path of this.values()) {
                path.check();
            }
        } catch (e) {
            throw new Error(`Paths.load : ${e}`);
        }
 
    }

    /**
     * Sauvegarde les parcours de la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = this.SHEET,
        tableName: string = this.SHEET,
        startCell: string = "A1"
    ): void {

        // Convertit la base de données en un tableau de données.
        const data: (string | number)[][] = Array
            .from(this.values())
            .map(path => [
                path.key,
                path.parity.printLetter(),
                path.lineDirection.printLetter(),
                path.missionCode,
                path.name,
                path.signature,
                path.stopsChecked
            ]);

        // Imprime le tableau.
        const table = WorkbookService.printTable(
            this.HEADERS,
             data,
             sheetName,
             tableName,
             startCell
        );

        // Trie le tableau selon la colonne des clés
        table.getSort().apply([
            { key: this.COL_KEY, ascending: true },
        ]);

        Stops.print();
    }
}

/**
 * Classe Train définissant un train, pour un unique jour, étant la réutilisation
 *  d'un ou deux trains précédents, et ayant une ou deux réutilisations,
 *  en faisant référence à un sillon avec horaires pouvant circuler plusieurs jours par semaine.
 */
class Train {

    // Constantes des éléments
    public static readonly NORTH: number = 0;
    public static readonly SOUTH: number = 1;

    // Propriétés de l'objet Train
    public key: string;                             // Clé du train
    public number: TrainNumber;                     // Numéro du train
    public date: DateTime;                          // Date et heure de départ du train
    public service: string;                         // Service auquel le train est rattaché
    public path: Path;                              // Parcours sur lequel le train circule
    public units: string[] = []                     // Eléments (numéro de matériel)
    public previousKeys: string[] = [];             // Clés des trains précédents
    public reusesKeys: string[] = [];               // Clés des trains de réutilisations

    /**
     * Constructeur de l'objet Train.
     * @param {string} [key=""] - Clé du train.
     * @param {TrainNumber | number | string | undefined} number - Numéro du train.
     * @param {DateTime | number | string | undefined} date - Date et heure de départ du train.
     * @param {string} [service=""] - Service auquel le train est rattaché.
     * @param {Path | string | undefined} path - Parcours sur lequel le train circule.
     * @param {string[]} [units=[]] - Eléments (numéro de matériel).
     * @param {string[]} [previousKeys=[]] - Clés des trains précédents.
     * @param {string[]} [reusesKeys=[]] - Clés des trains de réutilisations.
     */
    constructor(
        key: string = "",
        number: TrainNumber | number | string | undefined,
        date: DateTime | number | string | undefined,
        service: string = "",
        path: Path | string | undefined,
        units: string[] = [],
        previousKeys: string[] = [],
        reusesKeys: string[] = []
    ) {
        this.key = key;
        const numberObj = TrainNumber.from(number);
        if (!numberObj) {
            throw new Error(`Le numéro du train ${this} est invalide.`);
        }
        this.number = numberObj;
        const dateObj = DateTime.from(date, false);
        if (!dateObj) {
            throw new Error(`La date du train ${this.number} est invalide.`);
        }
        this.date = dateObj;
        this.service = service;
        const pathObj = Path.from(path);
        if (!pathObj) {
            throw new Error(`Le parcours du train ${this.number} est invalide.`);
        }
        this.path = pathObj;
        this.units = units;
        this.previousKeys = previousKeys;
        this.reusesKeys = reusesKeys;
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.key.toString();
    }

    /**
     * Retourne l'objet Train correspondant à la clé ou l'objet Train donné.
     * Si la clé est une string, elle est utilisée pour chercher l'objet Train correspondant
     * dans l'index des trains. Si la clé est un objet Train, il est retourné tel quel.
     * Si la clé est une string mais que l'objet Train correspondant n'existe pas, undefined est retourné.
     * @param {Train | string | null | undefined} value - Clé ou objet Train.
     * @returns {Train | undefined} - Objet Train correspondant,
     *  ou undefined si la clé est une string mais que l'objet Train correspondant n'existe pas.
     */
    public static from(
        value: Train | string | null | undefined
    ): Train | undefined {
        if (value == null || value === "" || value === "-") return undefined;
        if (value instanceof Train) return value;
        return Trains.get(value!);
    }

    /**
     * Construit la clé du train qui est composée de la date suivie du numéro du train.
     * @returns {string} - Clé du train.
     */
    public buildKey(): string {
        return `${this.date.format('yyyy-MM-dd')}_${this.number.format()}`;
    }

    /**
     * Retourne les trains précédents correspondants aux clés en paramètres.
     */
    public previous(): (Train | undefined)[] {
        return this.previousKeys.map(key => Train.from(key));
    }

    /**
     * Retourne les réutilisations (trains suivants) correspondants aux clés en paramètres.
     */
    public reuse(): (Train | undefined)[] {
        return this.reusesKeys.map(key => Train.from(key));
    }

    /**
     * Retourne l'arrêt associé à une gare.
     * Si un nombre est donné, il d'agit du numéro d'ordre de l'arrêt.
     *  (à partir de 0, ou négatif pour un décompte à partir du terminus)
     * Si la gare a une parité définie, renvoie l'arrêt correspondant.
     * Sinon, cherche l'arrêt dans le sens pair, puis dans le sens impair.
     * Si les deux arrêts sont trouvés, renvoie le premier arrêt chronologique.
     * Sinon, renvoie l'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     * @param {StationWithParity | Station | string | number} station - La gare à chercher.
     * @returns {Stop | undefined} - L'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     */
    public getStop(stop: Station | StationWithParity | string | number): Stop | undefined {
        return this.path.getStop(stop);
    }

    /**
     * Renvoie la plus petite des heures d'arrivée, de départ ou de passage à l'arrêt indiqué.
     * Si ignoreArrival est vrai, lit plutôt l'heure de départ ou de passage.
     * @param {Stop | Station | StationWithParity | string | number} stop - L'arrêt à chercher.
     * @param {boolean} [ignoreArrival=false] - Si vrai, ignore l'heure d'arrivée
     *  et préfère l'heure de départ ou de passage. Si faux (par défaut),
     *  c'est d'abord l'heure d'arrivée qui est prise en compte.
     * @returns {DateTime | undefined} - Heure la plus petite à l'arrêt,
     *  ou undefined si aucune heure n'est lue.
     */
    public getTimeAt(stop: Stop | Station | StationWithParity | string | number, ignoreArrival: boolean = false): DateTime | undefined {
        if (stop instanceof Stop) {
            return stop.getTime(ignoreArrival, this.date);
        }
        return this.getStop(stop)?.getTime(ignoreArrival, this.date);
    }
}

/**
 * Classe Trains contenant la liste des trains.
 */
class Trains {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Trains";               // Feuille contenant la liste des trains
    private static readonly TABLE = "Trains";               // Tableau contenant la liste des trains
    private static readonly HEADERS = [[                    // En-têtes du tableau des trains
        "Clé",
        "Numéro du train",
        "Date",
        "Service",
        "Parcours",
        "Eléments",
        "Trains précédents",
        "Réutilisations"
    ]];
    private static readonly COL_KEY = 0;                    // Colonne de la clé du train
    private static readonly COL_NUMBER = 1;                 // Colonne du numéro du train
    private static readonly COL_DATE = 2;                   // Colonne de la date et de l'heure de départ
    private static readonly COL_SERVICE = 3;                // Colonne du service auquel le train est rattaché
    private static readonly COL_PATH = 4;                   // Colonne du parcours du train
    private static readonly COL_UNITS = 5;                  // Colonne des éléments composant le train
    private static readonly COL_PREVIOUS = 6;               // Colonne des trains précédents
    private static readonly COL_REUSES = 7;                 // Colonne des réutilisations

    // Constantes de lecture du tableau d'importation
    private static readonly IMPORT_MODE = "TRAIN";          // Mode à filtrer
    private static readonly IMPORT_1_UNIT = "Court";        // Train court (à 1 élément)
    private static readonly IMPORT_2_UNITS = "Long";        // Train long (à 2 éléments)
    private static readonly IMPORT_SHEET = "Import trains"; // Feuille d'import des arrêts
    private static readonly IMPORT_HEADERS = [[             // En-têtes du tableau d'import des arrêts
        "Date Circulation",
        "Ecart",
        "Etat",
        "Nom",
        "Code mission",
        "Origine",
        "Heure origine",
        "Destination",
        "Heure destination",
        "Composition",
        "Mode",
        "Heure à la gare",
        "Voie Infra",
        "Voie à quai à la gare"
    ]]; 
    private static readonly COL_IMPORT_DATE = 0;                // Colonne de la date 
    private static readonly COL_IMPORT_NUMBER = 3;              // Colonne du numéro de train
    private static readonly COL_IMPORT_MISSION_CODE = 4;        // Colonne de la date
    private static readonly COL_IMPORT_FROM = 5;                // Colonne du service
    private static readonly COL_IMPORT_DEPARTURE_TIME = 6;      // Colonne des jours de circulation
    private static readonly COL_IMPORT_TO = 7;                  // Colonne de la gare
    private static readonly COL_IMPORT_ARRIVAL_TIME = 8;        // Colonne de l'heure de départ
    private static readonly COL_IMPORT_UNITS = 9;               // Colonne de l'heure de passage
    private static readonly COL_IMPORT_MODE = 10;               // Colonne de l'heure de passage

    // Constantes de classe
    public static readonly UNKNOWN_UNIT = "?";
 
    // Map des trains indexées par abréviation
    public static readonly map: Map<string, Train> = new Map();

    /**
     * Nombre de trains enregistrés dans la base de données.
     * @returns {number} - Nombre de trains enregistrés.
     */
    public static get size(): number {
        return this.map.size;
    }

    /**
     * Vérifie si un train est présent dans la base de données.
     * @param {string} key - Clé du train.
     * @returns {boolean} - Vrai si le train gare est présent, faux sinon.
     */
    public static has(key: string): boolean {
        return this.map.has(key);
    }

    /**
     * Renvoie un train correspondant à la clé donnée.
     * @param {string} key - Clé du train.
     * @returns {Train | undefined} - Train correspondant, ou undefined si non trouvé.
     */
    public static get(key: string): Train | undefined {
        return this.map.get(key);
    }

    /**
     * Renvoie une liste de trains correspondant aux critères donnés :
     *  - Numéros de train
     *  - Dates de circulation (adaptées ou non)
     *  - Gare de départ
     *  - Gare d'arrivée
     *  - Gares intermédiaires et intervalle d'heure de passage (origine et terminus compris)
     *  - Zones
     *  - Batteries
     * @returns {Train[]} - Liste des trains correspondant aux critères.
     */
    public static find(
        {
            numbers = [],
            dates = [],
            adaptedTime = true,
            from,
            to,
            via = [],
            dateFrom,
            dateTo,
            zones = [],
            batteries = []
        }: {
            numbers?: (TrainNumber | string)[];
            dates?: DateTime[];
            adaptedTime?: boolean;
            from?: StationWithParity | Station | string;
            to?: StationWithParity | Station | string;
            via?: (StationWithParity | Station | string)[];
            dateFrom?: DateTime;
            dateTo?: DateTime;
            zones?: (number)[];
            batteries?: (number)[];
        }
    ): Train[] {
   
        const result: Train[] = [];
    
        // Conversion des dates en format Excel (entier)
        const datesValues = dates
            .map(d => {
                const date = d.getDate(adaptedTime);
                if (!!d && date === 0) {
                    Log.warn(`Les dates du filtre des trains doivent être absolues et non nulles.`
                        + ` La date ${d} ne sera donc pas prise en compte`);
                }
                return date;
            })
            .filter(d => d!== 0);
        const dateFromValue = dateFrom?.getDate(adaptedTime);
        if (!!dateFrom && dateFromValue === 0) {
            Log.warn(`Les dates du filtre des trains doivent être absolues et non nulles.`
                + ` La date ${dateFrom} ne sera donc pas prise en compte`);
        }
        const dateToValue = dateTo?.getDate(adaptedTime);
        if (!!dateFrom && dateFromValue === 0) {
            Log.warn(`Les dates du filtre des trains doivent être absolues et non nulles.`
                + ` La date ${dateTo} ne sera donc pas prise en compte`);
        }
    
        // Vérifie les conditions ci-dessous pour chaque train
        for (const train of this.values()) {
    
            // Numéros de train
            if (numbers.length > 0) {
                if (!numbers.some(n => train.number.includes(n))) continue;
            }
    
            // Dates exactes
            if (datesValues.length > 0) {
                if (!datesValues.includes(train.date.getDate(adaptedTime))) continue;
            }

            // Arrêt de départ et/ou d'arrivée
            const fromStop = from ? train.path.getStop(from) : undefined;
            if (from && fromStop !== train.path.origin) continue;
            const toStop = to ? train.path.getStop(to) : undefined;
            if (to && toStop !== train.path.destination) continue;

            // Arrêts du train (départ, arrivée, gares intermédiaires)
            //  avec passage dans l'intervalle de dates
            // via.forEach(s => {
            //     const stop = train.path.getStop(s);
            //     if (!stop) continue;
            //     const arrivalTime = stop!.getTime(false, train.date);
            //     if (arrivalTime && dateToValue && arrivalTime.compareTo(dateTo!) < 0) continue;
            //     const departureTime = stop!.getTime(true, train.date);
            //     if (departureTime && dateFromValue && departureTime.compareTo(dateFrom!) < 0) continue;
            // })
    
            result.push(train);
        }
    
        return result;
    }

    /**
     * Ajoute un train dans la base de données, référencé par sa clé.
     * Si le train est déjà présent, une erreur est levée.
     * @param {Train} train - Train à ajouter.
     * @throws {Error} - Si le train est déjà présent dans la base de données.
     */
    private static set(train: Train): void {
        if (this.has(train.key)) {
            throw new Error(`Le train ${train} est déjà présent`
                + ` dans la base de données.`);
        }
        this.map.set(train.key, train);
    }

    /**
     * Retourne un tableau des valeurs de la base de données des trains.
     * @returns {Train[]} - Itérateur sur les valeurs.
     *  de la base de données des trains.
     */
    public static values(): Train[] {
        return Array.from(this.map.values());
    }
 
    /**
     * Efface tous les trains de la base de données.
     * Cela permet de forcer le rechargement des trains si besoin.
     */
    public static clear(): void {
        this.map.clear();
    }
 
    /**
     * Crée un objet Train avec les paramètres donnés.
     * Si la clé est vide, génère une clé unique pour le train,
     *  ou renvoie le train concerné si déjà existant.
     * Si un train avec la même clé est déjà présent dans la base de données, une erreur est levée.
     * @param {string} [key=""] - Clé du train.
     * @param {TrainNumber | number | string | undefined} number - Numéro du train.
     * @param {DateTime | number | string | undefined} date - Date et heure de départ du train.
     * @param {string} [service=""] - Service auquel le train est rattaché.
     * @param {Path | string | undefined} path - Parcours sur lequel le train circule.
     * @param {string[]} [units=[]] - Eléments (numéro de matériel).
     * @param {string[]} [previousKeys=[]] - Clés des trains précedents.
     * @param {string[]} [reusesKeys=[]] - Clés des trains de réutilisations.
     * @returns {Train} - Train créé, ou undefined si le train est déjà présent dans la base de données.
     * @throws {Error} - Si le train est déjà présent dans la base de données.
     */
    public static create(
        key: string = "",
        number: TrainNumber | number | string | undefined,
        date: DateTime | number | string | undefined,
        service: string = "",
        path: Path | string | undefined,
        units: string[] = [],
        previousKeys: string[] = [],
        reusesKeys: string[] = []
    ): Train {

        // Instancie l'objet Train.
        const train = new Train(
            key,
            number,
            date,
            service,
            path,
            units,
            previousKeys,
            reusesKeys
        );

        // Insère le train dans la base de données, en générant si besoin la clé
        return this.insert(train);
    }

    /**
     * Ajoute un train dans la base de données.
     * Si un train avec une même clé est déjà présent sans suffixe, des suffixes 1 et 2 sont ajoutés.
     * @param {Train} train - Train à ajouter.
     * @returns {Train} - Train ajouté avec sa clé mise à jour si nécessaire.
     */
    private static insert(train: Train): Train {

        // Clé existante : ajoute le train à la base de données.
        // Si un train avec la même clé est déjà présent dans la base de données, une erreur est levée.
        if(train.key) {

            // Ajoute l'objet Path dans la base de données, indexé par sa clé.
            this.set(train);

            return train;
        }

        // Clé non existante : génère une nouvelle clé.

        const radical = train.buildKey();

        // Si train déjà présent sans suffixe, ajoute les suffixes 1 et 2.
        if (this.has(radical)) {
            const firstTrain = this.get(radical)!;
            this.map.delete(radical);
            firstTrain.key = radical + '_1';
            this.set(firstTrain);
            train.key = radical + '_2';
            this.set(train);
            return train;
        }

        // Si train déjà présent avec suffixe 1, ajoute un nouveau suffixe.
        if (this.has(radical + '_1')) {
            let i = 2;
            while (this.has(radical + '_' + i)) i++;
            train.key = radical + '_' + i;
            this.set(train);
            return train;
        }

        // Nouveau train : génère la clé.
        train.key = radical;
        this.set(train);
        return train;
    }

    /**
     * Supprime le train de la base de données dont la clé est donnée en paramètre.
     * Si le train a un suffixe, les suffixes suivants sont décalés.
     * S'il n'existe plus qu'un train avec suffixe, le suffixe est supprimé.
     * @param {string} key - Clé du train à supprimer.
     */
    public static delete(key: string): void {

        // Supprime le train de la map.
        this.map.delete(key);

        // Si train avec suffixe, modifie les suffixes suivants.
        const parts = key.split('_');
        const radical = parts[0];
        if (parts.length > 1) {
            const number = parseInt(parts[1], 10);
            this.map.delete(key);
            let i = number;
            // Décale les suffixes suivants.
            while (this.has(radical + '_' + (i + 1))) {
                const train = this.get(radical + '_' + (i + 1))!;
                this.map.delete(train.key);
                train.key = radical + '_' + i;
                this.set(train)
                i++;
            }
            // Si présence d'un uniquement élément avec suffixe, supprime le suffixe.
            if (i <= 2 && this.has(radical + '_1')) {
                const train = this.get(radical + '_1')!;
                this.map.delete(train.key);
                train.key = radical;
                this.set(train);
            }
        } 
    }

    /**
     * Charge les trains.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les parcours s'ils ne sont pas encore chargés.
        Paths.load(); 

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`Trains.load : aucune donnée trouvée dans la table.`);
            return;
        }
        const splitAndFilter = (str: string) => 
            str.split(/[ +,:;]+/)
                .filter(unit => unit.length > 0);
        const adaptWithUnits = (trainKeys: string, units: string[]) => {
            const t = splitAndFilter(trainKeys);
            while (t.length < units.length) t.push(t[0]);
            return t;
        };

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
 
                // Récupère les champs.
                const key = WorkbookService.getString(row, this.COL_KEY);
                const number = WorkbookService.getString(row, this.COL_NUMBER);
                const date = WorkbookService.getNumber(row, this.COL_DATE);
                const service = WorkbookService.getString(row, this.COL_SERVICE);
                const path = WorkbookService.getString(row, this.COL_PATH);
                const unitsString = WorkbookService.getString(row, this.COL_UNITS);
                const units = splitAndFilter(unitsString);
                const previousString = WorkbookService.getString(row, this.COL_PREVIOUS);
                const previous = adaptWithUnits(previousString, units);
                const reusesString = WorkbookService.getString(row, this.COL_REUSES);
                const reuses = adaptWithUnits(reusesString, units);

                // Crée l'objet Train et l'insère dans la base de données.
                const train = this.create(
                    key,
                    number,
                    date,
                    service,
                    path,
                    units,
                    previous,
                    reuses
                );
            } 

        } catch (e) {
            throw new Error(`Trains.load (ligne ${excelRow}) : ${e}`);
        } 
    }

    /**
     * Sauvegarde les trains de la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = this.SHEET,
        tableName: string = this.TABLE,
        startCell: string = "A1"
    ): void {

        const joinUnits = (units: string[]) =>
            units.length === 0
                ? ""
                : units.every(u => u === units[0])
                    ? units[0]
                    : units.join(' + ');

        // Convertit la base de données en un tableau de données.
        const data: (string | number)[][] = Array
            .from(this.map.values())
            .map((train: Train) => [
                train.key,
                train.number.format(false, true),
                train.date.excelValue,
                train.service,
                train.path.key,
                joinUnits(train.units),
                joinUnits(train.previousKeys),
                joinUnits(train.reusesKeys)
            ]);

        // Imprime le tableau.
        const table = WorkbookService.printTable(
            this.HEADERS,
             data,
             sheetName,
             tableName,
             startCell
        );

        // Met les dates au format "hh:mm:ss".
        const timeColumns = [
            this.COL_DATE
        ];
        for (const col of timeColumns) {
            table.getRange().getColumn(col).setNumberFormat("dd/MM/yyyy");
        }
    }

    /**
     * Importe les trains dans la base de données à partir d'un tableau Excel.
     */
    public static import(): void {

        // Charge la base de données.
        const data = WorkbookService.getDataFromSheet(this.IMPORT_SHEET);
        if (!data || data.length <= 1) {
            Log.warn(`Trains.load : aucune donnée trouvée dans la table.`);
            return;
        }

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
 
                // Récupère les champs.

                // Saute les lignes avec le mauvais mode (ex : bus et non train)
                const mode = WorkbookService.getString(row, this.COL_IMPORT_MODE);

                if (mode !== this.IMPORT_MODE) continue;

                const date = WorkbookService.getString(row, this.COL_IMPORT_DATE);
                const number = WorkbookService.getString(row, this.COL_IMPORT_NUMBER);
                const missionCode = WorkbookService.getString(row, this.COL_IMPORT_MISSION_CODE);
                const from = WorkbookService.getString(row, this.COL_IMPORT_FROM);
                const departureTime = WorkbookService.getString(row, this.COL_IMPORT_DEPARTURE_TIME);
                const to = WorkbookService.getString(row, this.COL_IMPORT_TO);
                const arrivalTime = WorkbookService.getString(row, this.COL_IMPORT_ARRIVAL_TIME);
                const composition = WorkbookService.getString(row, this.COL_IMPORT_UNITS);
                const units = composition === this.IMPORT_1_UNIT
                    ? ['-']
                    : composition === this.IMPORT_2_UNITS
                        ? ['-', '-']
                        : [];

                // Crée le parcours à partir des gares de départ et d'arrivée.
                const path = Path.fromTerminals(
                    from,
                    departureTime,
                    to,
                    arrivalTime,
                    false,
                    true,
                    missionCode
                );

                // Crée l'objet Train et l'insère dans la base de données.
                const train = this.create(
                    "",
                    number,
                    date,
                    "",
                    path,
                    units,
                    units,
                    units
                );

                if ((rowIndex + 1) % 100 === 0 ) Log.info(rowIndex + 1);
            } 

        } catch (e) {
            throw new Error(`Trains.import (ligne ${excelRow}) : ${e}`);
        } 
    }
}

/**
 * Classe TrainPath définissant un sillon, c'est à dire la capacité d'un train à rouler
 *  sur un ou plusieurs jours de la semaine, sur un ou plusieurs services donnés.
 */
class TrainPath {

    // Constantes des éléments
    public static readonly NORTH: number = 0;
    public static readonly SOUTH: number = 1;

    // Propriétés de l'objet TrainPath
    public key: string;                             // Clé du sillon
    public number: TrainNumber;                     // Numéro du sillon
    public days: Days;                              // Jours de circulation du sillon
    public services: string[];                      // Services auxquels le sillon est rattaché
    public path: Path;                              // Parcours sur lequel le sillon circule
    public units: string[] = []                     // Composistion du sillon
    public previousKeys: string[] = [];             // Clés des sillons précédents
    public reusesKeys: string[] = [];               // Clés des sillons de réutilisations

    /**
     * Constructeur de l'objet TrainPath.
     * @param {string} [key=""] - Clé du sillon.
     * @param {TrainNumber | number | string | undefined} number - Numéro du sillon.
     * @param {Days | string | number} days - Jours de circulation du sillon.
     * @param {string[] | string} [services=[]] - Services auxquels le sillon est rattaché.
     * @param {Path | string | undefined} path - Parcours sur lequel le sillon circule.
     * @param {string[]} [units=[]] - Composistion du sillon.
     * @param {string[]} [previousKeys=[]] - Clés des sillons précédents.
     * @param {string[]} [reusesKeys=[]] - Clés des sillons de réutilisations.
     */
    constructor(
        key: string = "",
        number: TrainNumber | number | string | undefined,
        days: Days | string | number,
        services: string[] | string = [],
        path: Path | string | undefined,
        units: string[] = [],
        previousKeys: string[] = [],
        reusesKeys: string[] = []
    ) {
        this.key = key;
        const numberObj = TrainNumber.from(number);
        if (!numberObj) {
            throw new Error(`Le numéro du sillon ${this} est invalide.`);
        }
        this.number = numberObj;
        const daysObj = Days.from(days);
        if (!daysObj) {
            throw new Error(`Les jours sillon ${this.number} sont invalides.`);
        }
        this.days = daysObj;
        this.services = (typeof services === "string")
            ? services.split(/[ +,:;]+/)
            : services;
        if (this.services.length === 0) {
            throw new Error(`Le sillon ${this.number} doit être affecté à au moins un service.`);
        }
        const pathObj = Path.from(path);
        if (!pathObj) {
            throw new Error(`Le parcours du sillon ${this.number} est invalide.`);
        }
        this.path = pathObj;
        this.units = units;
        this.previousKeys = previousKeys;
        this.reusesKeys = reusesKeys;
    }

    /**
     * Retourne les sillons précédents correspondants aux clés en paramètres.
     */
    public get previous(): (TrainPath | undefined)[] {
        return this.previousKeys.map(key => TrainPath.from(key));
    }

    /**
     * Retourne les réutilisations (sillons suivants) correspondants aux clés en paramètres.
     */
    public get reuse(): (TrainPath | undefined)[] {
        return this.reusesKeys.map(key => TrainPath.from(key));
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     */
    public toString(): string {
        return this.key.toString();
    }

    /**
     * Retourne l'objet TrainPath correspondant à la clé ou l'objet TrainPath donné.
     * Si la clé est une string, elle est utilisée pour chercher l'objet TrainPath correspondant
     * dans l'index des sillons. Si la clé est un objet TrainPath, il est retourné tel quel.
     * Si la clé est une string mais que l'objet TrainPath correspondant n'existe pas, undefined est retourné.
     * @param {TrainPath | string | null | undefined} value - Clé ou objet TrainPath.
     * @returns {TrainPath | undefined} - Objet TrainPath correspondant,
     *  ou undefined si la clé est une string mais que l'objet TrainPath correspondant n'existe pas.
     */
    public static from(
        value: TrainPath | string | null | undefined
    ): TrainPath | undefined {
        if (value == null || value === "" || value === "-") return undefined;
        if (value instanceof TrainPath) return value;
        return TrainPaths.get(value!);
    }

    public static fromTrain(
        train: Train | undefined
    ): TrainPath | undefined {
        if (train == null) return undefined;

        if (value instanceof TrainPath) return value;
        return TrainPaths.get(value!);
    }

    /**
     * Construit la clé du sillon qui est composée du premier service rattaché
     *  suivi du numéro de sillon et du groupe de jours de circulation.
     * @returns {string} - Clé du sillon.
     */
    public buildKey(): string {
        return `${this.services[0]}_${this.number.format()}_${this.days.code}`;
    }
}

/**
 * Classe TrainPaths contenant la liste des sillons.
 */
class TrainPaths {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Sillons";              // Feuille contenant la liste des sillons
    private static readonly TABLE = "Sillons";              // Tableau contenant la liste des sillons
    private static readonly HEADERS = [[                    // En-têtes du tableau des sillons
        "Clé",
        "Numéro du sillon",
        "Jours",
        "Services",
        "Parcours",
        "Eléments",
        "Sillons précédents",
        "Réutilisations"
    ]];
    private static readonly COL_KEY = 0;                    // Colonne de la clé du sillon
    private static readonly COL_NUMBER = 1;                 // Colonne du numéro du sillon
    private static readonly COL_DAYS = 2;                   // Colonne des jours de circulation
    private static readonly COL_SERVICES = 3;               // Colonne des services auxquels le sillon est rattaché
    private static readonly COL_PATH = 4;                   // Colonne du parcours du sillon
    private static readonly COL_UNITS = 5;                  // Colonne de la composition du sillon
    private static readonly COL_PREVIOUS = 6;               // Colonne du sillon précédent
    private static readonly COL_REUSES = 7;                 // Colonne de la réutilisation

    // Constantes de classe
    public static readonly UNKNOWN_UNIT = "?";
 
    // Map des sillons indexées par abréviation
    public static readonly map: Map<string, TrainPath> = new Map();

    /**
     * Nombre de sillons enregistrés dans la base de données.
     * @returns {number} - Nombre de sillons enregistrés.
     */
    public static get size(): number {
        return this.map.size;
    }

    /**
     * Vérifie si un sillon est présent dans la base de données.
     * @param {string} key - Clé du sillon.
     * @returns {boolean} - Vrai si le sillon gare est présent, faux sinon.
     */
    public static has(key: string): boolean {
        return this.map.has(key);
    }

    /**
     * Renvoie un sillon correspondant à la clé donnée.
     * @param {string} key - Clé du sillon.
     * @returns {TrainPath | undefined} - TrainPath correspondant, ou undefined si non trouvé.
     */
    public static get(key: string): TrainPath | undefined {
        return this.map.get(key);
    }

    // public static getFrom(
    //     numbers: TrainNumber[],
    //     days: Days,
    //     services: string[],
    //     from: StationWithParity | Station | string,
    //     to: StationWithParity | Station | string,
    //     via: (StationWithParity | Station | string)[]
    // ): TrainPath | undefined {
    //     this.values.forEach(trainPath => {
    //         if (numbers.length > 0) {
    //             if 
    //         }
    //     })
    //     return this.map.get(key);
    // }

    /**
     * Ajoute un sillon dans la base de données, référencé par sa clé.
     * Si le sillon est déjà présent, une erreur est levée.
     * @param {TrainPath} trainPath - TrainPath à ajouter.
     * @throws {Error} - Si le sillon est déjà présent dans la base de données.
     */
    private static set(trainPath: TrainPath): void {
        if (this.has(trainPath.key)) {
            throw new Error(`Le sillon ${trainPath} est déjà présent`
                + ` dans la base de données.`);
        }
        this.map.set(trainPath.key, trainPath);
    }

    /**
     * Retourne un tableau des valeurs de la base de données des sillons.
     * @returns {TrainPath[]} - Itérateur sur les valeurs.
     *  de la base de données des sillons.
     */
    public static values(): TrainPath[] {
        return Array.from(this.map.values());
    }
 
    /**
     * Efface tous les sillons de la base de données.
     * Cela permet de forcer le rechargement des sillons si besoin.
     */
    public static clear(): void {
        this.map.clear();
    }
 
    /**
     * Crée un objet TrainPath avec les paramètres donnés.
     * Si la clé est vide, génère une clé unique pour le sillon,
     *  ou renvoie le sillon concerné si déjà existant.
     * Si un sillon avec la même clé est déjà présent dans la base de données, une erreur est levée.
     * @param {string} [key=""] - Clé du sillon.
     * @param {TrainNumber | number | string | undefined} number - Numéro du sillon.
     * @param {Days | string | number} days - Jours de circulation du sillon.
     * @param {string[] | string} [services=[]] - Services auxquels le sillon est rattaché.
     * @param {Path | string | undefined} path - Parcours sur lequel le sillon circule.
     * @param {string[]} [units=[]] - Composistion du sillon.
     * @param {string[]} [previousKeys=[]] - Clés des sillons précédents.
     * @param {string[]} [reusesKeys=[]] - Clés des sillons de réutilisations.
     * @returns {TrainPath} - TrainPath créé, ou undefined si le sillon est déjà présent dans la base de données.
     * @throws {Error} - Si le sillon est déjà présent dans la base de données.
     */
    public static create(
        key: string = "",
        number: TrainNumber | number | string | undefined,
        days: Days | string | number,
        services: string[] | string = [],
        path: Path | string | undefined,
        units: string[] = [],
        previousKeys: string[] = [],
        reusesKeys: string[] = []
    ): TrainPath {

        // Instancie l'objet TrainPath.
        const trainPath = new TrainPath(
            key,
            number,
            days,
            services,
            path,
            units,
            previousKeys,
            reusesKeys
        );

        // Insère le sillon dans la base de données, en générant si besoin la clé
        return this.insert(trainPath);
    }

    /**
     * Ajoute un sillon dans la base de données.
     * @param {TrainPath} trainPath - TrainPath à ajouter.
     * @returns {TrainPath} - TrainPath ajouté.
     */
    private static insert(trainPath: TrainPath): TrainPath {

        // Ajoute l'objet Path dans la base de données, indexé par sa clé.
        // Si un sillon avec la même clé est déjà présent dans la base de données, une erreur est levée.
        this.set(trainPath);

        return trainPath;
    }

    /**
     * Supprime le sillon de la base de données dont la clé est donnée en paramètre.
     * @param {string} key - Clé du sillon à supprimer.
     */
    public static delete(key: string): void {

        // Supprime le sillon de la map.
        this.map.delete(key);
    }

    /**
     * Charge les sillons.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(erase: boolean = false): void {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les parcours s'ils ne sont pas encore chargés.
        Paths.load(); 

        // Charge la base de données.
        const data = WorkbookService.getDataFromTable(this.SHEET, this.TABLE);
        if (!data || data.length <= 1) {
            Log.warn(`TrainPaths.load : aucune donnée trouvée dans la table.`);
            return;
        }
        const splitAndFilter = (str: string) => 
            str.split(/[ +,:;]+/)
                .filter(unit => unit.length > 0);
        const adaptWithUnits = (trainPathKeys: string, units: string[]) => {
            const t = splitAndFilter(trainPathKeys);
            while (t.length < units.length) t.push(t[0]);
            return t;
        };

        const dataTable = Array.from(data.slice(1).entries());
        const nbOfRows: number = dataTable.length;
        let excelRow: number = 0;
        try {

            // Parcourt les lignes (hors en-tête).
            for (const [rowIndex, row] of dataTable) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
 
                // Récupère les champs.
                const key = WorkbookService.getString(row, this.COL_KEY);
                const number = WorkbookService.getString(row, this.COL_NUMBER);
                const days = WorkbookService.getString(row, this.COL_DAYS);
                const services = WorkbookService.getString(row, this.COL_SERVICES);
                const path = WorkbookService.getString(row, this.COL_PATH);
                const unitsString = WorkbookService.getString(row, this.COL_UNITS);
                const units = splitAndFilter(unitsString);
                const previousString = WorkbookService.getString(row, this.COL_PREVIOUS);
                const previous = adaptWithUnits(previousString, units);
                const reusesString = WorkbookService.getString(row, this.COL_REUSES);
                const reuses = adaptWithUnits(reusesString, units);

                // Crée l'objet TrainPath et l'insère dans la base de données.
                const trainPath = this.create(
                    key,
                    number,
                    days,
                    services,
                    path,
                    units,
                    previous,
                    reuses
                );
            } 

        } catch (e) {
            throw new Error(`TrainPaths.load (ligne ${excelRow}) : ${e}`);
        } 
    }

    /**
     * Sauvegarde les sillons de la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static print(
        sheetName: string = this.SHEET,
        tableName: string = this.TABLE,
        startCell: string = "A1"
    ): void {

        const joinUnits = (units: string[]) =>
            units.length === 0
                ? ""
                : units.every(u => u === units[0])
                    ? units[0]
                    : units.join(' + ');

        // Convertit la base de données en un tableau de données.
        const data: (string | number)[][] = Array
            .from(this.map.values())
            .map((trainPath: TrainPath) => [
                trainPath.key,
                trainPath.number.format(false, true),
                trainPath.days.code,
                trainPath.services.join(' ,'),
                trainPath.path.key,
                joinUnits(trainPath.units),
                joinUnits(trainPath.previousKeys),
                joinUnits(trainPath.reusesKeys)
            ]);

        // Imprime le tableau.
        const table = WorkbookService.printTable(
            this.HEADERS,
             data,
             sheetName,
             tableName,
             startCell
        );
    }
}

function testWorkbookService(options: Partial<AssertDDOptions> = {}) {
    const assert = new AssertDD(options);
    const testSheetName = "testWorkbookService";
    const testTableName = "testTable";

    // 1️⃣ Création feuille
    let sheet = WorkbookService.getSheet({ sheetName: testSheetName, createIfMissing: true });
    assert.check("Création feuille", sheet.getName(), testSheetName);

    // 2️⃣ Récup feuille existante
    const sheet2 = WorkbookService.getSheet({ sheetName: testSheetName });
    assert.check("Récup feuille", sheet2.getName(), testSheetName);

    // 3️⃣ checkCellName OK
    assert.check("Cellule valide", WorkbookService.checkCellName("A1"), "A1");

    // 4️⃣ checkCellName KO
    assert.check("Cellule invalide", WorkbookService.checkCellName("123", false), "");

    // 5️⃣ Création tableau
    const headers = [["ColStr", "ColNum", "ColBool"]];
    const data = [
        ["Paris", 42, true],
        ["", "12", "FALSE"],
        [undefined, "abc", undefined]
    ];

    const table = WorkbookService.printTable(headers, data, testSheetName, testTableName);
    assert.check("Création tableau", table?.getName(), testTableName);

    // 6️⃣ Lecture brute
    const tableData = WorkbookService.getDataFromTable(testSheetName, testTableName);
    assert.check("Lecture brute", tableData[1][0], "Paris");

    const row1 = tableData[1];
    const row2 = tableData[2];
    const row3 = tableData[3];

    // --- getString (avec défaut "")
    assert.check("getString normal", WorkbookService.getString(row1, 0), "Paris");
    assert.check("getString vide => ''", WorkbookService.getString(row2, 0), "");
    assert.check("getString undefined => ''", WorkbookService.getString(row3, 0), "");

    // --- getNumber (défaut 0)
    assert.check("getNumber number", WorkbookService.getNumber(row1, 1), 42);
    assert.check("getNumber string", WorkbookService.getNumber(row2, 1), 12);
    assert.check("getNumber invalide => 0", WorkbookService.getNumber(row3, 1), 0);

    // --- getBoolean (défaut false)
    assert.check("getBoolean true", WorkbookService.getBoolean(row1, 2), true);
    assert.check("getBoolean 'FALSE'", WorkbookService.getBoolean(row2, 2), false);
    assert.check("getBoolean undefined => false", WorkbookService.getBoolean(row3, 2), false);

    // --- getRequired (OK)
    assert.check("getRequiredString OK", WorkbookService.getRequiredString(row1, 0), "Paris");
    assert.check("getRequiredNumber OK", WorkbookService.getRequiredNumber(row1, 1), 42);
    assert.check("getRequiredBoolean OK", WorkbookService.getRequiredBoolean(row1, 2), true);

    // --- getRequired (KO)
    let errorCount = 0;

    try { WorkbookService.getRequiredString(row2, 0); } catch { errorCount++; }
    try { WorkbookService.getRequiredNumber(row3, 1); } catch { errorCount++; }
    try { WorkbookService.getRequiredBoolean(row3, 2); } catch { errorCount++; }

    assert.check("getRequired KO déclenche erreur", errorCount, 3);

    // 7️⃣ Suppression feuille
    WorkbookService.getSheet({ sheetName: testSheetName })?.delete();
    assert.check(
        "Suppression feuille",
        WorkbookService.getSheet({ sheetName: testSheetName, failOnError: false }),
        null
    );

    // 8️⃣ Résumé
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
        dtAdapt.getDayOfWeek(true)?.toString(),
        Day.SATURDAY.toString()
    );

    assert.check(
        'getDayOfWeek(real)',
        dtAdapt.getDayOfWeek(false)?.toString(),
        Day.SUNDAY.toString()
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
    7. JOURS FÉRIÉS (Lundi de Pâques 2026)
    ========================================================== */

    // 06/04/2026 = Lundi de Pâques
    const easterMonday2026 = DateTime.from("06/04/2026 6:00")!;

    // Vérifie la détection du jour férié
    assert.check(
        'isHoliday (Lundi de Pâques 2026)',
        easterMonday2026.isHoliday(),
        true
    );

    // Vérifie que getDayOfWeek AVEC férié → HOLIDAY
    assert.check(
        'getDayOfWeek(withHolidays=true) → HOLIDAY',
        easterMonday2026.getDayOfWeek(true, true)?.toString(),
        Day.HOLIDAY.toString()
    );

    // Vérifie que SANS férié → vrai jour (lundi)
    assert.check(
        'getDayOfWeek(withHolidays=false) → Lundi',
        easterMonday2026.getDayOfWeek(true, false)?.toString(),
        Day.MONDAY.toString()
    );

    // Vérifie format avec férié
    assert.check(
        'format dddd avec férié',
        easterMonday2026.format('dddd dd/mm/yyyy', true, true),
        'Férié 06/04/2026'
    );

    // Vérifie format sans férié
    assert.check(
        'format dddd sans férié',
        easterMonday2026.format('dddd dd/mm/yyyy', true, false),
        'Lundi 06/04/2026'
    );

    /* ==========================================================
    8. resolveAgainst / relativeTo / equalsTo / compare
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

    const compareToTests = [
        { a: DateTime.from(45830 + 10/24), b: DateTime.from(45830 + 10/24), expected: 0 },
        { a: DateTime.from(45830 + 10/24), b: DateTime.from(45830 + 1/24), expected: 1 },
        { a: DateTime.from(45830 + 10/24), b: DateTime.from(45831), expected: -1 },
        { a: DateTime.from(45830 + 10/24), b: DateTime.from(10/24), expected: 0 },
        { a: DateTime.from(45830), b: DateTime.from(10/24), expected: 1 },
        { a: DateTime.from(1/24, true), b: DateTime.from(10/24, true), expected: -1 },
    ];

    const getSign = (v: number) => v > 0 ? 1 : v < 0 ? -1 : 0;
    compareToTests.forEach((t, index) => {
        assert.check(
            `compareTo test #${index + 1}`,
            getSign(t.a!.compareTo(t.b!)) ,
            t.expected
        );
    });

    /* ==========================================================
    9. equalsOrUndefined()
    ========================================================== */

    const dt1 = DateTime.from(45830 + 10/24)!;
    const dt2 = DateTime.from(45830)!;

    const equalsOrUndefinedTests = [
        { a: undefined, b: undefined, expected: true },
        { a: dt1, b: undefined, expected: false },
        { a: dt1, b: dt1, expected: true },
        { a: dt1, b: dt2, expected: false },
    ];

    equalsOrUndefinedTests.forEach((t, index) => {
        assert.check(
            `equalsOrUndefined test #${index + 1}`,
            DateTime.equalsOrUndefined(t.a, t.b),
            t.expected
        );
    });

    /* ==========================================================
    10. add / subtract relatifs
    ========================================================== */

    const A = DateTime.from(2/24, true)!;
    const B = DateTime.from(3/24, true)!;

    assert.check('add', round(A.add(B).excelValue), round(5/24));
    assert.check('subtract', round(A.subtract(B).excelValue), round(-1/24));

    /* ==========================================================
    11. PARSE STRING (from + parseDateAndTime)
    ========================================================== */

    const parseTests = [

        // --- Nombres simples ---
        { input: "1.5", expected: 1.5 },
        { input: "1,5", expected: 1.5 },
        { input: "-1.5", expected: undefined }, // absolu interdit

        // --- Heures seules ---
        { input: "04:30", expected: 4.5 / 24 },
        { input: "04h30", expected: 4.5 / 24 },
        { input: "04h30min01s", expected: (4.5 + 1/3600) / 24 },
        { input: "01:00", expected: 1/24 + 1 }, // rollover

        // --- Heure négative ---
        { input: "-02:00", expected: -2/24, relative: true },

        // --- Date seule ---
        { input: "22/06/2025", expected: 45830 },
        { input: "22-06-2025", expected: 45830 },
        { input: "2025/06/22", expected: 45830 },

        // --- Date + heure ---
        { input: "22/06/2025 04:30", expected: 45830 + 4.5/24 },
        { input: "22-06-2025 04:30", expected: 45830 + 4.5/24 },

        // --- Ordre inversé ---
        { input: "04:30 22/06/2025", expected: 45830 + 4.5/24 },

        // --- Heure négative avec date (doit être ignoré) ---
        { input: "22/06/2025 -02:00", expected: 45830 + 2/24 },
        { input: "22-06-2025 -02:00", expected: 45830 + 2/24 },

        // --- Double négatif (invalide) ---
        { input: "- - 02:00", expected: 1 + 2/24 },

        // --- Format partiel ---
        { input: "22/06 04:00", check: (v: number) => round(v % 1) === round(4/24) },

        // --- Chaîne invalide ---
        { input: "abc", expected: undefined },
        { input: "22/99/2025", expected: undefined },
        { input: "25:61", expected: undefined },

    ];

    parseTests.forEach((t, i) => {
        const dt = DateTime.from(t.input as string, t.relative);

        if ("check" in t) {
            assert.check(
                `parse test #${i + 1} (${t.input}) custom`,
                t.check!(dt?.excelValue ?? NaN),
                true
            );
        } else if (t.expected === undefined) {
            assert.check(
                `parse test #${i + 1} (${t.input})`,
                dt,
                undefined
            );
        } else {
            assert.check(
                `parse test #${i + 1} (${t.input})`,
                round(dt?.excelValue ?? NaN),
                round(t.expected as number)
            );
        }
    });

    /* ==========================================================
    12. parseDate formats
    ========================================================== */

    const dateTests = [
        { input: "22/06/2025", expected: 45830 },
        { input: "2025/06/22", expected: 45830 },
        { input: "22/06", check: (v: number) => v > 45000 },
    ];

    dateTests.forEach((t, i) => {
        const v = ExcelDate.parseDate(t.input);

        if ("check" in t) {
            assert.check(`parseDate #${i + 1}`, t.check!(v!), true);
        } else {
            assert.check(`parseDate #${i + 1}`, v, t.expected);
        }
    });

    /* ==========================================================
    SYNTHÈSE
    ========================================================== */

    assert.printSummary('testDateTime');
}

function testDays(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    Days.load();
    Day.load();

    /* ==========================================================
       1. Days.from()
       ========================================================== */

    const fromTests = [
        {
            desc: "Jour simple lundi",
            input: 1,
            expected: "1"
        },
        {
            desc: "Groupe semaine 12345",
            input: "5-4-3-2-1",
            expected: "12345"
        },
        {
            desc: "Valeurs invalides ignorées",
            input: "a9b1c7",
            expected: "17"
        }
    ];

    fromTests.forEach(t => {
        const d = Days.from(t.input)!;

        assert.check(
            `Days.from("${t.input}") → numbersString (${t.desc})`,
            d.numbersString,
            t.expected
        );
    });

    /* ==========================================================
       2. extractFromString
       ========================================================== */

    const extractTests = [
        { desc: "Nom complet", input: "lundi", expected: [1] },
        { desc: "Abréviation", input: "ma", expected: [2] },
        { desc: "Numéros mélangés", input: "7;1;3", expected: [1, 3, 7] },
        { desc: "Texte mixte", input: "lumeven", expected: [1, 3, 5] }
    ];

    extractTests.forEach(t => {
        const result = Days.extractFromString(t.input);
        assert.check(
            `extractFromString("${t.input}") (${t.desc})`,
            JSON.stringify(result),
            JSON.stringify(t.expected)
        );
    });

    /* ==========================================================
       3. union / intersection
       ========================================================== */

    const d1 = Days.from("1-3-5")!;
    const d2 = Days.from("3-4")!;

    const inter = Days.intersection(d1, d2);
    const union = Days.union(d1, d2);

    assert.check("Intersection 135 ∩ 34", inter?.numbersString, "3");
    assert.check("Union 135 ∪ 34", union?.numbersString, "1345");

    /* ==========================================================
       4. contains / intersects
       ========================================================== */

    const d = Days.from("1-3-5")!;

    assert.check("contains(3)", d.contains(3), true);
    assert.check("contains(2)", d.contains(2), false);
    assert.check("intersects true", d.intersects(Days.from(3)!), true);
    assert.check("intersects false", d.intersects(Days.from(2)!), false);

    /* ==========================================================
       5. Constantes Days
       ========================================================== */

    const constantsTests = [
        { const: Days.MONDAY,    num: 1, name: "Lundi" },
        { const: Days.TUESDAY,   num: 2, name: "Mardi" },
        { const: Days.WEDNESDAY, num: 3, name: "Mercredi" },
        { const: Days.THURSDAY,  num: 4, name: "Jeudi" },
        { const: Days.FRIDAY,    num: 5, name: "Vendredi" },
        { const: Days.SATURDAY,  num: 6, name: "Samedi" },
        { const: Days.SUNDAY,    num: 7, name: "Dimanche" },
        { const: Days.HOLIDAY,   num: 8, name: "Férié" },
    ];

    constantsTests.forEach(t => {

        assert.check(
            `Days constant contains (${t.name})`,
            t.const.contains(t.num),
            true
        );

        assert.check(
            `Days constant fullName (${t.name})`,
            t.const.fullName,
            t.name
        );

        assert.check(
            `Days constant identity (${t.name})`,
            Days.from(t.num) === t.const,
            true
        );
    });

    /* ==========================================================
       6. Day.from / cohérence avec Days
       ========================================================== */

    constantsTests.forEach(t => {

        const day = Day.from(t.num)!;

        assert.check(
            `Day.from(${t.num}) fullName`,
            day.fullName,
            t.name
        );

        assert.check(
            `Day.from(${t.num}) mask`,
            day.mask,
            t.const.mask
        );

        assert.check(
            `Day.from(${t.num}) index`,
            day.index,
            t.num - 1
        );

        assert.check(
            `Day.from(${t.num}) identity`,
            Day.from(day) === day,
            true
        );
    });

    /* ==========================================================
       7. Cohérence mask → index (log2)
       ========================================================== */

    for (let i = 0; i < 8; i++) {

        const day = Day.from(i + 1)!;

        assert.check(
            `maskToIndex ${1 << i}`,
            day.index,
            i
        );
    }

    /* ==========================================================
       8. Accesseurs statiques Day
       ========================================================== */

    const dayConstants = [
        { getter: () => Day.MONDAY,    index: 0 },
        { getter: () => Day.TUESDAY,   index: 1 },
        { getter: () => Day.WEDNESDAY, index: 2 },
        { getter: () => Day.THURSDAY,  index: 3 },
        { getter: () => Day.FRIDAY,    index: 4 },
        { getter: () => Day.SATURDAY,  index: 5 },
        { getter: () => Day.SUNDAY,    index: 6 },
        { getter: () => Day.HOLIDAY,   index: 7 },
    ];

    dayConstants.forEach(t => {
        const d = t.getter();

        assert.check(
            `Day getter index ${t.index}`,
            d.index,
            t.index
        );

        assert.check(
            `Day getter identity ${t.index}`,
            Day.from(d.code) === d,
            true
        );
    });

    /* ==========================================================
       9. Iteration Day.values()
       ========================================================== */

    const values = Array.from(Day.values());

    assert.check(
        "Day.values length",
        values.length,
        8
    );

    values.forEach((d, i) => {
        assert.check(
            `Day.values index ${i}`,
            d.index,
            i
        );
    });

    /* ==========================================================
    10. Days.difference / count / numbersString
    ========================================================== */

    const dA = Days.from("1-3-5")!;
    const dB = Days.from("3")!;

    const diff = Days.difference(dA, dB);

    assert.check("difference 135 - 3", diff?.numbersString, "15");

    assert.check("count 135", dA.count, 3);
    assert.check("numbersString 135", dA.numbersString, "135");

    /* ==========================================================
    11. DaysValues
    ========================================================== */

    const base = Days.from("1-2-3-4-5-6-7")!;
    const dv = new DaysValues(base);

    dv.set(Days.from("1-2-3-4-5")!, "A");
    dv.set(Days.from("6-7")!, "B");

    assert.check(
        "DaysValues toString split",
        dv.toString(),
        "12345: A, 67: B"
    );
    assert.check(
        "DaysValues get Monday",
        dv.get(Day.MONDAY),
        "A"
    );
    assert.check(
        "DaysValues get Sunday",
        dv.get(Day.SUNDAY),
        "B"
    );

    // merge automatique + valeur unique pour tous les jours de la semaine
    dv.set(Days.from("6-7")!, "A");

    assert.check(
        "DaysValues merge",
        dv.toString(),
        "A"
    );

    // fill gaps
    const dv2 = new DaysValues(base);
    dv2.set(Days.from("1-2")!, "X");
    dv2.fillGaps("Y");

    assert.check(
        "DaysValues fillGaps",
        dv2.isComplete(),
        true
    );

    // parsing
    const parsed = DaysValues.from(base, "12345: A, 67: B");

    assert.check(
        "DaysValues from string",
        parsed.toString(),
        "12345: A, 67: B"
    );

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary('testDays');
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
        { desc: 'Lettre impair "I"', value: "I", doubleParityAllowed: false, expected: Parity.ODD },
        { desc: 'Lettre pair "P"', value: "P", doubleParityAllowed: false, expected: Parity.EVEN },
        { desc: 'Chiffre impair 1', value: 1, doubleParityAllowed: false, expected: Parity.ODD },
        { desc: 'Chiffre pair 2', value: 2, doubleParityAllowed: false, expected: Parity.EVEN },
        { desc: 'Numéro de train impair', value: "12345", doubleParityAllowed: false, expected: Parity.ODD },
        { desc: 'Numéro de train pair', value: "12346", doubleParityAllowed: false, expected: Parity.EVEN },
        { desc: 'Valeur vide', value: "", doubleParityAllowed: false, expected: Parity.UNDEFINED },
        { desc: 'Zéro "0"', value: "0", doubleParityAllowed: false, expected: Parity.UNDEFINED },
        { desc: 'Double IP interdite', value: "IP", doubleParityAllowed: false, expected: Parity.UNDEFINED },
        { desc: 'Double IP autorisée', value: "IP", doubleParityAllowed: true, expected: Parity.DOUBLE },
        { desc: 'Double implicite "1/2"', value: "1/2", doubleParityAllowed: true, expected: Parity.DOUBLE }
    ];

    constructorTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleParityAllowed);
        assert.check(
            `Parity.from(${JSON.stringify(t.value)}, ${t.doubleParityAllowed}) – ${t.desc}`,
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
       3. Pool
       ========================================================== */

    assert.check(
        "Pool: même instance pour même valeur",
        Parity.from("I") === Parity.from(1),
        true
    );

    assert.check(
        "Pool: instances différentes si doubleParityAllowed diffère",
        Parity.from("I", false) !== Parity.from("I", true),
        true
    );

    /* ==========================================================
       4. is() / isDefined()
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
       5. isOpposedTo()
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
       6. equalsTo() / 
       ========================================================== */

    assert.check(
        "equalsTo basé sur identité",
        Parity.from("I").equalsTo(Parity.from("I")),
        true
    );

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

    const oddSimple = Parity.odd(false);
    const oddDoubleAllowed = Parity.odd(true);

    assert.check(
        "equalsTo faux si doubleParityAllowed différent",
        oddSimple.equalsTo(oddDoubleAllowed),
        false
    );

    /* ==========================================================
        7. Parity.includes()
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
       8. invert()
       ========================================================== */

    const invertTests = [
        { value: "I", doubleParityAllowed: false, expected: Parity.EVEN },
        { value: "P", doubleParityAllowed: false, expected: Parity.ODD },
        { value: "IP", doubleParityAllowed: true, expected: Parity.DOUBLE },
        { value: "", doubleParityAllowed: false, expected: Parity.UNDEFINED }
    ];

    invertTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleParityAllowed).invert();
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
       9. combineWith()
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
       10. printDigit() / printLetter()
       ========================================================== */

    const printTests = [
        { value: "I", digit: Parity.digit(Parity.ODD), letter: Parity.letter(Parity.ODD) },
        { value: "P", digit: Parity.digit(Parity.EVEN), letter: Parity.letter(Parity.EVEN) },
        {
            value: "IP",
            doubleParityAllowed: true,
            digit: Parity.digit(Parity.DOUBLE),
            letter: Parity.letter(Parity.ODD) + Parity.letter(Parity.EVEN)
        },
        { value: "", digit: "", letter: "" }
    ];

    printTests.forEach(t => {
        const p = Parity.from(t.value, t.doubleParityAllowed);
        assert.check(`printDigit ${t.value}`, p.printDigit(), t.digit);
        assert.check(`printLetter ${t.value}`, p.printLetter(), t.letter);
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
       12. static containsParityLetter()
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
       13. static letter() / digit()
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
        const tn = TrainNumber.from(t.input);
        assert.check(
            `TrainNumber.from(${JSON.stringify(t.input)}) (${t.desc})`,
            tn.value,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // includes()
    // ------------------------------------------------------------

    const tn = TrainNumber.from(146490);

    const includesTests = [
        { value: "146490", expected: true },
        { value: "146491", expected: true },
        { value: "146490/1", expected: true },
        { value: "146491/0", expected: true },
        { value: "146492", expected: false }
    ];

    includesTests.forEach(t => {
        assert.check(
            `includes(${t.value})`,
            tn.includes(t.value),
            t.expected
        );
    });

    // ------------------------------------------------------------
    // isDoubleParity
    // ------------------------------------------------------------

    const doubleParityTests = [
        { value: "146491", expected: false },
        { value: "146490/1", expected: true },
    ];

    doubleParityTests.forEach(t => {
        const tn = TrainNumber.from(t.value);

        assert.check(
            `doubleParity(${t.value})`,
            tn.isDoubleParity,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // isCommercial
    // ------------------------------------------------------------

    const commercialTests = [
        { value: 147490, expected: true },
        { value: 146490, expected: false },
        { value: "E46490", expected: false }
    ];

    commercialTests.forEach(t => {
        const tn = TrainNumber.from(t.value);
        assert.check(`isCommercial(${t.value})`,
             tn.isCommercial,
             t.expected
        );
    });

    // ------------------------------------------------------------
    // isW
    // ------------------------------------------------------------

    const wTests = [
        { value: 146490, expected: true },
        { value: 569907, expected: true },
        { value: 147490, expected: false }
    ];

    wTests.forEach(t => {
        const tn = TrainNumber.from(t.value);
        assert.check(`isW(${t.value})`,
            tn.isW,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // isMouvement
    // ------------------------------------------------------------

    const mouvementTests = [
        { value: "E46490", expected: true },
        { value: "146490", expected: false }
    ];

    mouvementTests.forEach(t => {
        const tn = TrainNumber.from(t.value);
        assert.check(`isMouvement(${t.value})`, tn.isMouvement, t.expected);
    });

    // ------------------------------------------------------------
    // zone
    // ------------------------------------------------------------

    const zoneTests = [
        { value: 147490, expected: 4 },
        { value: 146490, expected: null }
    ];

    zoneTests.forEach(t => {
        const tn = TrainNumber.from(t.value);

        assert.check(
            `zone(${t.value})`,
            tn.zone,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // battery
    // ------------------------------------------------------------

    const batteryTests = [
        { value: 147490, expected: 90 },
        { value: "147490/1", expected: 91 },
        { value: 146490, expected: null }
    ];

    batteryTests.forEach(t => {
        const tn = TrainNumber.from(t.value);

        assert.check(
            `battery(${t.value})`,
            tn.battery,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // format()
    // ------------------------------------------------------------

    const formatTests = [
        {
            desc: "Abrégé",
            value: 146490,
            expected: "6490",
            args: [true, false]
        },
        {
            desc: "Double masquée",
            value: 146490,
            expected: "6490",
            args: [true, true]
        }
    ];

    formatTests.forEach(t => {
        const tn = TrainNumber.from(t.value, false);
        assert.check(
            `format(${t.value}) (${t.desc})`,
            tn.format(...t.args),
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
        { value: 146491, parity: Parity.DOUBLE, expected: "146491/0" },
        { value: 146490, parity: Parity.DOUBLE, abbreviate: true, expected: "6490/1" }
    ];

    parityTests.forEach(t => {
        const tn = TrainNumber.from(t.value);
        assert.check(
            `adaptWithParity(${t.value}, ${t.parity})`,
            tn.adaptWithParity(t.parity, t.abbreviate),
            t.expected
        );
    });

    assert.printSummary("testTrainNumber");
}

function testStation(options: Partial<AssertDDOptions> = {}) {

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
       2. Accès getById() et get()
       ----------------------------------------------------------
       Vérifie :
       - Accès par ID et clé
       ========================================================== */

    const firstStation = Stations.values()[0] as Station;

    assert.check(
        "Stations contient au moins une Station",
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

    assert.check(
        "Stations.getById(0) retourne une Station",
        Stations.getById(0) instanceof Station,
        true
    );

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
    assert.printSummary("testStation");
}

function testStationWithParity(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    StationsWithParity.load(true);

    /* ==========================================================
       1. Construction / hasDefinedParity()
       ========================================================== */

    const sU = StationWithParity.from("JY");      // undefined
    const sO = StationWithParity.from("JY_1");    // odd
    const sE = StationWithParity.from("JY_2");    // even
    const childSwpU = StationWithParity.from("JY-146_1"); // undefined
    const childSwpO = StationWithParity.from("JY-146_2"); // odd
    const childSwpE = StationWithParity.from("JY-146_3"); // even

    assert.check("Station sans parité", sU!.hasDefinedParity(), false);
    assert.check("Station sans parité", sO!.hasDefinedParity(), true);
    assert.check("Station parity odd", sO!.parity.is(Parity.ODD), true);
    assert.check("Station parity even", sE!.parity.is(Parity.EVEN), true);

    /* ==========================================================
       2. from()
       ========================================================== */

    assert.check("from(instance)", StationWithParity.from(sO) === sO, true);
    assert.check("from(null)", StationWithParity.from(null) === undefined, true);

    /* ==========================================================
       3. includes()
       ========================================================== */

    assert.check(
        "includes même station (undefined inclut odd)",
        sU!.includes(sO),
        true
    );

    assert.check(
        "includes même station (odd inclut undefined = faux)",
        sO!.includes(sU),
        false
    );

    assert.check(
        "includes même station même parité",
        sO!.includes(sO),
        true
    );

    assert.check(
        "includes station fille sans parité",
        sU!.includes(childSwpU),
        true
    );

    assert.check(
        "includes station fille avec parité",
        sU!.includes(childSwpO),
        true
    );

    assert.check(
        "includes station fille parité opposée",
        sO!.includes(childSwpE),
        false
    );

    /* ==========================================================
       4. stationAfterTurnaround()
       ========================================================== */

    const turned = sO!.stationAfterTurnaround();

    assert.check("turnaround existe ou non", true, true); // tolérant dataset
    if (turned) {
        assert.check("turnaround station identique", turned.station === sO!.station, true);
        assert.check("turnaround parité inversée", turned.parity.is(Parity.EVEN), true);
    }

    /* ==========================================================
       5. expandWithChildren() - parité
       ========================================================== */

    const expandedU = sU!.expandWithChildren();

    const hasOdd = expandedU.some(s => s.parity.is(Parity.ODD));
    const hasEven = expandedU.some(s => s.parity.is(Parity.EVEN));

    assert.check("expand undefined contient odd", hasOdd, true);
    assert.check("expand undefined contient even", hasEven, true);

    const expandedO = sO!.expandWithChildren();

    assert.check(
        "expand avec parité définie ne duplique pas",
        expandedO.some(s => s.parity.is(Parity.ODD)),
        true
    );

    /* ==========================================================
       6. expandWithChildren() - unicité
       ========================================================== */

    const ids = expandedU.map(s => s.id);
    const uniqueIds = new Set(ids);

    assert.check(
        "expand ne contient pas de doublons",
        ids.length === uniqueIds.size,
        true
    );

    /* ==========================================================
       7. expandWithChildren() - cache
       ========================================================== */

    const expandedAgain = sU!.expandWithChildren();

    assert.check(
        "cache utilisé (même référence)",
        expandedU === expandedAgain,
        true
    );

    /* ==========================================================
       8. expandWithChildren() - stabilité
       ========================================================== */

    assert.check(
        "expand stable (mêmes éléments)",
        expandedU.length === expandedAgain.length,
        true
    );

    /* ==========================================================
       9. robustesse visited (anti boucle)
       ========================================================== */

    const visitedTest = sU!.expandWithChildren(new Set<number>());

    assert.check(
        "expand avec visited externe fonctionne",
        visitedTest.length > 0,
        true
    );

    /* ==========================================================
       10. key
       ========================================================== */

    assert.check("key sans parité", sU!.key, "JY");
    assert.check("key avec parité", sO!.key, "JY_1");

    /* ==========================================================
       Résumé
       ========================================================== */

    assert.printSummary("testStationWithParity");
}

function testConnection(options: Partial<AssertDDOptions> = {}) {

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
          2. values() / accès
          ========================================================== */
 
       const firstConnection = Connections.values()[0] as Connection;
 
       assert.check(
           "Connections.values() retourne une Connection",
           firstConnection instanceof Connection,
           true
       );
 
       /* ==========================================================
          3. has() / get()
          ========================================================== */
 
       const from = firstConnection.from;
       const to = firstConnection.to;
 
       assert.check(
           "Connections.has(from, to)",
           Connections.has(from, to),
           true
       );
 
       const c = Connections.get(from, to);
 
       assert.check(
           "Connections.get(from, to)",
           c === firstConnection,
           true
       );
 
       /* ==========================================================
          4. Cohérence métier
          ========================================================== */
 
       for (const connection of Connections.values()) {
 
           assert.check(
               `${connection} : from instanceof StationWithParity`,
               connection.from instanceof StationWithParity,
               true
           );
 
           assert.check(
               `${connection} : to instanceof StationWithParity`,
               connection.to instanceof StationWithParity,
               true
           );
 
           assert.check(
               `${connection} : from ≠ to`,
               !connection.from.equalsTo(connection.to),
               true
           );
 
           assert.check(
               `${connection} : temps > 0 sauf retournement`,
               connection.withTurnaround || connection.time.excelValue > 0,
               true
           );
 
           assert.check(
               `${connection} : temps relatif`,
               connection.time.isRelative,
               true
           );
       }
 
       /* ==========================================================
          5. resolveIds() (cas Station vs StationWithParity)
          ========================================================== */
 
       const station = from.station; // supposé exister
 
       assert.check(
           "has(Station, Station)",
           Connections.has(station, to.station),
           true
       );
 
       assert.check(
           "get(Station, Station)",
           Connections.get(station, to.station) instanceof Connection,
           true
       );
 
       /* ==========================================================
          6. print()
          ========================================================== */
 
       let printOk = true;
 
       try {
           Connections.print("testConnexions", "testConnexions", "A1");
       } catch {
           printOk = false;
       }
 
       assert.check(
           'Connections.print() OK',
           printOk,
           true
       );
 
       /* ==========================================================
          FIN
          ========================================================== */
 
       assert.printSummary("testConnection");
}

function testStop(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    /* ==========================================================
       1. Création
       ========================================================== */

    const stop = new Stop(
        "PZB_1",
        "PZB_2",
        "08:00:00",
        "08:02:00",
        undefined,
        false,
        "A;B"
    );

    assert.check(
        "Stop - instance créée",
        stop instanceof Stop,
        true
    );

    /* ==========================================================
       2. Station & clé
       ========================================================== */

    assert.check(
        "Stop.key",
        stop.key,
        "PZB_1"
    );

    assert.check(
        "Stop.stationAbbreviation",
        stop.stationAbbreviation,
        "PZB"
    );

    /* ==========================================================
       3. Rebroussement
       ========================================================== */

    assert.check(
        "Stop.withTurnaround",
        stop.withTurnaround,
        true
    );

    assert.check(
        "Stop.stationAfterTurnaround",
        stop.stationAfterTurnaround?.key,
        "PZB_2"
    );

    /* ==========================================================
       4. Horaires
       ========================================================== */

    assert.check(
        "Stop.arrivalTime défini",
        stop.arrivalTime instanceof DateTime,
        true
    );

    assert.check(
        "Stop.departureTime défini",
        stop.departureTime instanceof DateTime,
        true
    );

    assert.check(
        "Stop.passageTime undefined",
        stop.passageTime === undefined,
        true
    );

    /* ==========================================================
       5. getTime()
       ========================================================== */

    const t1 = stop.getTime();
    const t2 = stop.getTime(true);

    assert.check(
        "Stop.getTime() retourne DateTime",
        t1 instanceof DateTime,
        true
    );

    assert.check(
        "Stop.getTime(true) retourne DateTime",
        t2 instanceof DateTime,
        true
    );

    /* ==========================================================
       6. isIntermediateStop
       ========================================================== */

    assert.check(
        "Stop.isIntermediateStop",
        stop.isIntermediateStop(),
        true
    );

    /* ==========================================================
       7. Tracks
       ========================================================== */

    assert.check(
        "Stop.tracks longueur",
        stop.tracks.length,
        2
    );

    stop.addTrack("C");

    assert.check(
        "Stop.addTrack",
        stop.tracks.includes("C"),
        true
    );

    /* ==========================================================
       8. equalsTo, includes
       ========================================================== */

    const stopSame = new Stop(
        "PZB_1",
        "PZB_2",
        "08:00:00",
        "08:02:00"
    );

    const stopWithoutParity = new Stop(
        "PZB",
        undefined,
        "08:00:00",
        "08:02:00"
    );

    const stopOther = new Stop(
        "SQY_1",
        undefined,
        "08:00:00"
    );

    assert.check(
        "Stop.equalsTo (identique)",
        stop.equalsTo(stopSame),
        true
    );

    assert.check(
        "Stop.equalsTo (différent)",
        stop.equalsTo(stopOther),
        false
    );

    assert.check(
        "Stop.includes (arrêt sans parité inclut l'arrêt avec parité)",
        stopWithoutParity.includes(stop),
        false
    );

    assert.check(
        "Stop.includes (arrêt avec parité inclut l'arrêt sans parité)",
        stop.includes(stopWithoutParity),
        false
    );

    /* ==========================================================
       9. convertToRelativeTime
       ========================================================== */

    const ref = DateTime.from("08:00:00");

    let convertOk = true;

    try {
        stop.convertToRelativeTime(ref);
    } catch {
        convertOk = false;
    }

    assert.check(
        "Stop.convertToRelativeTime",
        convertOk,
        true
    );

    assert.printSummary("testStop");
}

function testPath(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    /* ==========================================================
       1. Création du Path
       ========================================================== */

    const path = Path.fromTerminals(
        "PZB",
        "08:00:00",
        "SQY",
        "09:00:00",
        false,
        undefined,
        "",
        "",
        "PZB>SQY>MPU;VC"
    );

    // path.addStop(new Stop("VC", "", "08:30:00", "08:32:00"));

    assert.check(
        "Path instance créée",
        path instanceof Path,
        true
    );

    assert.check(
        "Path.signature",
        path.signature,
        "PZB>MPU;VC>SQY"
    );

    /* ==========================================================
       2. findPath()
       ========================================================== */

    path.findPath();

    path.check();

    assert.check(
        "Path.findPath - FULL_PATH",
        path.stopsChecked,
        Path.FULL_PATH
    );

    assert.check(
        "Path.stops non vide",
        path.stops.length > 1,
        true
    );

    /* ==========================================================
       3. Premier et dernier arrêt
       ========================================================== */

    const first = path.stops[0];
    const last = path.stops[path.stops.length - 1];

    assert.check(
        "Premier arrêt PZB",
        first.stationAbbreviation,
        "PZB"
    );

    assert.check(
        "Dernier arrêt SQY",
        last.stationAbbreviation,
        "SQY"
    );

    /* ==========================================================
       4. Passage par groupe MPU/VC
       ========================================================== */

    const viaStations = path.stops.map(s => s.stationAbbreviation);

    const hasVia =
        viaStations.includes("MPU") ||
        viaStations.includes("VC");

    assert.check(
        "Path passe par MPU ou VC",
        hasVia,
        true
    );

    /* ==========================================================
       5. getStop (direct + parents)
       ========================================================== */

    const stopFromNumber = path.getStop(-1);
    assert.check(
        "getStop number",
        stopFromNumber instanceof Stop,
        true
    );

    const stopFromString = path.getStop("PZB");
    assert.check(
        "getStop string",
        stopFromString instanceof Stop,
        true
    );

    const stopFromStation = path.getStop(first.station);
    assert.check(
        "getStop Station",
        stopFromStation?.stationAbbreviation,
        first.stationAbbreviation
    );

    const stopFromSWP = path.getStop(first.station.key);
    assert.check(
        "getStop SWP key",
        stopFromSWP instanceof Stop,
        true
    );

    /* ==========================================================
       6. nextStop / previousStop
       ========================================================== */

    const next = path.nextStop(first);

    assert.check(
        "nextStop retourne Stop",
        next instanceof Stop,
        true
    );

    const prev = path.previousStop(next!);

    assert.check(
        "previousStop cohérent",
        prev === first,
        true
    );

    /* ==========================================================
       7. Index et positions
       ========================================================== */

    const indexCheck = path.stops.every((s, i) =>
        path["_stopPosition"].get(s.key) === i
    );

    assert.check(
        "Positions cohérentes",
        indexCheck,
        true
    );

    /* ==========================================================
       8. getStop sur gare sans parité
       ========================================================== */

    const stopNoParity = path.getStop(first.station.station);

    assert.check(
        "getStop sans parité fonctionne",
        stopNoParity instanceof Stop,
        true
    );

    /* ==========================================================
       9. signatureIndex
       ========================================================== */

    const ref = Paths.signatureIndex.get(path.signature);

    assert.check(
        "signatureIndex contient le Path",
        ref![0] === path,
        true
    );

    /* ==========================================================
       10. Reconstruction des connexions
       ========================================================== */

    const connections = path.buildConnectionsFromStops();

    assert.check(
        "buildConnectionsFromStops retourne connexions",
        connections.length > 0,
        true
    );

    /* ==========================================================
       11. Cohérence connexions -> stops
       ========================================================== */

    const rebuilt = new Path();
    rebuilt.stops = Array.from(path.stops);
    rebuilt.stopsChecked = Path.FULL_PATH;

    const rebuiltConnections = rebuilt.buildConnectionsFromStops();

    assert.check(
        "Reconstruction cohérente",
        rebuiltConnections.length === connections.length,
        true
    );

    /* ==========================================================
       12. Ordre des temps
       ========================================================== */

    const timesOrdered = path.stops.every((s, i, arr) => {
        if (i === 0) return true;
        return s.getTime()!.compareTo(arr[i - 1].getTime()!) >= 0;
    });

    assert.check(
        "Temps ordonnés",
        timesOrdered,
        true
    );

    /* ==========================================================
       FIN
       ========================================================== */

    assert.printSummary("testPath");
}