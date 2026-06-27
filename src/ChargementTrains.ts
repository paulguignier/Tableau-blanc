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

function main(
    workbook: ExcelScript.Workbook
) {
    const startTime = Date.now();
    WORKBOOK = workbook;
    CONSOLE = console;
    const sheet = WORKBOOK.getActiveWorksheet();

    let testMode = false;
    Log.configure({
        timer: true,
        debug: true,
        info: true,
        warn: true,
        bufferize: false
    });
    Log.startTimer();


    try {

        // Lance la fonction de tests.
        // Si les tests sont actifs, la suite du programme n'est pas exécuté.
        ///////////////////// Décommenter la ligne suivante pour activer les tests ////////////////////
        testMode = true;
        if (testMode) {
            runAllTests({ testMode, startTime });
            Log.print(`Fin des tests : ${Date.now() - startTime}ms`);
            return;
        }

        Log.info(`Chargement des paramètres`);
        Params.load();
        // Paths.load();
        // Trains.load();
        // // Trains.import();

        // const selection = Trains.find({ numbers: 142202 });
        // const train = selection[0];
 
        // const path = train.path;
        // const stop = path.getStop("CPM");
        // stop?.addTrack("A");
        // train.setReuse(0, { reuseKey: 'a', position: 1 });


        // Trains.printSelection({ trains: selection, stations: "CPM_2", sheetName: "TestTrains2", tableName: "TestTrains2"});

        // Trains.save({ sheetName: "TestTrains", tableName: "TestTrains"});
        // Stations.save({ sheetName: "TestStations", tableName: "TestStations"});
        // Connections.save({ sheetName: "TestConnections", tableName: "TestConnections"});

        // // Trains.import();
        // Trains.save();
        // Paths.save();




        // Paths.load("", "147500_J;148504_J;147201_J;148202_J;147402_J;
        //      148402_J;147601_J;148602_J;145801_J;145804_J");
        // Paths.load("2", "142446_J");
        // Log.debug(Paths.map);
        // const allCombinations = Paths.generateCombinations("MPU", "ETP", "".split(";"));
        // Log.info(allCombinations);
        // const shortestPath = Paths.findShortestPath(allCombinations);
        // Log.info(shortestPath);

        Log.print(`Fin du programme : ${Date.now() - startTime}ms`);

    } catch (e) {
        Log.warn(`Erreur lors de l'exécution du programme : ${e}`);
    }
    Log.print(`-------------`);
    Log.print(`Fin d'exécution.`);
    Log.flush();
}

/**
 * Fonction de tests pour les différentes parties du code.
 * Lorsqu'elle est appelée, toutes les autres fonctions ne sont pas exécutées.
 * Les tests sont actifs si la constante TEST_MODE est vrai.
 * @param {boolean} [testMode=false] Si vrai, les fonctions de test sont lancés,
 *  puis le programme est interrompu. Si faux (par défaut), le programme continue normalement.
 * @returns {boolean} Vrai si les tests sont actifs, faux sinon.
 */
function runAllTests(
    { 
        testMode = false,
        startTime = Date.now()
    }: { 
        testMode?: boolean,
        startTime?: number
    } = {}
): boolean {

    if (!testMode) return false;

    Log.info(`Chargement des paramètres`);
    Params.load();

    Log.info(`Début des tests`);

    try {
        testUtils({ printSuccess: false, printFailure: true });
        testWorkbookServices({ printSuccess: false, printFailure: true });
        testTableSerializer({ printSuccess: false, printFailure: true });
        testDateTime({ printSuccess: false, printFailure: true });
        testDays({ printSuccess: false, printFailure: true });
        testParity({ printSuccess: false, printFailure: true });
        testTrainNumber({ printSuccess: false, printFailure: true });
        testStation({ printSuccess: false, printFailure: true, catchErrors: false });
        testStationWithParity({ printSuccess: false, printFailure: true, catchErrors: false });
        testConnection({ printSuccess: false, printFailure: true, catchErrors: false });
        testStop({ printSuccess: false, printFailure: true });
        testPath({ printSuccess: false, printFailure: true });

    } catch (e) {
        Log.warn(`Erreur lors de l'exécution des tests : ${e}`);
    }

    AssertDD.printGlobalSummary();

    Log.print(`-------------`);
    Log.print(`Fin des tests.`);
    Log.flush();
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

/**
 * Type LogOptions comprenant les options de l'affichage des logs.
 * @param {boolean} timer Affiche les timers.
 * @param {boolean} debug Affiche les messages de debug.
 * @param {boolean} info Affiche les messages d'information.
 * @param {boolean} warn Affiche les messages d'avertissement.
 * @param {boolean} bufferize Stocke les logs avant affichage console.
 */
type LogOptions = {
    timer: boolean;          
    debug: boolean;
    info: boolean;
    warn: boolean;
    bufferize: boolean;
};

/*
 * Classe Log centralisant tous les affichages console, permettant :
 *  - l'affichage des timers,
 *  - le filtrage par niveau,
 *  - le buffering,
 *  - le flush manuel,
 *  - le formatage homogène.
 */
class Log {

    // Propriétés de la classe Log
    private static startTime: number = Date.now();              // Date du début du programme
    private static lastTimer: number = Date.now();              
    private static timers: Map<string, number> = new Map();     // Dates d'initialisations par intitulés
    private static options: LogOptions = {                      // Options de l'affichage des logs
        timer: false,                                           // Remontée des timers
        debug: false,                                           // Remontée des messages de debug
        info: true,                                             // Remontée des messages d'information
        warn: true,                                             // Remontée des messages d'avertissement
        bufferize: true                                         // Stockage des logs avant affichage console
    }; 

    // Tableau de stockage des logs
    private static buffer: unknown[][] = [];

    // Initialise le timer
    public static startTimer(
        name: string = ""
    ): void {
    
        const now = Date.now();

        // Initialise les timers généraux
        if (!this.startTime) {
            this.startTime = now;
            this.lastTimer = now;
        }

        // Si un nom de timer est donné, initialise le timer associé
        if (name) {
            this.timers.set(name, now);
            return;
        }

        // Sinon, initialise le timer par défaut
        this.lastTimer = now;
    }
    
    
    /**
     * Vérifie si une valeur est concatenable (null, undefined, string, number, boolean).
     * @param {unknown} value - Valeur à vérifier.
     * @returns {boolean} - Vrai si la valeur est concatenable, faux sinon.
     */
    private static isConcatable(value: unknown): boolean {
        return (
            value === null
            || value === undefined
            || typeof value === "string"
            || typeof value === "number"
            || typeof value === "boolean"
            || typeof value === "bigint"
            || value instanceof Date
        );
    }

    /**
     * Transforme les arguments en tableau propre à afficher.
     * Les primitives sont concaténées ensemble.
     * Les objets restent séparés.
     */
    private static buildOutput(
        prefix: string | undefined,
        args: unknown[]
    ): unknown[] {

        const output: unknown[] = [];
        let buffer = prefix !== undefined
            ? `[${prefix}]`
            : "";
 
        for (const arg of args) {

            if (this.isConcatable(arg)) {
                buffer += (buffer.length > 0 ? " " : "")
                    + String(arg);
                continue;
            }

            // Flush texte avant objet
            if (buffer.trim() !== "") {
                output.push(buffer);
                buffer = "";
            }

            output.push(arg);
        }
 
        // Flush final
        if (buffer.trim() !== "") {
            output.push(buffer);
        }
 
        return output;
    } 

    /**
     * Écrit un message dans le buffer ou directement dans la console.
     */
    private static write(
        prefix: string | undefined,
        args: unknown[]
    ): void {

        const output = this.buildOutput(prefix, args);

        if (this.options.bufferize) {
            this.buffer.push(output);
            return;
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
    public static configure(
        options: Partial<LogOptions>
    ): void {
        Object.assign(this.options, options);
    }

    /**
     * Affiche tous les logs bufferisés dans l'ordre.
     */
    public static flush(): void {

        for (const output of this.buffer) {
            CONSOLE.log(...output);
        }

        this.buffer = [];
    }

    /**
     * Vide le buffer sans affichage.
     */
    public static clear(): void {
        this.buffer = [];
    }

    public static timer(
        name?: string
    ): void {
    
        if (!this.options.timer || !this.startTime) return;
        const now = Date.now();

        let start = name ? this.timers.get(name) : 0;
        if (!start) {
            start = this.lastTimer;
            this.lastTimer = now;
        }
        const elapsed = now - start;
        const total = now - this.startTime;
        this.write("TIMER",[
            `${name ? name + " : " : ""}`
                + `${elapsed} ms `
                + `(${(total / 1000).toFixed(3)} s)`
        ]);
    }

    /**
     * Message DEBUG.
     * @param {...unknown[]} args - Arguments à afficher.
     */
    public static debug(
        ...args: unknown[]
    ): void {
        if (!this.options.debug) return;
        this.write("DEBUG", args);
    }
 
    /**
     * Message INFO.
     * @param {...unknown[]} args - Arguments à afficher.
     */
    public static info(
        ...args: unknown[]
    ): void {
        if (!this.options.info) return;
        this.write("INFO", args);
    }
 
    /**
     * Message WARN.
     * @param {...unknown[]} args - Arguments à afficher.
     */
    public static warn(
        ...args: unknown[]
    ): void {
        if (!this.options.warn) return;
        this.write("WARN", args);
    }

    /**
     * Affichage brut sans préfixe.
     */
    public static print(
        ...args: unknown[]
    ): void {
        this.write(undefined, args);
    }
 
}

/**
 * Type AssertDDOptions comprenant les options de comportement et d'affichage des tests.
 * @param {boolean} printSuccess - Affiche les tests réussis.
 * @param {boolean} printFailure - Affiche les tests échoués.
 * @param {boolean} catchErrors - Intercepte les erreurs durant les tests afin de continuer
 *  l'exécution des test suivants. Désactiver cette option améliore les performances,
 *  mais toute erreur interrompra immédiatement l'exécution.
 */
type AssertDDOptions = {
    printSuccess?: boolean;
    printFailure?: boolean;
    catchErrors?: boolean;
}

/**
 * Type AssertDDOptions permettant d'inclure au choix une valeur statique ou une fonction dynamique au test.
 * @template T - Type de l'élément testé.
 * @template TResult - Type de la valeur statique ou du renvoi de la fonction dynamique.
 */
type AssertDDValue<T, TResult> = 
    TResult | ((value: T, index: number) => TResult);

/**
 * Type AssertDDOptions configurant de manière dynamique la suite de tests data-driven.
 * Chaque propriété peut :
 * - être une valeur fixe,
 * - ou être générée dynamiquement à partir de l'élément testé.
 * Les propriétés présentes directement dans un élément du tableau
 * sont prioritaires sur celles définies dans cette configuration.
 * @template T - Type de l'élément testé.
 * @param {string} label - Description du test affichée dans les logs.
 * @param {unknown | (() => unknown)} actual - Valeur réelle obtenue, ou fonction à exécuter dynamiquement.
 * @param {unknown} expected - Valeur attendue.
 * @param {string} category - Catégorie logique du test.
 * @param {string} example - Exemple ou contexte supplémentaire affiché en cas d'échec.
 * @param {boolean} skip - Ignore le test.
 * @param {boolean} only - Si un test au moins contient only,
 *  seuls les test avec only vont s'exécuter.
 * @param {number} timeout - Timeout maximal theorique du test en millisecondes.
 * @param {(value: T, index: number) => void} beforeEach - Fonction exécutée avant chaque test.
 *  ex : - reset d'état,
 *       - préparation de mocks,
 *       - création d'objets temporaires
 * @param {(value: T, index: number) => void} afterEach - Fonction exécutée après chaque test, 
 *  même en cas d'erreur.
 *  ex : - nettoyage d'objets temporaires,
 *       - suppression de données temporaires,
 *       - fermeture de ressources
 */
type AssertDDConfig<T> = {
    label?: AssertDDValue<T, string>;
    actual?: unknown | (() => unknown);
    expected?: AssertDDValue<T, unknown>;
    category?: AssertDDValue<T, string>;
    example?: AssertDDValue<T, string>;
    skip?: AssertDDValue<T, boolean>;
    only?: AssertDDValue<T, boolean>;
    timeout?: AssertDDValue<T, number>;
    beforeEach?: (value: T, index: number) => void; 
    afterEach?: (value: T, index: number) => void;  
};

/**
 * Type AssertDDCheck représentant un test unitaire concret déjà résolu.
 * @param {string} label - Description du test affichée dans les logs.
 * @param {unknown | (() => unknown)} actual - Valeur réelle obtenue, ou fonction à exécuter dynamiquement.
 * @param {unknown} expected - Valeur attendue, ou AssertDD.THROWS lorsqu'une erreur est attendue.
 * @param {string} category - Catégorie logique du test.
 * @param {string} example - Exemple ou contexte supplémentaire affiché en cas d'échec.
 * @param {boolean} skip - Ignore le test.
 * @param {boolean} only - Si un test au moins contient only, seuls les test avec only vont s'exécuter.
 * @param {number} timeout - Timeout maximaltheid du test en millisecondes.
 */
type AssertDDCheck = {
    label: string;
    actual: unknown | (() => unknown);
    expected: unknown;
    category?: string;
    example?: string;
    skip?: boolean;
    only?: boolean;
    timeout?: number;
};

/**
 * Type AssertDDEntry représentant l'entrée utilisateur d'un test data-driven.
 * Chaque entrée peut :
 *  - contenir ses propres données métier,
 *  - et éventuellement surcharger la configuration globale.
 * @template T - Type de l'élément testé.
 */
type AssertDDEntry<T = unknown> =
    T & Partial<AssertDDCheck>;

/* 
 * Classe AssertDD contenant les options et les fonctions de tests Data-Driven.
 */
class AssertDD {

    public static readonly THROWS = Symbol("ASSERT_THROWS");    // Constante indiquant 
                                                                // qu'une erreur est attendue

    public static successfulSuites = 0;     // Nombre total d'ensemble de tests complets.
    public static failedSuites = 0;         // Nombre total d'ensemble de tests incomplets.

    private total = 0;                      // Décompte du nombre total de tests réalisés
    private success = 0;                    // Décompte du nombre total de tests réalisés avec succès
    private failure = 0;                    // Décompte du nombre total de tests en échec
    private skipped = 0;                    // Décompte du nombre total de tests sautés
    private startTime: number;              // Date de commencement des tests

    private options: AssertDDOptions;       // Options d'affichage des messages de succès et d'échecs

    /**
     * Constructeur de la classe AssertDD.
     * @param {AssertDDOptions} options - Options d'affichage des messages de succès et d'échecs
     */
    public constructor(
        options: AssertDDOptions = {}
    ) {
        this.startTime = Date.now();
        this.options = {
            printSuccess: options.printSuccess ?? true,
            printFailure: options.printFailure ?? true,
            catchErrors: options.catchErrors ?? true
        };
        // Le test est considéré en échec tant qu'il n'est pas validé avec printSummary
        AssertDD.failedSuites++;
    }

    /**
     * Résout un des éléments dynamiques du test (actual, expected...):
     * - priorité à la valeur locale (propre à chaque élément du test),
     * - sinon valeur du config (commune à tous les éléments du test)
     * - sinon fallback (valeur par défaut).
     * @template T - Type de l'élément testé.
     * @template TResult - Type de la valeur de retour.
     * @param {keyof AssertDDCheck} property - Nom de l'élément.
     * @param {Record<string, unknown>} local - Valeur locale de l'élément.
     * @param {Record<string, unknown>} config - Valeur commune de l'élément.
     * @param {number} index - Index de l'élément du test.
     * @param {TResult} fallback - Valeur de retour par défaut.
     */
    private resolveProperty<T, TResult>(
        {
            property,
            local,
            config,
            index,
            fallback
        }: {
            property: keyof AssertDDCheck;
            local: Record<string, unknown>,
            config: Record<string, unknown>,
            index: number;
            fallback: TResult;
        }
    ): TResult {

        let source: unknown;

        if (property in local) {
            source = local[property];
        } else if (property in config) {
            source = config[property];
        } else {
            return fallback;
        }

        return typeof source === "function"
            ? (
                source as (
                    local: Record<string, unknown>,
                    index: number
                ) => TResult
            )(local, index)

            : source as TResult;
    }

    /**
     * Exécute une suite de tests data-driven,
     *  imprime le résultat avec un symbol de réussite (✔) ou d'échec (✘).
     * Chaque entrée peut :
     *  - contenir des données métier,
     *  - définir directement certaines propriétés du test,
     *  - ou utiliser les valeurs/fonctions du config comme fallback.
     * Priorité de résolution :
     *  1. propriété présente dans l'entrée,
     *  2. propriété présente dans config,
     *  3. fallback interne.
     * @param {AssertDDEntry<T>[]} values - Entrées de chacun des tests.
     * @param {AssertDDConfig<T>} [config={}] - Configuration générale des tests.
     * @param {AssertDDOptions} options - Options d'affichage des succès et des échecs.
    */
    public check<T>(
        values: AssertDDEntry<T>[],
        config: AssertDDConfig<T> = {},
        options?: AssertDDOptions
    ): boolean {

        // Prend en compte les options d'exécution
        const printSuccess = options?.printSuccess ?? this.options.printSuccess;
        const printFailure = options?.printFailure ?? this.options.printFailure;
        const catchErrors = options?.catchErrors ?? this.options.catchErrors;

        // Génère les tests à partir des valeurs et de la configuration
        let checks: AssertDDCheck[] = values.map((entry, index): AssertDDCheck => ({
            label: this.resolveProperty({
                property: "label",
                local: entry,
                config,
                index,
                fallback: `Test #${index + 1}`
            }),
            actual: () => this.resolveProperty({
                property: "actual",
                local: entry,
                config,
                index,
                fallback: undefined
            }),
            expected: this.resolveProperty({
                property: "expected",
                local: entry,
                config,
                index,
                fallback: true
            }),
            category: this.resolveProperty({
                property: "category",
                local: entry,
                config,
                index,
                fallback: undefined
            }),
            example: this.resolveProperty({
                property: "example",
                local: entry,
                config,
                index,
                fallback: undefined
            }),
            skip: this.resolveProperty({
                property: "skip",
                local: entry,
                config,
                index,
                fallback: false
            }),
            only: this.resolveProperty({
                property: "only",
                local: entry,
                config,
                index,
                fallback: false
            }),
            timeout: this.resolveProperty({
                property: "timeout",
                local: entry,
                config,
                index,
                fallback: undefined
            })
        }));

        // Filtre les tests si l'option "only" est activée
        const hasOnly = checks.some(check => check.only);
        if (hasOnly) {
            checks = checks.filter(check => check.only);
        }

        // Exécute les tests et affiche les résultats
        let globalSuccess = true;
        for (const [index, check] of Array.from(checks.entries())) {

            const entry = values[index];
            
            // Saute le test si l'option "skip" est activée
            if (check.skip) {
                this.skipped++;
                continue;
            }

            // Exécute le test
            const run = () => {
                const start = Date.now();
                const actualValue = typeof check.actual === "function"
                    ? (check.actual as () => unknown)()
                    : check.actual;
                const elapsed = Date.now() - start;
                const ok = Utils.equals(actualValue, check.expected);
                this.total++;

                // Si le timeout est dépassé, le test est considéré comme échoué
                if (check.timeout !== undefined && elapsed > check.timeout) {
                    this.failure++;
                    globalSuccess = false;
                    if (printFailure) {
                        Log.print(`✘ ${check.label}`
                            + ` | timeout dépassé`
                            + ` (${elapsed}ms > ${check.timeout}ms)`
                        );
                    }
                    return;
                }

                // Test réussi
                if (ok) {
                    this.success++;
                    if (printSuccess) {
                        Log.print(`✔ ${check.label}`);
                    }
                    return;
                }

                // Test échoué
                this.failure++;
                globalSuccess = false;
                if (printFailure) {

                    let message = `✘ ${check.label}`
                        + ` | attendu: ${String(check.expected)}`
                        + ` | obtenu: ${String(actualValue)}`;

                    if (check.category) {
                        message += ` | catégorie: ${check.category}`;
                    }
                
                    if (check.example) {
                        message += ` | exemple: ${check.example}`;
                    }

                    Log.print(message);
                }
            };

            // Exécute les tests sans récupération des erreurs
            if (!catchErrors) {
                config.beforeEach?.(entry, index);
                run();
                config.afterEach?.(entry, index);
                continue;
            }

            // Exécute les tests avec récupération des erreurs
            try {
                config.beforeEach?.(entry, index);
                run();
            } catch (error) {

                const ok = check.expected === AssertDD.THROWS;
                this.total++;
                if (ok) {

                    this.success++;

                    if (printSuccess) {
                        Log.print(`✔ ${check.label}`);
                    }

                } else {

                    this.failure++;
                    globalSuccess = false;

                    if (printFailure) {
                        Log.print(
                            `✘ ${check.label}`
                            + ` | erreur inattendue: ${error}`
                        );
                    }
                }
                
            } finally {
                config.afterEach?.(entry, index);
            }
        };

        return globalSuccess;
    }

    /**
     * Imprime le resultat des tests.
     * @param {string} [title="Résultats des tests"] - Titre du test.
     */
    public printSummary(
        title: string = "Résultats des tests",
        reset: boolean = true
    ): void {

        const success = this.failure === 0;
        const elapsed = Date.now() - this.startTime;

        if (success) {
            AssertDD.failedSuites--;
            AssertDD.successfulSuites++; 
        }

        Log.print(
            `${title} : `
            + `${this.success} / ${this.total} réussis\n`
            + ` (échecs : ${this.failure})`
            + ` (ignorés : ${this.skipped})`
            + ` en : ${Math.round(elapsed / 1000)} s)`
        );

        if (reset) {
            this.reset();
        }
    }

    /**
     * Imprime le résultat global de toutes les suites exécutées.
     */
    public static printGlobalSummary(): void {

        const total =
            AssertDD.successfulSuites
            + AssertDD.failedSuites;

        Log.print(
            `Suites globales : `
            + `${AssertDD.successfulSuites} réussies / ${total}\n`
            + ` (échecs : ${AssertDD.failedSuites})`
        );
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

/**
 * Type PrimitiveValue reprenant les différents types de valeurs
 *  contenues dans une cellule de feuille de calcul Excel :
 *  - chaînes de caractères,
 *  - nombres,
 *  - booléens.
 */
type PrimitiveValue = string | number | boolean;

/**
 * Type PrimitiveType reprenant les types primitifs de valeurs Excel, sous forme de chaîne de caractères
 *  pour les utiliser comme paramètres de fonctions.
 */
type PrimitiveType = "string" | "number" | "boolean";

/**
 * Type ConvertedValue reprenant le type de la valeur Excel renvoyée selon le type demandé.
 * @template T - Type de la valeur Excel renvoyée, inclus dans les types primitifs.
 */
type ConvertedValue<T extends PrimitiveType | undefined> =
    T extends "string" ? string :
    T extends "number" ? number :
    T extends "boolean" ? boolean :
    PrimitiveValue;

/**
 * Type OneOrMany reprenant un ou plusieurs types appelés sous forme simple ou en tableau.
 * @template T - Type reçu sous forme simple ou sous forme de tableau.
 */
type OneOrMany<T> = T | T[];

/**
 * Type Nullable reprenant un type ou null ou undefined.
 * @template T - Type auquel est ajouté null ou undefined.
 */
type Nullable<T> = T | null | undefined;

/**
 * Type Input réunissant les types de données acceptés
 *  pour créer ou appeler un objet de classe T, y compris lui-même.
 * @template T - Types acceptés par la classe.
 */
type Input<T, TRaw> = T | TRaw;

/**
 * Type NullableOneOrMany reprenant un type ou un tableau de types ou null ou undefined.
 * @template T - Type auquel est ajouté null ou undefined, puis sous forme simple ou sous forme de tableau.
 */
type NullableOneOrMany<T> = Nullable<OneOrMany<T>>;

/**
 * Classe utilitaire Utils contenant des fonctions utilitaires.
 */
class Utils {

    /**
     * Compare deux valeurs, objets ou tableaux en analysant chacune de leurs valeurs individuelles.
     * @param {unknown} a - Première valeur à comparer. 
     * @param {unknown} b - Deuxième valeur à comparer.
     * @returns {boolean} Vrai si les deux valeurs sont identiques, faux sinon.
     */
    public static equals(
        a: unknown,
        b: unknown
    ): boolean {
    
        // Cas simples
        if (a === b) return true;
    
        // Un seul est null/undefined
        if (a == null || b == null) return false;
    
        // Tableaux
        if (Array.isArray(a) && Array.isArray(b)) {
    
            if (a.length !== b.length) return false;
    
            for (let i = 0; i < a.length; i++) {
                if (!this.equals(a[i], b[i])) {
                    return false;
                }
            }
    
            return true;
        }
    
        // Objets
        if (
            typeof a === "object"
            && typeof b === "object"
        ) {
    
            const keysA = Object.keys(a);
            const keysB = Object.keys(b);
    
            if (keysA.length !== keysB.length) {
                return false;
            }
    
            for (const key of keysA) {
                if (!this.equals(
                    (a as Record<string, unknown>)[key],
                    (b as Record<string, unknown>)[key]
                )) {
                    return false;
                }
            }
    
            return true;
        }
    
        return false;
    }

    /**
     * Indique si un objet possède une méthode.
     * @template TMethod - Nom de la méthode.
     * @param {unknown} value - Objet dont la méthode est recherchée.
     * @param {TMethod} method - Nom de la méthode.
     * @returns {boolean} - Vrai si l'objet possède la méthode, faux sinon.
     */
    public static hasMethod<TMethod extends string>(
        value: unknown,
        method: TMethod
    ): value is Record<TMethod, (...args: never[]) => unknown> {

        return (typeof value === "object"
            && value !== null
            && method in value
            && typeof (value as Record<string, unknown>)[method] === "function");
    }

    /**
     * Convertit une valeur Excel en un type primitif demandé.
     * Pour les objets, leurs méthodes toString, toNumber et toBoolean sont appelées si existantes.
     * Si le type n'est pas défini, les valeurs number et boolean sont renvoyées telles quelles,
     *  les autres sont converties en string.
     * @template {PrimitiveType | undefined} T - Type de la valeur Excel renvoyée, inclus dans les types primitifs.
     * @param {unknown} value - Valeur Excel.
     * @param {T} type - Type de la valeur Excel renvoyée.
     * @returns {ConvertedValue<T> | undefined} - Valeur primitive renvoyée.
     */
    public static convertValue<T extends PrimitiveType | undefined>(
        value: unknown,
        type?: T
    ): ConvertedValue<T> | undefined {
    
        if (value == null) return undefined;
    
        switch (type) {

            case undefined:

                // Conserve les nombres et booléens.
                if (typeof value === "number"
                    || typeof value === "boolean"
                ) {
                    return value as ConvertedValue<T>;
                }

                // Les chaînes (et objets) suivent le traitement "string".

            case "string":
    
                if (typeof value === "string") {
                    return (value.trim() || undefined) as ConvertedValue<T>;
                }
                return String(value) as ConvertedValue<T>;
    
            case "number":

                const filterFinite = (n: number) => Number.isFinite(n)
                    ? n as ConvertedValue<T>
                    : undefined;

                if (value === "" ||typeof value === "boolean") {
                    return undefined;
                }
                if (typeof value === "number") {
                    return filterFinite(value);
                }
                if (typeof value === "string") {
                    return filterFinite(Number(value.replace(",", ".")));
                }
                if (this.hasMethod(value, "toNumber")) {
                    return filterFinite(value.toNumber() as number);
                }
                return undefined;
    
            case "boolean":
    
                if (typeof value === "boolean") {
                    return value as ConvertedValue<T>;
                }
                if (typeof value === "number") {
                    return (value !== 0) as ConvertedValue<T>;
                }
                if (typeof value === "string") {
                    const normalized = value.trim().toLowerCase();
                    if (normalized === "") return undefined;
                    return ["true", "1", "oui", "yes"].includes(normalized) as ConvertedValue<T>;
                }
                if (this.hasMethod(value, "toBoolean")) {
                    return value.toBoolean() as ConvertedValue<T>;
                }
                return undefined;
    
            default:
                throw new Error(
                    `Type non pris en charge : ${type}`
                );
        }
    }

    /**
     * Convertit une valeur en un tableau,
     *  ou renvoie cette valeur s'il s'agit déjà d'un tableau.
     * @param {NullableOneOrMany<T>} value - Valeur à convertir en tableau.
     * @returns 
     */
    public static asArray<T>(
        value: Nullable<OneOrMany<T>>,
        {
            split,
            trim = false,
            filterNull = true,
            filterEmptyString = true
        }: {
            split?: string | RegExp,
            trim?: boolean,
            filterNull?: boolean,
            filterEmptyString?: boolean
        } = {}
    ): T[] {
    
        if (value == null) return [];
        let result: unknown[];
    
        // Sépare les éléments contenus dans les chaines,
        //  et assemble toutes les valeurs dans un seul tableau.
        if (split !== undefined && typeof value === "string") {
            result = value.split(split);
        } else {
            result = Array.isArray(value)
                ? [...value]
                : [value];
        }
    
        // Pour les chaines, supprime les espaces en tête et en queue.
        if (trim) {
            result = result.map(v =>
                typeof v === "string"
                    ? v.trim()
                    : v
            );
        }
    
        // Filtre les valeurs nulles.
        if (filterNull) {
            result = result.filter(v => v != null);
        }
    
        // Filtre les chaines vides.
        if (filterEmptyString) {
            result = result.filter(v =>
                !(typeof v === "string" && v === "")
            );
        }
    
        return result as T[];
    }

    /**
     * Convertit un tableau en une chaine de caractères, dont les éléments sont séparés par un symbole.
     * Si les éléments sont identiques, il n'est repris qu'une seule fois.
     * @param {unknown} values - Tableau d'objets ou de valeurs à convertir. Si values n'est pas
     *  un tableau, il est converti en chaine de caractères (si non nul) et renvoyé.
     * @param {string} [symbol=";"] - Symbole de séparation des éléments.
     * @param {boolean} [mergeEqualValues=true] - Si vrai, les valeurs identiques sont fusionnées.
     * @param {string} [defaultValue="?"] - Valeur par défaut pour les valeurs non définies.
     * @returns {string} - Chaine de caractères contenant les éléments du tableau.
     */
    public static joinArray(
        values: unknown,
        {
            symbol = ";",
            mergeEqualValues = true,
            defaultValue = ""
        }: {
            symbol?: string,
            mergeEqualValues?: boolean,
            defaultValue?: string
        } = {}
    ): string {
        
        if (values == null) return defaultValue;       
        if (!Array.isArray(values)) return String(values);

        const stringValues: string[] = values
            .map(v => v?.toString() ?? defaultValue)
            .filter(v => v !== "");

        if (stringValues.length === 0) return "";

        if (mergeEqualValues
            && stringValues.every(value => value === stringValues[0])
        ) return stringValues[0];
    
        return stringValues.join(symbol);
    }

    /**
     * Sépare une chaine de caractères en tableau de chaine de caractères.
     * @param {Nullable<string>} value - Chaine de caractères à séparer.
     * @param {RegExp} [separators=/[ +,:;]+/] - Expression réguliére des séparateurs. 
     * @param {(value: string) => T} parse - Fonction de conversion.
     * @returns {string[]} - Tableau de chaine de caractères.
     */
    public static splitArray<T>(
        value: PrimitiveValue | undefined,
        {
            separators = /[ +,:;]+/,
            parse = (v: string): T => v as unknown as T
        }: {
            separators?: RegExp | string,
            parse?: (value: string) => T
        } = {}
    ): T[] {
    
        if (value == null) return [];

        return String(value)
            .split(separators)
            .map(v => v.trim())
            .filter(v => v !== "")
            .map(v => parse(v));
    }

    /**
     * Sérialise une Map sous forme de chaîne de caractères.
     * @template K - Type des clés de la map.
     * @template V - Type des valeurs de la map.
     * @param {Map<K, V>} map - Map à sérialiser.
     * @param {(key: K) => string} [serializeKey=String] - Fonction de sérialisation des clés.
     * @param {(value: V) => string} [serializeValue=String] - Fonction de sérialisation des valeurs.
     * @param {string} [entrySeparator="|"] - Séparateur entre les entrées.
     * @param {string} [keyValueSeparator=">>"] - Séparateur entre la clé et la valeur.
     * @param {string} [defaultValue="?"] - Valeur par défaut pour les valeurs non définies.
     * @returns {string} - Chaîne sérialisée.
     */
    public static serializeMap<K, V>(
        map: Map<K, V>,
        {
            serializeKey = String,
            serializeValue = (value: V) => value ? String(value) : undefined,
            entrySeparator = "|",
            keyValueSeparator = ">>",
            defaultValue = "?"
        }: {
            serializeKey?: (key: K) => string,
            serializeValue?: (value: V) => string | undefined,
            entrySeparator?: string,
            keyValueSeparator?: string,
            defaultValue?: string
        } = {}
    ): string {
        return Array.from(map.entries())
            .map(([key, value]) => serializeKey(key) + keyValueSeparator + (serializeValue(value) ?? defaultValue))
            .join(entrySeparator);
    }

    /**
     * Désérialise une chaîne de caractères sous forme de Map.
     * @param {PrimitiveValue | undefined} value - Chaîne à désérialiser.
     * @template K - Type des clés de la map.
     * @template V - Type des valeurs de la map.
     * @param {(key: string) => K} parseKey - Fonction de conversion des clés.
     * @param {(value: string) => V} parseValue - Fonction de conversion des valeurs.
     * @param {string} [entrySeparator="|"] - Séparateur entre les entrées.
     * @param {string} [keyValueSeparator=">>"] - Séparateur entre la clé et la valeur.
     * @returns {Map<K, V>} - Map reconstruite.
     */
    public static deserializeMap<K, V>(
        value: PrimitiveValue | undefined,
        {
            parseKey,
            parseValue,
            entrySeparator = "|",
            keyValueSeparator = ">>",
            defaultValue = "?"
        }: {
            parseKey: (key: string) => K,
            parseValue: (value: string | undefined) => V,
            entrySeparator?: string,
            keyValueSeparator?: string,
            defaultValue?: string
        }
    ): Map<K, V> {

        const map = new Map<K, V>();
        if (value == null || value === "") return map;

        for (const entry of String(value).split(entrySeparator)) {
            const [key, val] = entry.split(keyValueSeparator);
            if (key === undefined || val === undefined) continue;
            map.set(parseKey(key), parseValue(val === defaultValue ? undefined : val));
        }

        return map;
    }
}

/**
 * Type CellValue reprenant une valeur de cellule de feuille de calcul Excel,
 *  qui est de type primitif ou indéfini.
 */
type CellValue = PrimitiveValue | undefined;

/**
 * Type ExcelWorksheet reprenant une feuille de calcul Excel.
 */
type ExcelWorksheet = ExcelScript.Worksheet;

/**
 * Type ExcelTable reprenant une table de calcul Excel.
 */
type ExcelTable = ExcelScript.Table;

/**
 * Type ExcelRange reprenant une plage de cellules de feuille de calcul Excel.
 */
type ExcelRange = ExcelScript.Range;

/**
 * Type HorizontalAlignment reprenant les différents types d'alignements horizontaux :
 *  - left : aligné à gauche,
 *  - center : aligné au centre,
 *  - right : aligné à droite,
 *  - fill : en remplissage,
 *  - justify : justifié,
 *  - centerAcrossSelection : aligné au centre de la sélection,
 *  - distributed : distribué,
 *  - general : mode général.
 */
type HorizontalAlignment = ExcelScript.HorizontalAlignment;

/**
 * Type VerticalAlignment reprenant les différents types d'alignements verticaux :
 *  - top : aligné en haut,
 *  - center : aligné au centre,
 *  - bottom : aligné en bas,
 *  - justify : justifié,
 *  - distributed : distribué.
 */
type VerticalAlignment = ExcelScript.VerticalAlignment;

/**
 * Type UnderlineStyle reprenant les styles de soulignement :
 *  - none : aucun,
 *  - single : simple,
 *  - double : double.
 */
type UnderlineStyle = ExcelScript.RangeUnderlineStyle;

/**
 * Type RangeFormatOptions reprenant les paramètres de formatage d'une cellule de feuille de calcul Excel.
 * @param {string} numberFormat - Format Excel de la colonne.
 * @param {number} width - Largeur de la colonne.
 * @param {boolean} autoFit - Largeur automatique.
 * @param {HorizontalAlignment} horizontalAlignment - Alignement horizontal.
 * @param {VerticalAlignment} verticalAlignment - Alignement vertical.
 * @param {boolean} bold - Gras.
 * @param {boolean} italic - Italique.
 * @param {UnderlineStyle} underline - Souligné.
 * @param {string} fontName - Nom de la police.
 * @param {number} fontSize - Taille de la police.
 * @param {string} fontColor - Couleur de la police (format "#FF0000").
 * @param {string} fillColor - Couleur de fond (format "#FFFF00").
 * @param {boolean} wrapText - Texte sur plusieurs lignes.
 * @param {boolean} hidden - Colonne cachée.
 */
type RangeFormatOptions = {
    numberFormat?: string;
    width?: number;
    autoFit?: boolean;
    horizontalAlignment?: HorizontalAlignment;
    verticalAlignment?: VerticalAlignment;
    bold?: boolean;
    italic?: boolean;
    underline?: UnderlineStyle;
    fontName?: string;
    fontSize?: number;
    fontColor?: string;
    fillColor?: string;
    wrapText?: boolean;
    hidden?: boolean;
};

/**
 * Type SortOptions reprenant les directions de tri.
 * @param {number} order - Ordre de tri.
 * @param {boolean} ascending - Sens du tri.
 */
type SortOptions = {
    order: number;
    ascending?: boolean;
}

/*
 * Classe utilitaire WorkbookServices de manipulation des feuilles de calcul Excel.
 */
class WorkbookServices {

    /**
     * Retourne la feuille de calcul Excel correspondant au nom donné.
     * Si la feuille n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * Si createIfMissing est vrai, crée la feuille si elle n'existe pas.
     * @param {string} sheetName - Nom de la feuille de calcul à chercher.
     * @param {boolean} createIfMissing - Si vrai, crée la feuille si elle n'existe pas (faux par défaut).
     * @param {boolean} failOnError - Si vrai (par défaut), lance une exception si la feuille n'existe pas.
     * @returns {Nullable<ExcelWorksheet>} - Feuille de calcul Excel correspondant au nom donné,
     *  ou null si elle n'existe pas.
     */
    public static getSheet(
        {
            sheetName,
            createIfMissing = false,
            failOnError = true
        }: {
            sheetName: string,
            createIfMissing?: boolean,
            failOnError?: boolean
        }
    ): Nullable<ExcelWorksheet> {
 
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
     * Retourne toutes les données non vides d'une feuille Excel.
     * Utilise la plage "utilisée" (used range).
     * @param {string} sheetName - Nom de la feuille.
     * @param {boolean} [failOnError=true] - Si vrai, lance une erreur si la feuille est vide ou inexistante.
     * @returns {PrimitiveValue[][]} - Données de la feuille.
     */
    public static getDataFromSheet(
        {
            sheetName,
            failOnError = true
        }: {
            sheetName: string,
            failOnError?: boolean
        }
    ): PrimitiveValue[][] {

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
     * Retourne le tableau Excel correspondant au nom donné dans la feuille de calcul donnée.
     * Si le tableau n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * @param {string} sheetName - Nom de la feuille de calcul où chercher le tableau.
     * @param {string} tableName - Nom du tableau à chercher.
     * @param {boolean} [failOnError=true] - Si vrai (par défaut), lance une exception
     *  si le tableau n'existe pas. Si faux, renvoie null.
     * @returns {Nullable<ExcelTable>} - Tableau Excel correspondant au nom donné,
     *  ou null si il n'existe pas.
     */
    public static getTable(
        {
            sheetName,
            tableName,
            failOnError = true
        }: {
            sheetName: string,
            tableName?: string
            failOnError?: boolean
        }
    ): Nullable<ExcelTable> {
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
     * Retourne les données du tableau Excel correspondant au nom donné
     *  dans la feuille de calcul donnée.
     * Si le tableau n'existe pas, renvoie null si failOnError est faux,
     *  sinon lance une exception.
     * @param {string} sheetName - Nom de la feuille de calcul où chercher le tableau.
     * @param {string} tableName - Nom du tableau à chercher.
     * @param {boolean} [failOnError=true] - Si vrai (par défaut),
     *  lance une exception si le tableau n'existe pas. Si faux, renvoie null.
     * @returns {PrimitiveValue[][]} - Données du tableau Excel
     *  correspondant au nom donné, ou null si il n'existe pas.
     */
    public static getDataFromTable(
        {
            sheetName,
            tableName,
            failOnError = true
        }: {
            sheetName: string,
            tableName?: string
            failOnError?: boolean
        }
    ): PrimitiveValue[][] {
        const table = this.getTable({ sheetName, tableName, failOnError });
        if (!table) return [];
        return table.getRange().getValues();
    }

    /**
     * Retourne toutes les lignes d'un tableau ou d'une feuille Excel.
     * @param {string} sheetName - Nom de la feuille.
     * @param {string} [tableName] - Nom du tableau.
     * @param {boolean} [failOnError=true] - Si vrai, lance une erreur si la feuille est vide ou inexistante.
     * @returns {PrimitiveValue[][]} - Données de la feuille.
     */
    public static getRows(
        {
            sheetName,
            tableName,
            failOnError = true
        }: {
            sheetName: string,
            tableName?: string
            failOnError?: boolean
        }
    ): [number, PrimitiveValue[]][] {
        const data = tableName
            ? WorkbookServices.getDataFromTable({ sheetName, tableName, failOnError })
            : WorkbookServices.getDataFromSheet({ sheetName, failOnError });
        if (!data || data.length <= 1) return [];
        return Array.from(data.slice(1).entries());
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne.
     * Si la valeur est null ou undefined, renvoie undefined.
     * Si la valeur est un nombre, le convertit en chaîne.
     * Si la valeur est une chaîne, la renvoie telle quelle, en supprimant les espaces inutiles.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @returns {string | undefined} - Valeur de la cellule sous forme de chaîne,
     *  ou undefined si elle est null ou undefined.
     */
    public static getStringOrUndefined(
        row: unknown[],
        col: number
    ): string | undefined {
        const v = row[col];
        if (v == null) return undefined;
        return String(v).trim() || undefined;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de nombre.
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
    public static getNumberOrUndefined(
        row: unknown[],
        col: number
    ): number | undefined {
        const v = row[col];
        if (typeof v === "number") return v;
        if (typeof v === "string" && v !== "") {
            const n = Number(v.replace(",", "."));
            return Number.isFinite(n) ? n : undefined;
        }
        return undefined;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de booléen.
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
    public static getBooleanOrUndefined(
        row: unknown[],
        col: number
    ): boolean | undefined {
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
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne,
     *  en supprimant les espaces inutiles, ou la valeur par défaut si la valeur est null ou undefined.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [defaultValue=""] - Valeur par défaut.
     * @returns {string} - Valeur de la cellule sous forme de chaîne.
     **/
    public static getString(
        row: unknown[],
        col: number,
        { defaultValue = "" }: { defaultValue?: string } = {}
    ): string {
        return this.getStringOrUndefined(row, col) ?? defaultValue;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de nombre,
     *  ou la valeur par défaut si la valeur est null ou undefined.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {number} [defaultValue=0] - Valeur par défaut.
     * @returns {number} - Valeur de la cellule sous forme de nombre.
     **/
    public static getNumber(
        row: unknown[],
        col: number,
        { defaultValue = 0 }: { defaultValue?: number } = {}
    ): number {
        return this.getNumberOrUndefined(row, col) ?? defaultValue;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de booléen,
     *  ou la valeur par défaut si la valeur est null ou undefined.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {boolean} [defaultValue=false] - Valeur par défaut.
     * @returns {boolean} - Valeur de la cellule sous forme de booléen.
     **/
    public static getBoolean(
        row: unknown[],
        col: number,
        { defaultValue = false }: { defaultValue?: boolean } = {}
    ): boolean {
        return this.getBooleanOrUndefined(row, col) ?? defaultValue;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de chaîne,
     *  ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @param {unknown[]} row - Ligne contenant la cellule. 
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {string} - Valeur de la cellule sous forme de chaîne.
     **/
    public static getRequiredString(
        row: unknown[],
        col: number,
        { errorMessage }: { errorMessage?: string } = {}
    ): string {
        const value = this.getStringOrUndefined(row, col);
        if (value === undefined) {
            throw new Error(errorMessage
                ?? `La chaine à récupérer est absente`
                    + ` dans la colonne ${col} de la ligne ${JSON.stringify(row)}.`);
        }
        return value;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de nombre,
     *  ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {number} - Valeur de la cellule sous forme de nombre.
     **/
    public static getRequiredNumber(
        row: unknown[],
        col: number,
        { errorMessage }: { errorMessage?: string } = {}
    ): number {
        const value = this.getNumberOrUndefined(row, col);
        if (value === undefined) {
            throw new Error(errorMessage
                ?? `Le nombre à récupérer est absent`
                    + ` dans la colonne ${col} de la ligne ${JSON.stringify(row)}.`);
        }
        return value;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{col}] sous forme de booléen,
     *  ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @param {unknown[]} row - Ligne contenant la cellule.
     * @param {number} col - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {boolean} - Valeur de la cellule sous forme de booléen.
     **/
    public static getRequiredBoolean(row: unknown[],
        col: number,
        { errorMessage }: { errorMessage?: string } = {}
    ): boolean {
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
    public static checkCellName(
        cellName: string,
        { failOnError = true }: { failOnError?: boolean } = {}
    ): string {
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
     * @param {(CellValue)[][]} data - Données du tableau.
     * @param {string} sheetName - Nom de la feuille de calcul où afficher le tableau.
     * @param {string} tableName - Nom du tableau à afficher.
     * @param {string} [startCell="A1"] - Cellule où commencer à afficher le tableau
     *  (par défaut: "A1").
     * @param {boolean} [failOnError=true] - Si vrai (par défaut), lance une exception
     *  si des erreurs surviennent. Si faux, renvoie null.
     * @returns {Nullable<ExcelTable>} - Tableau Excel créé, ou null si une erreur survient.
     */
    public static printTable(
        {
            headers,
            data,
            sheetName,
            tableName,
            startCell = "A1",
            failOnError = true
        }: {
            headers: string[],
            data: CellValue[][],
            sheetName: string,
            tableName: string,
            startCell?: string,
            failOnError?: boolean
        }
    ): Nullable<ExcelTable> {

        // Combine les en-têtes et les données.
        if (headers.length !== data[0].length) {
            throw new Error(`Les en-têtes et les données doivent avoir la même longueur.`);
        }
        const tableData: CellValue[][] = [headers];
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
        const existingTable = sheet.getTables().find((table: ExcelTable) => 
            table.getName() === tableName);
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

    /**
     * Applique le format fourni aux cellules fournies.
     * @param {ExcelRange} range - La plage de cellules à formater.
     * @param {RangeFormatOptions} format - Le format à appliquer aux cellules.
     */
    public static applyFormat(
        range: ExcelRange,
        format: RangeFormatOptions
    ): void {
    
        const rangeFormat = range.getFormat();
    
        // Applique le format de cellule.
        if (format.numberFormat) {
            range.setNumberFormat(format.numberFormat);
        }
    
        // Applique la largeur de colonne (valeur en pixels).
        if (format.width !== undefined) {
            rangeFormat.setColumnWidth(format.width / 2);
        }
    
        // Applique le formatage de texte.
        if (format.autoFit) {
            rangeFormat.autofitColumns();
        }
    
        // Cache la colonne.
        if (format.hidden !== undefined) {
            range.setColumnHidden(format.hidden);
        }

        // Applique l'alignement horizontal.
        if (format.horizontalAlignment !== undefined) {
            rangeFormat.setHorizontalAlignment(format.horizontalAlignment as HorizontalAlignment);
        }
        
        // Applique l'alignement vertical.
        if (format.verticalAlignment !== undefined) {
            rangeFormat.setVerticalAlignment(format.verticalAlignment as VerticalAlignment);
        }
    
        // Mets le texte sur plusieurs lignes.
        if (format.wrapText !== undefined) {
            rangeFormat.setWrapText(format.wrapText);
        }
    
        // Applique la couleur de fond.
        if (format.fillColor) {
            rangeFormat.getFill().setColor(format.fillColor);
        }
    
        // Récupère les propriétés de la police.
        const font = rangeFormat.getFont();

        // Applique le nom de la police.
        if (format.fontName) {
            font.setName(format.fontName);
        }
    
        // Applique la taille de la police.
        if (format.fontSize !== undefined) {
            font.setSize(format.fontSize);
        }
    
        // Applique la couleur de la police.
        if (format.fontColor) {
            font.setColor(format.fontColor);
        }
    
        // Mets le texte en gras.
        if (format.bold !== undefined) {
            font.setBold(format.bold);
        }
    
        // Mets le texte en italique.
        if (format.italic !== undefined) {
            font.setItalic(format.italic);
        }
    
        // Mets le texte en souligné.
        if (format.underline !== undefined) {
            font.setUnderline(format.underline as UnderlineStyle);
        }
    }

    /**
     * Applique les tris définis sur les colonnes d'une table.
     * @param {ExcelTable} table - Tableau à trier.
     * @param {TableColumn<unknown>[]} columns - Colonnes du tableau.
     */
    public static applySort(
        table: ExcelTable,
        columns: readonly ({ sort?: SortOptions } | undefined)[]
    ): void {
        const sortFields: ExcelScript.SortField[] = columns
            .map((column, index) => ({ column, index }))
            .filter(item => item.column?.sort)
            .sort((a, b) => a.column!.sort!.order - b.column!.sort!.order)
            .map(item => ({
                key: item.index,
                ascending: item.column!.sort!.ascending ?? true
            }));

        if (sortFields.length === 0) return;

        table.getSort().apply(sortFields);
    }
}

/**
 * Type TableColumn<TObject> représentant la définition d'une donnée à imprimer ou lire dans une colonne,
 *  incluant un titre, un type de donnée, une fonction de chargement et une fonction de sauvegarde.
 * @template TEntity - Entité à laquelle la définition fait référence.
 * @param {string} key - Clé de la colonne.
 * @param {string} header - En-tête de la colonne.
 * @param {number} columnIndex - Numéro de colonne Excel à lire ou écrire (emplacement à partir de 0).
 *  Si valeur négative, la colonne sera chargée à partir d'autres colonnes (contexte), et non écrite.
 *  Si valeur non définie ou en double, l'index sera déduit en fonction de l'ordre de création des colonnes.
 * @param {string} property - Nom de la propriété de l'objet (pour le chargement).
 * @param {PrimitiveType | undefined} type - Type de la propriété à charger.
 * @param {boolean} required - Lève une erreur si la donnée est undefined.
 * @param {PrimitiveValue} loadDefaultValue - Valeur par défaut si la donnée est undefined.
 * @param {(value: PrimitiveValue | undefined, context: LoadContext) => unknown} load - Fonction d'analyse
 *  de la donnée à charger, avec possibilité d'appeler une autre propriété grace à la fonction get("property").
 *  Si load ne fait pas appel à une valeur chargée, faire load: (_, { get }) => ...
 * @param {(value: unknown) => unknown} print - Fonction de conversion vers la valeur
 *  à sauvegarder dans un tableau Excel.
 * @param {CellValue} printDefaultValue - Valeur par défaut si la donnée est undefined.
 * @param {RangeFormatOptions} format - Paramètres de formatage Excel.
 * @param {SortOptions} sort - Paramètres de tri.
 */
type TableColumn<TEntity> = {
    key?: string;
    header?: string;
    columnIndex?: number;
    property?: string;
    type?: PrimitiveType | undefined;
    required?: boolean;
    loadDefaultValue?: CellValue;
    load?: (value: PrimitiveValue | undefined, context: LoadContext) => unknown;
    print?: (value: unknown, entity: TEntity) => unknown;
    printDefaultValue?: CellValue;
    format?: RangeFormatOptions;
    sort?: SortOptions;
};

/**
 * Type TableColumnFactory<TObject> représentant une fonction de fabrication de colonne,
 *  permettant de transmettre des paramètres au constructeur de la colonne.
 * @template TEntity - Entité à laquelle la fonction fait référence.
 */
type TableColumnFactory<TEntity> =
    (params?: Record<string, unknown>) => TableColumn<TEntity>;

/**
 * Type TableColumns<TObject> représentant la liste des définitions de colonnes d'un tableau Excel,
 *  qu'elle soit statique ou dynamique (avec des paramètres).
 * @template TEntity - Entité à laquelle la fonction fait référence.
 */
type TableColumns<TEntity> =
    Record<string, TableColumn<TEntity> | TableColumnFactory<TEntity>>;

/**
 * Type TableColumnParameters reprenant les paramètres d'une colonne dynamique, incluant :
 *  - le numéro de colonne (emplacement à partir de 0),
 *  - l'intitulé de la définition de la colonne,
 *  - les paramètres supplémentaires de la colonne.
 * colonne et le type.
 * @param {number} columnIndex - Numéro de colonne Excel à lire ou écrire (emplacement à partir de 0).
 *  Si valeur négative, la colonne sera chargée à partir d'autres colonnes (contexte), et non écrite.
 *  Si valeur non définie ou en double, l'index sera déduit en fonction de l'ordre de création des colonnes.
 * @param {string} [definition] - Intitulé de la colonne.
 * @param {Record<string, unknown>} [params] - Paramètres supplémentaires.
 */
type TableColumnParameters = {
        columnIndex?: number;
        definition?: string;
        [key: string]: unknown;
    };

/**
 * Type TableColumnReference reprenant le numéro de colonne Excel à lire ou écrire (emplacement à partir de 0).
 *  Si valeur négative, la colonne sera chargée à partir d'autres colonnes (contexte), et non écrite.
 *  Si valeur non définie ou en double, l'index sera déduit en fonction de l'ordre de création des colonnes. 
 *   ou les paramètres pour la fabrication de colonne.
 */
type TableColumnReference = number | TableColumnParameters;

/**
 * Type BuiltTableColumn reprenant la colonne construite avec ses paramètres obligatoires.
 * @template TEntity - Entité à laquelle la fonction fait référence.
 * @param {TableColumn<T>} column - Colonne à construire.
 * @param {number} excelColumn - Numéro de colonne Excel à lire ou écrire (emplacement à partir de 0).
 *  Si valeur négative, la colonne sera chargée à partir d'autres colonnes (contexte), et non écrite.
 *  Si valeur non définie ou en double, l'index sera déduit en fonction de l'ordre de création des colonnes.
 * @param {string} key - Clé de la colonne.
 * @param {string} property - Nom de la propriété.
 */
type BuiltTableColumn<TEntity> =
    TableColumn<TEntity> & {
        excelColumn: number;
        key: string;
        property: string;
    };

/**
 * Type RowFilter reprenant les paramètres d'un filtre du chargement de données à partir d'une des colonnes.
 * @param {string | number} column - Colonne servant de filtre (clé ou numéro).
 * @param {(value: unknown) => boolean} filter - Fonction de filtre   
 */
type RowFilter = {
    column: string | number;
    filter?: (value: unknown) => boolean;
}; 

/**
 * Interface LoadContext reprenant le contexte de chargement d'une propriété
 *  avec le nom de la propriété à appeler et son type
 *  si différent de celui donné dans la définition.
 * @param {unknown} get - Fonction de chargement de propriété, avec les 2 paramètres ci-dessous.
 *  @param {string} property - Nom de la propriété.
 *  @param {PrimitiveType} [type] - Type de la propriété.
 */
interface LoadContext {
    get(
        property: string,
        type?: PrimitiveType
    ): unknown;
}

/**
 * Classe TableSerializer permettant d'imprimer des tableaux Excel.
 */
class TableSerializer {

    public static readonly HIDDEN_COLUMN = 1;   // Valeur d'une colonne masquée

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{index}] convertie selon le type indiqué : 
     *  - string : convertit la valeur en chaîne, ou la renvoie en supprimant les espaces inutiles,
     *  - number : convertit la valeur en nombre, ou essaie de convertir une chaîne en nombre
     *     en remplaçant les virgules par des points,
     *  - boolean : convertit la valeur en boolean, ou essaie de convertir la valeur en booléen en remplaçant
     *     les valeurs 1, "true", "1", "oui" et "yes" par true, 0, "false", "0", "non" et "no" par false.
     *  - undefined : renvoie les valeurs number et boolean telles quelles,
     *   les autres sont converties en string.
     * Si la valeur est null ou undefined, ou si la conversion échoue, renvoie undefined.
     * @template {PrimitiveType | undefined} T - Type vers lequel la conversion se fait.
     * @param {T} type - Type de conversion. Si non défini, les valeurs de type number et boolean
     *  sont conservées telles quelles, les autres valeurs sont converties en string.
     * @param {CellValue[]} row - Ligne contenant la cellule.
     * @param {number} index - Colonne contenant la cellule.
     * @returns {ConvertedValue<T> | undefined} - Valeur de la cellule selon le type indiqué,
     *  ou undefined si la conversion échoue.
     */
    public static getValueOrUndefined<T extends PrimitiveType | undefined>(
        {
            type,
            row,
            index
        }: {
            type?: T,
            row: CellValue[],
            index: number
        }
    ): ConvertedValue<T> | undefined {
        return Utils.convertValue(row[index], type);
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{index}] selon le type indiqué,
     *  ou la valeur par défaut si la valeur est null ou undefined.
     * @template {PrimitiveType | undefined} T - Type vers lequel la conversion se fait.
     * @param {T} type - Type de conversion. Si non défini, les valeurs de type number et boolean
     *  sont conservées telles quelles, les autres valeurs sont converties en string.
     * @param {CellValue[]} row - Ligne contenant la cellule.
     * @param {number} index - Colonne contenant la cellule.
     * @param {PrimitiveValue} [defaultValue] - Valeur par défaut,
     *  convertie si besoin en fonction du type.
     * @returns {ConvertedValue<T>} - Valeur de la cellule selon le type indiqué.
     **/
    public static getValue<T extends PrimitiveType | undefined>(
        { 
            type,
            row,
            index,
            defaultValue
        }: {
            type?: T,
            row: CellValue[],
            index: number,
            defaultValue?: PrimitiveValue;
        }
    ): ConvertedValue<T> {
        
        const typeValues: Record<string, PrimitiveValue> = {
            string: "",
            number: 0,
            boolean: false
        };
        const value = this.getValueOrUndefined({ type, row, index });
        if (value !== undefined) return value;

        const defaultConverted = Utils.convertValue(defaultValue, type);

        if (defaultConverted !== undefined) return defaultConverted;

        return (
            type !== undefined
                ? typeValues[type]
                : ""
        ) as ConvertedValue<T>;
    }

    /**
     * Retourne la valeur de la cellule à l'adresse {row}[{index}] sous forme de chaîne,
     *  ou lance une exception si la valeur est null ou undefined, avec un message d'erreur personnalisé.
     * @template {PrimitiveType | undefined} T - Type vers lequel la conversion se fait.
     * @param {PrimitiveType} type - Type de conversion. Si non défini, les valeurs de type number et boolean
     *  sont conservées telles quelles, les autres valeurs sont converties en string. 
     * @param {CellValue[]} row - Ligne contenant la cellule. 
     * @param {number} index - Colonne contenant la cellule.
     * @param {string} [errorMessage] - Message d'erreur personnalisé.
     * @returns {PrimitiveValue} - Valeur de la cellule.
     **/
    
    public static getRequiredValue<T extends PrimitiveType | undefined>(
        { 
            type,
            row,
            index,
            errorMessage
        }: {
            type?: T,
            row: CellValue[],
            index: number,
            errorMessage?: string;
        }
    ): ConvertedValue<T> {

        const value = this.getValueOrUndefined({ type, row, index });

        let message: string;
        switch (type) {
            case "string":
                message = "La chaine à récupérer est absente.";
                break;
            case "number":
                message = "Le nombre à récupérer est absent.";
                break;
            case "boolean":
                message = "Le booléen à récupérer est absent.";
                break;
            default:
                message = "La valeur à récupérer est absente.";
        }

        if (value == undefined || value === "") {
            throw new Error(errorMessage
                ?? `${message} dans la colonne ${index} de la ligne ${JSON.stringify(row)}.`
            );
        }
        return value;
    }

    /**
     * Helper de joinArray avec paramètres par défaut.
     * Convertit un tableau en une chaine de caractères,
     *  dont les éléments sont séparés par un symbole.
     * @param {unknown} value - Tableau d'objets ou de valeurs à convertir.
     * @returns {string} - Chaine de caractères contenant les éléments du tableau.
     */
    public static readonly printArray: (value: unknown) => string
        = Utils.joinArray;

    /**
     * Helper de splitArray avec paramètres par défaut.
     * Sépare une chaine de caractères en tableau de chaines de caractères
     *  pour le chargement d'un tableau de valeurs depuis Excel.
     * @param {PrimitiveValue | undefined} value - Chaine de caractères à séparer.
     * @returns {string[]} - Tableau de chaines de caractères.
     */
    public static readonly loadArray: (value: PrimitiveValue | undefined) => string[]
        = Utils.splitArray;

    /**
     * Construit la liste des colonnes à imprimer à partir
     * d'une définition de colonnes et d'une configuration.
     * @template TEntity - Entité à laquelle la colonne fait réfrence.
     * @param {Record<string, number>} columns - Colonnes demandées.
     * @param {TableColumns<TEntity>} definitions - Définitions disponibles.
     * @returns {(TableColumn<TEntity> | undefined)[]} - Colonnes à imprimer.
     */
    public static buildColumns<TEntity>(
        columns: Record<string, TableColumnReference>,
        definitions: TableColumns<TEntity>
    ): (BuiltTableColumn<TEntity> | undefined)[] {
    
        const result: (BuiltTableColumn<TEntity> | undefined)[] = [];
    
        const findFreeColumn = (preferred?: number): number => {
            if (preferred !== undefined
                && result[preferred] === undefined)
            {
                return preferred;
            } 
            const sign = (!preferred || preferred >=0) ? 1 : -1;
            let col = 0;
            while (result[sign*col] !== undefined) col++;
            return col;
        };

        for (const [key, columnDefinition] of Object.entries(columns)) {
    
            // Vérifie si la colonne est déjà définie où si elle doit être calculée avec des paramètres.
            const isSimpleColumn = typeof columnDefinition === "number";
    
            // Recherche l'emplacement de la colonne.
            const preferredColumn: number | undefined = isSimpleColumn
                ? columnDefinition
                : columnDefinition.columnIndex;
            const excelColumn = findFreeColumn(preferredColumn);
    
            // Recherche la définition associée.
            const definitionName: string = isSimpleColumn
                ? key
                : columnDefinition.definition ?? key;
            const definition = definitions[definitionName];
            if (!definition) continue;

            // Récupère la définition de la colonne
            //  selon si elle est définie de manière constante ou dépend de paramètres (fonction).
            let builtColumn: TableColumn<TEntity>;
            if (typeof definition === "function") {
                // Ajoute la définition avec ses paramètres
                builtColumn = definition(
                    isSimpleColumn
                        ? undefined
                        : columnDefinition
                );
            } else {
                // Ajoute simplement la colonne sans paramètre (valeurs fixes) avec sa clé.
                builtColumn = definition;
            }

            // Ajoute le numéro de colonne, la propriété associée et la clé
            //  comme composantes de la définition.
            result[excelColumn] = {
                key,
                property: builtColumn.property ?? key,
                excelColumn,
                ...builtColumn
            };
        }

        return result;
    }

    /** Charge à partir d'une ligne Excel la valeur d'une propriété,
     *   y compris si celle-ci fait appel à une autre propriété avec get(property)
     *   dans la fonction load.
     * @template TEntity - Entité à laquelle les propriétés font référence.
     * @param {string} property - Nom de la propriété.
     * @param {CellValue[]} row - Ligne Excel.
     * @param {Record<string, BuiltTableColumn<TEntity>>} builtColumnsByProperty - Dictionnaire
     *  des colonnes par propriété.
     * @param {Record<string, unknown>} result - Valeurs déjà récupérées.
     * @param {string[]} stack - Pile des propriétés en cours de traitement pour éviter les boucles infinies.
     * @returns {unknown} - Valeur de la propriété.
     */
    private static resolveProperty<TEntity>(
        property: string,
        {
            row,
            builtColumnsByProperty,
            result,
            stack = [],
            ignoreRequired = false
        }: {
            row: CellValue[],
            builtColumnsByProperty: Record<string, BuiltTableColumn<TEntity>>,
            result: Record<string, unknown>,
            stack?: string[],
            ignoreRequired?: boolean
        }
    ): unknown {
    
        // Renvoie la valeur de la propriété si déjà calculée.
        if (property in result) {
            return result[property];
        }
    
        // Détecte des références circulaires pour éviter les boucles infinies.
        if (stack.includes(property)) {
            throw new Error(
                `Référence circulaire détectée : `
                + [...stack, property].join(" -> ")
            );
        }
        
        // Recherche la colonne correspondant à la propriété.
        const builtColumn = builtColumnsByProperty[property];
        if (!builtColumn) {
            return undefined;
        }

        // Charge la valeur brute de la cellule, si le numéro de colonne est connu.
        let value: CellValue = row[builtColumn.excelColumn];
        if (builtColumn.type) {
            if (builtColumn.loadDefaultValue !== undefined) {
                value = this.getValue({
                    type: builtColumn.type,
                    row,
                    index: builtColumn.excelColumn,
                    defaultValue: builtColumn.loadDefaultValue
                });
            } else if (builtColumn.required && !ignoreRequired) {
                value = this.getRequiredValue({
                    type: builtColumn.type,
                    row,
                    index: builtColumn.excelColumn,
                    errorMessage: `La valeur de ${builtColumn.header ?? property} est requise.`
                });
            } else {
                value = this.getValueOrUndefined({
                    type: builtColumn.type,
                    row,
                    index: builtColumn.excelColumn,
                });
            }
        }
        
        // Si les erreurs sont empêchées (filtre) alors que la valeur est requise et undefined,
        //  arrête la résolution de la propriété et renvoie undefined.
        if (ignoreRequired && builtColumn.required && value === undefined) return undefined;

        // Récupère le contexte de chargement.
        const context: LoadContext = {

            get: (
                requestedProperty: string,
                forcedType?: PrimitiveType
            ): unknown => {

                // Vérifie que la colonne appelée foruni bien une valeur.
                const requestedColumn = builtColumnsByProperty[requestedProperty];
                if (!requestedColumn) {
                    return undefined;
                }
    
                // Récupère directement la valeur demandée en cas de surcharge de type
                //  (type différent du type initial de la propriété demandée)
                if (forcedType) {
                    return this.getValueOrUndefined({
                        type: forcedType,
                        row,
                        index: requestedColumn.excelColumn
                    });
                }
    
                // Sans surcharge, récupère la valeur de la propriété demandée.
                return this.resolveProperty(
                    requestedProperty,
                    {
                        row,
                        builtColumnsByProperty,
                        result,
                        stack: [...stack, property]
                    }
                );
            }
        };
    
        // Charge la valeur finale avec l'appel éventuel de la fonction de chargement
        //  qui peut faire appel à une autre propriété avec la fonction get() grâce au contexte.
        const loadedValue = builtColumn.load
            ? builtColumn.load(value, context)
            : value;
    
        result[property] = loadedValue;

        return loadedValue;
    }

    /**
     * Charge à partir d'une ligne Excel les données nécessaires à la construction d'un objet,
     *  en en renvoyant les paramètres.
     * @template TEntity - Entité à laquelle les données à charger font référence.
     * @template TResult - Type renvoyé par la fonction de chargement.
     * @param {CellValue[]} row - Ligne de données extraite d'Excel.
     * @param {Record<string, number>} databaseColumns - Liste des colonnes à analyser
     *  avec leur position (à partir de 0).
     * @param {Record<string, any>} definitions - Définitions des colonnes.
     * @param {(RowFilter | string)[]} filters - Filtres sur les lignes.
     * @returns {Record<string, unknown>} - Paramètres de construction de l'objet.
     */
    public static loadRow<TEntity, TResult>({
        row,
        columns,
        definitions,
        filters = []
    }: {
        row: CellValue[],
        columns: Record<string, TableColumnReference>,
        definitions: TableColumns<TEntity>,
        filters?: OneOrMany<RowFilter | string>
    }): TResult | undefined {

        // Construit les colonnes avec la liste des colonnes appelées et leur définitions.
        const builtColumns = this.buildColumns(columns, definitions);

        // Construit le dictionnaire des colonnes par propriétés.
        const builtColumnsByProperty: Record<string, BuiltTableColumn<TEntity>> = {};
        for (const column of builtColumns) {
            if (!column) continue;
            builtColumnsByProperty[column.property] = column;
        }

        // Initialisation du resultat final qui va être constitué
        //  par les différents appels de resolveProperty().
        const result: Record<string, unknown> = {};
        
        // Filtre les lignes selon les filtres fournis
        const filtersArray = Utils.asArray(filters);
        for (const filter of filtersArray) {

            // Normalise le filtre selon s'il est déjà un RowFilter
            //  ou simplement le nom de la colonne à filtrer.
            const rowFilter: RowFilter = typeof filter === "string"
                ? { column: filter }
                : filter;

            // Récupère la propriété à filtrer 
            //  selon si la colonne est donnée par son nom ou son numéro de colonne.
            const property = typeof rowFilter.column === "string"
                ? rowFilter.column
                : builtColumns[rowFilter.column]?.property;
            if (!property) return undefined;
        
            const value = this.resolveProperty(
                property,
                {
                    row,
                    builtColumnsByProperty,
                    result,
                    ignoreRequired: true
                }
            );

            // Défini la fonction de filtre :
            //  par défaut la valeur doit être différente de "".
            const filterFunction = rowFilter.filter ?? ((v: unknown) => String(v) !== "");

            // Si la valeur est non définie ou ne passe pas le filtre, renvoie undefined.
            if (value === undefined || !filterFunction(value)) return undefined;
        }
    
        // Récupère les données de la ligne et les affecte aux propriétés de l'objet
        for (const column of builtColumns) {
            if (!column) continue;
            this.resolveProperty(
                column.property,
                {
                    row,
                    builtColumnsByProperty,
                    result
                }
            );
        }
    
        return result as TResult;
    }

    /**
     * Imprime dans un tableau Excel les données d'une classe à partir de la liste des colonnes à imprimer.
     * @template TEntity - Entité à laquelle les données font réfrence.
     * @param {entities: TEntity[]} [entities] - Données à imprimer.
     * @param {(TableColumn<TEntity> | undefined)[]} [columns] - Colonnes à imprimer.
     * @param {string} [sheetName] - Nom de la feuille.
     * @param {string} [tableName] - Nom du tableau.
     * @param {string} [startCell="A1"] - Cellule de départ.
     * @returns {Nullable<ExcelTable>} - Tableau Excel.
     */
    public static print<TEntity>({
        entities,
        columns,
        definitions,
        sheetName,
        tableName,
        startCell = "A1"
    }: {
        entities: TEntity[],
        columns: Record<string, TableColumnReference>,
        definitions: TableColumns<TEntity>
        sheetName: string,
        tableName: string,
        startCell?: string
    }): Nullable<ExcelTable> {

        // Construit les colonnes avec la liste des colonnes appelées et leur définitions.
        const builtColumns: (BuiltTableColumn<TEntity> | undefined)[] = this.buildColumns(columns, definitions);

        // Génère l'en-tête.
        const headers: string[] = builtColumns.map(column => column?.header ?? "");

        // Génère les données.
        const data: CellValue[][] = entities.map(entity => {
            const row: CellValue[] = Array(builtColumns.length); //.fill("");
            for (const [index, column] of Array.from(builtColumns.entries())) {
                if (!column || column.excelColumn < 0) continue;
                const propertyValue = (entity as Record<string, unknown>)[column.property];
                const value = Utils.convertValue(
                        column.print?.(propertyValue, entity) ?? propertyValue,
                        column.type)
                row[column.excelColumn] = value ?? column.printDefaultValue;
            }
            return row;
        });

        // Imprime le tableau.
        const table = WorkbookServices.printTable({
            headers,
            data,
            sheetName,
            tableName,
            startCell
        });
    
        // Formate les colonnes.
        for (const [index, column] of Array.from(builtColumns.entries())) {
            if (!column?.format) continue;
            WorkbookServices.applyFormat(
                table.getRange().getColumn(index),
                column.format
            );
        }

        // Trie les colonnes.
        WorkbookServices.applySort(
            table,
            builtColumns
        );

        return table;
    }
}

/**
 * Type ParameterDefinition représentant les éléments de la définition d'un paramètre.
 * @template TParam - Type final du paramètre.
 * @param {string} title - Intitule du paramètre.
 * @param {PrimitiveType} type - Type du paramètre dans le tableau Excel.
 * @param {unknown | (() => unknown)} defaultValue - Valeur par défaut du paramètre,
 *  ou appel de cette valeur si présente dans une autre classe.
 * @param {(value: PrimitiveValue | undefined) => boolean} validate - Condition d'acceptation
 *  de la valeur brute du paramètre (avant analyse par load).
 * @param {(value: PrimitiveValue | undefined) => unknown} deserialize - Fonction d'analyse
 *  de la valeur brute du paramètre.
 * @param {(value: TParam) => PrimitiveValue | undefined} serialize - Fonction de sérialisation
 *  vers la valeur à sauvegarder dans un tableau Excel.
 * @param {string} [numberFormat] - Format numérique.
 */
interface ParameterDefinition<TParam> {
    title: string;
    type?: PrimitiveType;
    defaultValue: PrimitiveValue | (() => PrimitiveValue);
    validate?: (value: PrimitiveValue) => boolean;
    deserialize?: (value: PrimitiveValue) => TParam | undefined;
    serialize?: (value: TParam) => PrimitiveValue | undefined;
    numberFormat?: string;
}

/**
 * Type ParameterValue extrayant le type final d'une définition de paramètre.
 */
type ParameterValue<T> =
    T extends ParameterDefinition<infer TValue>
        ? TValue
        : never;
        
/**
 * Type ParamsValues représentant les valeurs des paramètres.
 */
type ParamsValues = {
    [K in ParamKey]:
        ParameterValue<typeof Params.PARAMETER_DEFINITIONS[K]>;
};

/**
 * Type ParamKey représentant les clés des paramètres.
 */
type ParamKey = keyof typeof Params.PARAMETER_DEFINITIONS;

/*
 * Classe utilitaire Params contenant les paramètres globaux.
 */
class Params {  

    // Constantes de lecture des paramètres sur Excel
    private static readonly SHEET = "Paramètres";           // Nom de la feuille
    private static readonly TABLE = "Paramètres";           // Nom du tableau
    private static readonly START_CELL = "A1";              // Première cellule
    private static readonly DATABASE_COLUMNS = {            // Liste des colonnes avec leur emplacement
        title: 0,
        value: 1
    } as const;
    private static readonly COLUMN_DEFINITIONS: TableColumns<unknown> = {          // Définitions des colonnes
        title: { header: "Paramètre", type: "string" },
        value: { header: "Valeur", type: undefined }    
    };

    // Liste des paramètres
    public static readonly PARAMETER_DEFINITIONS = {                  // Liste des paramètres
        maxConnectionNumber: {
            title: "Nombre maximum de connexions par gare",
            type: "number",
            defaultValue: 6,    // 6 connexions maximum
            validate: (value: PrimitiveValue | undefined) => 
                typeof value === "number" 
                    && value > 0
        } as ParameterDefinition<number>,
    
        turnaroundTime: {
            title: "Temps de retournement (minutes)",
            type: "number",
            defaultValue: 10,    // 10 minutes
            validate: value => Number(value) >= 0,
            deserialize: value => {
                let excelTime = Number(value);
                if (excelTime >= 1) excelTime = (excelTime / 60 / 24) % 1;
                return DateTime.from(
                    excelTime,
                    { isRelative: true }
                );
            },
            serialize: value => value.excelValue,
            numberFormat: "hh:mm:ss",
        } as ParameterDefinition<DateTime>,
    
        maxTrainUnits: {
            title: "Nombre maximal d'éléments par train",
            type: "number",
            defaultValue: 2,    // 2 unités par train maximum
            validate: value => 
                typeof value === "number" 
                    && value > 0
        } as ParameterDefinition<number>,
    
        stationsSuffixes: {
            title: "Suffixes des gares",
            type: "string",
            defaultValue: "BV;00",  // Suffixes 00 et BV
            deserialize: TableSerializer.loadArray,
            serialize: TableSerializer.printArray
        } as ParameterDefinition<string[]>,
        
        rolloverHour: {
            title: "Heure de changement de journée",
            type: "number",
            defaultValue: () => DateTime.ROLLOVER_HOUR,
            validate: value => Number(value) >= 0,
            deserialize: value => {
                let excelTime = Number(value);
                if (excelTime >= 1) excelTime = (excelTime / 24) % 1;
                return DateTime.from(
                    excelTime,
                    { isRelative: true }
                );
            },
            serialize: value => value.excelValue,
            numberFormat: "hh:mm"
        } as ParameterDefinition<DateTime>,
    
        gainTimeWithoutStop: {
            title: "Gain de temps d'un passage sans arrêt (minutes)",
            type: "number",
            defaultValue: 2,    // 2 minutes
            deserialize: value => {
                let excelTime = Math.abs(Number(value));
                if (excelTime >= 1) excelTime = (excelTime / 60 / 24) % 1;
                return DateTime.from(
                    excelTime,
                    { isRelative: true }
                );
            },
            serialize: value => value.excelValue,
            numberFormat: "hh:mm:ss"
        } as ParameterDefinition<DateTime>

    } as const;

    // Map des paramètres
    public static readonly map: Map<ParamKey, ParamsValues[ParamKey]> = new Map();

    /**
     * Retourne le nombre de paramètres enregistrés dans la base de données
     * @returns {number} - Nombre de paramètres enregistrés
     */
    public static get size(): number {
        return this.map.size;
    }

    /**
     * Vérifie si un paramètre est présent dans la base de données.
     * @template { ParamKey } K - Type des clés de la map des paramètres.
     * @param {K} key - Clé du paramètre.
     * @returns {boolean} - Vrai si le paramètre est présent, faux sinon.
     */
    public static has<K extends ParamKey>(
        key: K
    ): boolean {
        return this.map.has(key);
    }

    /**
     * Retourne le paramètre correspondant à la clé donnée.
     * @template { ParamKey } K - Type des clés de la map des paramètres.
     * @param {K} key - Clé du paramètre.
     * @returns {unknown | undefined} - Paramètre correspondant, ou undefined si la clé n'existe pas.
     */
    public static get<K extends ParamKey>(
        key: K
    ): ParamsValues[K] {
        return this.map.get(key) as ParamsValues[K];
    }

    /**
     * Ajoute un nouveau paramètre dans la base de données, référencé par sa clé.
     * @template { ParamKey } K - Type des clés de la map des paramètres.
     * @param {K} key - Nom du paramètre.
     * @param {unknown} value - Valeur du paramètre.
     */
    public static set<K extends ParamKey>(
        key: K,
        value: ParamsValues[K]
    ): void {
        this.map.set(key, value);
    }
 
    /**
     * Retourne un tableau des valeurs de la base de données des paramètres.
     * @returns {unknown[]} - Itérateur sur les valeurs.
     *  de la base de données des paramètres.
     */
    public static values(): unknown[] {
        return Array.from(this.map.values());
    }

    /**
     * Efface toutes les paramètres de la base de données.
     * Cela permet de forcer le rechargement des paramètres si besoin.
     */
    public static clear() {
        this.map.clear();
    }

    /**
     * Analyse la valeur brute d'un paramètre récupérée dans la base de données.
     * @template TParam - Type du paramètre.
     * @param {PrimitiveValue} value - Valeur brute.
     * @param {ParameterDefinition} definition - Définition du paramètre.
     * @returns {unknown} - Valeur analysée.
     */
    private static deserializeParameter<TParam>(
        value: PrimitiveValue | undefined,
        definition: ParameterDefinition<TParam>        
    ): TParam | undefined {
        const primitiveValue = Utils.convertValue(value, definition.type);
        if (primitiveValue === undefined) return undefined;
        if (definition.validate && !definition.validate(primitiveValue)) return undefined;
        return definition.deserialize 
            ? definition.deserialize(primitiveValue) 
            : primitiveValue as unknown as TParam;
    }

    /**
     * Sérialise la valeur d'un paramètre pour l'enregistrer dans la base de données.
     * @template TParam - Type du paramètre.
     * @param {unknown} value - Valeur du paramètre.
     * @param {ParameterDefinition} definition - Définition du paramètre.
     * @
     */
    private static serializeParameter<TParam>(
        value: TParam,
        definition: ParameterDefinition<TParam>
    ): PrimitiveValue | undefined {
        return Utils.convertValue(definition.serialize
            ? definition.serialize(value)
            : value);
    }

    /**
     * Charge les paramètres.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les paramètres avec valeurs par défaut.
        const paramsTitleToKey: Record<string, ParamKey> = Object.create(null);
        for (const [key, definition] of Object.entries(this.PARAMETER_DEFINITIONS)) {
            const defaultValue:PrimitiveValue = (typeof definition.defaultValue === "function")
                ? definition.defaultValue()
                : definition.defaultValue;
            this.set(key as ParamKey, this.deserializeParameter(
                defaultValue,
                definition as ParameterDefinition<ParamsValues[ParamKey]>)!);
            paramsTitleToKey[definition.title] = key as ParamKey;
        }

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE,
            failOnError: false
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table. Celle-ci va être créée.`);
            this.save();
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        let loadedParameters: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;
    
                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2;
    
                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    unknown,
                    { title: string, value: PrimitiveValue | undefined }
                >({
                    row,
                    columns:this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: [ "title", "value" ],
                });
                if (!params) continue;
    
                // Analyse la valeur récupérée.
                if (!(params.title in paramsTitleToKey)) {
                    Log.warn(`${this.name}.load() : paramètre '${params.title}' inexistant.`);
                    continue;
                }
                const key = paramsTitleToKey[params.title];
                const definition = this.PARAMETER_DEFINITIONS[key];
                const value = this.deserializeParameter(
                    params.value,
                    definition as ParameterDefinition<ParamsValues[ParamKey]>);
                if (value === undefined) continue;
                

                // Enregistre la valeur du paramètre.
                this.set(key, value);
                loadedParameters++;
            }

        } catch (e) {
            throw new Error(`Paths.load (ligne ${excelRow}) : ${e}`);
        }

        // Sauvegarde les paramètres si le tableau était incomplet.
        if (loadedParameters < this.size) {
            this.save();
        }

        // Charge les paramètres des classes utilitaires.
        Days.load({ erase });
        Parity.load({ erase });
        TrainNumber.load({ erase });

        Log.timer(`${this.name}.load()`);
    }

    /**
     * Sauvegarde la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell=this.START_CELL] - Adresse de la cellule de départ pour le tableau.
     */
    public static save(
        {
            sheetName = this.SHEET,
            tableName = this.TABLE,
            startCell = this.START_CELL
        }: {
            sheetName?: string,
            tableName?: string,
            startCell?: string
        } = {}
    ): void {
    
        Log.startTimer(`${this.name}.save()`);

        // Récupère les informations de formatage des paramètres
        let i = 1;
        const values = Object.entries(this.PARAMETER_DEFINITIONS)
            .map(([key, definition]) => ({
                title: definition.title,
                value: this.serializeParameter(
                    this.get(key as ParamKey),
                    definition as ParameterDefinition<ParamsValues[ParamKey]>),
                rowIndex: i++,
                numberFormat: definition.numberFormat
            }));

        // Imprime les paramètres
        const table = TableSerializer.print({
            entities: values,
            columns: this.DATABASE_COLUMNS,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        // Formate les cellules des paramètres concernés
        for (const value of values) {
            if (!value.numberFormat) continue;
            WorkbookServices.applyFormat(
                table.getRange().getCell(value.rowIndex, this.DATABASE_COLUMNS.value),
                { numberFormat: value.numberFormat }
            );
        }

        Log.timer(`${this.name}.save()`);
    
        Stops.save();
    }
}

/*
 * Classe utilitaire immuable contenant les valeurs d'une date Excel.
 */
class ExcelDate {

    // Constante de valeur initiale (epoch) des dates Excel
    public static readonly EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30));

    // Cache des jours fériés par année
    private static holidayCache: Map<number, Set<number>> = new Map();
 
    // Propriétés de la classe DateTime
    public readonly value: number;          // Valeur Excel
    public readonly year: number;           // Année
    public readonly month: number;          // Mois
    public readonly day: number;            // Jour
    public readonly dayOfWeek: Day;         // Jour de la semaine
    public readonly isHoliday: boolean;     // Indique si le jour est férié

    /**
     * Constructeur de la classe ExcelDate.
     * @param {number} excelValue - Valeur Excel du jour, qui représente le nombre de jours
     *  écoulés depuis le 30 décembre 1899.
     */
    public constructor(
        excelValue: number
    ) {

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
    public static parseDate(
        value: string
    ): number | undefined {

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
    public static getHolidays(
        year: number
    ): Set<number> {
 
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

    // Propriétés de la classe DateTime
    public readonly value: number;      // Valeur Excel
    public readonly hour: number;       // Heure
    public readonly minute: number;     // Minutes
    public readonly second: number;     // Secondes

    /**
     * Constructeur de la classe ExcelTime.
     * @param {number} excelValue - Valeur Excel du temps, dont la fraction de jour représente l'heure.
     */
    public constructor(
        excelValue: number
    ) {
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
    public static parseTime(
        value: string
    ): number | undefined {
        const separatorRegex = /[^\d]/; // Toute caractère ou chaine de caractère non numérique
                                        //  est considérée comme un séparateur (ex : 'h', 'min' ...)
        const parts = value.split(separatorRegex).filter(t => t !== "");
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
 * Type DateTimeInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet DateTime.
 */
type DateTimeInput = Input<DateTime, number | string>;

/**
 * Classe utilitaire immuable DateTime pour la gestion des dates et horaires Excel.
 *  Si le temps est absolu et non daté, et que l'heure est inférieure à l'heure de changement de journée,
 *  elle est incrémentée de 1 pour rester comparable aux autres heures de la journée précédente.
 */
class DateTime {

    // Constantes pour le calcul des dates et heures
    public static readonly ROLLOVER_HOUR = 3 / 24;          // Heure de changement de journée par défaut
                                                            //  qui correspond à 3h00
    private static readonly MIN_EXCEL_DATE = 2;             // Valeur minimale d'un temps absolu daté
    private static readonly MAX_GAP: number = 3/24/3600;    // Différence maximale entre 2 horaires
                                                            //  pour les considérer comme égaux
                                            
    // Constantes du formatage des dates et heures
        
    // Constantes des formats des dates et heures prédéfinis
    public static readonly DATE_FORMAT_FOR_ID: string = "yymmdd";
    public static readonly DATE_FORMAT_WITH_YEAR: string = "dd/mm/yyyy";
    public static readonly DATE_FORMAT_WITHOUT_YEAR: string = "dd/mm";
    public static readonly TIME_FORMAT_WITH_SECONDS: string = "hh:nn:ss";
    public static readonly TIME_FORMAT_WITHOUT_SECONDS: string = "hh:nn";
    public static readonly FULL_DATE_TIME_FORMAT: string = "dd/mm/yyyy hh:nn:ss";                                                           
    private static readonly MONTHS = [                      // Tableau des noms et des abréviations des mois
        { number: 1,  fullName: "Janvier",   abbreviation: "Jan" },
        { number: 2,  fullName: "Février",   abbreviation: "Fev" },
        { number: 3,  fullName: "Mars",      abbreviation: "Mar" },
        { number: 4,  fullName: "Avril",     abbreviation: "Avr" },
        { number: 5,  fullName: "Mai",       abbreviation: "Mai" },
        { number: 6,  fullName: "Juin",      abbreviation: "Juin" },
        { number: 7,  fullName: "Juillet",   abbreviation: "Juil" },
        { number: 8,  fullName: "Août",      abbreviation: "Août" },
        { number: 9,  fullName: "Septembre", abbreviation: "Sep" },
        { number: 10, fullName: "Octobre",   abbreviation: "Oct" },
        { number: 11, fullName: "Novembre",  abbreviation: "Nov" },
        { number: 12, fullName: "Décembre",  abbreviation: "Dec" }
    ];
    private static readonly FORMAT_TOKENS = [           // Tokens de formatage de date et heure
        "yyyy", // Année à 4 chiffres
        "yy",   // Année à 2 chiffres
        "mmmm", // Mois en lettres
        "mmm",  // Mois abbrégé
        "mm",   // Mois à 2 chiffres
        "m",    // Mois à 1 ou 2 chiffres
        "dddd", // Jour de la semaine en lettres
        "ddd",  // Jour de la semaine abbrégé
        "dd",   // Jour à 2 chiffres
        "d",    // Jour à 1 ou 2 chiffres
        "hh",   // Heure à 2 chiffres
        "h",    // Heure à 1 ou 2 chiffres
        "nn",   // Minutes à 2 chiffres
        "n",    // Minutes à 1 ou 2 chiffres
        "ss",   // Secondes à 2 chiffres
        "s"     // Secondes à 1 ou 2 chiffres
    ] as const;                                         
    private static readonly FORMAT_REGEX = new RegExp(
        [...DateTime.FORMAT_TOKENS]
            .sort((a, b) => b.length - a.length)
            .join("|"),
        "g"
    );                                                  // Regex des tokens de formatage

    // Propriétés de la classe DateTime
    public readonly excelValue: number;                 // Valeur du temps en format Excel
                                                        //  à partir du 01/01/1900 00:00:00
    public readonly isRelative: boolean = false;        // Indique si le temps est relatif
                                                        //  (différence entre 2 horaires)
    private _computed: boolean = false;                 // Indique si les éléments de la date sont calculés
    private _formats: Map<string, string> = new Map();  // Cache des formats de la représentation textuelle

    // Valeurs des éléments
    private _realDate?: ExcelDate;                      // Date réelle (uniquement pour les temps absolus
                                                        //  datés, donc avec excelValue >= MIN_EXCEL_DATE)
    private _adaptedDate?: ExcelDate;                   // Date adaptée (jour suivant) si l'heure de la date
                                                        //  est inférieure à l'heure de changement de jour
    private _time?: ExcelTime;                          // Heure de la journée 
                                                        //  (undefined si le temps n'est qu'une date)

    /**
     * Constructeur privé de la classe DateTime.
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
    private constructor(
        excelValue: number,
        {
            isRelative = false,
            adaptTime = true
        }: {
            isRelative?: boolean,
            adaptTime?: boolean
        } = {}
    ) {
        this.isRelative = isRelative;
        this.excelValue = (!isRelative && adaptTime) ? DateTime.adaptTime(excelValue) : excelValue;
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Date et heure au format 'dd/mm/yyyy hh:nn:ss'
     */
    public toString(): string {
        return this.format(DateTime.FULL_DATE_TIME_FORMAT);
    }

    /**
     * Retourne la valeur du temps en nombre décimal.
     * @returns {number} - Valeur Excel du temps en nombre décimal
     */
    public toNumber(): number {
        return this.excelValue;
    }

    /**
     * Crée une instance objet DateTime à partir d'une valeur.
     * Si la valeur est déjà un objet DateTime, il est retourné tel quel.
     * Sinon, un nouvel objet DateTime est créé avec la valeur fournie.
     * @param {Nullable<DateTimeInput>} value Valeur du temps en nombre décimal
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
        value: Nullable<DateTimeInput>,
        {
            isRelative = false,
            adaptTime = true
        }: {
            isRelative?: boolean,
            adaptTime?: boolean
        } = {}
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
                const parsed = this.parseDateAndTime(trimmed, {
                    isRelative,
                    adaptTime
                });
                return parsed;
            }
        } else {
            v = value;
        }
 
        if (v === undefined || isNaN(v)) return undefined;
 
        // Un temps absolu doit être >= 0
        if (!isRelative && v < 0) return undefined;
 
        return new DateTime(v, { isRelative, adaptTime });
    } 

    /**
     * Retourne l'objet ExcelDate de la date, adaptée ou non.
     * @param {boolean} adapted - Indique si l'on souhaite avoir l'objet ExcelDate adapté ou non.
     * @returns {ExcelDate | undefined} - L'objet ExcelDate correspondant au temps adapté ou non,
     *  ou undefined si le temps n'est pas défini.
     */
    private getDateObj(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): ExcelDate | undefined {
        return (adaptTime && this._adaptedDate) ? this._adaptedDate : this._realDate;
    }

    /**
     * Retourne la valeur Excel de la date, adaptée ou non.
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir la valeur Excel
     *  de la date adaptée ou non.
     * @returns {number} - La valeur Excel de la date, adaptée ou non, ou 0 si le temps n'est pas défini.
     */
    public getDate(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj({ adaptTime });
        return dateObj?.value ?? 0;
    }

    /**
     * Retourne l'année de la date, adaptée ou non.
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir l'année
     *  de la date adaptée ou non.
     * @returns {number} - L'année de la date, adaptée ou non, ou 0 si le temps n'est pas défini.
     */
    public getYear(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj({ adaptTime });
        return dateObj?.year ?? 0;
    }

    /**
     * Retourne le mois de la date, adapté ou non.
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir le mois
     *  de la date adaptée ou non.
     * @returns {number} - Le mois de la date, adapté ou non, ou 0 si le temps n'est pas défini.
     */
    public getMonth(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj({ adaptTime });
        return dateObj?.month ?? 0;
    }
 
    /**
     * Retourne le jour du mois de la date, adapté ou non.
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir le jour du mois
     *  de la date adaptée ou non.
     * @returns {number} - Le jour du mois de la date, adapté ou non, ou 0 si le temps n'est pas défini.
     */
    public getDay(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): number {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj({ adaptTime });
        return dateObj?.day ?? 0;
    }
 
    /**
     * Retourne le jour de la semaine correspondant à la date, adapté ou non.
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir le jour de la semaine
     *  de la date adaptée ou non.
     * @param {boolean} [withHolidays=true] - Indique si l'on souhaite indiquer les jours fériés (jour 8).
     * @returns {Day | undefined} - Le jour de la semaine correspondant à la date,
     *  adapté ou non, ou undefined si le temps n'est pas défini.
     */
    public getDayOfWeek(
        {
            adaptTime = true,
            withHolidays = true
        }: {
            adaptTime?: boolean,
            withHolidays?: boolean
        } = {}
    ): Day | undefined {
        if (!this._computed) this.compute();
        const dateObj = this.getDateObj({ adaptTime });
        return dateObj?.isHoliday && withHolidays ? Day.HOLIDAY : dateObj?.dayOfWeek;
    }

    /**
     * Retourne l'heure correspondant au temps, adaptée ou non.
     * Si le temps est relatif, renvoie l'heure relative.
     * Si le temps est absolu (daté ou non), renvoie la fraction d'une journée (tronquée modulo 1).
     * Si le temps est adapté, il est incrémenté de 1 (ex : 25h00).
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir l'heure adaptée ou non.
     * @returns {number}- L'heure correspondant au temps, adaptée ou non, ou 0 si le temps n'est pas défini.
     */
    public getTime(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        if (!timeObj) return 0;
        if (!this.isRelative && !adaptTime) {
            return timeObj.value % 1;
        }
        return timeObj.value;
    }
 
    /**
     * Retourne le nombre d'heures de l'heure correspondant au temps, adapté ou non.
     * Si le temps est relatif, renvoie le nombre d'heures relatif.
     * Si le temps est absolu et que adaptTime est faux,
     *  renvoie le nombre d'heures de l'objet DateTime tronqué modulo 24.
     * Si le temps est adapté, l'heure est incrémentée de 24 (ex : 25h00). 
     * @param {boolean} [adaptTime=false] - Indique si l'on souhaite avoir
     *  le nombre d'heures adapté ou non.
     * @returns {number} - Le nombre d'heures correspondant au temps, adapté ou non,
     *  ou 0 si le temps n'est pas défini.
     */
    public getHours(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): number {
        if (!this._computed) this.compute();
        const timeObj = this._time;
        if (!timeObj) return 0;
        if (!this.isRelative && !adaptTime) {
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

    public hasDate(): boolean {
        return this.excelValue > DateTime.MIN_EXCEL_DATE;
    }

    /**
     * Retourne si la date correspondant au temps est un jour férié.
     * @param {boolean} [adaptTime=true] - Indique si l'on souhaite avoir la date adaptée ou non.
     * @returns {boolean} - Vrai si la date est un jour férié, faux sinon ou si la date n'est pas définie.
     */
    public isHoliday(
        { adaptTime = true }: { adaptTime?: boolean } = {}
    ): boolean {
        if (!this._computed) this.compute();
        return this.getDateObj({ adaptTime })?.isHoliday ?? false;
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
        {
            isRelative = false,
            adaptTime = true
        }: {
            isRelative?: boolean,
            adaptTime?: boolean
        } = {}
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

        return new DateTime(result, { isRelative, adaptTime });
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
        if (!this.isRelative && this.hasDate()) {
            this._realDate = new ExcelDate(this.excelValue);
            timeOfDays = this.excelValue % 1;
            if (timeOfDays < Params.get('rolloverHour').excelValue) {
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
     * Retourne un nouvel objet DateTime égal au temps courant
     *  résolu par rapport à une référence.
     * Si le temps courant est relatif, il est ajouté à la référence.
     * Sinon, le temps courant est renvoyé tel quel.
     * @param {DateTime} reference - Référence à utiliser pour résoudre le temps courant.
     * @returns {DateTime} - Nouvel objet DateTime égal au temps courant résolu par rapport à la référence.
     */
    public resolveAgainst(
        reference: DateTime
    ): DateTime {
        if (this.isRelative) {
            return new DateTime(
                reference.excelValue + this.excelValue,
                { isRelative: reference.isRelative }
            );
        }

        return this;
    }

    /**
     * Retourne un nouvel objet DateTime égal au temps courant relatif par rapport à une référence.
     * Les deux temps doivent être absolus.
     * @param {DateTime} reference - Référence à utiliser pour résoudre le temps courant.
     * @returns {DateTime} - Nouvel objet DateTime égal au temps courant relatif par rapport à la référence.
     * @throws {Error} - Si l'un des deux temps est relatif.
     */
    public relativeTo(
        reference: DateTime
    ): DateTime {
        if (this.isRelative || reference.isRelative) {
            throw new Error(`Les deux temps doivent être absolus`);
        }

        return new DateTime(
            this.excelValue - reference.excelValue,
            { isRelative: true }
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
    public compareTo(
        other: DateTime
    ): number {

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
     * @param {Nullable<DateTime>} other - Temps à comparer.
     * @returns {boolean} - Vrai si les deux temps sont égaux, faux sinon.
     */
    public equalsTo(
        other: Nullable<DateTime>
    ): boolean {
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
    public add(
        other: DateTime
    ): DateTime {
        if (!this.isRelative || !other.isRelative) {
            throw new Error(`L'addition n'est possible qu'entre temps relatifs`);
        }

        return new DateTime(
            this.excelValue + other.excelValue,
            { isRelative: true }
        );
    }

    /**
     * Soustrait au temps relatif un autre temps relatif.
     * @param {DateTime} other - Temps relatif à soustraire.
     * @returns {DateTime} - Nouvel objet DateTime égal à la différence entre les deux temps relatifs.
     * @throws {Error} - Si l'un des deux temps n'est pas relatif.
     */
    public subtract(
        other: DateTime
    ): DateTime {
        if (!this.isRelative || !other.isRelative) {
            throw new Error(`La soustraction n'est possible qu'entre temps relatifs`);
        }

        return new DateTime(
            this.excelValue - other.excelValue,
            { isRelative: true }
        );
    }

    /**
     * Elimine la partie date du format de la date et de l'heure :
     *  - si un élément de date est trouvé avant un élément de l'heure,
     *  toute la partie précédant le premier élément de l'heure est supprimée.
     * - si un élément de date est trouvé après un élément de l'heure,
     *  toute la partie suivant le dernier élément de l'heure est supprimée.
     * @param {string} format - Format de la date et de l'heure
     * @returns {string} - Nouveau format de la date et de l'heure
     */
    public static removeDatePartFormat(format: string): string {

        const ymdRegex = /y|m|d/g;
        const hnsRegex = /h|n|s/g;
        const regexFormat = format
            .replace(ymdRegex, "y")
            .replace(hnsRegex, "h");

        const firstYmdIndex = regexFormat.indexOf("y");
        const lastYmdIndex = regexFormat.lastIndexOf("y");
        const firstHnsIndex = regexFormat.indexOf("h");
        const lastHnsIndex = regexFormat.lastIndexOf("h");

        // Si pas d'élément de date, renvoie le format en entier.
        if (firstYmdIndex === -1) return format;
        // Si pas d'élément d'heure, renvoie une chaîne vide.
        if (firstHnsIndex === -1) return "";

        let newFormat = format;
        // Si un élément date est rencontré après un élément de l'heure,
        //  la partie suivant l'heure est enlevée.
        if (lastYmdIndex > lastHnsIndex) {
            newFormat = newFormat.slice(0, lastHnsIndex + 1);
        }
        // Si un élément de date est trouvé avant un élément de l'heure,
        //  la partie précédent l'heure est enlevée.
        if (firstYmdIndex < firstHnsIndex) {
            newFormat = newFormat.slice(firstHnsIndex);
        }
    
        return newFormat;
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
    public format(
        format: string,
        {
            adaptTime = true,
            withHolidays = true
        }: {
            adaptTime?: boolean,
            withHolidays?: boolean
        } = {}
    ): string {
    
        const key = `${format}|${adaptTime ? 1 : 0}|${withHolidays ? 1 : 0}`;
    
        const cached = this._formats.get(key);
        if (cached !== undefined) return cached;
    
        this.compute();

        // Si pas de date, enlève les éléments de la date dans le format
        const newFormat = this.hasDate() ? format : DateTime.removeDatePartFormat(format);

        const prefix = this.excelValue < 0 ? "-" : "";
        const pad = (v: number) => v.toString().padStart(2, "0");

        const tokens: Record<string, string> = {
            // Année
            yyyy: !this.hasDate() ? '' : this.getYear({ adaptTime }).toString(),
            yy: !this.hasDate() ? '' : pad(this.getYear({ adaptTime }) % 100),
            // Mois
            mmmm: !this.hasDate() ? '' : DateTime.MONTHS[this.getMonth({ adaptTime }) - 1].fullName,
            mmm: !this.hasDate() ? '' : DateTime.MONTHS[this.getMonth({ adaptTime }) - 1].abbreviation,
            mm: !this.hasDate() ? '' : pad(this.getMonth({ adaptTime })),
            m: !this.hasDate() ? '' : this.getMonth({ adaptTime }).toString(),
            // Jour
            dd: !this.hasDate() ? '' : pad(this.getDay({ adaptTime })),
            d: !this.hasDate() ? '' : this.getDay({ adaptTime }).toString(),
            // Jour de la semaine
            dddd: !this.hasDate() ? '' : this.getDayOfWeek({ adaptTime, withHolidays })?.fullName ?? "",
            ddd: !this.hasDate() ? '' : this.getDayOfWeek({ adaptTime, withHolidays })?.abbreviation ?? "",
            // Heure
            hh: pad(this.getHours({ adaptTime })),
            h: this.getHours({ adaptTime }).toString(),
            // Minute
            nn: pad(this.getMinutes()),
            n: this.getMinutes().toString(),
            // Seconde
            ss: pad(this.getSeconds()),
            s: this.getSeconds().toString(),
        };
    
        const result = prefix + newFormat
            .toLowerCase()
            .replace(DateTime.FORMAT_REGEX, t => tokens[t]);
    
        this._formats.set(key, result);
    
        return result;
    }

    /**
     * Ajuste une heure pour tenir compte du changement de journée.
     * Si l'heure est inférieure à l'heure de changement de journée,
     *  ajoute 1 pour passer à la journée suivante.
     *  Par exemple : 01:00 → 25:00 si changement de journée à 03:00
     * Cela ne s'applique que sur les heures non datées (valeur < 1).
     * @param {number} time - Heure à ajuster.
     * @returns {number} - Heure ajustée.
     */
    public static adaptTime(
        time: number
    ): number {
        return (time < Params.get('rolloverHour').excelValue) ? time + 1 : time;
    }
}

/**
 * Type DayInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Day.
 */
type DayInput = Input<Day, number | string>;

class Day { 

    // Indicateur de chargement
    private static loaded = false;

    // Liste des jours de la semaine identifiés par leur indice
    private static list: Day[] = new Array(8);

    // Propriétés de la classe Day
    public readonly mask: number;           // Masque de bits contenant les numéro du jour de la semaine
                                            //  (identique à Days)
                                            //  (de lundi > bit 0 à dimanche > bit 6, férié > bit 7)
    public readonly code: string;           // Code chiffre du jour de la semaine
    public readonly fullName: string;       // Nom du jour de la semaine
    public readonly abbreviation: string;   // Abréviation du jour de la semaine

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
        this.abbreviation = days.abbreviation;
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
     * @returns {string} - Nom complet du jour de la semaine.
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
     * Retourne une instance Day correspondant au numéro de jour fourni.
     * Si le numéro de jour n'existe pas, renvoie undefined.
     * @param {Nullable<DayInput>} value - Valeur à analyser
     *  pour le jour correspondant.
     * @returns {Days | undefined} - Instance de Day correspondante.
     */
    public static from(
        value: Nullable<DayInput>
    ): Day | undefined {
 
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
    private static maskToIndex(
        mask: number
    ): number {
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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

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
        Log.timer(`${this.name}.load()`);
    }
}

/**
 * Type DaysInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Days.
 */
type DaysInput = Input<Days, number | string>;
/**
 * Classe utilitaire Days pour la gestion des jours de la semaine, individuels ou groupés. 
 *  (JOB du lundi au vendredi, WE pour samedi et dimanche...).
 */
class Days {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Paramètres";        // Feuille contenant les paramètres des jours de la semaine
    private static readonly TABLE = "Jours";        // Tableau contenant les paramètres des jours de la semaine
    private static readonly COL_NUMBERS = 0;        // Colonne contenant le numéro du jour 
    private static readonly COL_CODE_LETTER = 1;    // Colonne contenant la lettre de code du groupe de jours
    private static readonly COL_FULL_NAME = 2;      // Colonne contenant le nom complet du jour de la semaine
    private static readonly COL_ABBREVIATION = 3;   // Colonne contenant l'abréviation du jour de la semaine

    // Valeur par défaut des jours de la semaine individuels
    //  (si non renseignés dans le tableau des paramètres)
    private static readonly WEEKDAYS = [
        { number: 1, fullName: "Lundi",    abbreviation: "Lu" },
        { number: 2, fullName: "Mardi",    abbreviation: "Ma" },
        { number: 3, fullName: "Mercredi", abbreviation: "Me" },
        { number: 4, fullName: "Jeudi",    abbreviation: "Je" },
        { number: 5, fullName: "Vendredi", abbreviation: "Ve" },
        { number: 6, fullName: "Samedi",   abbreviation: "Sa" },
        { number: 7, fullName: "Dimanche", abbreviation: "Di" },
        { number: 8, fullName: "Férié",    abbreviation: "Fer", code: "F" }
    ];

    // Indicateur de chargement
    private static loaded = false;

    // Liste des groupes de jours de la semaine identifiés par leur masque de bits,
    //  en assurant l'unicité de chaque groupe de jours
    private static masksList: (Days | undefined)[] = new Array(256);
    // Map des groupes de jours de la semaine identifiés par leur code, nom ou abréviation,
    //  (plusieurs codes possibles par groupe de jour, y compris ceux non optimisés)
    private static mapByString: Map<string, Days> = new Map();
 
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

    // Propriétés de la classe Days
    public readonly mask: number;           // Masque de bits contenant les numéro(s) du ou des jours
                                            //  du groupe de jours de la semaine
                                            //  (de 1 : lundi > bit 0 à 7 : dimanche > bit 6, 8 : férié > bit 7)
    public readonly code: string;           // Code alphanumérique du groupe de jours de la semaine
    public readonly fullName: string;       // Nom du jour ou du groupe de jours de la semaine
    public readonly abbreviation: string;   // Abréviation du jour ou du groupe de jours de la semaine
    private _numberString?: string;         // Chaîne de caractères contenant les numéros des jours du groupe de jours
    private _count: number = -1;            // Nombre de jours contenus dans le groupe de jours (-1 si non calculé)

    /**
     * Constructeur privé de la classe Days.
     * @param {number} mask - Masque de bits des numéro(s) des jours du groupe de jours de la semaine. 
     * @param {string} code - Code alphanumérique du groupe de jours de la semaine.
     * @param {string} fullName - Nom complet du jour ou du groupe de jours de la semaine.
     * @param {string} abbreviation - Abréviation du jour ou du groupe de jours de la semaine.
     */
    private constructor(
        {
            mask,
            code,
            fullName = "",
            abbreviation = ""
        }: {
            mask: number,
            code: string,
            fullName?: string,
            abbreviation?: string
        }
    ) {
        this.mask = mask;
        this.code = code;
        this.fullName = fullName;
        this.abbreviation = abbreviation;
    }

    /**
     * Retourne la concaténation des numéros des jours du groupe de jours.
     * @returns {string} Chaîne de caractères contenant les numéros des jours de la semaine.
     */
    public get numbersString(): string {
        if (this._numberString === undefined) {
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
     * @returns {string} - Nom complet du jour ou du groupe de jours de la semaine.
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
     * Retourne une instance Days correspondant au numéro de jour fourni.
     * Si le numéro de jour n'existe pas, renvoie undefined.
     * Charge les paramètres des jours de la semaine si ce n'est pas déjà fait.
     * @param {Nullable<DaysInput>} value - Valeur à analyser
     *  pour le groupe de jours correspondant.
     * @returns {Days | undefined} - Instance de Days correspondante.
     */
    public static from(
        value: Nullable<DaysInput>
    ): Days | undefined {
 
        if (value == null || value === '') return undefined;
        if (value instanceof Days) return value;

        const code = String(value);
        const existing = this.get(code);
        if (existing) return existing;

        // Analyse de la chaine de caractères pour trouver le ou les jours correspondants.
        const numbers = this.extractFromString(code);
        if (numbers.length === 0) {
            return undefined;
        }
        const mask = this.numbersToMask(numbers);

        // Recherche ou création du groupe de jours.
        const days = this.get(mask) ?? this.create({ numbers });

        // Ajout du code fourni en entrée pour éviter une nouvelle analyse,
        //  y compris si ce code n'est pas optimisé (code optimisé calculé dans create()).
        this.mapByString.set(code, days);

        return days;
    }

    /**
     * Retourne une instance Days correspondant au masque de bits fourni.
     * Crée le groupe de jours si celui-ci n'existe pas encore.
     * @param {number} value - Masque de bits correspondant au groupe de jours.
     * @returns {Days | undefined} - Instance de Days correspondante.
     */
    public static fromMask(
        mask: number
    ): Days | undefined{
        if (mask <= 0 || mask >= 256) return undefined
        return this.masksList[mask] ?? this.create({ numbers: this.maskToNumbers(mask) });
    }
 
    /**
     * Transforme un tableau de numéros de jours en un masque de bits.
     * Chaque numéro correspond à un bit dans le masque, allant de 1 (lundi) à 7 (dimanche), et 8 (férié).
     * @param {number[]} numbers - Tableau de numéros de jours.
     * @returns {number} - Masque de bits correspondant aux numéros de jours.
     */
    private static numbersToMask(
        numbers: number[]
    ): number {
        return numbers.reduce((mask, n) => mask | (1 << (n - 1)), 0);
    }

    /**
     * Retourne un tableau de numéros de jours à partir d'un masque de bits.
     * Chaque bit dans le masque correspond à un numéro de jour,
     *  allant de 1 (lundi) à 7 (dimanche), et 8 (férié).
     * @param {number} mask - Masque de bits correspondant aux numéros de jours.
     * @returns {number[]} - Tableau de numéros de jours correspondant au masque de bits.
     */
    private static maskToNumbers(
        mask: number
    ): number[] {
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
     * @param {DayInput} day - Jour de la semaine (1 : lundi, 2 : mardi, ..., 7 : dimanche, 8 : férie)
     * @returns {boolean} - Vrai si le groupe de jours contient le jour, faux sinon.
     */
    public contains(
        day: DayInput
    ): boolean {
        const dayObj = Day.from(day);
        return (this.mask & dayObj!.mask) !== 0;
    }

    /**
     * Retourne true si le groupe de jours contient au moins un jour commun
     *  avec le groupe de jours fourni en paramètre.
     * @param {Days} other - Autre groupe de jours.
     * @returns {boolean} - Vrai si le groupe de jours contient au moins un jour commun, faux sinon.
     */
    public intersects(
        other: Days
    ): boolean {
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
    public static difference(
        days1: Days,
        days2: Days
    ): Days | undefined {
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
    private static cleanAndSortNumbers(
        numbersString: string
    ): number[] {
        return Array.from(new Set(
            numbersString
                .replace(/[^1-8]/g, '')     // Supprime les caractères non numériques
                                            //  et non compris entre 1 et 8
                .split('')                  // Divise la chaîne en un tableau de chiffres
                .map(x => Number(x))      // Convertit les caractères en nombres
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
    public static extractFromString(
        value: string
    ): number[] {

        let processed = String(value).toUpperCase();

        // Analyse avec Regex des chaines servant pour l'extraction.
        for (const s of this.extractionPatterns) {
            const regex = new RegExp(s.pattern, 'g');
            processed = processed.replace(regex, s.numbersString);
        }

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
    private static optimiseMask(
        mask: number
    ): Days[] {

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
    public static has(
        value: string | number
    ): boolean {
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
    public static get(
        value: string | number
    ): Days | undefined {
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
     * @param {string} abbreviation - Abréviation du groupe de jours.
     * @returns {Days} - Nouveau groupe de jours créé.
     * @throws {Error} - Si le groupe de jours est déjà présent dans la base de données.
     */
    private static create(
        {
            numbers,
            code = "",
            fullName = "",
            abbreviation = ""
        }: {
            numbers: number[],
            code?: string,
            fullName?: string,
            abbreviation?: string
        }
    ): Days {

        if (numbers.length === 0) {
            throw new Error(`Le groupe de jours ${code} ne contient pas de jours.`);
        }
 
        // Vérifie que le groupe de jour n'existe pas déjà.
        const mask = this.numbersToMask(numbers);
        const existing = this.get(mask);
        if (existing) return existing;

        // Si toutes les valeurs sont fournies (chargement depuis les paramètres),
        //  récupère toutes les valeurs.
        let finalCode = code;
        let finalFullName = fullName;
        let finalAbbreviation = abbreviation;

        // Si seul les numéros de jours sont fournis,
        //  un code optimisé est calculé par assemblage de groupes de jours connus.
        //  Cet assemblage définit également un nom complet et une abréviation par concaténation.
        if (!code) {
            const parts = this.optimiseMask(mask);
            finalCode = parts.map(d => d.code).join('');
            finalFullName = parts.map(d => d.fullName).join(' + ');
            finalAbbreviation = parts.map(d => d.abbreviation).join('');
        }

        // Instancie le nouveau groupe de jours.
        const days = new Days({
            mask,
            code: finalCode,
            fullName: finalFullName,
            abbreviation: finalAbbreviation
        });

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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);
        
        // Vérifie si la table à charger existe déjà.
        if (this.loaded) {
            if (!erase) return;
            this.clear();
        }

        // Charge la base de données.
        const data = WorkbookServices.getDataFromTable({ sheetName: this.SHEET, tableName: this.TABLE });

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
                const fullName = WorkbookServices.getRequiredString(
                    row,
                    this.COL_FULL_NAME,
                    { errorMessage : `Nom complet du groupe de jours`
                        + ` non renseigné dans le tableau des jours.` }
                );
                if (this.has(fullName)) {
                    throw new Error(`Nom complet ${fullName} déjà utilisé.`);
                }
                const abbreviation = WorkbookServices.getRequiredString(
                    row,
                    this.COL_ABBREVIATION,
                    { errorMessage : `Groupe de jours du ${fullName} :`
                        + ` abréviation non renseignée dans le tableau des jours.` }
                );
                if (this.has(abbreviation)) {
                    throw new Error(`Groupe de jours du ${fullName} :`
                        + ` abbreviation ${abbreviation} déjà utilisée.`);
                }
                const numbersString = WorkbookServices.getString(row, this.COL_NUMBERS);
                const numbers = this.cleanAndSortNumbers(numbersString);
                if (numbers.length === 0) {
                    throw new Error(`Groupe de jours du ${fullName} :`
                        + ` numéros des jours non renseignés ou invalides dans le tableau des jours.`);
                }
                const codeLetter = WorkbookServices.getString(row, this.COL_CODE_LETTER)
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
                const days = this.create({
                    numbers,
                    code,
                    fullName,
                    abbreviation
                });
            }

        } catch (e) {
            throw new Error(`Days.load (ligne ${excelRow}) : ${e}`);
        } 

        // Si non renseignés dans le tableau, charge les jours individuels par défaut.
        for (const d of this.WEEKDAYS) {
            const numbersString = String(d.number);
            if (!this.has(numbersString)) {
                this.create({
                    numbers: [d.number],
                    code: d.code ?? numbersString,
                    fullName: d.fullName,
                    abbreviation: d.abbreviation
                });
            }
        };

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
                pattern: days.abbreviation.toUpperCase(),
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
        Log.startTimer(`${this.name}.load()`);

        // Charge les jours individuels.
        Day.load();
    }
}

/*
 * Classe utilitaire DaysValues pour la gestion des valeurs associées aux jours.
 */
class DaysValues {

    // Propriétés de la classe Days
 
    public readonly days: Days;                             // Groupe de jours total concerné
    private _entries: { days: Days, value: string }[] = []; // Liste de règles (Days → valeur)
    private _toString? : string;                            // Cache de la représentation textuelle de l'objet

    /**
     * Constructeur de la classe DaysValues.
     * @param {Days} days - Groupe de jours total concerné.
     */
    public constructor(
        days: Days
    ) {
        this.days = days;
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Chaine au format "jour1: valeur1, jour2: valeur2, ..."
     *  où chaque jour est représenté par un ensemble de numéros de jours
     *  (ex: "1-2,4,6" pour les jours lundi, mardi, jeudi et samedi).
     */
    public toString(): string {
        
        if (this._toString === undefined) {
            const entries = this._entries;
            this._toString =
                entries.length === 0 ? ""
                : entries.length === 1 && entries[0].days.mask === this.days.mask
                    ? entries[0].value
                    : entries.map(e => `${e.days.numbersString}: ${e.value}`).join(", ");
        }
        return this._toString;
    }

    /**
     * Crée une instance DaysValues à partir d'un groupe de jours et d'une chaîne de valeurs.
     * La chaîne de valeurs peut contenir des valeurs uniques ou multiples, séparées par des virgules.
     * Si la chaîne de valeurs contient une seule valeur, le groupe de jours est considéré comme un seul jour.
     * Si la chaîne de valeurs contient plusieurs valeurs, chaque valeur est associée à un groupe de jours.
     * @param {Days} days - Groupe de jours total concerné.
     * @param {string} input - Chaîne de valeurs.
     * @returns {DaysValues} - L'objet DaysValues créé.
     */ 
    public static from(
        days: Days,
        input: string
    ): DaysValues {

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
     * Retourne la valeur associée au jour de la semaine donné.
     * Si le jour n'est pas couvert par une règle, renvoie une chaîne vide.
     * @param {Day} day - Jour de la semaine.
     * @returns {string} - Valeur associée.
     */
    public get(
        day: Day
    ): string {

        for (const entry of this._entries) {
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
    public set(
        days: Days,
        value: string
    ): void {

        // Restreint le groupe de jours à affecter au périmètre autorisé
        const validDays = Days.intersection(days, this.days);
        if (!validDays || validDays.mask === 0) return;

        // Supprime les parties déjà couvertes dans les autres sous-groupes de jours (valeur modifiée)
        this._entries = this._entries
            .map(e => {
                const remaining = Days.difference(e.days, validDays);
                return remaining ? { days: remaining, value: e.value } : null;
            })
            .filter(e => e !== null)

        // Ajoute la nouvelle valeur
        this._entries.push({ days: validDays, value });

        this.merge();
        this._toString = undefined
    }

    /**
     * Fusionne les valeurs associées aux groupes de jours.
     * Regroupe par valeur, puis reconstruit les entrées avec les valeurs fusionnées.
     */
    private merge(): void {

        const map: Map<string, number> = new Map();

        // Regroupe par valeur
        for (const e of this._entries) {
            const prev = map.get(e.value) ?? 0;
            map.set(e.value, prev | e.days.mask);
        }

        // Reconstruit _entries
        this._entries = [];
        for (const [value, mask] of Array.from(map.entries())) {
            this._entries.push({
                days: Days.get(mask)!,
                value
            });
        }
        this._entries.sort((a, b) => 
            a.days.numbersString.localeCompare(b.days.numbersString));
    }

    /**
     * Vérifie si toutes les parties du groupe de jours sont couvertes.
     * @returns {boolean} - Vrai si toutes les parties sont couvertes, faux sinon.
     */
    public isComplete(): boolean {

        let mask = 0;
        for (const e of this._entries) {
            mask |= e.days.mask;
        }

        return mask === this.days.mask;
    }

    /**
     * Ajoute les valeurs manquantes dans le groupe de jours.
     * Pour chaque jour non couvert par une règle, ajoute une entrée avec la valeur par défaut.
     * @param {string} defaultValue - Valeur associée aux jours non couverts.
     */
    public fillGaps(
        defaultValue: string
    ): void {

        let covered = 0;
        for (const e of this._entries) covered |= e.days.mask;
        const missingMask = this.days.mask & ~covered;

        if (missingMask) {
            this._entries.push({
                days: Days.fromMask(missingMask)!,
                value: defaultValue
            });
            this.merge();
            this._toString = undefined;
        }
    }

    /**
     * Vérifie si les deux objets DaysValues sont égaux.
     * La comparaison est faite en fonction de la représentation textuelle des objets.
     * @param {DaysValues} other - L'objet DaysValues à comparer.
     * @returns {boolean} - Vrai si les deux objets sont égaux, faux sinon.
     */
    public equals(
        other: DaysValues
    ): boolean {
        return this.toString() === other.toString();
    }
}

/**
 * Type ParityInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Parity.
 */
type ParityInput = Input<Parity, string | number>;

/*
 * Classe utilitaire Parity immuable qui permet de manipuler la parité
 *  d'un train, d'un parcours ou d'un arrêt.
 */
class Parity {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Paramètres";        // Feuille contenant les paramètres de parité
    private static readonly TABLE = "Parité";       // Tableau contenant les paramètres de parité
    private static readonly START_CELL = "A1";              // Première cellule de la liste des gares
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

    // Tableau des parités possibles (pool[parity][doubleAllowed])
    private static pool: {
        [key: number]: {
            standard: Parity;
            double: Parity;
        };
    } = {};

    // Map des lettres et nombres désignants les parités
    private static letters: Map<number, string> = new Map();
    private static digits: Map<number, number> = new Map();
 
    // Indicateur de chargement
    private static loaded = false;

    // Propriétés de la classe Parity
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
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
    ) {
        this.value = value;
        this.doubleParityAllowed = doubleParityAllowed;
    }
 
    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Parité sous forme de texte.
     */
    public toString(): string {
        return this.value.toString();
    }

    /**
     * Retourne l'instance de Parity qui correspond à la valeur de parité spécifiée.
     * @param {number} [value] - Valeur de parité.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la double parité est autorisée.
     *  Si faux (par défaut), la double parité est impossible.
     * @returns {Parity} - Instance de Parity correspondante.
     */
    private static get(
        value: number,
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
    ): Parity {
        return this.pool[value]?.[
            doubleParityAllowed ? "double" : "standard"
        ] ?? this.pool[this.UNDEFINED][
            doubleParityAllowed ? "double" : "standard"
        ];
    }

    /**
     * Retourne une instance de Parity à partir d'une valeur qui peut être :
     *  - une lettre de parité (ou la concaténation des deux lettres sans ordre si double parité),
     *  - un chiffre de parité (format chaîne ou nombre),
     *  - un numéro de train (pair, impair ou double s'il contient un '/'),
     *  - une instance de Parity (retourne la même instance),
     *  - null ou undefined (retourne une instance de Parity avec valeur this.UNDEFINED).
     * @param {Nullable<ParityInput>} value - Valeur à analyser pour la parité.
     * @returns {Parity} - Instance de Parity correspondante.
     */
    public static from(
        value: Nullable<ParityInput>,
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
    ): Parity {
        if (value instanceof Parity) {
            if (value.doubleParityAllowed === doubleParityAllowed) return value;
            return this.get(value.value, { doubleParityAllowed });
        }
 
        const normalized = this.normalize(value, { doubleParityAllowed });
        return this.get(normalized, { doubleParityAllowed });
    } 

    /**
     * Normalise en une valeur de parité une valeur, qui peut être :
     *  - la lettre de parité (ou la concaténation des deux lettres sans ordre si double parité),
     *  - le chiffre de parité (format chaîne ou nombre),
     *  - un numéro de train (pair, impair ou double s'il contient un '/').
     * @param {Nullable<string | number>} value - Valeur à normaliser.
     * @param {boolean} doubleParityAllowed - Si vrai, la double parité est autorisée.
     * @returns {number} - Valeur de parité normalisée.
     */
    private static normalize(
        value: Nullable<string | number>,
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
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
            return this.normalize(numeric, { doubleParityAllowed });
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
     * Retourne l'objet Parité qui n'a pas de valeur définie.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} - Parité sans valeur définie.
     */
    public static undefined(
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
    ): Parity {
        return this.get(this.UNDEFINED, { doubleParityAllowed });
    }

    /**
     * Retourne l'objet Parité qui correspond à une parité impaire.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} - Parité impaire.
     */
    public static odd(
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
    ): Parity {
        return this.get(this.ODD, { doubleParityAllowed });
    }

    /**
     * Retourne l'objet Parité qui correspond à une parité paire.
     * @param {boolean} [doubleParityAllowed=false] - Si vrai, la parité
     *  accepte les parités doubles, sinon elle les refuse.
     * @returns {Parity} - Parité paire.
     */
    public static even(
        { doubleParityAllowed = false }: { doubleParityAllowed?: boolean } = {}
    ): Parity {
        return this.get(this.EVEN, { doubleParityAllowed });
    }

    /**
     * Retourne l'objet Parité qui correspond à une parité double.
     * @returns {Parity} - Parité double.
     */
    public static double(): Parity {
        return this.get(this.DOUBLE, { doubleParityAllowed: true });
    }

    /**
     * Vérifie si la parité est identique à une valeur de parité donnée.
     * @param {number} parity - Autre valeur de parité à comparer.
     * @returns {boolean} - Vrai si les deux parités sont identiques, faux sinon.
     */
    public is(
        parity: number
    ): boolean {
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
    public isOpposedTo(
        other: Parity | undefined
    ): boolean {
        return !!other
            && (this.value === Parity.ODD && other.value === Parity.EVEN
                || this.value === Parity.EVEN && other.value === Parity.ODD);
    }

    /**
     * Vérifie si la parité est définie et identique à celle d'une autre variable de parité.
     * @param {Parity | undefined} other - Autre variable de parité à comparer.
     * @returns {boolean} - Vrai si les deux parités sont identiques, faux sinon.
     */
    public equalsTo(
        other: Parity | undefined
    ): boolean {
        return this === other;
    }

    /**
     * Vérifie si la parité inclut une autre valeur de parité.
     * @param {Nullable<ParityInput>} other - Autre valeur de parité à inclure.
     * @returns {boolean} - Vrai si la parité inclut la valeur de parité, faux sinon.
     */
    public includes(
        other: Nullable<ParityInput>
    ): boolean {
        const requested = Parity.from(other, { doubleParityAllowed: this.doubleParityAllowed });
 
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
                return Parity.even({ doubleParityAllowed: this.doubleParityAllowed });
            case Parity.EVEN:
                return Parity.odd({ doubleParityAllowed: this.doubleParityAllowed });
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
     * Si la parité de départ n'est pas définie, utilise la parité fournie en paramètre.
     * Si la parité fournie en paramètre n'est pas définie, utilise la parité de départ.
     * Si les deux parités sont identiques, retourne la parité de départ.
     * Sinon, combine ces deux parités en une parité double.
     * @param {Parity} other - Parité à combiner avec la parité actuelle.
     * @returns {Parity} - Parité combinée.
     */
    public combineWith(
        other: Parity
    ): Parity {

        if (!this.doubleParityAllowed) throw new Error(`Il n'est pas possible de combiner une`
            + ` parité à une autre si celle de départ n'autorise pas les parités doubles.`
            + ` Le résultat est forcément une parité qui accepte les parités doubles.`);
 
        if (!this.isDefined()) {
            return Parity.from(other.value, { doubleParityAllowed: true });
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
    public static digit(
        parity: number
    ): number {
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
    public static letter(
        parity: number
    ): string {
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
    public static containsParityLetter(
        string: string,
        parity: number
    ): boolean {
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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        // Vérifie si la table à charger existe déjà.
        if (this.loaded) {
            if (!erase) return;
            this.letters.clear();
            this.digits.clear();
        }

        // Crée le tableau des parités possibles
        for (const parity of [this.UNDEFINED, this.ODD, this.EVEN, this.DOUBLE]) {
            this.pool[parity] = {
                standard: parity === this.DOUBLE
                    ? this.pool[this.UNDEFINED].standard
                    : new Parity(parity, { doubleParityAllowed: false}),
            
                double: new Parity(parity, { doubleParityAllowed: true})
            };
        }

        // Charge la base de données.
        const data = WorkbookServices.getDataFromTable({ sheetName: this.SHEET, tableName: this.TABLE });

        const getParityLetter = (
            row: number,
            fallback: string
        ): string =>
            WorkbookServices.getString(data[row], this.COL_LETTER, { defaultValue: fallback })
                .toUpperCase();

        this.letters.set(this.ODD, getParityLetter(this.ROW_ODD, "I"));
        this.letters.set(this.EVEN, getParityLetter(this.ROW_EVEN, "P"));

        const getParityDigit = (
            row: number,
            fallback: number
        ): number =>
            WorkbookServices.getNumber(data[row], this.COL_NUMBER, { defaultValue: fallback });

        this.digits.set(this.ODD, getParityDigit(this.ROW_ODD, 1));
        this.digits.set(this.EVEN, getParityDigit(this.ROW_EVEN, 2));
        this.digits.set(this.DOUBLE, getParityDigit(this.ROW_DOUBLE, -2));

        this.loaded = true;
        Log.timer(`${this.name}.load()`);
    }
}

/**
 * Type TrainNumberParity représentant la parité paire ou impaire
 *  d'un numéro de train.
 */
type TrainNumberParity = typeof Parity.EVEN | typeof Parity.ODD;

/**
 * Type TrainNumberInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet TrainNumber.
 */
type TrainNumberInput = Input<TrainNumber, string | number>;

/**
 * Classe TrainNumber définissant un numéro de train.
 * Il est alphanumérique, sans ponctuation et sans espaces, avec un chiffre pour dernier caractère.
 * La double parité est marquée par ######/#.
 */
class TrainNumber {

    // Constantes de lecture des paramètres sur Excel
    private static readonly SHEET = "Paramètres";                // Nom de la feuille
    private static readonly TABLE = "TrainNumberRegex";     // Nom du tableau
    private static readonly START_CELL = "D21";             // Première cellule
    private static readonly REGEX_COLUMNS = {               // Liste des colonnes avec leur emplacement
        commercial: 0,
        emptyPassenger: 1,
        mouvement: 2,
        abbreviate4: 3
    } as const;
    private static regex: {                                 // Liste des regex
        commercial: RegExp;                                 //  - trains commerciaux
        emptyPassenger: RegExp;                             //  - trains W
        mouvement: RegExp;                                  //  - évolutions
        abbreviate4: RegExp;                                //  - trains abrégeables à 4 chiffres
    };           
    private static loaded = false;                          // Indicateur de chargement

    // Etat courant du numéro de train
    private _parity: TrainNumberParity;             // Parité initiale du numéro du train,
                                                    //  convention celle au départ du train
                                                    //  (ODD ou EVEN uniquement)
    private _doubleParity: boolean;                 // Indique si le train circule sous double parité

    // Cache immutable des variantes
    private readonly variants: Record<TrainNumberParity,
            { single: string; double: string;}>;

    // Cache interne
    private _zone?: Nullable<number>;              // Zone du train si train commercial
                                                   //  (de 0 à 9 : 4ème chiffre du numéro)
    private _battery?: Nullable<number>;           // Batterie du train si train commercial
                                                   //  (de 0 à 99 : 5ème et 6ème chiffres)

    /**
     * Constructeur privé de la classe TrainNumber.
     * Garde uniquement les caractères alphanumériques
     *  et met les lettres en majuscules.
     * @param {TrainNumberInput} value - Numéro de train.
     * @param {boolean} [doubleParity=false] - Avec double parité.
     */
    private constructor(
        value: TrainNumberInput,
        { doubleParity = false }: { doubleParity?: boolean } = {}
    ) {

        const raw = value.toString();
        const normalized = TrainNumber.normalize(raw);

        if (!TrainNumber.isValidTrainNumber(normalized)) {
            throw new Error(
                `Numéro de train invalide : ${value}.`
                + ` Il doit être constitué d'au moins 4 caractères`
                + ` alphanumériques, le dernier étant un chiffre.`
            );
        }

        const lastDigit = normalized.charCodeAt(normalized.length - 1) - 48;
        const prefix = normalized.slice(0, -1);

        const evenDigit = lastDigit - (lastDigit % 2);
        const oddDigit = evenDigit + 1;

        const even = prefix + evenDigit;
        const odd = prefix + oddDigit;

        this.variants = {
            [Parity.EVEN]: {
                single: even,
                double: `${even}/${oddDigit}`
            },
            [Parity.ODD]: {
                single: odd,
                double: `${odd}/${evenDigit}`
            }
        };

        this._parity = lastDigit % 2
            ? Parity.ODD
            : Parity.EVEN;
        this._doubleParity = doubleParity || raw.includes("/");
    }

    /**
     * Numéro de référence du train, par convention le numéro pair
     *  sans double parité, pour réaliser toutes les analyses métier.
     * @returns 
     */
    private get referenceValue(): string {
        return this.variants[Parity.EVEN].single;
    }

    /**
     * Parité du numéro du train,
     *  par convention celle au départ du train.
     */
    public get parity(): TrainNumberParity {
        return this._parity;
    }

    /**
     * Modifie la parité du numéro de train (paire ou impaire) à son départ,
     *  ou lève une erreur si celle-ci est invalide.
     * @param {Parity | number} value - Parité du numéro de train.
     */
    public set parity(value: Parity | number) {
        this._parity = TrainNumber.normalizeTrainNumberParity(value, { raiseError: true })!;
    }

    /**
     * Indique si le train circule sous double parité.
     * @returns {boolean} - Vrai si le train circule sous double parité, faux sinon.
     */
    public get doubleParity(): boolean {
        return this._doubleParity;
    }

    /**
     * Modifie l'indicateur de double parité.
     * @param {boolean} value - Indicateur de double parité.
     */
    public set doubleParity(value: boolean) {
        this._doubleParity = value;
    }

    /**
     * Valeur courante du numéro de train,
     *  selon la parité initiale et la double parité.
     * @returns {string} - Numéro de train.
     */
    public get value(): string {
        return this._doubleParity
            ? this.variants[this._parity].double
            : this.variants[this._parity].single;
    }

    /**
     * Teste si le train est commercial.
     * @returns {boolean} - Vrai si le train est commercial, faux sinon.
     */
    public get isCommercial(): boolean {
        return TrainNumber.regex.commercial.test(this.referenceValue) ?? false;
    }

    /**
     * Teste si le train est W (vide voyageur).
     * @returns {boolean} - Vrai si le train est W, faux sinon.
     */
    public get isEmptyPassenger(): boolean {
        return TrainNumber.regex.emptyPassenger.test(this.referenceValue) ?? false;
    }

    /**
     * Teste si le train est une évolution.
     * @returns {boolean} - Vrai si le train est une évolution, faux sinon.
     */
    public get isMouvement(): boolean {
        return TrainNumber.regex.mouvement.test(this.referenceValue) ?? false;
    }

    /**
     * Retourne la zone du train (de 0 à 9 : 4ème chiffre du numéro du train).
     * Retourne null si le train n'est pas commercial.
     * @returns {Nullable<number>} Zone du train.
     */
    public get zone(): Nullable<number> {

        if (this._zone === undefined) {
            this._zone = this.isCommercial
                ? this.referenceValue.charCodeAt(3) - 48
                : null;
        }

        return this._zone;
    }

    /**
     * Retourne la batterie du train (de 0 à 99 : 5ème et 6ème chiffres du numéro du train)
     * Si le train a une parité double, c'est le 3ème chiffre qui donne la parité.
     * Retourne null si le train n'est pas commercial.
     * @returns {Nullable<number>} - Batterie du train.
     */ 
    public get battery(): Nullable<number> {

        if (this._battery === undefined) {

            this._battery = this.isCommercial
                ? parseInt(this.value.slice(4, 6), 10)
                    + (this.doubleParity ? parseInt(this.value[2], 10) % 2 : 0)
                : null;
        }

        return this._battery;
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string.
     * @returns {string} - Numéro de train.
     */
    public toString(): string {
        return this.format();
    }

    /**
     * Crée une instance de TrainNumber à partir d'un numéro de train ou d'un nombre.
     * La méthode normalise le numéro de train en supprimant les caractères non-alphanumériques.
     * La double parité est marquée par ######/#.
     * @param {Nullable<TrainNumberInput>} value - Numéro de train (nombre ou chaine de caractères).
     * @param {boolean} doubleParity - Si vrai, force la double parité. Si faux (par défaut),
     *  la double parité est détectée avec la présence de "/" dans le numéro de train.
     * @returns {TrainNumber | undefined} - Instance de TrainNumber correspondant au numéro de train.
     */
    public static from(
        value: Nullable<TrainNumberInput>,
        { doubleParity = false }: { doubleParity?: boolean } = {}
    ): TrainNumber | undefined {

        if (value == null || value === "") return undefined;
        if (value instanceof TrainNumber) return value;

        return new TrainNumber(value, { doubleParity });
    }

    /**
     * Normalise un numéro de train en supprimant les caractères non-alphanumériques
     *  et en ne gardant que la partie précédent un "/".
     * @param {string} value - Numéro de train à normaliser.
     * @returns {string} - Numéro de train normalisé.
     */
    private static normalize(
        value: string
    ): string {
        return value
            .split("/")[0]
            .toUpperCase()
            .replace(/[^A-Z0-9]/g, '');
    }

    /**
     * Normalise la parité d'un numéro de train, pour qu'elle soit impaire ou paire uniquement.
     * @param {Parity | number} parity - Parité du numéro de train à normaliser.
     * @param {boolean} raiseError - Indique si une erreur doit être levée si la parité est invalide.
     * @returns {TrainNumberParity | undefined} - Parité normalisée, ou undefined si la parité est invalide.
     */
    private static normalizeTrainNumberParity(
        parity: Parity | number,
        { raiseError = false }: { raiseError?: boolean } = {}
    ): TrainNumberParity | undefined {
    
        const result = parity instanceof Parity
            ? parity.value
            : parity;
        
        if (result === Parity.EVEN || result === Parity.ODD) return result;
        
        if (raiseError) throw new Error(`La parité d'un numéro de train`
            + ` doit être impaire ou paire (ODD ou EVEN).`);

        return undefined;
    }

    /**
     * Retourne le numéro de train selon les paramètres demandés :
     *  - abrégé de 6 à 4 chiffres si autorisé,
     *  - avec ou sans double parité,
     *  - avec parité imposée.
     * @param {boolean} [abbreviate=false] - Si vrai, le numéro du train est abrégé
     *  de 6 à 4 chiffres pour les trains commerciaux. Si faux (par défaut), le numéro n'est pas abrégé.
     * @param {boolean} [withDoubleParity=false] - Si vrai, le numéro est renvoyé avec double parité
     *  (format ######/#) si le numéro de train inclue la double parité.
     * @param {Parity | number} [forceParity] - Parité imposée. Si la parité imposée est différente
     *  de la parité du numéro de train et que celui-ci n'inclue pas à double parité,
     *  renvoie une chaîne vide.
     * @returns {string} - Chaîne représenant le numéro de train. 
     */
    public format(
        {
            abbreviate = false,
            withDoubleParity = true,
            forceParity
        }: {
            abbreviate?: boolean;
            withDoubleParity?: boolean;
            forceParity?: Parity | number;
        } = {}
    ): string {

        let parity = this._parity;
        let doubleParity = this._doubleParity;

        let normalizedForceParity = forceParity
            ? TrainNumber.normalizeTrainNumberParity(forceParity) 
            : undefined;
        if (normalizedForceParity
            && normalizedForceParity !== this._parity
        ) {
            if (!this._doubleParity) return "";
            parity = normalizedForceParity;
        }

        let result =
            doubleParity && withDoubleParity
                ? this.variants[parity].double
                : this.variants[parity].single;

        if (abbreviate) {
            result = TrainNumber.abbreviate(result);
        }

        return result;
    }

    /**
     * Teste si une valeur correspond à ce train (toutes formes confondues).
     * @param {Nullable<TrainNumberInput>} value - Valeur à tester.
     * @returns {boolean} - Vrai si la valeur correspond au train, faux sinon.
     */
    public includes(
        value: Nullable<TrainNumberInput>
    ): boolean {

        if (value == null) return false;

        const normalized = value instanceof TrainNumber
            ? value.value
            : TrainNumber.normalize(String(value));

        return Object.values(this.variants)
            .some(v =>
                normalized === v.single
                || normalized === v.double
            );
    }

    /**
     * Vérifie si un numéro de train est valide : 
     *  code alphanumérique de 4 caractères au moins, le dernier étant un chiffre.
     * @param {string} value - Numéro de train à vérifier.
     * @returns {boolean} - Vrai si le numéro de train est valide, faux sinon.
     */
    private static isValidTrainNumber(
        value: string
    ): boolean {

        if (!value) return false;

        const lastChar = value.slice(-1);

        return value.length >= 4
            && /^[0-9]$/.test(lastChar);
    }

    /**
     * Abrège le numéro de train à 4 chiffres si possible.
     * La méthode teste si le numéro de train correspond à une expression régulière
     *  définie dans la classe TrainNumber.
     * Si le numéro de train correspond, il est abrégé en supprimant les 2 premiers chiffres.
     * Si le numéro de train ne correspond pas, il est renvoyé inchangé.
     * @param {string} value - Numéro de train à abréger.
     * @returns {string} - Numéro de train abrégé de 6 à 4 chiffres s'il est abrégeable.
     */
    private static abbreviate(
        value: string
    ): string {

        return this.regex.abbreviate4.test(value.split("/")[0])
            ? value.substring(2)
            : value;
    }

    /**
     * Charge les paramètres des numéros de train
     *  - regex des numéros de train W,
     *  - regex des numéros de train abrégeables à 4 chiffres.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        if (this.loaded && !erase) return;

        // Récupère les lignes de la table des paramètres.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }

        // Crée la table de récupérations des motifs qui vont constituer chaque Regex
        const patterns = Object.fromEntries(
            Object.keys(this.REGEX_COLUMNS).map(key => [key, [] as string[]])
        ) as Record<keyof typeof this.REGEX_COLUMNS, string[]>;

        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                for (const [key, column] of 
                    Object.entries(this.REGEX_COLUMNS) as [keyof typeof this.REGEX_COLUMNS, number][]) {

                    const value = TableSerializer.getValueOrUndefined({ type: "string", row, index: column })?.trim();
                    if (!!value) patterns[key].push(value.trim());
                }
            } 

        } catch (e) {
            throw new Error(`${this.name}.load (ligne ${excelRow}) : ${e}`);
        }

        // Construit chaque Regex
        this.regex = Object.fromEntries(
            Object.keys(patterns).map(key => [
                key,
                this.buildRegex(
                    patterns[key as keyof typeof patterns]
                )
            ])
        ) as typeof this.regex;

        this.loaded = true;

        Log.timer(`${this.name}.load()`);
    }

    /**
     * Construit une Regex à partir des motifs des trains spécifiques.
     * @param {string[]} patterns - Motifs des trains.
     * @returns {RegExp} - Regex construite.
     */
    private static buildRegex(
        patterns: string[]
    ): RegExp {
    
        const parts = patterns.map(pattern => "^"
            + pattern.replace(/#/g, "\\d")
            + "$"
        );
    
        return parts.length
            ? new RegExp(parts.join("|"))
            : /^$/;
    }
}

/**
 * Type StationInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Station ou StationWithParity.
 */
type StationInput = Input<Station, StationWithParity | string>;

/**
 * Interface StationParams contenant les paramètres d'une gare.
 * @param {string} abbreviation - Abréviation de la gare.
 * @param {string} name - Nom de la gare.
 * @param {ParityInput} turnaround - Parité d'un rebroussement possible
 *  (la parité est celle du train avant rebroussement).
 * @param {boolean} reverseLineDirection - Parité de la ligne inversée sur cette gare.
 */
interface StationParams {
    abbreviation: string;
    name: string;
    turnaround?: ParityInput;
    reverseLineDirection: boolean;
}

/* 
 * Classe Station définissant une gare.
 */
class Station {

    // Propriétés de la classe Station
    public readonly id: number;                         // Id de la gare
    public readonly abbreviation!: string;              // Abréviation de la gare
    public readonly name: string;                       // Nom de la gare
    public referenceStation: Nullable<Station>;         // Gare de rattachement
    public childStations: Station[];                    // Sous-gares
    public readonly turnaround: Parity;                 // Parité d'un rebroussement possible
                                                        //  (la parité est celle du train avant rebroussement)
    public readonly reverseLineDirection: boolean;      // Parité de la ligne inversée sur cette gare

    /**
     * Constructeur d'une gare.
     * @param {StationParams} params - Paramètres de la gare.
     * @param {number} id - Id de la gare.
     * @param {Station} referenceStation - Gare de rattachement.
     */
    public constructor(params: StationParams & {
        id: number,
        referenceStation: Nullable<Station>
    }) {
        this.id = params.id;
        if (!params.abbreviation) {
            throw new Error(`Une gare ne peut pas avoir une abréviation vide.`);
        }
        this.abbreviation = params.abbreviation;
        if (!params.name) {
            throw new Error(`La gare ${params.abbreviation} ne peut pas avoir un nom vide.`);
        }
        this.name = params.name;
        this.referenceStation = params.referenceStation ?? null;
        this.childStations = [];
        this.turnaround = Parity.from(params.turnaround, { doubleParityAllowed: true });
        this.reverseLineDirection = params.reverseLineDirection;
    }
 
    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Abréviation de la gare.
     */
    public toString(): string {
        return this.abbreviation;
    }

    /**
     * Retourne une instance de Station à partir d'une valeur qui peut être :
     *  - une instance de Station,
     *  - une instance de StationWithParity,
     *  - un nom de gare ou une clé d'arrêt (avec ou sans suffixe parité),
     *  - null ou undefined (lève une erreur).
     * @param {Nullable<StationInput>} value - Valeur à analyser
     *  pour la gare.
     * @returns {StationWithParity | undefined} - Instance de Station correspondante.
     */
    public static from(
        value: Nullable<StationInput>,
    ): Station | undefined {

        if (value == null || value === "") return undefined;

        // Instance de Station
        if (value instanceof Station) return value;
           

        // Instance de StationWithParity
        if (value instanceof StationWithParity) return value.station;

        // Chaîne qui correspond à la clé d'une gare ou d'une gare avec parité
        return Stations.get(value.split("_")[0]);
    }
}

/**
 * Classe Stations contenant la liste des gares.
 */
class Stations {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Gares";                // Nom de la feuille
    private static readonly TABLE = "Gares";                // Nom du tableau
    private static readonly START_CELL = "A1";              // Première cellule
    private static readonly DATABASE_COLUMNS = {            // Liste des colonnes avec leur emplacement
        abbreviation: 0,
        name: 1,
        referenceStationAbbreviation: 2,
        turnaround: 3,
        reverseLineDirection: 4
    } as const;
    
    /**
     * Liste des définitions générales :
     *  - abréviation,
     *  - nom,
     *  - gare de rattachement,
     *  - gare de rebroussement,
     *  - parité de ligne inversée.
     */
    private static readonly COLUMN_DEFINITIONS: TableColumns<Station> = {

        abbreviation: {
            header: "Abréviation",
            type: "string",
            required: true,
            load: value => String(value).toUpperCase(),
            format: { width: 100 }
        },
    
        name: {
            header: "Nom",
            type: "string",
            required: true,
            format: { width: 300 }
        },
    
        referenceStationAbbreviation: {
            header: "Gare de rattachement",
            type: "string",
            // print: (value, station) => station.referenceStation?.abbreviation,
            format: { width: 100 }
        },
    
        turnaround: {
            header: "Gare de rebroussement",
            type: "string",
            print: value => (value as Parity).printLetter(),
            format: { width: 40 }
        },
    
        reverseLineDirection: {
            header: "Parité de ligne inversée",
            type: "boolean",
            format: { width: 40 }
        }
    };

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
    public static has(
        value: string
    ): boolean {
        return value in this.abbrMap || value in this.nameMap;
    }

    /**
     * Retourne la gare correspondant à l'ID donné.
     * @param {number} id - ID de la gare.
     * @returns {Station} - Gare correspondante.
     */
    public static getById(
        id: number
    ): Station {
        const s = this.list[id];
        if (!s) throw new Error(`Gare : ID ${id} inconnue`);
        return s;
    }

    /**
     * Retourne une gare correspondant à l'abréviation ou au nom donné.
     * @param {string} value - Abréviation ou nom de la gare.
     * @returns {Station | undefined} - Gare correspondante, ou undefined si non trouvée.
     */
    public static get(
        value: string
    ): Station | undefined {

        let adaptValue = value;

        for (const suffix of Params.get('stationsSuffixes')) {
            if (adaptValue.endsWith(`-${suffix}`)) {
                adaptValue = adaptValue.slice(0, -suffix.length - 1);
                break; // S'arrête au premier match.
            }
        }

        return this.abbrMap[adaptValue] ?? this.nameMap[adaptValue];
    }

    /**
     * Crée une nouvelle gare et l'ajoute à la base de données, 
     *  référencée par son ID, sa clé et son nom.
     * Si une gare avec la même clé ou le même nom existe déjà, une erreur est levée.
     * @param {StationParams} params - Paramètres de la gare.
     * @returns {Station} - La nouvelle gare créée.
     * @throws {Error} - Si une gare avec la même clé ou le même nom existe déjà.
     */
    private static create(params: StationParams): Station {

        // Vérifie que la gare n'existe pas déjà.
        for (const value of [params.abbreviation, params.name]) {
            if (this.has(value)) {
                throw new Error(`La gare ${value} est déjà présente dans la base de données.`);
            }
        }

        // Calcule l'ID.
        const id = this.list.length;

        // Instancie la nouvelle gare.
        const station = new Station({
            ...params,
            id,
            referenceStation: null,
        });

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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        const referenceStationPairs: [Station, string][] = [];
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    Station,
                    StationParams & { referenceStationAbbreviation?: string; }
                >({
                    row,
                    columns: this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: "abbreviation"
                });
                if (!params) continue;

                // Crée l'objet et l'insère dans la base de données.
                const station = this.create(params);

                // Mémorise les paires gare/gare de rattachement.
                if (params.referenceStationAbbreviation) {
                    referenceStationPairs.push([
                        station,
                        params.referenceStationAbbreviation
                    ]);
                }
            } 

        } catch (e) {
            throw new Error(`${this.name}.load (ligne ${excelRow}) : ${e}`);
        } 

        // Parcourt les paires pour ajouter les objets des gares de réference à chaque gare.
        for (const [station, referenceStationAbbv] of referenceStationPairs) {
            const referenceStation = this.get(referenceStationAbbv);

            if (referenceStation) {
                station.referenceStation = referenceStation;
                referenceStation.childStations.push(station);
            }
        }

        Log.timer(`${this.name}.load()`);
    }

    /**
     * Sauvegarde la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell=this.START_CELL] - Adresse de la cellule de départ pour le tableau.
     */
    public static save(
        {
            sheetName = this.SHEET,
            tableName = this.TABLE,
            startCell = this.START_CELL
        }: {
            sheetName?: string,
            tableName?: string,
            startCell?: string
        } = {}
    ): void {
    
        Log.startTimer(`${this.name}.save()`);

        TableSerializer.print({
            entities: Array.from(this.values()),
            columns: this.DATABASE_COLUMNS,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        Log.timer(`${this.name}.save()`);
    }
}

/**
 * Type StationInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Station ou StationWithParity.
 */
type StationWithParityInput = Input<StationWithParity, Station | string>;

/**
 * Interface StationWithParityParams contenant les paramètres d'une gare avec parité.
 * @param {Station} station - Gare.
 * @param {ParityInput} parity - Parité.
 */
interface StationWithParityParams {
    station: Station;
    parity: ParityInput;
}

/**
 * Classe StationWithParity immuable définissant une gare d'arrêt ou de passage d'un train
 *  et sa parité associée.
 */
class StationWithParity {

    // Constantes de parité
    public static readonly UNDEFINED = 0;           // Valeur de parité undefined pour le calcul de l'ID
    public static readonly ODD = 1;                 // Valeur de parité impaire pour le calcul de l'ID
    public static readonly EVEN = 2;                // Valeur de parité paire pour le calcul de l'ID

    // Propriétés de la classe StationWithParity
    public readonly id: number;                     // Identifiant unique
    public readonly key: string;                    // Clé (en cache)

    private _expandedCache?: StationWithParity[];   // Cache des gares avec parité rattachées

    /**
     * Constructeur d'une gare avec parité.
     * @param {StationWithParityParams} params - Paramètres de la gare avec parité.
     * @param {number} id - Id de la gare.
     * @param {Station} referenceStation - Gare de rattachement.
     */
    public constructor(params: StationWithParityParams & {
        id: number
    }) {
        this.id = params.id;

        const parityObj = Parity.from(params.parity, { doubleParityAllowed: false });
        this.key = StationWithParity.keyOf(params.station, parityObj); 
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
     * @returns {string} - Clé de l'objet, 
     *  c'est-a-dire l'abbréviation et la parité de la gare sous la forme GARE_#.
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
     * @param {Nullable<StationWithParityInput>} value - Valeur à analyser
     *  pour la gare.
     * @param {ParityInput} [parity] - Parité optionnelle imposée,
     *  qui remplace celle potentiellement présente dans value.
     * @returns {StationWithParity | undefined} - Instance de StationWithParity correspondante.
     */
    public static from(
        value: Nullable<StationWithParityInput>,
        { parity = Parity.UNDEFINED }: { parity?: ParityInput } = {}
    ): StationWithParity | undefined {

        if (value == null || value === "") return undefined;

        const parityObj = Parity.from(parity, { doubleParityAllowed: false });

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
     * Retourne la valeur unique de la parité pour le calcul de l'ID, qui est :
     *  - this.ODD si la parité est impaire,
     *  - this.EVEN si la parité est paire,
     *  - this.UNDEFINED sinon.
     * @param {Parity} parity - La parité à transformer.
     * @returns {number} - Valeur unique de la parité. pour le calcul.
     */
    public static parityValue(
        parity: Parity
    ): number {
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
            parity: Parity.from(parityPart, { doubleParityAllowed: false })
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
     * @param {Nullable<StationWithParity>} other - Autre objet StationWithParity à comparer.
     * @returns {boolean} - Vrai si les deux objets ont la même gare, faux sinon.
     */
    public hasSameStationTo(
        other: Nullable<StationWithParity>
    ): boolean {
        return !!other && Math.floor(this.id / 3) === Math.floor(other.id / 3);
    }
 
    /**
     * Vérifie si la gare avec parité a parmi ses gares rattachées une seconde gare, c'est à dire :
     *  - que cette seconde gare est identique ou est une gare fille de la première,
     *  - et que si la parité de la première est définie, celle de la seconde est identique.
     * @param {Nullable<StationWithParity>} other - Autre objet StationWithParity qui doit être inclus ou non.
     * @returns {boolean} - Vrai si l'objet inclut l'autre, faux sinon.
     */
    public includes(
        other: Nullable<StationWithParity>
    ): boolean {
        return !!other
            && this.expandWithChildren().includes(other);
    }
 
    /**
     * Vérifie si l'objet StationWithParity est identique à l'autre.
     * @param {Nullable<StationWithParity>} other - Autre objet StationWithParity à comparer.
     * @returns {boolean} - Vrai si les deux objets sont identiques, faux sinon.
     */
    public equalsTo(
        other: Nullable<StationWithParity>
    ): boolean {
        return this === other;
    }

    /**
     * Retourne une chaîne représentant l'objet StationWithParity sous la forme
     *  GARE_PARITE, où GARE est l'abbréviation de la gare et
     *  PARITE est la parité sous forme de chiffre.
     * @param {Station} station - Gare.
     * @param {Parity} parity - Parité.
     * @returns {string} - Chaîne représentant l'objet StationWithParity.
     */
    private static keyOf(
        station: Station,
        parity: Parity
    ): string {
        return parity.isDefined() 
            ? `${station.abbreviation}_${parity.printDigit()}` 
            : `${station.abbreviation}`;
    }

    /**
     * Retourne un tableau de toutes les gares avec parités rattachées en renvoyant : 
     *  - les 3 parités si elle n'est pas définie,
     *  - les gares filles.
     * La méthode prend en paramètre un ensemble de gares déjà visitées pour éviter les boucles infinies.
     * @param {Set<number>} [visited=new Set<number>()] - Ensemble des gares déjà visitées.
     * @returns {StationWithParity[]} - Tableau contenant toutes les gares visitées.
     */
    public expandWithChildren(
        visited: Set<number> = new Set<number>()
    ): StationWithParity[] {

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
    public static hasKey(
        key: string
    ): boolean {
        return key in this.keyMap;
    }

    /**
     * Retourne la gare avec parité correspondant à l'ID donné.
     * @param {number} id - ID de la gare avec parité.
     * @returns {StationWithParity} - Gare correspondant si elle existe, undefined sinon.
     */
    public static getById(
        id: number
    ): StationWithParity {
        const s = this.list[id];
        if (!s) throw new Error(`Gare avec parité : ID ${id} inconnue`);
        return s;
    }
 
    /**
     * Retourne une gare avec parité correspondant à la clé donnée.
     * @param {string} key - Clé de la gare avec parité.
     * @returns {StationWithParity | undefined} - Gare correspondant si elle existe, undefined sinon.
     */
    public static getByKey(
        key: string
    ): StationWithParity | undefined {
        return this.keyMap[key];
    }
 
    /**
     * Retourne la gare avec parité correspondant à la gare et la parité données.
     * @param {Station} station - Gare à trouver.
     * @param {ParityInput} parity - Parité à trouver.
     * @returns {StationWithParity | undefined} - Gare correspondant si elle existe, undefined sinon.
     */
    public static getFromStationAndParity(
        station: Station,
        parity: ParityInput
    ): StationWithParity | undefined {
        const base = station.id * 3;
        const parityObj = Parity.from(parity, { doubleParityAllowed: false });
        const id = base + StationWithParity.parityValue(parityObj);
        return this.list[id];
    }

    /**
     * Crée une gare avec parité et l'ajoute à la base de données,
     *  référencée par son ID ou sa clé.
     * Si la gare avec parité est déjà présente, une erreur est levée.
     * @param {StationWithParityParams} params - Paramètres de la gare avec parité.
     * @returns {StationWithParity} - Gare avec parité crée.
     * @throws {Error} - Si la gare avec parité est déjà présente dans la base de données.
     */
    private static create(params: StationWithParityParams): StationWithParity {

        const parityObj = Parity.from(params.parity, { doubleParityAllowed: false });
        const id = params.station.id * 3 + StationWithParity.parityValue(parityObj);
        const swp = new StationWithParity({ ...params, id });
        if (this.hasKey(swp.key)) {
            throw new Error(`La gare avec parité ${swp} est déjà présente`
                + ` dans la base de données.`);
        }
        this.list[swp.id] = swp;
        this.keyMap[swp.key] = swp;

        return swp;
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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

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
                this.create({ station, parity });
            }
        }

        Log.timer(`${this.name}.load()`);
    }
}

/**
 * Interface ConnectionParams contenant les paramètres de la classe Connection.
 * @param {StationWithParity} from - Gare de départ.
 * @param {StationWithParity} to - Gare d'arrivée.
 * @param {DateTimeInput} [time] - Temps de trajet.
 * @param {boolean} [withMovement] - Connexion sous régime de l'évolution.
 * @param {boolean} [changeParity] - Connexion avec changement de parité.
 */
interface ConnectionParams {
    from: StationWithParity;
    to: StationWithParity;
    time?: DateTimeInput;
    withMovement?: boolean;
    changeParity?: boolean;
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

    // Propriétés de la classe Connection
    public readonly from: StationWithParity;            // Gare de départ
    public readonly to: StationWithParity;              // Gare d'arrivée
    private _time: DateTime;                            // Temps de trajet
    public readonly withTurnaround: boolean;            // Connexion impliquant un rebroussement
    public readonly withMovement: boolean;              // Connexion sous régime de l'évolution
    public readonly changeParity: boolean;              // Connexion avec changement de parité

    /**
     * Constructeur d'une connexion.
     * @param {ConnectionParams} params - Paramètres de la connexion.
     */
    public constructor(params: ConnectionParams) {

        const fromObj = StationWithParity.from(params.from);
        if (!fromObj) throw new Error(`La gare de départ ${params.from} est inconnue.`);
        const toObj = StationWithParity.from(params.to);
        if (!toObj) throw new Error(`La gare d'arrivée ${params.to} est inconnue.`);
        if (fromObj.equalsTo(toObj)) {
            throw new Error(
                `Une connexion ne peut pas relier ${fromObj} à elle-même`
                + ` sans changement de gare ou de parité.`
            );
        }
        this.from = fromObj;
        this.to = toObj;
        
        this.withTurnaround = this.from.hasSameStationTo(this.to);
        let timeObj: DateTime | undefined;
        if (this.withTurnaround) {
            timeObj = DateTime.from(0, { isRelative: true })!
        } else {
            timeObj = DateTime.from(params.time, { isRelative: true });
            if (!timeObj || timeObj.excelValue <= 0) {
                timeObj = DateTime.from(Connection.DEFAULT_CONNECTION_TIME,
                { isRelative: true })!;
            }
        }
        this._time = timeObj;
        
        this.withMovement = params.withMovement ?? false;
        this.changeParity = params.changeParity ?? false;
    }

    /**
     * Retourne le temps de trajet de la connexion.
     * @returns {DateTime} - Temps de trajet de la connexion.
     */
    public get time(): DateTime {
        return this._time;
    }

    /**
     * Modifie le temps de trajet de la connexion.
     * @param {DateTimeInput} value - Nouveau temps de trajet de la connexion.
     * @throws {Error} - Si le temps de trajet est inférieur ou égal à 0 ou n'est pas relatif.
     */
    public set time(
        value: DateTimeInput
    ) {
        const timeObj = DateTime.from(value, { isRelative: true });
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
     * @returns {string} - Connexion sous  la forme "GARE_DEPART_# -> GARE_ARRIVEE_#".
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
    private static readonly SHEET = "Paramètres";                // Nom de la feuille 
    private static readonly TABLE = "Connexions";           // Nom du tableau
    private static readonly START_CELL = "I1";              // Première cellule
    private static readonly DATABASE_COLUMNS = {            // Liste des colonnes avec leur emplacement
        from: 0,
        to: 1,
        time: 2,
        withMovement: 3,
        changeParity: 4
    } as const;

    /**
     * Liste des définitions générales :
     *  - gares de départ et d'arrivée,
     *  - temps de trajet,
     *  - durée de la connexion,
     *  - sous régime de l'évolution,
     *  - avec changement de parité.
     */
    private static readonly COLUMN_DEFINITIONS: TableColumns<Connection> = {

        from: {
            header: "Origine",
            type: "string",
            required: true,
            // print: (value, connection) => connection.from.key,
            format: { width: 100 }
        },
    
        to: {
            header: "Destination",
            type: "string",
            required: true,
            // print: (value, connection) => connection.to.key,
            format: { width: 100 }
        },
    
        time: {
            header: "Temps",
            type: "number",
            format: {
                numberFormat: "hh:mm:ss",
                width: 60
            }
        },
    
        withMovement: {
            header: "Mouvement",
            type: "boolean",
            format: { width: 40 }
        },
    
        changeParity: {
            header: "Changement parité",
            type: "boolean",
            format: { width: 40 }
        }
    };
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
     * Retourne la connexion correspondant aux gares de départ et d'arrivée données.
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
    public static getNeighbors(
        id: number
    ): Connection[] {
        return this.list[id] ?? [];
    }

    /**
     * Crée une nouvelle connexion et l'ajoute à la base de données,
     *  référencée par ses gares de départ et d'arrivée.
     * Si la connexion est déjà présente dans la base de données, une erreur est levée.
     * @param {ConnectionParams} params - Paramètres de la connexion.
     * @returns {Connection} - La connexion ajoutée.
     * @throws {Error} - Si la connexion est déjà présente dans la base de données.
     */
    private static create(params: ConnectionParams): Connection {

        const connection = new Connection(params);

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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les gares si elles n'ont pas encore été chargées.
        StationsWithParity.load(); 

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        const referenceStationPairs: [Station, string][] = [];
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    Connection,
                    ConnectionParams
                >({
                    row,
                    columns: this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: "from"
                });
                if (!params) continue;

                // Crée l'objet et l'insère dans la base de données.
                const station = this.create(params);
            }

        } catch (e) {
            throw new Error(`Connections.load (ligne ${excelRow}) : ${e}`);
        }

        Log.timer(`${this.name}.load()`);
    }

    /**
     * Sauvegarde la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell=this.START_CELL] - Adresse de la cellule de départ pour le tableau.
     */
    public static save(
        {
            sheetName = this.SHEET,
            tableName = this.TABLE,
            startCell = this.START_CELL
        }: {
            sheetName?: string,
            tableName?: string,
            startCell?: string
        } = {}
    ): void {
    
        Log.startTimer(`${this.name}.save()`);

        TableSerializer.print({
            entities: Array.from(this.values()),
            columns: this.DATABASE_COLUMNS,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        Log.timer(`${this.name}.save()`);
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

        const queue: ConnectionState[] = [];
        const visited: Map<number, number> = new Map();
 
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
                queue.push(new ConnectionState(
                    {
                        stationId,
                        cost: 0,
                        groupIndex: 1,
                        visitedMask: 0
                    }
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
     * @param {ConnectionState[]} queue - File d'attente contenant les états à visiter.
     * @param {Map<number, number>} visited - Carte des états déjà visités.
     * @param {number[][][]} routeStations - Trajet avec les gares intermédiaires
     *  et les parités possibles.
     * @returns {Connection[] | undefined} - Chemin le plus court entre le départ
     *  et l'arrivée du trajet. Si aucun chemin n'est trouvé, undefined est renvoyé.
     */
    private static runGroupedDijkstra(
        queue: ConnectionState[],
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
     * @param {ConnectionState} state - État actuel.
     * @param {number[][][]} routeStations - Trajet avec les gares intermédiaires
     *  et les parités possibles.
     * @returns {ConnectionState[]} - Liste des états suivants.
     */
    private static expandNeighbors(
        state: ConnectionState,
        routeStations: number[][][]
    ): ConnectionState[] {
 
        const result: ConnectionState[] = [];
 
        // Donne les gares voisines.
        const neighbors = this.getNeighbors(state.stationId);
 
        for (const connection of neighbors) {

            // Donne l'id de la gare voisine.
            const nextStationId = connection.to.id;
 
            // Ajoute le coût de la connection : temps de parcours, ou temps de retournement.
            const nextCost = state.cost
                + (connection.withTurnaround
                    ? Params.get('turnaroundTime').excelValue
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
            result.push(new ConnectionState(
                {
                    stationId: nextStationId,
                    cost: nextCost,
                    groupIndex: nextGroup,
                    visitedMask: nextMask,
                    previous: state,
                    via: connection
                }
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
    public static saveConnectionsTimes(
        paths: (Train | TrainPath | PathInput)[]
    ) {
        const pathsList = Utils.asArray(paths).map(value => {
            if (value instanceof Train || value instanceof TrainPath) return value.path;
            if (value instanceof Path) return value;
            return Paths.get(value);
        }).filter(v => v !== undefined) as Path[];

        for (const path of pathsList) {
            for (const stop of path.stops) {
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
            }
        }
    }
}

/**
 * Interface ConnectionStateParams contenant les paramètres de la classe ConnectionState.
 * @param {number} stationId - Identifiant de la gare.
 * @param {number} cost - Coût du chemin.
 * @param {number} groupIndex - Index du groupe de gares.
 * @param {number} visitedMask - Masque des gares visitées.
 * @param {ConnectionState} previous - Etat précedent.
 * @param {Connection} via - Connection ajoutée.
 */
interface ConnectionStateParams {
    stationId: number;
    cost: number;
    groupIndex: number;
    visitedMask: number;
    previous?: ConnectionState;
    via?: Connection;
}

/**
 * Classe ConnectionState définissant un état de recherche de l'algorithme Dijkstra.
 */
class ConnectionState {

    public readonly stationId: number;              // Identifiant de la gare
    public readonly cost: number;                   // Coût du chemin
    public readonly groupIndex: number;             // Index du groupe de gares
    public readonly visitedMask: number;            // Masque des gares visitées
    public readonly previous?: ConnectionState;     // Etat précedent
    public readonly via?: Connection;               // Connection ajoutée

    /**
     * Constructeur de l'état de recherche de l'algorithme Dijkstra.
     * @param {ConnectionStateParams} params - Paramètres de l'état de recherche.
     */
    public constructor(params: ConnectionStateParams) {

        this.stationId = params.stationId,
        this.cost = params.cost,
        this.groupIndex = params.groupIndex,
        this.visitedMask = params.visitedMask,
        this.previous = params.previous,
        this.via = params.via
    }

    /**
     * Retourne une clé unique pour l'état de recherche, composée de l'identifiant de la gare,
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
        let current: ConnectionState | undefined = this;

        while (current?.via) {
            path.push(current.via);
            current = current.previous;
        }

        return path.reverse();
    }
}

/**
 * Type StopInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Stop.
 */
type StopInput = Input<Stop, StationInput>;

/**
 * Interface StopParams contenant les paramètres de la classe Stop.
 * @param {PathInput} path - Parcours de l'arrêt.
 * @param {StationInput} station - Gare de l'arrêt.
 * @param {StationWithParity | string} stationAfterTurnaround - Gare avec parité après rebroussement
 * @param {DateTimeInput} arrivalTime - Temps / Heure d'arrivée de l'arrêt.
 * @param {DateTimeInput} departureTime - Temps / Heure de départ de l'arrêt.
 * @param {DateTimeInput} passageTime - Temps / Heure de passage à l'arrêt (sans arrêt).
 * @param {boolean} areRelativeTimes - Indique si les heures sont relatives.
 * @param {OneOrMany<string>} tracks - Voies de l'arrêt.
 */
interface StopParams {
    path?: PathInput;
    station: StationInput;
    stationAfterTurnaround?: StationWithParity | string;
    arrivalTime?: DateTimeInput;
    departureTime?: DateTimeInput;
    passageTime?: DateTimeInput;
    areRelativeTimes?: boolean;
    tracks?: OneOrMany<string>;
}

/*
 * Classe Stop définissant l'arrêt ou le passage d'un train dans une gare.
 */
class Stop {

    // Propriétés de la classe Stop
    public  path?: Path;                            // Parcours de l'arrêt
    public readonly station: StationWithParity;     // Gare de l'arrêt
    private _withTurnaround: boolean = false;       // Arrêt avec rebroussement
    private _arrivalTime?: DateTime;                // Temps / Heure d'arrivée de l'arrêt
    private _departureTime?: DateTime;              // Temps / Heure de départ de l'arrêt
    private _passageTime?: DateTime;                // Temps / Heure de passage à l'arrêt (sans arrêt)
    private _tracks: string[];                      // Voies de l'arrêt
 
    /**
     * Constructeur d'un arrêt.
     * @param {StopParams} params - Paramètres de l'arrêt.
     */
    public constructor(params: StopParams) {

        // Récupère le parcours de l'arrêt
        this.path = Path.from(params.path);
        // Détermine la gare d'arrêt
        const stationObj = StationWithParity.from(params.station)
        if (!stationObj) throw new Error(`La gare ${params.station} du parcours ${params.path} est inconnue.`);
        this.station = stationObj;

        // Détermine le rebroussement
        this._withTurnaround = this.canTurnaroundTo(params.stationAfterTurnaround);

        // Détermine les horaires de l'arrêt
        this.setTimes({ 
            arrivalTime: params.arrivalTime,
            departureTime: params.departureTime,
            passageTime: params.passageTime,
            areRelativeTimes: params.areRelativeTimes
        });

        // Détermine les voies de l'arrêt
        this._tracks = Utils.asArray(params.tracks, {split: ';', trim: true, filterEmptyString: true});
    }

    /**
     * Retourne une clé unique pour l'arrêt, composée du nom de la gare et de la parité (si connue).
     * @returns {string} - Clé unique.
     */
    public get key(): string {
        return this.station.key;
    }

    /**
     * Retourne le nom de la gare associée à cet arrêt.
     * @returns {string} - Nom de la gare.
     */
    public get stationName(): string {
        return this.station!.station.name;
    }

    /**
     * Retourne l'abréviation de la gare associée à cet arrêt.
     * @returns {string} - Abréviation de la gare.
     */
    public get stationAbbreviation(): string {
        return this.station!.station.abbreviation;
    }

    /**
     * Retourne vrai si l'arrêt à un rebroussement possible, faux sinon.
     * @returns {boolean} - Vrai si l'arrêt a un rebroussement possible, faux sinon.
     */
    public get withTurnaround(): boolean {
        return this._withTurnaround;
    }

    /**
     * Retourne vrai si l'arrêt est un passage sans arrêt, faux sinon.
     * @returns {boolean} - Vrai si l'arrêt est un passage sans arrêt, faux sinon.
     */
    public get withNonStopPassage(): boolean {
        return !!this._passageTime;
    }

    /**
     * Indique si l'arrêt est un arrêt intermédiaire,
     *  avec une heure d'arrivée et une heure de départ, ou une heure de passage.
     * @returns {boolean} - Vrai si l'arrêt est un arrêt intermédiaire, faux sinon.
     */
    public get isIntermediateStop(): boolean {
        return this.withNonStopPassage
            || (!!this._arrivalTime && !!this._departureTime);
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
    public set tracks(
        value: string[]
    ) {
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
    public set stationAfterTurnaround(
        value: StationWithParity | string | undefined
    ) {
        this._withTurnaround = this.canTurnaroundTo(value);
        if (!!this._passageTime) {
            this.setTimes({
                arrivalTime: this._passageTime,
                departureTime: Params.get('turnaroundTime').resolveAgainst(this._passageTime!)
            });
        }
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Clé de l'arrêt, c'est à dire abbréviation de la gare.
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
    private canTurnaroundTo(
        stationAfterTurnaround : StationWithParity | string | undefined
    ): boolean {

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
     * @param {DateTimeInput} [arrivalTime] - Heure d'arrivée à l'arrêt.
     * @param {DateTimeInput} [departureTime] - Heure de départ à l'arrêt.
     * @param {DateTimeInput} [passageTime] - Heure de passage à l'arrêt.
     * @param {boolean} [areRelativeTimes=undefined] - Vrai si les heures sont relatives, faux sinon.
     */
    public setTimes(
        {
            arrivalTime,
            departureTime,
            passageTime,
            areRelativeTimes
        }: {
            arrivalTime?: DateTimeInput,
            departureTime?: DateTimeInput,
            passageTime?: DateTimeInput,
            areRelativeTimes?: boolean
        }
    ) {
        this._arrivalTime = DateTime.from(arrivalTime, { isRelative: areRelativeTimes });
        this._departureTime = DateTime.from(departureTime, { isRelative: areRelativeTimes });
        this._passageTime = (!arrivalTime && !departureTime)
            ? DateTime.from(passageTime, { isRelative: areRelativeTimes })
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
     * Retourne la plus petite des heures d'arrivée, de départ ou de passage à l'arrêt.
     * Si ignoreArrival est vrai, lit plutôt l'heure de départ ou de passage.
     * @param {boolean} [ignoreArrival=false] - Si vrai, ignore l'heure d'arrivée
     *  et préfère l'heure de départ ou de passage. Si faux (par défaut),
     *  c'est d'abord l'heure d'arrivée qui est prise en compte.
     * @param {boolean} [ignorePassage=false] - Si vrai, renvoie undefined si l'arrêt n'est qu'un passage.
     * @param {boolean} [ignoreDeparture=false] - Si vrai, ignore l'heure de départ.
     * @param {DateTime} [reference] - Heure de référence pour les heures relatives.
     * @returns {DateTime | undefined} - Heure la plus petite à l'arrêt,
     *  ou undefined si aucune heure n'est lue.
     */
    public getTime(
        {
            ignoreArrival = false,
            ignorePassage = false,
            ignoreDeparture = false,
            reference
        }: {
            ignoreArrival?: boolean,
            ignorePassage?: boolean,
            ignoreDeparture?: boolean,
            reference?: DateTime
        } = {}
    ): DateTime | undefined {

        if (ignorePassage && !!this._passageTime) return undefined;
        let time = this._arrivalTime;
        if (ignoreArrival || !this._arrivalTime) {
            time = (ignoreDeparture || !this._departureTime)
                ? this._passageTime
                : this._departureTime;
        }
        return (time && time!.isRelative && reference) ? time.resolveAgainst(reference) : time;
    }

    /**
     * Convertit les heures d'arrivée, de départ et de passage
     *  en temps relatifs par rapport à une référence.
     * Lève une erreur si le temps de référence est déjà relatif.
     * Lève un avertissement si les temps à convertir sont déjà relatifs.
     * Cependant pas d'erreur levée si le temps de référence et toutes les heures sont déjà relatives.
     * @param {DateTime} reference - Référence à utiliser pour convertir les heures.
     */
    public convertToRelativeTime(
        reference: DateTime,
        { throwErrorIfAlreadyRelative = false }: { throwErrorIfAlreadyRelative?: boolean } = {}
    ): void {
 
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
     * @param {Nullable<Stop>} other - Autre arrêt à comparer.
     * @returns {boolean} - Vrai si les arrêts sont égaux, faux sinon.
     */
    public equalsTo(
        other: Nullable<Stop>
    ): boolean {
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
     * @param {Nullable<Stop>} other - Autre arrêt à comparer.
     * @returns {boolean} - Vrai si les arrêts sont égaux, faux sinon.
     */
    public includes(
        other: Nullable<Stop>
    ): boolean {
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
    public addTrack(
        track: string
    ): void {
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
    private static readonly SHEET = "Arrêts";               // Nom de la feuille 
    private static readonly TABLE = "Arrêts";               // Nom du tableau
    private static readonly START_CELL = "A1";              // Première cellule
    private static readonly DATABASE_COLUMNS = {            // Liste des colonnes avec leur emplacement
        path: 0,
        station: 1,
        stationAfterTurnaround: 2,
        arrivalTime: 3,
        departureTime: 4,
        passageTime: 5,
        tracks: 6,
        nextStation: 7
    } as const;
    
    // Constantes de lecture du tableau à importer
    private static readonly IMPORT_SHEET = "Import arrêts";     // Nom de la feuille 
    private static readonly IMPORT_TABLE = "Import_arrêts";     // Nom du tableau
    private static readonly IMPORT_COLUMNS = {                  // Liste des colonnes avec leur emplacement
        trainNumber: 0,
        date: 1,
        service: 2,
        days: 3,
        station: 4,
        arrivalTime: 5,
        departureTime: 6,
        passageTime: 7,
        tracks: 8
    } as const;

    // Constantes des définitions des données de la classe chargées et imprimées dans les colonnes Excel

    /**
     * Liste des définitions générales :
     *  - clé,
     *  - gare,
     *  - gare après retournement,
     *  - heure d'arrivée,
     *  - heure de départ,
     *  - heure de passage,
     *  - voies,
     *  - gare suivante.
     */
    private static readonly COLUMN_DEFINITIONS: TableColumns<Stop> = {

        path: {
            header: "Parcours",
            type: "string",
            required: true,
            // print: value => (value as Path)?.key,
            format: { width: 40 },
            sort: { order: 1 }
        },
    
        station: {
            header: "Gare",
            type: "string",
            required: true,
            format: { width: 100 }
        },
    
        stationAfterTurnaround: {
            header: "Gare après rebroussement",
            type: "string",
            format: { width: 100 }
        },
    
        arrivalTime: {
            header: "Arrivée",
            type: "number",
            format: { 
                numberFormat: "hh:mm:ss",
                width: 120
            }
        },
    
        departureTime: {
            header: "Départ",
            type: "number",
            format: {
                numberFormat: "hh:mm:ss",
                width: 120
            }
        },
    
        passageTime: {
            header: "Passage",
            type: "number",
            format: {
                numberFormat: "hh:mm:ss",
                width: 120
            }
        },
    
        tracks: {
            header: "Voie",
            type: "string",
            load: TableSerializer.loadArray,
            print: TableSerializer.printArray,
            format: { width: 100 }
        },
    
        nextStation: {
            header: "Gare suivante",
            type: "string",
            format: { width: 100 }
        }
    };

    /**
     * Liste des définitions spécifiques à l'import,
     *  en surcharge des définitions générales :
     *  - numéro de train,
     *  - date,
     *  - service,
     *  - jours de circulation,
     *  - gare après retournement.
     */
    private static readonly IMPORT_COLUMN_DEFINITIONS: TableColumns<Stop> = {

        ...this.COLUMN_DEFINITIONS,
    
        trainNumber: {
            header: "N° origine",
            type: "string",
            required: true
        },
    
        date: {
            header: "Date",
            type: "number",
            required: true
        },
    
        service: {
            header: "Service",
            type: "string"
        },
    
        days: {
            header: "Jours de circulation",
            type: "string"
        }
    };

    /**
     * Charge les arrêts.
     * Les arrêts sont stockés dans la propriété "stops" des parcours correspondants.
     * Si un train n'existe pas, un message d'erreur est affiché.
     */
    public static load(): void {

        Log.startTimer(`${this.name}.load()`);

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête

                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    Stop,
                    StopParams
                >({
                    row,
                    columns: this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: "path"
                });
                if (!params) continue;
                
                // Instancie l'objet Stop.
                const stop = new Stop({
                    ...params,
                    areRelativeTimes: true }
                );

                // Ajoute l'arrêt au parcours.
                stop.path?.stops.push(stop);
            }

        } catch (e) {
            throw new Error(`Stops.load (ligne ${excelRow}) : ${e}`);
        }

        Log.timer(`${this.name}.load()`);
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

        Log.startTimer(`${this.name}.import()`);

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
;
                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    Stop,
                    StopParams
                >({
                    row,
                    columns: this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: "path"
                });
                if (!params) continue;
    
                // Instancie l'objet Stop.
                // const stop = new Stop({
                //     ...params,
                //     areRelativeTimes: true }
                // );

                // // Ajoute l'arrêt au parcours.
                // stop.path.stops.push(stop);
            }
        } catch (e) {
            throw new Error(`Stops.import (ligne ${excelRow}) : ${e}`);
        }

        Log.timer(`${this.name}.import()`);
    }
 
    /**
     * Sauvegarde la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell=this.START_CELL] - Adresse de la cellule de départ pour le tableau.
     */
    public static save(
        {
            sheetName = this.SHEET,
            tableName = this.TABLE,
            startCell = this.START_CELL
        }: {
            sheetName?: string,
            tableName?: string,
            startCell?: string
        } = {}
    ): void {

        Log.startTimer(`${this.name}.save()`);

        // Récupère les arrêts dans les parcours.
        const stops: Stop[] = [];
    
        for (const path of Paths.values()) {
            stops.push(...path.stops);
        }
    
        // Sauvegarde la base de données.
        TableSerializer.print({
            entities: stops,
            columns: this.DATABASE_COLUMNS,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        Log.timer(`${this.name}.save()`);
    }
}

/**
 * Type PathInput contenant les types de données acceptés
 *  pour créer ou appeler un objet Stop.
 */

type PathInput = Input<Path, string>;

/**
 * Interface PathParams contenant les paramètres de la classe Path.
 * @param {string} key - Clé du parcours.
 * @param {ParityInput} parity - Parité du parcours.
 * @param {ParityInput} lineDirection - Parité de la ligne.
 * @param {string} missionCode - Code mission.
 * @param {string} name - Nom du parcours.
 * @param {string} signature - Signature du parcours.
 * @param {Stop[]} stops - Arrêts du parcours.
 * @param {number} stopsChecked - Nombre d'arrêts du parcours.
 */
interface PathParams {
    key?: string;
    parity?: ParityInput;
    lineDirection?: ParityInput;
    missionCode?: string;
    name?: string;
    signature?: string;
    stops?: Stop[];
    stopsChecked?: number;
}

/**
 * Interface PathFromTerminalsParams contenant les paramètres pour créer un parcours
 *  à partir de sa gare origine et sa gare destination.
 * @param {StationInput} origin - Gare origine.
 * @param {DateTimeInput} departureTime - Heure de départ.
 * @param {StationInput} destination - Gare de destination.
 * @param {DateTimeInput} arrivalTime - Heure d'arrivée.
 * @param {boolean} areRelativeTimes - Indique si les heures sont relatives.
 * @param {boolean} findPath - Indique si le parcours doit être calculé.
 * @param {string} missionCode - Code mission.
 * @param {string} name - Nom du parcours.
 * @param {string} signature - Signature du parcours.
 */
interface PathFromTerminalsParams {
    origin: StationInput;
    departureTime: DateTimeInput;
    destination: StationInput;
    arrivalTime: DateTimeInput;
    areRelativeTimes?: boolean;
    findPath?: boolean;
    missionCode?: string;
    name?: string;
    signature?: string;
}

/**
 * Classe Path définissant le parcours d'un train, avec ses gares et temps de passage
 *  par rapport à la gare origine.
 */
class Path {

    // Résultats de la vérification du parcours
    public static readonly  UNCHECKED = 0;          // Parcours non vérifié
    public static readonly  ONLY_ORIGIN_AND_DESTINATION = 1;   
                                                    // Parcours avec uniquement les gares origine et destination
    public static readonly  WITH_VIA_STOPS = 2;     // Parcours avec gares intermédiaires
    public static readonly  FULL_PATH = 3;          // Parcours complet calculé par chainage de connexions 
    public static readonly  ERROR_WITH_STOPS = -1;  // Parcours avec erreur
 
    // Propriétés de la classe Path
    public key: string;                             // Clé du parcours
    public parity: Parity;                          // Parité du parcours
                                                    //  (synthèse des parités pour chaque gare)
    public lineDirection: Parity;                   // Direction du parcours sur la ligne
                                                    //  (donnée par une parité globale)
    public missionCode: string;                     // Code mission des trains du parcours (facultatif)
    public name: string;                            // Nom du parcours (facultatif)
    private _signature: string;                     // Signature du parcours : gares définissant le parcours
                                                    //  séparées par '>' pour les arrêts ordonnés et
                                                    //  par ';' si leur ordre de parcours est laissé libre
    private _routeStations?: StationWithParity[][] = [];     
                                                    // Tableau des gares ou groupes de gares d'arrêts
                                                    //  du parcours définis dans la signature 
    public stops: Stop[] = [];                      // Gares d'arrêt ou gares de passage du parcours
    private _stopsIndex: Map<string, Stop> = new Map();
                                                    // Dictionnaire des arrêts référencés
                                                    //  par leur clé (abréviation_parité)
    private _stopPosition: Map<string, number> = new Map();
                                                    // Dictionnaire de la position des arrêts
                                                    //  dans le parcours (référencés par leur clé)
    public stopsChecked: number = Path.UNCHECKED;   // Résultat de la vérification du parcours
                                                    //  (0 si non vérifié)

    /**
     * Constructeur d'un parcours.
     * @param {PathParams} params - Paramètres du parcours.
     */
    public constructor(params: PathParams = {}) {
        this.key = params.key ?? "";
        this.parity = Parity.from(params.parity, { doubleParityAllowed: true });
        this.lineDirection = Parity.from(params.lineDirection, { doubleParityAllowed: true });
        this.missionCode = params.missionCode ?? "";
        this.name = params.name ?? "";
        this._signature = params.signature ?? "";
        this.stops = params.stops ?? [];
        this.stopsChecked = params.stopsChecked ?? Path.UNCHECKED;

        for (const stop of this.stops) {
            stop.path = this;
        }
    }

    /**
     * Retourne l'arrêt d'origine du parcours.
     * @returns {Stop | undefined} - L'arrêt d'origine, ou undefined si le parcours n'a pas d'arrêt.
     */
    public get origin(): Stop | undefined {
        return this.stops[0];
    }

    /**
     * Retourne l'arrêt de destination du parcours.
     * @returns {Stop | undefined} - L'arrêt de destination,
     *  ou undefined si le parcours n'a pas d'arrêt de destination.
     */
    public get destination(): Stop | undefined {
        return this.stops[this.stops.length - 1] ;
    }

    /**
     * Retourne la signature du parcours, qui est la concaténation
     *  des noms des gares d'arrêt du parcours, précédés de "@"
     *  si l'ordre de passage n'est pas imposé.
     * @returns {string} - Signature du parcours.
     */
    public get signature(): string {
        return this._signature;
    }

    /**
     * Retourne le tableau des gares d'arrêt du parcours.
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
                        .filter(s => s !== undefined)
                );
        }
 
        return this._routeStations ?? [];
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Clé du parcours.
     */
    public toString(): string {
        return this.key;
    }

    /**
     * Retourne le parcours Path correspondant au paramètre path.
     * Si path est déjà un objet Path, il est retourné tel quel.
     * Si path est un string, il est considéré comme le clé du parcours et
     *  l'objet Path correspondant est retourné s'il existe, sinon undefined est retourné.
     * @param {Nullable<PathInput>} value - Parcours à retourner,
     *  sous forme d'objet Path ou de clé string.
     * @returns {Path | undefined} - Parcours Path correspondant, ou undefined si le clé n'existe pas.
     */
    public static from(
        value: Nullable<PathInput>
    ): Path | undefined {
        if (value == null || value === "") return undefined;
        if (value instanceof Path) return value;
        return Paths.get(value!);
    }

    /**
     * Crée un parcours Path à partir des gares d'origine et de destination,
     *  ainsi que de leur heures de départ et d'arrivée.
     * @param {StationInput} origin - Nom de la gare d'origine.
     * @param {DateTimeInput} departureTime - Heure de départ à la gare d'origine.
     * @param {StationInput} destination - Nom de la gare de destination.
     * @param {DateTimeInput} arrivalTime - Heure d'arrivée à la gare de destination.
     * @param {boolean} [areRelativeTimes=false] - Si vrai, les heures de départ et d'arrivée
     *  sont considérées comme relatives.
     * @param {string} [missionCode=""] - Code mission des trains du parcours (facultatif).
     * @param {string} [name=""] - Nom du parcours (facultatif).
     * @param {string} [signature=""] - Signature du parcours (facultatif).
     * @returns {Path} - Un objet Path représentant le parcours.
     */
    public static fromTerminals(params: PathFromTerminalsParams): Path {

        const originObj = StationWithParity.from(params.origin);
        if (!originObj) throw new Error(`Gare d'origine ${params.origin} incorrecte.`);
        const departureTimeObj = DateTime.from(params.departureTime, { isRelative: params.areRelativeTimes });
        if (!departureTimeObj) throw new Error(`Heure de départ ${params.departureTime} incorrecte.`);
        const destinationObj = StationWithParity.from(params.destination);
        if (!destinationObj) throw new Error(`Gare de destination ${params.destination} incorrecte.`);
        const arrivalTimeObj = DateTime.from(params.arrivalTime, { isRelative: params.areRelativeTimes });
        if (!arrivalTimeObj) throw new Error(`Heure d'arrivée ${params.arrivalTime} incorrecte.`);

        const s1 = new Stop({
            station: originObj,
            stationAfterTurnaround: undefined,
            arrivalTime: undefined,
            departureTime: departureTimeObj,
            passageTime: undefined,
            areRelativeTimes: params.areRelativeTimes
        });
        const s2 = new Stop({
            station: destinationObj,
            stationAfterTurnaround: undefined,
            arrivalTime: arrivalTimeObj,
            departureTime: undefined,
            passageTime: undefined,
            areRelativeTimes: params.areRelativeTimes
        });
 
        const path = Paths.create({
            missionCode: params.missionCode,
            name: params.name,
            signature: params.signature,
            stops: [s1, s2]
        });
 
        // Retourne directement le parcours s'il existait déjà
        if (path.stopsChecked !== Path.UNCHECKED) return path;

        // Convertit les horaires en relatifs
        path.convertStopsToRelative();

        // Finalise le parcours en constituant les index, signature et en le vérifiant
        path.stopsChecked = Path.ONLY_ORIGIN_AND_DESTINATION;
        path.finalize();

        // Calcule le parcours si demandé
        if (params.findPath) path.findPath();

        return path;
    }

    /**
     * Retourne le radical de la clé du parcours constitué de
     *  origine_destination_codeMission_nomDuParcours (si ces valeurs existent).
     * @returns {string} - Radical de la clé du parcours.
     */
    public buildRadical(): string {
        const origin = this.origin?.stationAbbreviation ?? "";
        const dest = this.destination?.stationAbbreviation ?? "";
        if (!origin || !dest) throw new Error(`Pas d'origine ou de destination valide dans le parcours.`);
 
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
     * @throws {Error} - Si les trains du parcours sont déjà passé par l'arrêt
     *  et que erase est faux.
     */
    public addStop(
        stop: Stop,
        {
            finalize = true,
            erase = false
        }: {
            finalize?: boolean,
            erase?: boolean
        } = {}
    ): void {

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

        // Finalise le parcours en triant les arrêts, en constituant index et signature et en le vérifiant
        if (finalize) this.finalize();
    } 

    /**
     * Supprime un arrêt du parcours.
     */
    private removeStop(
        station: Station | StationWithParity | string
    ): void {
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
    private isStopInSignature(
        stop: Stop
    ): boolean {
        if (!this._signature) return false;
        return this.routeStations.some(group =>
            group.some(station => stop.station.includes(station))
        );
    }
 
    /**
     * Finalise le parcours et ses arrêts en :
     *  - triant les arrêts,
     *  - reconstruisant l'index et la signature,
     *  - recalculant les parités du parcours,
     *  - vérifiant le parcours.
     */
    public finalize(): void {

        if (this.stops.length === 0) return;

        this.orderStops();
        this.rebuildStopIndex();
        this.rebuildStopPosition();
 
        if (this.stopsChecked === Path.FULL_PATH) {
            this.recomputeParities();
        }
 
        if (!this._signature) {
            this.buildSignatureFromStops();
        }

        this.check();
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

        this.parity = Parity.undefined({ doubleParityAllowed: true });
        this.lineDirection = Parity.undefined({ doubleParityAllowed: true });
 
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
     * @param {StationInput | number} station - La gare à chercher.
     * @returns {Stop | undefined} - L'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     */
    public getStop(
        station: StationInput | number | undefined
    ): Stop | undefined {

        if (station == null || station === "") return undefined;

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

                const odd = StationWithParity.from(swp, { parity: Parity.odd() })!;
                const even = StationWithParity.from(swp, { parity: Parity.even() })!;

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
            .filter(s => s !== undefined);
        const parents: StationWithParity[] = parentStations
            .map(s => StationWithParity.from(s, { parity: stationObj.parity }))
            .filter(s => !!s);
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
     * @param {StopInput} stop - L'arrêt ou la gare à chercher.
     * @returns {Stop | undefined} - L'arrêt suivant, ou undefined si la gare est la dernière.
     */
    public nextStop(
        stop: StopInput
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
     * @param {StopInput} stop - L'arrêt ou la gare à chercher.
     * @returns {Stop | undefined} - L'arrêt précédent, ou undefined si la gare est la première.
     */
    public previousStop(
        stop: StopInput
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
            if (this.stopsChecked === Path.ONLY_ORIGIN_AND_DESTINATION) {
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
        if (firstStop.isIntermediateStop) {
            throw new Error(`Le premier arrêt ne peut pas contenir d'heure d'arrivée`
                + ` mais uniquement une heure de départ.`);
        }
        // Vérifie l'absence d'heure de départ dans le dernier arrêt.
        if (lastStop.isIntermediateStop) {
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
                case Path.ONLY_ORIGIN_AND_DESTINATION:
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
        let stopFromSigToFind: Map<string, StationWithParity> = new Map();

        for (let i = 1; i < this.stops.length; i++) {
            switch (this.stopsChecked) {
                case Path.ERROR_WITH_STOPS:
                case Path.UNCHECKED:
                    return;
                case Path.ONLY_ORIGIN_AND_DESTINATION:
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
            if ((i < this.stops.length - 1) && !this.stops[i].isIntermediateStop) {
                throw new Error(`L'arrêt à la gare de ${this.stops[i]} doit comporter`
                    + ` une heure d'arrivée et une heure de départ, ou une heure de passage.`);
            }
            // Vérifie que l'heure de passage est postérieure au passage précedent.
            if (this.stops[i].getTime()!.compareTo(this.stops[i - 1].getTime({ ignoreArrival: true })!) <= 0) {
                throw new Error(`L'heure d'arrivée ou de passage`
                    + ` ${this.stops[i].getTime()!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
                    + ` à la gare de ${this.stops[i]}`
                    + ` doit être postérieure à l'heure de passage ou de départ`
                    + ` ${this.stops[i - 1].getTime({ ignoreArrival: true })!.format(DateTime.TIME_FORMAT_WITH_SECONDS)}`
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
        let ref = refList ? refList[0] : null;
        try {
            if (ref && ref.stopsChecked !== Path.FULL_PATH) {
                Log.warn(`Le parcours de référence ${ref}`
                    + ` n'est pas complet et ne peut pas servir de base de calcul pour ${this}.`);
                Paths.signatureIndex.delete(this.signature);
                ref = null;
            }
            
            connections = ref 
                ? ref.buildConnectionsFromStops()
                : this.shortestPathThrough();

            this.buildStopsFromConnections(connections);
        } catch (e) {
            throw new Error(`Calcul du parcours ${this} : ${e}`);
        }

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
            throw new Error(`RouteStations invalide`);
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
    public buildStopsFromConnections(
        connections: Connection[]
    ): void {

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
        const firstStop = new Stop({
            path: this,
            station: firstConnection.from,
            stationAfterTurnaround: undefined,
            arrivalTime: undefined,
            departureTime: firstExisting.departureTime,
            passageTime: undefined,
            areRelativeTimes,
            tracks: firstExisting.tracks
        });
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
                // Si l'arrêt est un arrêt intermédiaire
                //  et que le temps d'arrêt est inférieur au temps minimal de retournement,
                //  ajuste l'heure de départ pour respecter cette durée.
                if (lastStop.isIntermediateStop 
                    && lastStop.departureTime!
                        .relativeTo(lastStop.arrivalTime!)
                        .compareTo(Params.get('turnaroundTime'))
                    < 0)
                {
                    lastStop.setTimes({
                        departureTime: Params.get('turnaroundTime').resolveAgainst(lastStop.arrivalTime!)
                    });
                }
                if (lastStop.withNonStopPassage) {
                    lastStop.setTimes({
                        arrivalTime: lastStop.passageTime,
                        passageTime: undefined,
                        departureTime: Params.get('turnaroundTime').resolveAgainst(lastStop.passageTime!)
                })};
            }

            buffer.push(c);
            const stop = this.getStop(c.to);

            // Si l'arrêt est connu avec horaires,
            //  calcule toutes les connexions depuis le dernier arrêt connu
            //  pour déterminer l'horaire de passage dans chacune des gares parcourues.
            if (stop) {

                // Récupère les horaires aux deux arrêts connus.
                const startStop = newStops[newStops.length - 1];
                const startTime = startStop.getTime({ ignoreArrival: true });
                if (!startTime) throw new Error(`L'arrêt ${startStop} n'a pas d'heure de départ.`);
                const endTime = stop.getTime({ ignoreArrival: false });
                if (!endTime) throw new Error(`L'arrêt ${stop} n'a pas d'heure d'arrivée.`);
                if (endTime.excelValue <= startTime.excelValue) {
                    // Arrêt trouvé déjà parcouru (dans le cas d'un deuxième passage dans l'autre sens).
                    continue;
                }

                // Calcule le(s) temps de retournement à retrancher du temps de parcours total,
                //  sauf pour la dernière connexion du buffer (le temps de retournement sera pris
                //  en compte dans le temps d'arrêt du dernier arrêt connu trouvé)
                const totalTurnarounds = buffer
                    .slice(0, -1)
                    .filter(x => x.withTurnaround).length;

                // Calcule le nombre d'accélérations ou de décélérations économisées
                //  dans les gares où le train passe sans arrêt, 
                //  gagnant ainsi du temps par rapport au temps de parcours d'une connexion.
                const stationLeavingWithoutStop = buffer.length - 1 - totalTurnarounds + (startStop.withNonStopPassage ? 1 : 0);
                const stationArrivingWithoutStop = buffer.length - 1 - totalTurnarounds + (lastStop.withNonStopPassage ? 1 : 0);   

                // Calcule le temps de parcours entre les deux arrêts connus,
                //  déduction faite des temps de retournement.
                const interpolatedTime = endTime.excelValue
                    - startTime.excelValue
                    - totalTurnarounds * Params.get('turnaroundTime').excelValue;
                    // + (stationLeavingWithoutStop + stationArrivingWithoutStop)
                    //     * Params.accelerationTime.excelValue
 
                // Calcule la somme des temps de parcours de chaque connexion.
                // Cette base de temps permet de calculer les temps de parcours réels au prorata
                //  des temps de parcours des connexions.
                const totalTime = buffer.reduce((sum, x) =>
                    sum + x.time.excelValue, 0);
                const ratio = interpolatedTime / totalTime;
                
                // Calcule le gain de temps pour chaque gare avec passage,
                //  gagné sur le fait qu'il n'y a ni décélération, ni accélération.
                // const timeSaving = (totalTime - interpolatedTime) / stationsWithNonStopService;

                // Parcourt les connexions du buffer pour créer les arrêts.
                let elapsed = 0;
                let previousStopIsNonStop = false;
                for (let i = 0; i < buffer.length; i++) {
                    const bc = buffer[i];

                    // Vérifie si la connexion suivante est un rebroussement
                    //  (donc à partir de la deuxième connexion du buffer).
                    // La première connexion du buffer ne peut pas être une connexion de rebroussement,
                    //  sinon la connexion aurait été prise en compte
                    //  au début de la première boucle for (buffer vide).
                    const stopWithTurnaround = buffer[i + 1]?.withTurnaround;

                    // Vérifie si la connexion est un rebroussement
                    if (bc.withTurnaround) {
                        // Si la connexion est un retournement, le dernier arrêt est forcement
                        //  un arrêt calculé, donc avec une heure de passage
                        //  (sinon la connexion aurait été sautée au début de la première boucle for).
                        // Donc le dernier arrêt est transformé en arrêt avec rebroussement, 
                        //  avec pour durée le temps de retournement par défaut.
                        lastStop.stationAfterTurnaround = bc.to;
                        elapsed += Params.get('turnaroundTime').excelValue;
                        continue;
                    } else {
                        elapsed += bc.time.excelValue * ratio;
                    }
 
                    // Calcule l'heure de passage.
                    // Chaque connexion voit son temps de parcours réduit de la valeur de timeSaving,
                    //  sauf la première et la dernière connexion réduites de la moitié de timeSaving
                    const interpolated = startTime.excelValue + elapsed;
                    lastStop = new Stop({
                        path: this,
                        station: bc.to,
                        stationAfterTurnaround: undefined,
                        arrivalTime: undefined,
                        departureTime: undefined,
                        passageTime: interpolated,
                        areRelativeTimes
                    }); 
                    newStops.push(lastStop);
                }

                // Modifie le dernier arrêt avec les horaires d'arrivée et de départ, et les voies.
                if (stop.arrivalTime) {
                    lastStop.setTimes({
                        arrivalTime: stop.arrivalTime.excelValue,
                        departureTime: stop.departureTime?.excelValue,
                        areRelativeTimes
                    });
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
            this.addStop(stop, { finalize: false });
        }

        // Finalise le parcours en triant les arrêts, en constituant index et signature et en le vérifiant
        this.finalize();
    }
}

/**
 * Classe Paths contenant la liste des parcours.
 */
class Paths {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Parcours";             // Nom de la feuille 
    private static readonly TABLE = "Parcours";             // Nom du tableau
    private static readonly START_CELL = "A1";              // Première cellule
    private static readonly DATABASE_COLUMNS = {            // Liste des colonnes avec leur emplacement
        key: 0,
        parity: 1,
        lineDirection: 2,
        missionCode: 3,
        name: 4,
        signature: 5,
        stopsChecked: 6
    } as const;

    /**
     * Liste des définitions générales :
     *  - clé,
     *  - parité du parcours,
     *  - parité de ligne,
     *  - code mission,
     *  - nom du parcours,
     *  - signature du parcours,
     *  - état de vérification.
     */
    private static readonly COLUMN_DEFINITIONS: TableColumns<Path> = {

        key: {
            header: "Clé",
            type: "string",
            required: true,
            format: { width: 40 },
            sort: { order: 1 }
        },

        parity: {
            header: "Parité du parcours",
            type: "string",
            required: true,
            format: { width: 40 }
        },

        lineDirection: {
            header: "Parité de ligne",
            type: "string",
            required: true,
            format: { width: 40 }
        },

        missionCode: {
            header: "Code mission",
            type: "string",
            format: { width: 80 }
        },

        name: {
            header: "Nom",
            type: "string",
            format: { width: 120 }
        },

        signature: {
            header: "Route",
            type: "string",
            format: { width: 120 }
        },

        stopsChecked: {
            header: "Etat de vérification",
            type: "number",
            format: { width: 40 }
        }
    };

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
    public static has(
        key: string
    ): boolean {
        return this.map.has(key);
    }

    /**
     * Retourne le parcours correspondant à la clé donnée.
     * @param {string} key - Clé du parcours.
     * @returns {Path | undefined} - Parcours correspondant, ou undefined si la clé n'existe pas.
     */
    public static get(
        key: string
    ): Path | undefined {
        return this.map.get(key);
    }

    /**
     * Ajoute un nouveau parcours dans la base de données, référencé par sa clé.
     * Si le parcours est déjà présent, une erreur est levée.
     * @param {Path} path - Parcours à enregistrer.
     * @throws {Error} - Si le parcours est déjà présent dans la base de données.
     */
    private static set(
        value: Path
    ): void {
        if (this.has(value.key)) {
            throw new Error(`Le parcours ${value} est déjà présent`
                + ` dans la base de données.`);
        }
        this.map.set(value.key, value);
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
     * @param {PathParams} params - Paramètres du parcours.
     * @returns {Path} - Parcours créé.
     */
    public static create(params: PathParams): Path {

        // Instancie l'objet Path.
        const path = new Path(params);

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
    private static insert(
        path: Path
    ): Path {

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
 
            const numberMap: Map<number, Path> = new Map();
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
        if (letterKey == null) {
            letterKey = this.nextLetter(radicalMap);
 
            const numberMap: Map<number, Path> = new Map();
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
    public static delete(
        path: Path
    ): void {

        // Supprime l'objet Path de la base de données, indexé par sa clé.
        this.map.delete(path.key);

        // Détermine les composantes de la clé
        const radical = this.extractRadical(path.key);
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
    private static indexToLetters(
        index: number
    ): string {

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
    private static extractRadical(
        key: string
    ): string {
        return key.split("~")[0].split("#")[0];
    }

    /**
     * Extrait la lettre de la clé d'un parcours (chaîne de la forme "~X" où X est la lettre du suffixe).
     * @param {string} key - Clé du parcours.
     * @returns {string} - Lettre du suffixe (ou une chaîne vide si la clé n'a pas de suffixe lettre).
     */
    private static extractLetter(
        key: string
    ): string {
        const m = key.match(/~([A-Z]+)/);
        return m ? m[1] : "";
    }
 
    /**
     * Extrait le numéro de la clé d'un parcours
     *  (chaîne de la forme "#X" où est le numéro du suffixe numérique).
     * @param {string} key - Clé du parcours.
     * @returns {number} - Numéro du suffixe numérique (ou 0 si la clé n'a pas de suffixe numérique).
     */
    private static extractNumber(
        key: string
    ): number {
        const m = key.match(/#(\d+)/);
        return m ? Number(m[1]) : 0;
    }

    /**
     * Construit une clé de parcours à partir d'un radical,
     *  d'une lettre de suffixe et d'un numéro de suffixe.
     * La clé est composée de la forme "radical~lettre#nombre" avec
     *  - un suffixe lettre optionnel précédé de "~" pour une signature différente avec le même radical,
     *  - un suffixe numérique optionnel précédé de "#" pour des arrêts avec horaires différents
     *     pour un radical et une signature identique.
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
        if (number> 0) key += `#${number}`;
 
        return key;
    }

    /**
     * Cherche si un parcours existe déjà avec un même radical et une même signature.
     * Si oui donne le suffixe lettre de ce parcours, sinon renvoie null.
     * @param {Map<string, Map<number, Path>>} radicalMap - Map des parcours ayant le même radical
     *  que celui du parcours pour lequel la recherche est faite.
     * @param {string} signature - Signature du parcours à chercher.
     * @returns {Nullable<string>} - Lettre du suffixe de la clé du parcours trouvé
     *  (même radical et même signature).
     */
    private static findLetterBySignature(
        radicalMap: Map<string, Map<number, Path>>,
        signature: string
    ): Nullable<string> {
 
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
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les connexions si elles ne sont pas encore chargées.
        Connections.load();

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;
    
                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2;
    
                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    Path,
                    PathParams
                >({
                    row,
                    columns: this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: "key"
                });    
                if (!params) continue;
                
                // Crée l'objet et l'insère dans la base de données.
                const path = this.create(params);
            }

        } catch (e) {
            throw new Error(`Paths.load (ligne ${excelRow}) : ${e}`);
        }

        Log.timer(`${this.name}.load()`);

        // Charge les arrêts des parcours.
        Stops.load();

        // Vérifie si les parcours sont valides.
        Log.startTimer(`${this.name}.finalize()`);
        try {
            for (const path of this.values()) {
                path.finalize();
            }
        } catch (e) {
            throw new Error(`Paths.load : ${e}`);
        }

        Log.timer(`${this.name}.finalize()`);
    }

    /**
     * Sauvegarde la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell=this.START_CELL] - Adresse de la cellule de départ pour le tableau.
     */
    public static save(
        {
            sheetName = this.SHEET,
            tableName = this.TABLE,
            startCell = this.START_CELL
        }: {
            sheetName?: string,
            tableName?: string,
            startCell?: string
        } = {}
    ): void {
    
        Log.startTimer(`${this.name}.save()`);

        TableSerializer.print({
            entities: this.values(),
            columns: this.DATABASE_COLUMNS,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        Log.timer(`${this.name}.save()`);
    
        Stops.save();
    }
}
/**
 * Type ReuseTarget reprenant les deux types Train et TrainPath qui peuvent constituer une réutilisation.
 */
type ReuseTarget = Train | TrainPath;

/**
 * Type ReuseParams contenant les paramètres d'une réutilisation.
 * @param {string} reuseKey - Clé du train ou du sillon de réutilisation.
 * @param {number} position - Position de l'élément.
 * @param {Nullable<StationInput>} maintenanceCenter - Garage en technicentre.
 */
type ReuseParams = {
    reuseKey?: string;
    position?: number;
    maintenanceCenter?: Nullable<StationInput>;
};

/**
 * Classe Reuse contenant la réutilisation d'un élément d'un train ou d'un sillon. 
 */
class Reuse<T extends ReuseTarget> {

    public readonly type: typeof Train | typeof TrainPath;  // Type de réutilisation (train ou sillon)
    public readonly reuseKey?: string;                      // Clé du train ou du sillon de réutilisation
    public readonly position: number;                       // Position de l'élément
                                                            //  dans le train/sillon de réutilisation :
                                                            //  - positif pour la réutilisation suivante
                                                            //  - négatif pour la réutilisation précédente
                                                            //  - 0 pour un garage en technicentre
    private _target?: Nullable<T>;                          // Train ou sillon de réutilisation     
    private _isMouvement?: Nullable<boolean>;               // Indique si le train/sillon de réutilisation
                                                            //  est une évolution
    private _isEmptyPassenger?: Nullable<boolean>;          // Indique si le train/sillon de réutilisation
                                                            //  est vide voyageur
    private _reuse?: Nullable<Reuse<T>>;                    // Réutilisation suivante
    public readonly maintenanceCenter?: Nullable<Station>;  // Gare du technicentre si la réutilisation
                                                            //  est un garage au technicentre

    /**
     * Constructeur privé de la classe Reuse.
     * @param {typeof Train | typeof TrainPath} type - Type de réutilisation (train ou sillon).
     * @param {ReuseParams} params - Paramètres de la réutilisation.
     */
    private constructor(
        type: typeof Train | typeof TrainPath,
        params: ReuseParams
    ) {
        this.type = type;
        this.reuseKey = params.reuseKey;
        
        if (!Number.isInteger(params.position)) {
            throw new Error(`La position de réutilisation doit être un entier.`);
        }
        if (this.reuseKey) {
            if (!params.position) throw new Error(`Une réutilisation associée à un train`
                + ` doit comporter une position non nulle de l'élément dans le train`
                + ` (positive pour la réutilisation suivante, negative pour la réutilisation précédente).`);
            this.position = params.position;
            this.maintenanceCenter = null;
        } else {
            this.position = 0;
            this.maintenanceCenter = Station.from(params.maintenanceCenter);
            if (!this.maintenanceCenter) {
                throw new Error(`Une réutilisation de garage doit comporter un technicentre valide.`);
            }
        }
    }

    /**
     * Retourne une instance de Reuse à partir d'une valeur qui peut être :
     *  - une clé de réutilisation,
     *  - les paramètres ReuseParams d'une réutilisation.
     * @param {Nullable<PathInput>} value - Parcours à retourner,
     *  sous forme d'objet Path ou de clé string.
     * @returns {Path | undefined} - Parcours Path correspondant, ou undefined si le clé n'existe pas.
     */
    public static from(
        { type, key, params }: { type: (typeof Train), key?: string, params?: ReuseParams }
    ): Reuse<Train> | undefined;
    
    public static from(
        { type, key, params }: { type: (typeof TrainPath), key?: string, params?: ReuseParams }
    ): Reuse<TrainPath> | undefined;

    public static from(
        {
            type,
            key,
            params
        }: {
            type: (typeof Train | typeof TrainPath),
            key?: string,
            params?: ReuseParams
        }
    ): Reuse<ReuseTarget> | undefined {

        // Instancie la réutilisation à partir de la clé.
        if (!!key) {

            if (params?.reuseKey !== undefined
                || params?.position !== undefined
                || params?.maintenanceCenter !== undefined
            ) {
                throw new Error(`Impossible de fournir simultanément key`
                    + ` et les propriétés d'un réutilisation.`);
            }

            const parsed = Reuse.parseKey(key);
            if (!parsed) return undefined;
            
            return new Reuse(
                type,
                parsed
            );
        }

        // Instancie la réutilisation à partir des paramètres de la réutilisation
        if (!params ||(!params.reuseKey && !params.maintenanceCenter)) return undefined;
        return new Reuse(
            type,
            params
        );
    }

    /**
     * Retourne les paramètres d'une réutilisation à partir d'une clé de réutilisation.
     * @param {string} key - Clé de réutilisation.
     * @returns {ReuseParams | undefined} - Paramètres de la réutilisation.
     */
    private static parseKey(
        key: string
    ): ReuseParams | undefined {

        if (!key) return undefined;

        // Clé d'un garage
        if (key!.slice(0, 2) === "G@") {
            return { maintenanceCenter: key.slice(2) };
        }

        // Clé d'une réutilisation avec train ou sillon
        const [reuseKey, position] = key.split('#');
        if (!position) return undefined;
        const parsedPosition = Number(position);
        if (!Number.isInteger(parsedPosition)) return undefined;
        return { reuseKey, position: parsedPosition };
    }

    /**
     * Retourne l'objet du train ou du sillon de réutilisation.
     * @returns {Nullable<T>} - Train ou sillon de réutilisation.
     */
    public get target(): Nullable<T> {
        if (this._target === undefined) {
            this._target = this.type.from(this.reuseKey) as T
                ?? null;
        }
        return this._target;
    }

    /**
     * Indique si le train ou le sillon de réutilisation est une évolution.
     * @returns {Nullable<boolean>} - Indique si le train ou le sillon de réutilisation est une évolution.
     */
    public get isMouvement(): Nullable<boolean> {
        if (this._isMouvement === undefined) {
            this._isMouvement = this.target?.isMouvement
                ?? null;
        }
        return this._isMouvement;
    }

    /**
     * Indique si le train ou le sillon de réutilisation est W.
     * @returns {Nullable<boolean>} - Indique si le train ou le sillon de réutilisation est W.
     */
    public get isEmptyPassenger(): Nullable<boolean> {
        if (this._isEmptyPassenger === undefined) {
            this._isEmptyPassenger = this.target?.isEmptyPassenger
                ?? null;
        }
        return this._isEmptyPassenger;
    }

    /**
     * Retourne la réutilisation suivante.
     * @returns {Nullable<Reuse<T>>} - Réutilisation suivante.
     */
    public get reuse(): Nullable<Reuse<T>> {
        if (this._reuse === undefined) {
            this._reuse = this.target?.reusesMap?.get(this.position) as Reuse<T>
                ?? null;
        }
        return this._reuse;
    }

    /**
     * Retourne la clé du train ou du sillon de réutilisation,
     *  composée de la clé de la réutilisation et de la position de l'élément dans la réutilisation.
     * Si la réutilisation est un garage au technicentre, la clé est composée
     *  du code de la gare du technicentre et de la position de l'élément, précédés de "G@".
     * @returns {string} - Clé du train ou du sillon de réutilisation.
     */
    public get key(): string {
        return this.maintenanceCenter
            ? "G@" + this.maintenanceCenter.abbreviation
            : this.reuseKey 
                ? this.reuseKey + "#" + this.position
                : "?";
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Numéro de réutilisation, ou nom du technicentre, ou "?" si non défini.
     */
    public toString(): string {
        return this.maintenanceCenter
            ? this.maintenanceCenter.abbreviation
            : this.target?.number.format()
                ?? "?";
    }

    /**
     * Retourne la nième réutilisation valide, selon le rang indiqué,
     *  et selon si les évolutions ou les W sont acceptés ou non.
     * @param {number} [occurrence=1] - Occurence de la réutilisation à donner.
     * @param {boolean} [excludeMouvements=false] - Exclut les évolutions.
     * @param {boolean} [excludeEmptyPassenger=false] - Exclut les W.
     * @returns {Nullable<Reuse<T>>} - Première réutilisation valide.
     */
    public resolve(
        {
            occurrence = 1,
            excludeMouvements = false,
            excludeEmptyPassenger = false
        }: {
            occurrence?: number,
            excludeMouvements?: boolean,
            excludeEmptyPassenger?: boolean
        } = {}
    ): Reuse<T> | undefined {

        if (!Number.isInteger(occurrence) || occurrence < 1) {
            throw new Error(`Le nombre d'occurrences doit être un entier positif.`);
        }

        let reuse: Nullable<Reuse<T>> = this;

        const visited = new Set<Reuse<T>>();

        let o = occurrence;
        while (reuse) {

            // Protection anti-boucle
            if (visited.has(reuse)) {
                return undefined;
            }

            visited.add(reuse);

            const excluded = (excludeMouvements && reuse.isMouvement === true)
                || (excludeEmptyPassenger && reuse.isEmptyPassenger === true);

            if (!excluded) o--;
            if (o === 0 || reuse.maintenanceCenter) return reuse;

            reuse = reuse.reuse;
        }

        return undefined;
    }
}

/**
 * Type TrainInput réunissant les types de données acceptés
 *  pour créer ou appeler un objet Train.
 */
type TrainInput = Input<Train, string>;

interface TrainParams {
    key?: string;
    number: TrainNumberInput;
    date: DateTimeInput;
    service?: string;
    path: PathInput;
    units?: string[];
    reusesMap?: Map<number, Reuse<Train> | undefined>;
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

    // Propriétés de la classe Train
    public test = 123;
    public key: string;             // Clé du train
    public number: TrainNumber;     // Numéro du train
    public date: DateTime;          // Date et heure de départ du train
    public service: string;         // Service auquel le train est rattaché
    public path: Path;              // Parcours sur lequel le train circule
    public units: string[];         // Eléments (numéro de matériel) comptés à partir de 0
    public readonly reusesMap: Map<number, Reuse<Train> | undefined>;        
                                    // Map des réutilisations, selon la position de chaque élément
                                    //  - positif pour la réutilisation suivante
                                    //  - négatif pour la réutilisation précédente
    /**
     * Constructeur de la classe Train.
     * @param {TrainParams} params - Paramètres du train.
     */
    public constructor(params: TrainParams) {
       
        this.key = params.key ?? "";
        const numberObj = TrainNumber.from(params.number);
        if (!numberObj) {
            throw new Error(`Le numéro du train ${this} est invalide.`);
        }
        this.number = numberObj;
        const dateObj = DateTime.from(params.date, { isRelative: false });
        if (!dateObj) {
            throw new Error(`La date du train ${this.number} est invalide.`);
        }
        if (!dateObj.hasDate()) {
            throw new Error(`La date du train ${this.number} ne contient d'une heure non datée.`);
        }
        this.date = dateObj;
        this.service = params.service ?? "";
        const pathObj = Path.from(params.path);
        if (!pathObj) {
            throw new Error(`Le parcours du train ${this.number} est invalide.`);
        }
        this.path = pathObj;
        this.units = params.units ?? [];
        this.reusesMap = params.reusesMap ?? new Map();
    }

    /**
     * Indique si le train est une évolution.
     * @returns {boolean} - Vrai si le train est une évolution, faux sinon.
     */
    public get isMouvement(): boolean {
        return this.number.isMouvement;
    }
    
    /**
     * Indique si le train est W.
     * @returns {boolean} - Vrai si le train est W, faux sinon.
     */
    public get isEmptyPassenger(): boolean {
        return this.number.isEmptyPassenger;
    }

    /**
     * Retourne la gare d'origine du train.
     * @returns {Stop | undefined} - Gare d'origine du train.
     */
    public get origin(): Stop | undefined {
        return this.getStop(0);
    }

    /**
     * Retourne l'heure de départ du train.
     * @returns {DateTime | undefined} - Heure de départ du train.
     */
    public get departureTime(): DateTime | undefined {
        return this.getTime(0);
    }

    /**
     * Retourne la gare de destination du train.
     * @returns {Stop | undefined} - Gare de destination du train.
     */
    public get destination(): Stop | undefined {
        return this.getStop(-1);
    }

    /**
     * Retourne l'heure d'arrivée du train.
     * @returns {DateTime | undefined} - Heure d'arrivée du train.
     */
    public get arrivalTime(): DateTime | undefined {
        return this.getTime(-1);
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @return {string} - Clé du train.
     */
    public toString(): string {
        return this.key.toString();
    }

    /**
     * Retourne l'objet Train correspondant à la clé ou l'objet Train donné.
     * Si la clé est une string, elle est utilisée pour chercher l'objet Train correspondant
     *  dans l'index des trains. Si la clé est un objet Train, il est retourné tel quel.
     * Si la clé est une string mais que l'objet Train correspondant n'existe pas, undefined est retourné.
     * @param {Nullable<TrainInput>} value - Clé ou objet Train.
     * @returns {Train | undefined} - Objet Train correspondant,
     *  ou undefined si la clé est une string mais que l'objet Train correspondant n'existe pas.
     */
    public static from(
        value: Nullable<TrainInput>
    ): Train | undefined {
        if (value == null || value === "" || value === "-") return undefined;
        if (value instanceof Train) return value;
        return Trains.get(value!);
    }

    /**
     * Retourne le radical de la clé du train constitué de date_numéroOrigine.
     * @returns {string} - Radical de la clé du parcours.
     */
    public buildRadical(): string {
        return `${this.date.format('yyyy-MM-dd')}_${this.number.format({ withDoubleParity: false })}`;
    }

    /**
     * Retourne les réutilisations (trains précédents et suivants) correspondants aux clés en paramètres.
     * @param {number} [offset=1] - Indice de la réutilisation par rapport au train en cours
     *  (ex: -1 pour le train précédent, 0 pour le train en cours, 2 pour le deuxième train suivant, etc.).
     * @param {boolean} [excludeMouvements=false] - Indique si les évolutions sont exclues.
     * @param {boolean} [excludeEmptyPassenger=false] - Indique si les W sont exclues.
     * @returns {(Reuse<Train> | undefined)[]} - Tableau des réutilisations
     *  pour chaque élément compté à partir de 0.
     */
    public reuses(
        offset: number,
        {
            excludeMouvements = false,
            excludeEmptyPassenger = false
        }: {
            excludeMouvements?: boolean,
            excludeEmptyPassenger?: boolean
        } = {}
    ): (Reuse<Train> | undefined)[] {

        // Si offset = 0, retourne le train actuel pour tous les éléments.
        if (offset === 0) return new Array(this.units.length).fill(this);

        // Détermine la direction et le nombre d'occurrence.
        const direction = offset > 0 ? 1 : -1;
        const occurrence = offset * direction;
    
        const reuses: (Reuse<Train> | undefined)[] = [];
        for (let p = 0; p < this.units.length; p++) {
            reuses.push(this.reusesMap.get(direction *(p + 1))
                ?.resolve({ occurrence, excludeMouvements, excludeEmptyPassenger })
                    ?? undefined);
        }
        return reuses;
    }

    /**
     * Affecte une réutilisation à un élément du train.
     * @param {number} unit - Numéro de l'élément du train (à partir de 0).
     */
    public setReuse(
        unit: number,
        reuseParams: ReuseParams
    ): void {
        const reuse = Reuse.from({
            type: Train,
            params: reuseParams
        });
        if (!reuse) return;
        this.reusesMap.set(unit, reuse);
    }

    /**
     * Affecte la réutilisation inverse au train de réutilisation de l'élément demandé.
     * @param {number} unit - Numéro de l'élément du train (à partir de 0).
     */
    public setReverseReuse(
        unit: number,
    ): void {
        const reuse = this.reusesMap.get(unit);
        const reuseTrain = reuse?.target;
        if (!reuseTrain) return;
        reuseTrain.setReuse(-reuse.position, { reuseKey: this.key, position: -unit });
    }

    /**
     * Retourne l'arrêt associé à une gare.
     * Si un nombre est donné, il d'agit du numéro d'ordre de l'arrêt.
     *  (à partir de 0, ou négatif pour un décompte à partir du terminus)
     * Si la gare a une parité définie, renvoie l'arrêt correspondant.
     * Sinon, cherche l'arrêt dans le sens pair, puis dans le sens impair.
     * Si les deux arrêts sont trouvés, renvoie le premier arrêt chronologique.
     * Sinon, renvoie l'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     * @param {StationInput | number} station - La gare à chercher.
     * @returns {Stop | undefined} - L'arrêt trouvé, ou undefined si aucun arrêt n'est trouvé.
     */
    public getStop(
        station: StationInput | number
    ): Stop | undefined {
        return this.path.getStop(station);
    }

    /**
     * Retourne la plus petite des heures d'arrivée, de départ ou de passage à l'arrêt indiqué.
     * Si ignoreArrival est vrai, lit plutôt l'heure de départ ou de passage.
     * @param {Stop | StationInput | number} stop - L'arrêt à chercher.
     * @param {boolean} [ignoreArrival=false] - Si vrai, ignore l'heure d'arrivée
     *  et préfère l'heure de départ ou de passage. Si faux (par défaut),
     *  c'est d'abord l'heure d'arrivée qui est prise en compte.
     * @param {boolean} [ignorePassage=false] - Si vrai, renvoie undefined si l'arrêt n'est qu'un passage.
     * @returns {DateTime | undefined} - Heure la plus petite à l'arrêt,
     *  ou undefined si aucune heure n'est lue.
     */
    public getTime(
        stop: Stop | StationInput | number,
        { 
            ignoreArrival = false,
            ignorePassage = false,
            ignoreDeparture = false
        }: {
            ignoreArrival?: boolean,
            ignorePassage?: boolean,
            ignoreDeparture?: boolean
        } = {}
    ): DateTime | undefined {
        return (stop instanceof Stop ? stop : this.getStop(stop))
            ?.getTime({ ignoreArrival, ignorePassage, ignoreDeparture, reference: this.date });
    }
}

/**
 * Classe Trains contenant la liste des trains.
 */
class Trains {

    // Constantes de classe
    public static readonly UNKNOWN_UNIT = "?";

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Trains";               // Nom de la feuille 
    private static readonly TABLE = "Trains";               // Nom du tableau 
    private static readonly START_CELL = "A1";              // Première cellule
    private static readonly DATABASE_COLUMNS = {            // Liste des colonnes avec leur emplacement
        key: 0,
        number: 1,
        date: 2,
        service: 3,
        path: 4,
        units: 5,
        reusesMap: 6
    } as const;

    // Constantes de lecture du tableau à importer
    private static readonly IMPORT_MODE = "TRAIN";          // Mode à filtrer
    private static readonly IMPORT_1_UNIT = "Court";        // Train court (à 1 élément)
    private static readonly IMPORT_2_UNITS = "Long";        // Train long (à 2 éléments)
    private static readonly IMPORT_SHEET = "Import trains"; // Nom de la feuille
    private static readonly IMPORT_COLUMNS = {              // Liste des colonnes avec leur emplacement
        number: 3,
        missionCode: 4,
        origin: 5,
        departureTime: 6,
        destination: 7,
        arrivalTime: 8,
        units: 9,
        mode: 10,
        date: -1
    } as const;

    // Constantes des définitions des données de la classe chargées et imprimées dans les colonnes Excel

    /**
     * Liste des définitions générales :
     *  - clé,
     *  - numéro,
     *  - date,
     *  - service,
     *  - parcours,
     *  - code mission,
     *  - gare de départ,
     *  - heure de départ,
     *  - gare d'arrivée,
     *  - heure d'arrivée,
     *  - éléments,
     *  - trains précédents,
     *  - réutilisations,
     *  - heures de passage à un arrêt
     */
    private static readonly COLUMN_DEFINITIONS: TableColumns<Train> = {

        key: {                                              // Clé du train
            header: "Clé",
            type: "string",
            required: true,
            format: { width: 40 }
        },
    
        number: {                                           // Numéro du train
            header: "Numéro du train",
            type: "string",
            required: true,
            print: value => (value as TrainNumber).format({ abbreviate: false, withDoubleParity: false }),
            format: { width: 100 }
        },
    
        date: {                                             // Date et heure de départ du train
            header: "Date",
            type: "number",
            required: true,
            format: {
                numberFormat: "dd/MM/yyyy",
                width: 120
            }
        },

        dayOfWeek: {                                        // Jour de la semaine de circulation
            header: "Jour",
            type: "number",
            print: (value, train) => train.date.getDayOfWeek()?.abbreviation,
            format: { width: 100 }
        },
    
        service: {                                          // Service auquel le train est rattaché
            header: "Service",
            type: "string",
            format: { width: 120 }
        },
    
        path: {                                             // Clé du parcours du train
            header: "Parcours",
            type: "string",
            required: true,
            // print: (value, train) => train.path.key,
            format: { width: 120 }
        },
            
        missionCode: {                                      // Code mission
            header: "Code mission",
            type: "string",
            print: (value, train) => train.path.missionCode,
            format: { width: 60 }
        },
    
        origin: {                                             // Origine
            header: "Origine",
            type: "string",
            required: true,
            format: { width: 100 }
        },

        departureTime: {                                    // Heure de départ de l'origine
            header: "Heure du départ",
            type: "number",
            required: true,
            format: {
                numberFormat: "hh:mm:ss",
                width: 120
            }
        },

        destination: {                                               // Destination
            header: "Terminus",
            type: "string",
            required: true,
            format: { width: 100 }
        },

        arrivalTime: {                                      // Heure d'arrivée à destination
            header: "Heure d'arrivée",
            type: "number",
            required: true,
            format: {
                numberFormat: "hh:mm:ss",
                width: 120
            }
        },
    
        units: {                                            // Eléments du train
            header: "Eléments",
            type: "string",
            load: TableSerializer.loadArray,
            print: TableSerializer.printArray,
            format: { width: 120 }
        },

        reusesMap: {                                        // Clés des réutilisations
            header: "Clés des réutilisations",
            type: "string",
            load: value =>
                Utils.deserializeMap(value, {
                    parseKey: Number,
                    parseValue: key => Reuse.from({ type: Train, key })
                }),
            print: value => 
                Utils.serializeMap(
                    value as Map<number, Reuse<Train>>, 
                    { serializeValue: reuse => reuse?.key }),
            format: { width: 120 }
        },
    
        previousTrains: ({                                  // Trains précédents
            excludeMouvements = true,
            excludeEmptyPassenger = false
        }: {
            excludeMouvements?: boolean,
            excludeEmptyPassenger?: boolean
        } = {}) => ({
            header: "Trains précédents",
            type: "string",
            print: (value, train) => Utils.joinArray(
                train.reuses(-1, { excludeMouvements, excludeEmptyPassenger })
                    .map((reuse?: Reuse<Train>) => reuse?.toString()),
                { symbol: " + ", defaultValue: "?" }
            ),
            format: { width: 120 }
        }),
    
        nextTrains: ({                                      // Trains suivants
            excludeMouvements = true,
            excludeEmptyPassenger = false
        }: {
            excludeMouvements?: boolean,
            excludeEmptyPassenger?: boolean
        } = {}) => ({
    
            header: "Trains suivants",
            type: "string",
            print: (value, train) => Utils.joinArray(
                train.reuses(1, { excludeMouvements, excludeEmptyPassenger })
                    .map((reuse?: Reuse<Train>) => reuse?.toString()),
                { symbol: " + ", defaultValue: "?" }
            ),
            format: { width: 120 }
        }),
    
        stopTime: ({                                        // Heure d'arrivée ou de départ
            station,
            ignoreArrival = false,
            ignorePassage = false,
            ignoreDeparture = false,
            arrivalReplacementValue = "",
            passageReplacementValue = "",
            departureReplacementValue = ""
        }: {
            station?: Station | StationWithParity | string,
            ignoreArrival?: boolean,
            ignorePassage?: boolean,
            ignoreDeparture?: boolean,
            arrivalReplacementValue?: string,
            passageReplacementValue?: string,
            departureReplacementValue?: string
        } = {}) => ({
            header: 
                [
                    station,
                    [
                        (!ignoreArrival && !arrivalReplacementValue) ? "Arrivée" : "",
                        (!ignorePassage && !passageReplacementValue) ? "Passage" : "",
                        (!ignoreDeparture && !departureReplacementValue) ? "Départ" : "",
                    ].filter(v => v !== "").join("/")
                ].filter(v => v !== "").join(" "),
            type: "number",
            print: (value, train) => {
                if (!station) return undefined;
                const stop = train.getStop(station);
                if (!stop) return undefined;
                if (arrivalReplacementValue && stop.arrivalTime) return arrivalReplacementValue;
                if (passageReplacementValue && stop.passageTime) return passageReplacementValue;
                if (departureReplacementValue && stop.departureTime) return departureReplacementValue;
                return train.getTime(stop, { ignoreArrival, ignorePassage, ignoreDeparture });
            },
            format: {
                numberFormat: "hh:mm:ss",
                width: 120
            }
        })
    };

    /**
     * Liste des définitions spécifiques à l'import,
     *  en surcharge des définitions générales :
     *  - éléments,
     *  - code mission,
     *  - gare de départ,
     *  - heure de départ,
     *  - gare d'arrivée,
     *  - heure d'arrivée.
     */
    private static readonly IMPORT_COLUMN_DEFINITIONS: TableColumns<Train> = {

        ...Trains.COLUMN_DEFINITIONS,

        date: {                                             // Date et heure de départ du train
            ...Trains.COLUMN_DEFINITIONS.date,
            type: "number",
            load: (_, { get }) => get("departureTime")
        },

        units: {                                            // Eléments
            ...Trains.COLUMN_DEFINITIONS.units,
            type: "string",    
            load: value =>
                value === this.IMPORT_1_UNIT
                    ? ["-"]
                    : value === this.IMPORT_2_UNITS
                        ? ["-", "-"]
                        : []
        },

        mode: {                                      // Code mission
            type: "string"
        },
    };
 
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
    public static has(
        key: string
    ): boolean {
        return this.map.has(key);
    }

    /**
     * Retourne un train correspondant à la clé donnée.
     * @param {string} key - Clé du train.
     * @returns {Train | undefined} - Train correspondant, ou undefined si non trouvé.
     */
    public static get(
        key: string
    ): Train | undefined {
        return this.map.get(key);
    }

    /**
     * Ajoute un train dans la base de données, référencé par sa clé.
     * Si le train est déjà présent, une erreur est levée.
     * @param {Train} train - Train à ajouter.
     * @throws {Error} - Si le train est déjà présent dans la base de données.
     */
    private static set(
        value: Train
    ): void {
        if (this.has(value.key)) {
            throw new Error(`Le train ${value} est déjà présent`
                + ` dans la base de données.`);
        }
        this.map.set(value.key, value);
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
     * @param {TrainParams} params - Paramètres de création du train.
     * @throws {Error} - Si le train est déjà présent dans la base de données.
     */   
    public static create(params: TrainParams): Train {
        
        // Instancie l'objet Train.
        const train = new Train(params);

        // Insère le train dans la base de données, en générant si besoin la clé
        return this.insert(train);
    }

    /**
     * Ajoute un train dans la base de données.
     * Si un train avec une même clé est déjà présent sans suffixe, des suffixes 1 et 2 sont ajoutés.
     * @param {Train} train - Train à ajouter.
     * @returns {Train} - Train ajouté avec sa clé mise à jour si nécessaire.
     */
    private static insert(
        train: Train
    ): Train {

        // Clé existante : ajoute le train à la base de données.
        // Si un train avec la même clé est déjà présent dans la base de données, une erreur est levée.
        if(train.key) {

            // Ajoute l'objet Path dans la base de données, indexé par sa clé.
            this.set(train);

            return train;
        }

        // Clé non existante : génère une nouvelle clé.

        const radical = train.buildRadical();

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
    public static delete(
        key: string
    ): void {

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
     * Retourne une liste de trains correspondant aux critères donnés :
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
            numbers,
            dates,
            adaptTime = true,
            origin,
            destination,
            via,
            timeFrom,
            timeTo,
            zones,
            batteries
        }: {
            numbers?: OneOrMany<TrainNumberInput>,
            dates?: NullableOneOrMany<DateTime>,
            adaptTime?: boolean,
            origin?: StationInput,
            destination?: StationInput,
            via?: OneOrMany<StationInput>,
            timeFrom?: DateTime,
            timeTo?: DateTime,
            zones?: OneOrMany<number>,
            batteries?: OneOrMany<number>
        } = {}
    ): Train[] {

        // Normalise les filtres
        const numbersArray = Utils.asArray(numbers, {split: ";"})
            .map(n => TrainNumber.from(n))
            .filter(n => n !== null);
    
        const datesArray = Utils.asArray(dates, {split: ";"})
            .map(d => {
                const date = DateTime.from(d);
                if (date === null) return 0;
                const value = date!.getDate({ adaptTime });
                if (value === 0) {
                    Log.warn(`Les dates du filtre des trains doivent être absolues et non nulles. `
                        + `La date ${date} ne sera pas prise en compte.`);
                }
                return value;
            })
            .filter(value => value !== 0);
    
        const viaArray = Utils.asArray(via, {split: ";"});
    
        const zonesArray = Utils.asArray(zones, {split: ";"})
            .filter(z => Number.isInteger(z) && z >= 0 && z <= 9);
    
        const batteriesArray = Utils.asArray(batteries, {split: ";"})
            .filter(b => Number.isInteger(b) && b >= 0 && b <= 99);

        // Filtre les trains selon les critères données
        return this.values().filter(train => {
    
            // Numéros de train
            if (
                numbersArray.length > 0
                    && !numbersArray.some(number => train.number.includes(number))
            ) {
                return false;
            }

            // Dates
            if (
                datesArray.length > 0
                && !datesArray.includes(train.date.getDate({ adaptTime }))
            ) {
                return false;
            }

            // Gare de départ
            const originStop = train.path.getStop(origin);
            if (originStop && originStop !== train.path.origin) {
                return false;
            }

            // Gare d'arrivée
            const destinationStop = train.path.getStop(destination);
            if (destinationStop && destinationStop !== train.path.destination) {
                return false;
            }

            // Arrêts du train (départ, arrivée, gares intermédiaires)
            //  avec passage dans l'intervalle de dates
            for (const station of viaArray) {
    
                const stop = train.path.getStop(station);
    
                // Le train ne passe pas par cette gare
                if (!stop) return false;
    
                // Heure d'arrivée
                if (timeFrom) {
                    const arrivalTime = stop.getTime({ ignoreArrival: false, reference: train.date });
                    if (arrivalTime && arrivalTime.compareTo(timeFrom) < 0) return false;
                }
    
                // Heure de départ
                if (timeTo) {
                    const departureTime = stop.getTime({ ignoreArrival: true, reference: train.date });
                    if (departureTime && departureTime.compareTo(timeTo) > 0) return false;
                }
            }
    
            // Zones
            if (
                zonesArray.length > 0
                    && train.number.zone !== null
                    && !zonesArray.includes(train.number.zone!)
            ) {
                return false;
            }

            // Batteries
            if (
                batteriesArray.length > 0
                    && train.number.battery !== null
                    && !batteriesArray.includes(train.number.battery!)
            ) {
                return false;
            }
    
            return true;
        });
    }

    /**
     * Charge les trains.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        Log.startTimer(`${this.name}.load()`);

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les parcours s'ils ne sont pas encore chargés.
        Paths.load(); 

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
 
                // Récupère les champs.
                const params = TableSerializer.loadRow<
                    Train,
                    TrainParams
                >({
                    row,
                    columns: this.DATABASE_COLUMNS,
                    definitions: this.COLUMN_DEFINITIONS,
                    filters: "key"
                });
                if (!params) continue;
                
                // Crée l'objet et l'insère dans la base de données.
                const train = this.create(params);
            } 

        } catch (e) {
            throw new Error(`${this.name}.load (ligne ${excelRow}) : ${e}`);
        } 

        Log.timer(`${this.name}.load()`);
    }

    /**
     * Importe les trains dans la base de données à partir d'un tableau Excel.
     */
    public static import(): void {

        Log.startTimer(`${this.name}.import()`);

        // Récupère les lignes de la base de données.
        const rows = WorkbookServices.getRows({
            sheetName: this.SHEET,
            tableName: this.TABLE
        });
        if (!rows.length) {
            Log.warn(`${this.name}.load() : aucune donnée trouvée dans la table.`);
            return;
        }
        
        // Parcourt les lignes (hors en-tête).
        let excelRow: number = 0;
        try {

            for (const [rowIndex, row] of rows) {

                // Vérifie si la ligne est vide.
                if (row.length === 0) continue;

                // Calcule le numéro de ligne Excel.
                excelRow = rowIndex + 2; // +1 pour slice, +1 pour en-tête
;
                // Récupère les champs.
                const params = TableSerializer.loadRow({
                    row,
                    columns: this.IMPORT_COLUMNS,
                    definitions: this.IMPORT_COLUMN_DEFINITIONS,
                    filters: [{ column: "mode", filter: v => v === this.IMPORT_MODE }]
                }) as {
                    date: number,
                    number: number,
                    missionCode: string,
                    origin: string,
                    departureTime: number,
                    destination: string,
                    arrivalTime: number,
                    units: string[],
                    mode: string
                };
                if (!params) continue;

                // Crée le parcours à partir des gares de départ et d'arrivée.
                const path = Path.fromTerminals({
                    ...params,
                    areRelativeTimes: false,
                    findPath: true
                });

                // Crée l'objet et l'insère dans la base de données.
                const train = this.create({
                    ...params,
                    path
                });

                if ((rowIndex + 1) % 100 === 0 ) Log.info(rowIndex + 1);
            } 

        } catch (e) {
            throw new Error(`${this.name}.import (ligne ${excelRow}) : ${e}`);
        } 

        Log.timer(`${this.name}.import()`);
    }

    /**
     * Sauvegarde la base de données dans un tableau.
     * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
     * @param {string} [tableName=this.TABLE] - Nom du tableau.
     * @param {string} [startCell=this.START_CELL] - Adresse de la cellule de départ pour le tableau.
     */
    public static save(
        {
            sheetName = this.SHEET,
            tableName = this.TABLE,
            startCell = this.START_CELL
        }: {
            sheetName?: string,
            tableName?: string,
            startCell?: string
        } = {}
    ): void {
    
        Log.startTimer(`${this.name}.save()`);

        TableSerializer.print({
            entities: Array.from(this.map.values()),
            columns: this.DATABASE_COLUMNS,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        Log.timer(`${this.name}.save()`);
    }

    /**
     * Imprime une sélection de trains dans un tableau,
     *  en affichant l'horaire de passage dans les gares choisies.
     * @param {Train[]} trains - Sélection de trains à afficher.
     * @param {{ string, boolean }[]} stops - Gares dont les horaires sont à afficher.
     * @param {string} sheetName - Nom de la feuille de calcul.
     * @param {string} tableName - Nom du tableau.
     * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
     */
    public static printSelection(
        {
            trains,
            stations,
            sheetName,
            tableName,
            startCell = "A1"
        }: {
            trains: Train[],
            stations?: OneOrMany<StationInput>,
            sheetName: string,
            tableName: string,
            startCell?: string
        }
    ): void {

        Log.startTimer(`${this.name}.printSelection()`);

        const columns: Record<string, TableColumnReference> = {
            key: {},
            number: {},
            missionCode: {},
            date: {},
            origin: {},
            departureTime: {},
            destination: {},
            arrivalTime: {}
        };
        const stationsArray = Utils.asArray(stations, {split: ";"})
            .map(s => StationWithParity.from(s))
            .filter(s => s !== null);

        for (const [index, stop] of stationsArray.entries()) {
        
            columns[`arrival_${index}`] = {
                definition: "stopTime",
                station: stop,
                ignoreDeparture: true
            };
        
            columns[`departure_${index}`] = {
                column: 9 + index * 2,
                definition: "stopTime",
                station: stop,
                ignoreArrival: true,
                passageReplacementValue: ">"
            };
        }

        TableSerializer.print({
            entities: trains,
            columns,
            definitions: this.COLUMN_DEFINITIONS,
            sheetName,
            tableName,
            startCell
        });

        Log.timer(`${this.name}.printSelection()`);
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

    // Propriétés de la classe TrainPath
    public key: string;             // Clé du sillon
    public number: TrainNumber;     // Numéro du sillon
    public days: Days;              // Jours de circulation du sillon
    public services: string[];      // Services auxquels le sillon est rattaché
    public path: Path;              // Parcours sur lequel le sillon circule
    public units: string[] = []     // Composistion du sillon
    public readonly reusesMap: Map<number, Reuse<Train>> = new Map();              
                                    // Map des réutilisations, selon la position de chaque élément
                                    //  - positif pour la réutilisation suivante
                                    //  - négatif pour la réutilisation précédente

    /**
     * Constructeur de la classe TrainPath.
     * @param {string} [key=""] - Clé du sillon.
     * @param {TrainNumberInput| undefined} number - Numéro du sillon.
     * @param {DaysInput} days - Jours de circulation du sillon.
     * @param {string[] | string} [services=[]] - Services auxquels le sillon est rattaché.
     * @param {PathInput | undefined} path - Parcours sur lequel le sillon circule.
     * @param {string[]} [units=[]] - Composistion du sillon.
     * @param {string[]} [previousKeys=[]] - Clés des sillons précédents.
     * @param {string[]} [reuseRelations=[]] - Clés des sillons de réutilisations.
     */
    public constructor(
        {
            key = "",
            number,
            days,
            services = [],
            path,
            units = [],
            previousKeys = [],
            reuseRelations = []
        }: {
            key?: string,
            number?: TrainNumberInput| undefined,
            days?: DaysInput,
            services?: string[] | string,
            path?: PathInput | undefined,
            units?: string[],
            previousKeys?: string[],
            reuseRelations?: string[]
        }
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
        this.reuseRelations = reuseRelations;
    }

    public get isMouvement(): boolean {
        return this.number.isMouvement;
    }
    
    public get isEmptyPassenger(): boolean {
        return this.number.isEmptyPassenger;
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
        return this.reuseRelations.map(key => TrainPath.from(key));
    }

    /**
     * Retourne une représentation textuelle simple et stable de l'objet,
     *  utilisée implicitement dans les conversions string (ex: `${obj}`).
     * @returns {string} - Clé du sillon.
     */
    public toString(): string {
        return this.key.toString();
    }

    /**
     * Retourne l'objet TrainPath correspondant à la clé ou l'objet TrainPath donné.
     * Si la clé est une string, elle est utilisée pour chercher l'objet TrainPath correspondant
     *  dans l'index des sillons. Si la clé est un objet TrainPath, il est retourné tel quel.
     * Si la clé est une string mais que l'objet TrainPath correspondant n'existe pas, undefined est retourné.
     * @param {TrainNullable<PathInput>} value - Clé ou objet TrainPath.
     * @returns {TrainPath | undefined} - Objet TrainPath correspondant,
     *  ou undefined si la clé est une string mais que l'objet TrainPath correspondant n'existe pas.
     */
    public static from(
        value: TrainNullable<PathInput>
    ): TrainPath | undefined {
        if (value == null || value === "" || value === "-") return undefined;
        if (value instanceof TrainPath) return value;
        return TrainPaths.get(value!);
    }

    // public static fromTrain(
    //     train: Train | undefined
    // ): TrainPath | undefined {
    //     if (train == null) return undefined;

    //     if (value instanceof TrainPath) return value;
    //     return TrainPaths.get(value!);
    // }

    /**
     * Construit la clé du sillon qui est composée du premier service rattaché
     *  suivi du numéro de sillon et du groupe de jours de circulation.
     * @returns {string} - Clé du sillon.
     */
    public buildKey(): string {
        return `${this.services[0]}_${this.number.format({ withDoubleParity: false })}_${this.days.code}`;
    }
}

/**
 * Classe TrainPaths contenant la liste des sillons.
 */
class TrainPaths {

    // Constantes de lecture de la base de données Excel
    private static readonly SHEET = "Sillons";              // Feuille contenant la liste des sillons
    private static readonly TABLE = "Sillons";              // Tableau contenant la liste des sillons
    private static readonly START_CELL = "A1";              // Première cellule de la liste des sillons
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
    public static has(
        key: string
    ): boolean {
        return this.map.has(key);
    }

    /**
     * Retourne un sillon correspondant à la clé donnée.
     * @param {string} key - Clé du sillon.
     * @returns {TrainPath | undefined} - TrainPath correspondant, ou undefined si non trouvé.
     */
    public static get(
        key: string
    ): TrainPath | undefined {
        return this.map.get(key);
    }

    // public static getFrom(
    //     numbers: TrainNumber[],
    //     days: Days,
    //     services: string[],
    //     from: StationInput,
    //     to: StationInput,
    //     via: (StationInput)[]
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
    private static set(
        trainPath: TrainPath
    ): void {
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
     * @param {TrainNumberInput| undefined} number - Numéro du sillon.
     * @param {DaysInput} days - Jours de circulation du sillon.
     * @param {string[] | string} [services=[]] - Services auxquels le sillon est rattaché.
     * @param {PathInput | undefined} path - Parcours sur lequel le sillon circule.
     * @param {string[]} [units=[]] - Composistion du sillon.
     * @param {string[]} [previousKeys=[]] - Clés des sillons précédents.
     * @param {string[]} [reuseRelations=[]] - Clés des sillons de réutilisations.
     * @returns {TrainPath} - TrainPath créé, ou undefined si le sillon est déjà présent dans la base de données.
     * @throws {Error} - Si le sillon est déjà présent dans la base de données.
     */
    public static create(
        {
            key = "",
            number,
            days,
            services = "",
            path,
            units = [],
            previousKeys = [],
            reuseRelations = []
        }: {
            key?: string,
            number?: TrainNumberInput| undefined,
            days: DaysInput,
            services?: string,
            path?: PathInput | undefined,
            units?: string[],
            previousKeys?: string[],
            reuseRelations?: string[]
        }
    ): TrainPath {

        // Instancie l'objet TrainPath.
        const trainPath = new TrainPath({
            key,
            number,
            days,
            services,
            path,
            units,
            previousKeys,
            reuseRelations
        });

        // Insère le sillon dans la base de données, en générant si besoin la clé
        return this.insert(trainPath);
    }

    /**
     * Ajoute un sillon dans la base de données.
     * @param {TrainPath} trainPath - TrainPath à ajouter.
     * @returns {TrainPath} - TrainPath ajouté.
     */
    private static insert(
        trainPath: TrainPath
    ): TrainPath {

        // Ajoute l'objet Path dans la base de données, indexé par sa clé.
        // Si un sillon avec la même clé est déjà présent dans la base de données, une erreur est levée.
        this.set(trainPath);

        return trainPath;
    }

    /**
     * Supprime le sillon de la base de données dont la clé est donnée en paramètre.
     * @param {string} key - Clé du sillon à supprimer.
     */
    public static delete(
        key: string
    ): void {

        // Supprime le sillon de la map.
        this.map.delete(key);
    }

    /**
     * Charge les sillons.
     * @param {boolean} [erase=false] - Si vrai, force le rechargement de la base de données.
     *  Si faux (par défaut), ne recharge pas si déjà chargé.
     */
    public static load(
        { erase = false }: { erase?: boolean } = {}
    ): void {

        // Vérifie si la table à charger existe déjà.
        if (this.size > 0) {
            if (!erase) return;
            this.clear();
        }

        // Charge les parcours s'ils ne sont pas encore chargés.
        Paths.load(); 

        // Charge la base de données.
        const data = WorkbookServices.getDataFromTable({ sheetName: this.SHEET, tableName: this.TABLE });
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
                const key = WorkbookServices.getString(row, this.COL_KEY);
                const number = WorkbookServices.getString(row, this.COL_NUMBER);
                const days = WorkbookServices.getString(row, this.COL_DAYS);
                const services = WorkbookServices.getString(row, this.COL_SERVICES);
                const path = WorkbookServices.getString(row, this.COL_PATH);
                const unitsString = WorkbookServices.getString(row, this.COL_UNITS);
                const units = splitAndFilter(unitsString);
                const previousString = WorkbookServices.getString(row, this.COL_PREVIOUS);
                const previous = adaptWithUnits(previousString, units);
                const reusesString = WorkbookServices.getString(row, this.COL_REUSES);
                const reuses = adaptWithUnits(reusesString, units);

                // Crée l'objet et l'insère dans la base de données.
                const trainPath = this.create({
                    key,
                    number,
                    days,
                    services,
                    path,
                    units,
                    previousKeys: previous,
                    reuseRelations: reuses
                });
            } 

        } catch (e) {
            throw new Error(`TrainPaths.load (ligne ${excelRow}) : ${e}`);
        } 
    }

//     /**
//      * Sauvegarde les sillons de la base de données dans un tableau.
//      * @param {string} [sheetName=this.SHEET] - Nom de la feuille de calcul.
//      * @param {string} [tableName=this.TABLE] - Nom du tableau.
//      * @param {string} [startCell="A1"] - Adresse de la cellule de départ pour le tableau.
//      */
//     public static print(
//         sheetName: string = this.SHEET,
//         tableName: string = this.TABLE,
//         startCell: string = "A1"
//     ): void {

//         // Convertit la base de données en un tableau de données.
//         const data: PrimitiveValue | undefined[][] = Array
//             .from(this.map.values())
//             .map((trainPath: TrainPath) => [
//                 trainPath.key,
//                 trainPath.number.format({ abbreviate: false, withoutDoubleParity: true }),
//                 trainPath.days.code,
//                 trainPath.services.join(' ,'),
//                 trainPath.path.key,
//                 TableSerializer.joinArray(trainPath.units),
//                 joinUnits(trainPath.previousKeys),
//                 joinUnits(trainPath.reuseRelations)
//             ]);

//         // Imprime le tableau.
//         const table = WorkbookServices.printTable({
//             headers: this.HEADERS, 
//             data,
//             sheetName, 
//             tableName, 
//             startCell
//         });
//     }
}

function testUtils(
    options: Partial<AssertDDOptions> = {}
): void {

    const assert = new AssertDD(options);

    /* ==========================================================
       1. convertValue()
       ========================================================== */

    const testObject = {
        toNumber: () => 123,
        toBoolean: () => true,
        toString: () => "OBJ"
    };

    const convertTests = [

        // String
        { desc: "string",            value: "Paris",    type: "string",  expected: "Paris" },
        { desc: "string trim",       value: " Paris ",  type: "string",  expected: "Paris" },
        { desc: "empty string",      value: "",         type: "string",  expected: undefined },
        { desc: "number -> string",  value: 42,         type: "string",  expected: "42" },
        { desc: "boolean -> string", value: true,       type: "string",  expected: "true" },

        // Number
        { desc: "number",            value: 42,         type: "number",  expected: 42 },
        { desc: "string -> number",  value: "12.5",     type: "number",  expected: 12.5 },
        { desc: "string virgule",    value: "12,5",     type: "number",  expected: 12.5 },
        { desc: "string invalide",   value: "abc",      type: "number",  expected: undefined },
        { desc: "boolean -> number", value: true,       type: "number",  expected: undefined },
        { desc: "toNumber()",        value: testObject, type: "number",  expected: 123 },

        // Boolean
        { desc: "boolean true",  value: true,           type: "boolean", expected: true },
        { desc: "boolean false", value: false,          type: "boolean", expected: false },
        { desc: "number 0",      value: 0,              type: "boolean", expected: false },
        { desc: "number 42",     value: 42,             type: "boolean", expected: true },
        { desc: "string oui",    value: "oui",          type: "boolean", expected: true },
        { desc: "string yes",    value: "yes",          type: "boolean", expected: true },
        { desc: "string false",  value: "false",        type: "boolean", expected: false },
        { desc: "string vide",   value: "",             type: "boolean", expected: undefined },
        { desc: "toBoolean()",   value: testObject,     type: "boolean", expected: true },

        // Sans conversion
        { desc: "raw string",   value: "Paris",         type: undefined, expected: "Paris" },
        { desc: "raw number",   value: 42,              type: undefined, expected: 42 },
        { desc: "raw boolean",  value: true,            type: undefined, expected: true },
        { desc: "raw object",   value: {},              type: undefined, expected: "[object Object]" },
        { desc: "raw object",   value: {},              type: undefined, expected: "[object Object]" },

        // Raw Null | Undefined
        { desc: "raw null, type string",        value: null,      type: "string",  expected: undefined },
        { desc: "raw undefined, type number",   value: undefined, type: "number",  expected: undefined },
        { desc: "raw null, type undefined",     value: null,      type: undefined, expected: undefined },
        { desc: "raw undefined, type undefined",value: undefined, type: undefined, expected: undefined },
    ];

    assert.check(convertTests, {
        category: "Utils",
        label: t => `convertValue(${t.type}) ${t.desc}`,
        actual: (t: typeof convertTests[number]) =>
            Utils.convertValue(
                t.value,
                t.type as PrimitiveType
            )
    });

    /* ==========================================================
       2. asArray()
       ========================================================== */

    const asArrayTests = [

        {  desc: "valeur simple", value: 42,            expected: [42] },
        {  desc: "tableau",       value: [1, 2, 3],     expected: [1, 2, 3] },
        {  desc: "null",          value: null,          expected: [] },
        {  desc: "split",         value: "A;B;C",       expected: ["A", "B", "C"],
            options: { split: ";" } },
        {  desc: "trim",          value: " A ; B ; C ", expected: ["A", "B", "C"],
            options: { split: ";", trim: true  } },
        {  desc: "filtre vide",   value: "A;;B",        expected: ["A", "B"],
            options: { split: ";"  } },
        {  desc: "conserve vide", value: "A;;B",        expected: ["A", "", "B"],
            options: { split: ";", filterEmptyString: false  } }
    ];

    assert.check(asArrayTests, {
        category: "Utils",
        label: t => `asArray() ${t.desc}`,
        actual: (t: typeof asArrayTests[number]) =>
            Utils.asArray(
                t.value,
                t.options
            )
    });

    /* ==========================================================
       3. equals()
       ========================================================== */

    const equalsTests = [

        // Primitifs
        { desc: "string égales",                value1: "Paris",   value2: "Paris",   expected: true },
        { desc: "string différentes",           value1: "Paris",   value2: "Lyon",    expected: false },
        { desc: "number égaux",                 value1: 42,        value2: 42,        expected: true },
        { desc: "boolean différents",           value1: true,      value2: false,     expected: false },

        // Null / undefined
        { desc: "null null",                    value1: null,      value2: null,      expected: true },
        { desc: "undefined undefined",          value1: undefined, value2: undefined, expected: true },
        { desc: "null undefined",               value1: null,      value2: undefined, expected: false },

        // Tableaux
        { desc: "tableaux égaux",               value1: [1, 2, 3], value2: [1, 2, 3], expected: true },
        { desc: "tableaux ordre différent",     value1: [1, 2, 3], value2: [3, 2, 1], expected: false },
        { desc: "tableaux tailles différentes", value1: [1, 2],    value2: [1, 2, 3], expected: false },

        // Objets
        { desc: "objets égaux",        value1: { a: 1, b: 2 }, value2: { a: 1, b: 2 }, expected: true },
        { desc: "objets différents",   value1: { a: 1, b: 2 }, value2: { a: 1, b: 3 }, expected: false },
        { desc: "propriété manquante", value1: { a: 1, b: 2 }, value2: { a: 1 },       expected: false },

        // Imbrications
        { desc: "objets imbriqués égaux",      
            value1: { a: [1, 2], b: { c: 3 } }, value2: { a: [1, 2], b: { c: 3 } }, expected: true },
        { desc: "objets imbriqués différents", 
            value1: { a: [1, 2], b: { c: 3 } }, value2: { a: [1, 2], b: { c: 4 } }, expected: false }
    ];

    assert.check(equalsTests, {
        category: "Utils",
        label: t => `equals() ${t.desc}`,
        actual: (t: typeof equalsTests[number]) =>
            Utils.equals(
                t.value1,
                t.value2
            )
    });

    /* ==========================================================
       4. joinArray()
       ========================================================== */

    const joinTests = [
        { input: ["a", "b", "c"],                                       expected: "a;b;c" },
        { input: ["x", "x", "x"],                                       expected: "x" },
        { input: ["a", null, ""],                                       expected: "a" },
        { input: ["x", "x", "x"],       mergeEqualValues: false,        expected: "x;x;x" },
        { input: ["a", "b", "c"],       symbol: ",",                    expected: "a,b,c" },
        { input: [],                    symbol: "-", defaultValue: "-", expected: "" },
        { input: [null, undefined, ""], symbol: "-", defaultValue: "-", expected: "-" },
    ];

    assert.check(joinTests, {
        category: "Utils",
        label: t => `joinArray(${t.input})`,
        actual: (t: typeof joinTests[number]) => Utils.joinArray(t.input, t)
    });

    /* ==========================================================
       5. splitArray()
       ========================================================== */

    const splitTests = [
        { input: "a;b;c",                       expected: ["a", "b", "c"] },
        { input: "a b c",                       expected: ["a", "b", "c"] },
        { input: "a,,b;;c",                     expected: ["a", "b", "c"] },
        { input: " a ; b ; c ",                 expected: ["a", "b", "c"] },
        { input: "a;a;b",                       expected: ["a", "a", "b"] },
        { input: "a;;;b;;;;c",                  expected: ["a", "b", "c"] },
        { input: "a|b|c",      separators: "|", expected: ["a", "b", "c"] },
        { input: "",                            expected: [] },
        { input: undefined,                     expected: [] },
        { input: 123,                           expected: ["123"] },
    ];

    assert.check(splitTests, {
        category: "Utils",
        label: t => `splitArray(${String(t.input)})`,
        actual: (t: typeof splitTests[number]) => Utils.splitArray(t.input, t)
    });

    /* ==========================================================
       6. SYMETRIE joinArray() & splitArray()
       ========================================================== */

    const roundTripTests = [
        { input: ["a", "b", "c"]},
        { input: ["x"]},
        { input: ["Paris", "Lyon"]},
        { input: []}
    ];
    
    assert.check(roundTripTests, {
        category: "Utils",
        label: t => `split(join(${JSON.stringify(t)}))`,
        actual: (t: typeof roundTripTests[number])  =>
            Utils.splitArray(Utils.joinArray(t.input)),
        expected: (t: typeof roundTripTests[number]) => t.input
    });

    /* ==========================================================
       7. serializeMap()
       ========================================================== */

       const serializeTests = [
        { map: new Map([[1, "A"]]),           expected: "1>>A" },
        { map: new Map([[1, "A"], [2, "B"]]), expected: "1>>A|2>>B" },
        { map: new Map([[1, null]]),          expected: "1>>?" },
        { map: new Map([[1, undefined]]),     expected: "1>>?" }
    ];

    assert.check(serializeTests, {
        category: "TableSerializer",
        label: t => `serializeMap(size=${t.map.size})`,
        actual: (t: typeof serializeTests[number]) => Utils.serializeMap(t.map)
    });

    /* ==========================================================
       8. deserializeMap()
       ========================================================== */

    const deserializeTests = [
        { value: "1>>A",      expected: new Map([[1, "A"]]) },
        { value: "1>>A|2>>B", expected: new Map([[1, "A"], [2, "B"]]) },
        { value: "",          expected: new Map() },
        { value: undefined,   expected: new Map() }
    ];

    assert.check(deserializeTests, {
        category: "TableSerializer",
        label: t => `deserializeMap(${String(t.value)})`,
        actual: (t: typeof deserializeTests[number]) => Utils.deserializeMap(
            t.value,
            {
                parseKey: Number,
                parseValue: v => v ?? ""
            }
        )
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testUtils");
}

function testWorkbookServices(
    options: Partial<AssertDDOptions> = {}
) {
    
    const assert = new AssertDD(options);

    const testSheetName = "testWorkbookServices";
    const testTableName = "testTable";

    /* ==========================================================
       1. FEUILLES
       ========================================================== */
    
    // La fonction getSheet ne peut pas être incorporée dans assert.check
    //  car elle nécessite un lancement en fonction async.
    const sheetName = WorkbookServices.getSheet({ sheetName: testSheetName, createIfMissing: true }).getName();
    const sheetName2 = WorkbookServices.getSheet({ sheetName: testSheetName }).getName();

    const getSheetTests = [
        { label: "Création feuille",     actual: sheetName },
        { label: "Récupération feuille", actual: sheetName2 }
    ];

    assert.check(getSheetTests, {
        category: "WorkbookServices",
        expected: testSheetName
    });

    /* ==========================================================
       2. VALIDATION CELLULES
       ========================================================== */

    const cellTests = [
        { input: "A1",  failOnError: true,  expected: "A1" },
        { input: "123", failOnError: false, expected: "" }
    ];

    assert.check(cellTests, {
        category: "WorkbookServices",
        label: t => `checkCellName("${t.input}")`,
        actual: (t: typeof cellTests[number]) =>
            WorkbookServices.checkCellName(
                t.input,
                { failOnError: t.failOnError }
            )
    });

    /* ==========================================================
       3. TABLEAU EXCEL
       ========================================================== */

    const headers = ["ColStr", "ColNum", "ColBool"];
    const data: CellValue[][] = [
        ["Paris", 42, true],
        ["", "12", "FALSE"],
        [undefined, "abc", undefined]
    ];

    const table = WorkbookServices.printTable({
            headers, 
            data,
            sheetName: testSheetName, 
            tableName: testTableName
    });
    const tableName = table.getName();

    assert.check([{
            label: "Création tableau",
            actual: tableName,
            expected: testTableName
    }], {
        category: "WorkbookServices"
    });

    /* ==========================================================
       4. LECTURE BRUTE
       ========================================================== */

    // La fonction getDataFromTable ne peut pas être incorporée dans assert.check
    //  car elle nécessite un lancement en fonction async.
    const tableData = WorkbookServices.getDataFromTable({ sheetName: testSheetName, tableName: testTableName });

    assert.check([{
        label: "Lecture brute ligne 1 colonne 1",
        actual: () => tableData[1][0],
        expected: "Paris"
    }], {
        category: "WorkbookServices"
    });

    /* ==========================================================
       5. SUPPRESSION FEUILLE
       ========================================================== */

    // La fonction getSheet ne peut pas être incorporée dans assert.check
    //  car elle nécessite un lancement en fonction async.
    WorkbookServices.getSheet({ sheetName: testSheetName })?.delete();
    const deletedSheet = WorkbookServices.getSheet({ sheetName: testSheetName, failOnError: false });
       
    assert.check([{
        label: "Suppression feuille",
        actual: deletedSheet,
        expected: null
    }], {
        category: "WorkbookServices"
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testWorkbookServices");
}

function testTableSerializer(
    options: Partial<AssertDDOptions> = {}
) {
    const assert = new AssertDD(options);

    /* ==========================================================
       1. BASE DE DONNÉES SIMULÉE
       ========================================================== */

    const row = [
        "Paris",     // 0 string
        "42",        // 1 number (string)
        "TRUE",      // 2 boolean
        "",          // 3 empty string
        undefined,   // 4 undefined
        "abc"        // 5 invalid number
    ] as (CellValue)[];

    /* ==========================================================
       2. getValueOrUndefined()
       ========================================================== */

    const getValueOrUndefinedTests = [
        { type: "string",  index: 0, expected: "Paris" },
        { type: "string",  index: 4, expected: undefined },

        { type: "number",  index: 1, expected: 42 },
        { type: "number",  index: 3, expected: undefined },
        { type: "number",  index: 5, expected: undefined },

        { type: "boolean", index: 2, expected: true },

        { type: undefined, index: 0, expected: "Paris" },
        { type: undefined, index: 1, expected: "42" },
        { type: undefined, index: 2, expected: "TRUE" },
        { type: undefined, index: 3, expected: undefined },
        { type: undefined, index: 4, expected: undefined },
    ];

    assert.check(getValueOrUndefinedTests, {
        category: "TableSerializer",
        label: t => `getValueOrUndefined(${t.type}) col ${t.index}`,
        actual: (t: typeof getValueOrUndefinedTests[number]) => TableSerializer.getValueOrUndefined({
            type: t.type as PrimitiveType,
            row,
            index: t.index
        })
    });

    /* ==========================================================
       3. getValue() (avec valeur par défaut)
       ========================================================== */

    const getValueTests = [
        { type: "string", index: 4,                       expected: "" },
        { type: "string", index: 4, defaultValue: "X",    expected: "X" },
        { type: "string", index: 4, defaultValue: 1,      expected: "1" },

        { type: "number", index: 5,                       expected: 0 },
        { type: "number", index: 5, defaultValue: -1,     expected: -1 },
        { type: "number", index: 5, defaultValue: "de",   expected: 0 },
        { type: "number", index: 4, defaultValue: "12",   expected: 12 },

        { type: "boolean", index: 4,                      expected: false },
        { type: "boolean", index: 4, defaultValue: true,  expected: true },
        { type: "boolean", index: 4, defaultValue: 1,     expected: true },
        { type: "boolean", index: 4, defaultValue: "oui", expected: true },

        // Sans type
        { type: undefined, index: 0,                      expected: "Paris" },
        { type: undefined, index: 4,                      expected: "" },
        { type: undefined, index: 4, defaultValue: "X",   expected: "X" },
        { type: undefined, index: 4, defaultValue: 123,   expected: 123 },
    ];

    assert.check(getValueTests, {
        category: "TableSerializer",
        label: t => `getValue(${t.type}) col ${t.index}`,

        actual: (t: typeof getValueTests[number]) => TableSerializer.getValue({
            type: t.type as PrimitiveType,
            row,
            index: t.index,
            defaultValue: t.defaultValue
        })
    });

    /* ==========================================================
       4. getRequiredValue()
       ========================================================== */

    const getRequiredValueTests = [
        { type: "string",  index: 0, expected: "Paris" },
        { type: "number",  index: 1, expected: 42 },
        { type: "boolean", index: 2, expected: true },

        { type: "string",  index: 3, expected: AssertDD.THROWS },
        { type: "string",  index: 4, expected: AssertDD.THROWS },
        { type: "number",  index: 5, expected: AssertDD.THROWS },

        // Sans type
        { type: undefined, index: 0, expected: "Paris" },
        { type: undefined, index: 1, expected: "42" },
        { type: undefined, index: 2, expected: "TRUE" },

        { type: undefined, index: 3, expected: AssertDD.THROWS },
        { type: undefined, index: 4, expected: AssertDD.THROWS },
    ];

    assert.check(getRequiredValueTests, {
        category: "TableSerializer",
        label: t => `getRequiredValue(${t.type}) col ${t.index}`,
        actual: (t: typeof getRequiredValueTests[number]) => TableSerializer.getRequiredValue({
            type: t.type as PrimitiveType,
            row,
            index: t.index
        })
    });

    /* ==========================================================
       5. printArray()
       ========================================================== */

    const joinTests = [
        { input: ["a", "b", "c"], expected: "a;b;c" },
        { input: ["x", "x", "x"], expected: "x" },
        { input: ["a", null, ""], expected: "a" },
    ];

    assert.check(joinTests, {
        category: "TableSerializer",
        label: t => `printArray(${t.input})`,
        actual: (t: typeof joinTests[number]) => TableSerializer.printArray(t.input)
    });

    /* ==========================================================
       6. loadArray()
       ========================================================== */

    const splitTests = [
        { input: "a;b;c",   expected: ["a", "b", "c"] },
        { input: "a b c",   expected: ["a", "b", "c"] },
        { input: "a,,b;;c", expected: ["a", "b", "c"] },
        { input: "",        expected: [] },
        { input: undefined, expected: [] },
        { input: 123,       expected: ["123"] },
    ];

    assert.check(splitTests, {
        category: "TableSerializer",
        label: t => `loadArray(${String(t.input)})`,
        actual: (t: typeof splitTests[number]) => TableSerializer.loadArray(t.input)
    });

    /* ==========================================================
       7. buildColumns()
       ========================================================== */

    const columns = {
        name: 0,
        age: 1
    } as const;

    const definitions = {
        name: { header: "Nom", type: "string" },
        age:  { header: "Age", type: "number" }
    } as const;

    const buildColumnsTests = [
        { index: 0, expected: "name" },
        { index: 1, expected: "age" }
    ];

    const builtColumns = TableSerializer.buildColumns(
        columns,
        definitions
    );

    assert.check(buildColumnsTests, {
        category: "TableSerializer",
        label: t => `buildColumns property : ${t.expected}`,
        actual: (t: typeof buildColumnsTests[number]) => builtColumns[t.index]?.property
    });

    /* ==========================================================
       8. loadRow()
       ========================================================== */

    const rowLoad = [
        "PAR",
        "Paris",
        "1"
    ] as PrimitiveValue[];

    const columnsLoad = {
        abbreviation: 0,
        name: 1,
        active: 2
    } as const;

    const definitionsLoad = {
        abbreviation: { type: "string", required: true },
        name:         { type: "string" },
        active:       { type: "boolean" }
    } as const;

    const result = TableSerializer.loadRow<
        unknown,
        {
            abbreviation: string;
            name: string;
            active: boolean;
        }
    >({
        row: rowLoad,
        columns: columnsLoad,
        definitions: definitionsLoad
    });

    assert.check([{
        label: "loadRow abbreviation",
        actual: () => result?.abbreviation,
        expected: "PAR"
    }], {
        category: "TableSerializer"
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testTableSerializer");
}

function testDateTime(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    Params.load();

    // Arrondi pour comparer deux dates sans prendre en compte des différences inférieures à la seconde
    const round = (v: number) => 
        Math.round(v * 1e10) / 1e10;

    /* ==========================================================
       1. CONSTRUCTION & ROLLOVER
       ========================================================== */

    // Test réalisé avec une heure de changement de jour fixée à 3h00. 
    // Si une erreur se déclenche, merci de vérifier ce paramètre.
    const constructorTests = [
        { value: 4/24, isRelative: false, expectedValue: 4/24,     desc: "04:00" },
        { value: 1/24, isRelative: false, expectedValue: 1/24 + 1, desc: "01:00 → 25:00" },
        { value: 0,    isRelative: false, expectedValue: 1,        desc: "00:00 → 24:00" },
        { value: 1/24, isRelative: true,  expectedValue: 1/24,     desc: "Durée relative" }
    ];

    assert.check(constructorTests, {
        category: "DateTime",
        label: t => `from(${t.value}, ${t.isRelative}) (${t.desc})`,
        actual: (t: typeof constructorTests[number]) =>
            round(DateTime.from(t.value, { isRelative: t.isRelative })!.excelValue),
        expected: (t: typeof constructorTests[number]) =>
            typeof t.expectedValue === "number"
               ? round(t.expectedValue)
               : t.expectedValue
    });

    /* ==========================================================
       2. GETTERS HEURE
       ========================================================== */

    const time1 = DateTime.from(4.5 / 24)!;

    const gettersTests = [
        { label: "getHours()",   actual: () => time1.getHours(),   expected: 4 },
        { label: "getMinutes()", actual: () => time1.getMinutes(), expected: 30 },
        { label: "getSeconds()", actual: () => time1.getSeconds(), expected: 0 }
    ];

    assert.check(gettersTests, {
        category: "DateTime"
    });

    /* ==========================================================
       3. GETTERS DATE & ADAPTATION
       ========================================================== */

    const dtAdapt = DateTime.from(45830 + 1/24)!;

    const getDayTests = [
        { adapted: true,  expected: 21 },
        { adapted: false, expected: 22 },
    ];

    assert.check(getDayTests, {
        category: "DateTime",
        label: t => `getDay(${t.adapted ? "adapted" : "real"})`,
        actual: (t: typeof getDayTests[number]) => dtAdapt.getDay({ adaptTime: t.adapted })
    });

    const getDayOfWeekTests = [
        { adapted: true,  expected: Day.SATURDAY.toString() },
        { adapted: false, expected: Day.SUNDAY.toString() },
    ];


    assert.check(getDayOfWeekTests, {
        label: t => `getDayOfWeek(${t.adapted ? "adapted" : "real"})`,
        actual: (t: typeof getDayOfWeekTests[number]) => dtAdapt
            .getDayOfWeek({ adaptTime: t.adapted })?.toString()
    });
    
    /* ==========================================================
       4. getTime() adapté / non adapté
       =========================================================== */

    const time2 = DateTime.from(1/24)!;

    const getTimeTests = [
        { adapted: true,  expected: round(1 + 1/24) },
        { adapted: false, expected: round(1/24) },
    ];

    assert.check(getTimeTests, {
        category: "DateTime",
        label: t => `getTime(${t.adapted ? "adapted" : "real"})`,
        actual: (t: typeof getTimeTests[number]) =>
            round(time2.getTime({ adaptTime: t.adapted }))
    });

    /* ==========================================================
       5. removeDatePartFormat()
       ========================================================== */

    const removeDatePartFormatTests = [

        // Date avant heure
        { format: "dd/mm/yyyy hh:nn",                   expected: "hh:nn" }, 
        { format: "dddd dd mmmm yyyy hh:nn:ss",         expected: "hh:nn:ss" },

        // Date après heure 
        { format: "hh:nn dd/mm/yyyy",                   expected: "hh:nn" }, 
        { format: "hh:nn:ss (dd/mm/yyyy)",              expected: "hh:nn:ss" },

        // Date avant ET après heure 
        { format: "dd/mm/yyyy hh:nn dd/mm/yyyy",        expected: "hh:nn" },

        // Heure seule 
        { format: "hh:nn",                              expected: "hh:nn" }, 
        { format: "hh:nn:ss",                           expected: "hh:nn:ss" },

        // Date seule 
        { format: "dd/mm/yyyy",                         expected: "" }, 
        { format: "dddd dd mmmm yyyy",                  expected: "" },

        // Texte autour de l'heure 
        { format: "Départ hh:nn",                       expected: "Départ hh:nn" }, 
        { format: "hh:nn arrivée",                      expected: "hh:nn arrivée" }, 
        { format: "Départ hh:nn arrivée le dd/mm/yyyy", expected: "Départ hh:nn" }, 
        { format: "Le dd/mm/yyyy départ hh:nn arrivée", expected: "hh:nn arrivée" },

        // Aucun token reconnu 
        { format: "texte libre",                        expected: "texte libre" },

        // Cas limites 
        { format: "",                                   expected: "" }, 
        { format: "yyyy",                               expected: "" },
        { format: "hh",                                 expected: "hh"}
    ];

    assert.check(removeDatePartFormatTests, {
        category: "DateTime",
        label: t => `removeDatePartFormat("${t.format}")`,
        actual: (t: typeof removeDatePartFormatTests[number]) =>
            DateTime.removeDatePartFormat(t.format)
    });

    /* ==========================================================
       5. format() heure
       ========================================================== */

    const formatTimeTests = [
        { value: 4.5/24,  fmt: DateTime.TIME_FORMAT_WITH_SECONDS,    expected: "04:30:00" },
        { value: 4.5/24, fmt: DateTime.FULL_DATE_TIME_FORMAT,        expected: "04:30:00" },
        { value: 4.5/24,  fmt: DateTime.TIME_FORMAT_WITHOUT_SECONDS, expected: "04:30" },
        { value: -4.5/24, fmt: DateTime.TIME_FORMAT_WITHOUT_SECONDS, expected: "-04:30" },
    ];

    assert.check(formatTimeTests, {
        category: "DateTime",
        label: t => `format("${t.fmt}")`,
        actual: (t: typeof formatTimeTests[number]) =>
            DateTime.from(t.value, { isRelative: true })!.format(t.fmt)
    });

    /* ==========================================================
       6. format() date
       ========================================================== */

    const dt = DateTime.from(45860.75)!;

    const formatDateTests = [
        { fmt: DateTime.DATE_FORMAT_WITH_YEAR,    expected: "22/07/2025" },
        { fmt: DateTime.FULL_DATE_TIME_FORMAT,    expected: "22/07/2025 18:00:00" },
        { fmt: DateTime.DATE_FORMAT_WITHOUT_YEAR, expected: "22/07" },
        { fmt: "ddd. dd mmm. yyyy",               expected: "Ma. 22 Juil. 2025" },
        { fmt: "dddd dd mmmm yyyy",               expected: "Mardi 22 Juillet 2025" },
        { fmt: DateTime.DATE_FORMAT_FOR_ID,       expected: "250722" },
    ];

    assert.check(formatDateTests, {
        category: "DateTime",
        label: t => `format("${t.fmt}")`,
        actual: (t: typeof formatDateTests[number]) =>
            dt.format(t.fmt)
    });

    /* ==========================================================
       7. JOURS FÉRIÉS
       ========================================================== */

    // 06/04/2026 = Lundi de Pâques
    const easterMonday2026 = DateTime.from("06/04/2026 6:00")!;

    const holidayTests = [
        { label: 'isHoliday (Lundi de Pâques 2026)',
            actual: () => easterMonday2026.isHoliday(), expected: true },
        { withHolidays: true,                           expected: Day.HOLIDAY.toString() },
        { withHolidays: false,                          expected: Day.MONDAY.toString() }
    ];

    assert.check(holidayTests, {
        category: "DateTime",
        label: t => `getDayOfWeek(withHolidays=${t.withHolidays})`,
        actual: (t: typeof holidayTests[number]) => {
            return easterMonday2026
                .getDayOfWeek({ adaptTime: true, withHolidays: t.withHolidays })?.toString();
        }
    });

    const holidayFormatTests = [
        { withHolidays: true,  expected: "Férié 06/04/2026" },
        { withHolidays: false, expected: "Lundi 06/04/2026" }
    ];

    assert.check(holidayFormatTests, {
        category: "DateTime",
        label: t => `format(withHolidays=${t.withHolidays})`,
        actual: (t: typeof holidayFormatTests[number]) =>
            easterMonday2026.format(
                "dddd dd/mm/yyyy",
                { adaptTime: true, withHolidays: t.withHolidays }
            )
    });

    /* ==========================================================
       8. resolveAgainst()
       ========================================================== */

    const ref = DateTime.from(45830 + 10/24)!;
    const rel = DateTime.from(3/24, { isRelative: true })!;
    const abs = DateTime.from(45830 + 10/24)!;
    const abs2 = DateTime.from(45830 + 13/24)!;

    const resolveAgainstTests = [
        { a: rel, b: ref, expected: round(45830 + 13/24) },
        { a: abs, b: ref, expected: round(abs.excelValue) },
    ];

    assert.check(resolveAgainstTests, {
        category: "DateTime",
        label: t => `resolveAgainst(${t.a}, ${t.b})`,
        actual: (t: typeof resolveAgainstTests[number]) =>
            round(t.a.resolveAgainst(t.b).excelValue)
    });

    /* ==========================================================
       9. relativeTo()
       ========================================================== */
    
    const relativeToTests = [
        { a: rel,  b: ref, expected: AssertDD.THROWS },
        { a: abs2, b: ref, expected: round(3/24) },
    ];

    assert.check(relativeToTests, {
        category: "DateTime",
        label: t => `relativeTo(${t.a}, ${t.b})`,
        actual: (t: typeof relativeToTests[number]) =>
            round(t.a.relativeTo(t.b).excelValue)
    });

    /* ==========================================================
       10. equalsTo()
       ========================================================== */

       const equalsToTests = [
           { a: rel, b: ref, expected: false },
           { a: abs, b: ref, expected: true },
       ];

       assert.check(equalsToTests, {
           category: "DateTime",
           label: t => `equalsTo(${t.a}, ${t.b})`,
           actual: (t: typeof equalsToTests[number]) =>
               t.a.equalsTo(t.b)
       });

    /* ==========================================================
       11. compareTo()
       ========================================================== */

    const compareToTests = [
        { a: 45830 + 10/24, b: 45830 + 10/24, expected: 0 },
        { a: 45830 + 10/24, b: 45830 + 1/24,  expected: 1 },
        { a: 45830 + 10/24, b: 45831,         expected: -1 },
        { a: 45830 + 10/24, b: 10/24,         expected: 0 },
        { a: 45830,         b: 10/24,         expected: 1 },
        { a: 1/24,          b: 10/24,         expected: -1, isRelative: true },
    ];

    const getSign = (v: number) =>
        v > 0 ? 1 : v < 0 ? -1 : 0;

    assert.check(compareToTests, {
        category: "DateTime",
        label: t => `compareTo(${t.a}, ${t.b})`,
        actual: (t: typeof compareToTests[number]) =>
            getSign(DateTime.from(t.a, { isRelative: "isRelative" in t })!
                    .compareTo(DateTime.from(t.b, { isRelative: "isRelative" in t })!)
            )
    });

    /* ==========================================================
       12. equalsOrUndefined()
       ========================================================== */

    const dt1 = DateTime.from(45830 + 10/24)!;
    const dt2 = DateTime.from(45830)!;

    const equalsOrUndefinedTests = [
        { a: undefined, b: undefined, expected: true },
        { a: dt1,       b: undefined, expected: false },
        { a: dt1,       b: dt1,       expected: true },
        { a: dt1,       b: dt2,       expected: false },
    ];

    assert.check(equalsOrUndefinedTests, {
        category: "DateTime",
        label: t => `equalsOrUndefined(${t.a ? "dt" : "undefined"}, ${t.b ? "dt" : "undefined"})`,
        actual: (t: typeof equalsOrUndefinedTests[number]) =>
            DateTime.equalsOrUndefined(t.a, t.b)
    });

    /* ==========================================================
       13. add() / subtract()
       ========================================================== */

    const A = DateTime.from(2/24, { isRelative: true})!;
    const B = DateTime.from(3/24, { isRelative: true })!;

    const operationTests = [
        { label: "add()",      actual: () => round(A.add(B).excelValue),      expected: round(5/24) },
        { label: "subtract()", actual: () => round(A.subtract(B).excelValue), expected: round(-1/24) }
    ];

    assert.check(operationTests, {
        category: "DateTime"
    });

    /* ==========================================================
       14. PARSE STRING
       ========================================================== */

    const parseTests = [

        // Nombres simples
        { input: "1.5",         expected: 1.5 },
        { input: "1,5",         expected: 1.5 },
        { input: "-1.5",        expected: undefined }, // absolu interdit

        // Heures seules
        { input: "04:30",       expectedValue: 4.5 / 24 },
        { input: "04h30",       expectedValue: 4.5 / 24 },
        { input: "04h30min01s", expectedValue: (4.5 + 1/3600) / 24 },
        { input: "01:00",       expectedValue: 1/24 + 1 }, // rollover

        // Heure négative
        { input: "-02:00",      expectedValue: -2/24, relative: true },

        // Date seule
        { input: "22/06/2025",  expected: 45830 },
        { input: "22-06-2025",  expected: 45830 },
        { input: "2025/06/22",  expected: 45830 },

        // Date + heure
        { input: "22/06/2025 04:30",  expectedValue: 45830 + 4.5/24 },
        { input: "22-06-2025 04:30",  expectedValue: 45830 + 4.5/24 },

        // Ordre inversé
        { input: "04:30 22/06/2025",  expectedValue: 45830 + 4.5/24 },

        // Heure négative avec date (doit être ignoré)
        { input: "22/06/2025 -02:00", expectedValue: 45830 + 2/24 },
        { input: "22-06-2025 -02:00", expectedValue: 45830 + 2/24 },

        // Double négatif (invalide)
        { input: "- - 02:00",         expectedValue: 1 + 2/24 },

        // Format partiel
        { input: "22/06 04:00", timeOnly: true, expectedValue: 4/24 },

        // Chaîne invalide
        { input: "abc",         expected: undefined },
        { input: "22/99/2025",  expected: undefined },
        { input: "25:61",       expected: undefined },

    ];

    assert.check(parseTests, {
        category: "DateTime",
        label: t =>`parse("${t.input}")`,
        actual: (t: typeof parseTests[number]) => {

            const dt = DateTime.from(t.input, { isRelative: t.relative});
            if (dt === undefined) return undefined;
            
            const value = ("timeOnly" in t)
                ? dt?.excelValue % 1
                : dt?.excelValue;

            return round(value);
        },
        expected: (t: typeof parseTests[number]) =>
             typeof t.expectedValue === "number"
                ? round(t.expectedValue)
                : t.expectedValue
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testDateTime");
}

function testDays(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    Days.load();
    Day.load();

    /* ==========================================================
       1. Days.from()
       ========================================================== */

    const fromTests = [
        { input: 1,           expected: "1",     desc: "Jour simple" },
        { input: "5-4-3-2-1", expected: "12345", desc: "Tri automatique" },
        { input: "a9b1c7",    expected: "17",    desc: "Valeurs invalides ignorées" }
    ];

    assert.check(fromTests, {
        category: "Days",
        label: t => `Days.from("${t.input}") (${t.desc})`,
        actual: (t: typeof fromTests[number]) => Days.from(t.input)!.numbersString
    });

    /* ==========================================================
       2. extractFromString()
       ========================================================== */

    const extractTests = [
        { input: "lundi",   expectedValue: [1] },
        { input: "ma",      expectedValue: [2] },
        { input: "7;1;3",   expectedValue: [1, 3, 7] },
        { input: "lumeven", expectedValue: [1, 3, 5] }
    ];

    assert.check(extractTests, {
        category: "Days",
        label: t => `extractFromString("${t.input}")`,
        actual: (t: typeof extractTests[number]) =>
            JSON.stringify(Days.extractFromString(t.input)),
        expected: (t: typeof extractTests[number]) =>
            JSON.stringify(t.expectedValue)
    });

    /* ==========================================================
       3. union() / intersection()
       ========================================================== */

    const d1 = Days.from("1-3-5")!;
    const d2 = Days.from("3-4")!;

    const intersectionUnionTests = [
        { label: "intersection 135 ∩ 34",
            actual: () => Days.intersection(d1, d2)?.numbersString, expected: "3" },
        { label: "union 135 ∪ 34",
            actual: () => Days.union(d1, d2)?.numbersString,        expected: "1345" }
    ];

    assert.check(intersectionUnionTests, {
        category: "Days"
    });

    /* ==========================================================
       4. contains()
       ========================================================== */

    const d = Days.from("1-3-5")!;

    const containsTests = [
        { day: 3, expected: true },
        { day: 2, expected: false }
    ];

    assert.check(containsTests, {
        category: "Days",
        label: t => `contains(${t.day})`,
        actual: (t: typeof containsTests[number]) =>
            d.contains(t.day)
    });

    /* ==========================================================
       5. intersects()
       ========================================================== */

    const intersectsTests = [
        { day: 3, expected: true },
        { day: 2, expected: false }
    ];

    assert.check(intersectsTests, {
        category: "Days",
        label: t => `intersects(${t.day})`,
        actual: (t: typeof intersectsTests[number]) =>
            d.intersects(Days.from(t.day)!)
    });

    /* ==========================================================
       6. Days.difference() / count() / numbersString()
       ========================================================== */

    const dA = Days.from("1-3-5")!;
    const dB = Days.from("3")!;

    const differenceCountNumbersStringTests = [
        { label: "difference 135 - 3", actual: () => Days.difference(dA, dB)?.numbersString, expected: "15"},
        { label: "count 135",          actual: () => dA.count,                               expected: 3 },
        { label: "numbersString 135",  actual: () => dA.numbersString,                       expected: "135" }
    ];
    
    assert.check(differenceCountNumbersStringTests, {
        category: "Days"
    });

    /* ==========================================================
       7. ACCESSEURS STATIQUES Day
       ========================================================== */

    const constantsTests = [
        { days: Days.MONDAY,    day: Days.MONDAY,       num: 1, name: "Lundi" },
        { days: Days.TUESDAY,   day: Days.TUESDAY,      num: 2, name: "Mardi" },
        { days: Days.WEDNESDAY, day: Days.WEDNESDAY,    num: 3, name: "Mercredi" },
        { days: Days.THURSDAY,  day: Days.THURSDAY,     num: 4, name: "Jeudi" },
        { days: Days.FRIDAY,    day: Days.FRIDAY,       num: 5, name: "Vendredi" },
        { days: Days.SATURDAY,  day: Days.SATURDAY,     num: 6, name: "Samedi" },
        { days: Days.SUNDAY,    day: Days.SUNDAY,       num: 7, name: "Dimanche" },
        { days: Days.HOLIDAY,   day: Days.HOLIDAY,      num: 8, name: "Férié" }
    ];

    assert.check(constantsTests, {
        category: "Days",
        label: t => `Constante ${t.name}`,
        actual: (t: typeof constantsTests[number]) => 
            [
                t.days.contains(t.num),
                t.days.fullName,
                Days.from(t.num) === t.days
            ].join("|"),

        expected: (t: typeof constantsTests[number]) =>
            [true, t.name, true].join("|")
    });

    /* ==========================================================
       8. Day.from()
       ========================================================== */

    assert.check(constantsTests, {
        category: "Day",
        label: t => `Day.from(${t.num})`,
        actual: (t: typeof constantsTests[number]) => {
            const day = Day.from(t.num)!;
            return [
                day,
                day.fullName,
                day.mask,
                day.mask,
                day.index,
                Day.from(day) === day
            ].join("|");
        },
        expected: (t: typeof constantsTests[number]) =>
            [
                t.day,
                t.name,
                t.day.mask,
                t.days.mask,
                t.num - 1,
                true
            ].join("|")
    });

    /* ==========================================================
       9. Day.values()
       ========================================================== */

    const values = Array.from(Day.values());

    assert.check([{
        label: "Day.values length",
        actual: () => values.length,
        expected: 8
    }], {
        category: "Day"
    });

    /* ==========================================================
       10. DaysValues
       ========================================================== */

    const base = Days.from("1-2-3-4-5-6-7")!;
    const dv = new DaysValues(base);

    dv.set(Days.from("1-2-3-4-5")!, "A");
    dv.set(Days.from("6-7")!, "B");

    const daysValuesTests = [
        { label: "DaysValues split",  actual: () => dv.toString(),      expected: "12345: A, 67: B" },
        { label: "DaysValues Monday", actual: () => dv.get(Day.MONDAY), expected: "A" },
        { label: "DaysValues Sunday", actual: () => dv.get(Day.SUNDAY), expected: "B" }
    ];

    assert.check(daysValuesTests, {
        category: "Days"
    });

    dv.set(Days.from("6-7")!, "A");
    const dv2 = new DaysValues(base);
    dv2.set(Days.from("1-2")!, "X");
    dv2.fillGaps("Y");
    const parsed = DaysValues.from(base, "12345: A, 67: B");

    const daysValuesTests2 = [
        { label: "DaysValues merge",    actual: () => dv.toString(),     expected: "A" },
        { label: "DaysValues fillGaps", actual: () => dv2.isComplete(),  expected: true },
        { label: "DaysValues.from()",   actual: () => parsed.toString(), expected: "12345: A, 67: B" }
    ];

    assert.check(daysValuesTests2, {
        category: "Days"
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testDays");
}

function testParity(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    Parity.load();

    /* ==========================================================
       1. CONSTRUCTEUR & normalizeParityValue()
       ========================================================== */

    const constructorTests = [
        {  desc: 'Lettre impair "I"',      value: "I",     doubleParityAllowed: false, expected: Parity.ODD },
        {  desc: 'Lettre pair "P"',        value: "P",     doubleParityAllowed: false, expected: Parity.EVEN },
        {  desc: "Chiffre impair 1",       value: 1,       doubleParityAllowed: false, expected: Parity.ODD },
        {  desc: "Chiffre pair 2",         value: 2,       doubleParityAllowed: false, expected: Parity.EVEN },
        {  desc: "Numéro de train impair", value: "12345", doubleParityAllowed: false, expected: Parity.ODD },
        {  desc: "Numéro de train pair",   value: "12346", doubleParityAllowed: false, expected: Parity.EVEN },
        {  desc: "Valeur vide",            value: "",      doubleParityAllowed: false, expected: Parity.UNDEFINED },
        {  desc: 'Zéro "0"',               value: "0",     doubleParityAllowed: false, expected: Parity.UNDEFINED },
        {  desc: "Double IP interdite",    value: "IP",    doubleParityAllowed: false, expected: Parity.UNDEFINED },
        {  desc: "Double IP autorisée",    value: "IP",    doubleParityAllowed: true,  expected: Parity.DOUBLE },
        {  desc: 'Double implicite "1/2"', value: "1/2",   doubleParityAllowed: true,  expected: Parity.DOUBLE  }
  ];

  assert.check(constructorTests, {
        category: "Parity",
        label: t => `Parity.from(${JSON.stringify(t.value)}, ${t.doubleParityAllowed}) (${t.desc})`,
        actual: (t: typeof constructorTests[number]) =>
            Parity.from(t.value, { doubleParityAllowed: t.doubleParityAllowed }).value
    });

    /* ==========================================================
       2. null / undefined
       ========================================================== */

    const nullTests = [
        { value: null },
        { value: undefined }
    ];

    assert.check(nullTests, {
        category: "Parity",
        label: t => `Parity.from(${t.value})`,
        actual: (t: typeof nullTests[number]) =>
            Parity.from(t.value).value,
        expected: () => Parity.UNDEFINED
    });

    /* ==========================================================
       3. Pool
       ========================================================== */

    const poolTests = [
        { label: "Pool même instance", actual: () => Parity.from("I"), expected: Parity.from(1) },
        { label: "Pool différent selon doubleParityAllowed",
            actual: () => Parity.from("I", { doubleParityAllowed: false })
                !== Parity.from("I", { doubleParityAllowed: true }),   expected: true }
    ];

    assert.check(poolTests, {
        category: "Parity"
    });

    /* ==========================================================
       4. is()
       ========================================================== */

    const isTests = [
        { value: "I", parity: Parity.ODD,  expected: true },
        { value: "I", parity: Parity.EVEN, expected: false }
    ];

    assert.check(isTests, {
        category: "Parity",
        label: t => `is ${t.value} === ${t.parity}`,
        actual: (t: typeof isTests[number]) =>
            Parity.from(t.value).is(t.parity)
    });

    /* ==========================================================
       5. isDefined()
       ========================================================== */

    const isDefinedTests = [
        { parity: Parity.UNDEFINED, expected: false },
        { parity: Parity.EVEN,      expected: true }
    ];

    assert.check(isDefinedTests, {
        category: "Parity",
        label: t => `isDefined ${t.parity}`,
        actual: (t: typeof isDefinedTests[number]) =>
            Parity.from(t.parity).isDefined()
    });

    /* ==========================================================
       6. isOpposedTo()
       ========================================================== */

    const isOpposedTests = [
        { a: "I",       b: "P",         expected: true },
        { a: "P",       b: "I",         expected: true },
        { a: "I",       b: "I",         expected: false },
        { a: undefined, b: "I",         expected: false },
        { a: undefined, b: undefined,   expected: false }
    ];

    assert.check(isOpposedTests, {
        category: "Parity",
        label: t => `isOpposedTo ${t.a} / ${t.b}`,
        actual: (t: typeof isOpposedTests[number]) => {

            const a = t.a !== undefined ? Parity.from(t.a) : undefined;
            const b = t.b !== undefined ? Parity.from(t.b) : undefined;

            return a?.isOpposedTo(b) ?? false;
        }
    });

    const isOpposedTests2 = [
        { label: "Parity.double() est opposé à rien",
            actual: () => Parity.double().isOpposedTo(Parity.odd()),  expected: false },
        { label: "Parity.undefined() est opposé à rien",
            actual: () => Parity.double().isOpposedTo(Parity.even()), expected: false }
    ];

    assert.check(isOpposedTests2, {
        category: "Parity"
    });

    /* ==========================================================
       7. equalsTo()
       ========================================================== */

    const equalsTests = [
        { a: "I", b: 1,   expected: true },
        { a: "I", b: "P", expected: false }
    ];

    assert.check(equalsTests, {
        category: "Parity",
        label: t => `equalsTo ${t.a} / ${t.b}`,
        actual: (t: typeof equalsTests[number]) =>
            Parity.from(t.a).equalsTo(Parity.from(t.b))
    });

    const oddSimple = Parity.odd({ doubleParityAllowed: false });
    const oddDoubleAllowed = Parity.odd({ doubleParityAllowed: true });

    const equalsToTests2 = [
        { label: "equalsTo basé sur identité",
            actual: () => Parity.from("I").equalsTo(Parity.from("I")),  expected: true },
        { label: "equalsTo faux si doubleParityAllowed_different",
            actual: () => oddSimple.equalsTo(oddDoubleAllowed),         expected: false }
    ];

    assert.check(equalsToTests2, {
        category: "Parity"
    });

    /* ==========================================================
       8. includes()
       ========================================================== */

    const includesTests = [

        // Parités simples
        { a: "I",  b: "I",  expected: true },
        { a: "I",  b: 1,    expected: true },
        { a: "I",  b: "P",  expected: false },
        { a: "P",  b: "I",  expected: false },

        // Parités doubles
        { a: "IP", b: "I",  expected: true },
        { a: "IP", b: "P",  expected: true },
        { a: "IP", b: 1,    expected: true },
        { a: "IP", b: 2,    expected: true },

        // Parités doubles non autorisées
        { a: "I",  b: "IP", expected: false },
        { a: "P",  b: "IP", expected: false },

        // Parités indéfinies
        { a: null, b: "I",  expected: false },
        { a: "I",  b: null, expected: false },
        { a: null, b: null, expected: false }
    ];

    assert.check(includesTests, {
        category: "Parity",
        label: t => `includes ${t.a} ⊇ ${t.b}`,
        actual: (t: typeof includesTests[number]) =>
            Parity.from(t.a, { doubleParityAllowed: true }).includes(t.b)
    });

    const simpleParity = Parity.odd({ doubleParityAllowed: false });

    assert.check([{
        label: "includes refuse double si non autorisée",
        actual: () => simpleParity.includes("IP"),
        expected: false
    }], {
        category: "Parity"
    });

    /* ==========================================================
       9. invert()
       ========================================================== */

    const invertTests = [
        { value: "I",   doubleParityAllowed: false, expected: Parity.EVEN },
        { value: "P",   doubleParityAllowed: false, expected: Parity.ODD },
        { value: "IP",  doubleParityAllowed: true,  expected: Parity.DOUBLE },
        { value: "",    doubleParityAllowed: false, expected: Parity.UNDEFINED }
    ];

    assert.check(invertTests, {
        category: "Parity",
        label: t => `invert ${t.value}`,
        actual: (t: typeof invertTests[number]) =>
            Parity.from(t.value, { doubleParityAllowed: t.doubleParityAllowed })
                .invert()
                .value
    });

    const pDouble = Parity.double();
    const pUndefined = Parity.undefined();

    const invert2Tests = [
        { label: "invert DOUBLE retourne même instance",    value: pDouble,     expected: pDouble },
        { label: "invert UNDEFINED retourne même instance", value: pUndefined,  expected: pUndefined }
    ];

    assert.check(invert2Tests, {
        category: "Parity",
        actual: (t: typeof invert2Tests[number]) =>
            t.value.invert()
    });

    /* ==========================================================
       10. combineWith()
       ========================================================== */

    const combineTests = [
        { a: Parity.undefined({ doubleParityAllowed: true }), b: Parity.odd(),       expected: Parity.ODD },
        { a: Parity.odd({ doubleParityAllowed: true }),       b: Parity.undefined(), expected: Parity.ODD },
        { a: Parity.odd({ doubleParityAllowed: true }),       b: Parity.odd(),       expected: Parity.ODD },
        { a: Parity.even({ doubleParityAllowed: true }),      b: Parity.even(),      expected: Parity.EVEN },
        { a: Parity.odd({ doubleParityAllowed: true }),       b: Parity.even(),      expected: Parity.DOUBLE },
        { a: Parity.even({ doubleParityAllowed: false }),     b: Parity.odd(),       expected: AssertDD.THROWS }
    ];

    assert.check(combineTests, {
        category: "Parity",
        label: t => `combineWith ${t.a.value} + ${t.b.value}`
            + `${t.expected === AssertDD.THROWS ? " interdit si doubleParityAllowed = false" : ""}`,
        actual: (t: typeof combineTests[number]) =>
            t.a.combineWith(t.b).value
    });

    const original = Parity.odd({ doubleParityAllowed: true });
    const combined = original.combineWith(Parity.even());

    const combineTests2 = [
        { label: "combineWith ne modifie pas l'origine", 
            actual: () => original.value,        expected: Parity.ODD },
        { label: "combineWith retourne nouvelle instance", 
            actual: () => combined !== original, expected: true }
    ];

    assert.check(combineTests2, {
        category: "Parity"
    });

    /* ==========================================================
       11. printDigit() / printLetter()
       ========================================================== */

    const printTests = [
        { value: "I",  digit: Parity.digit(Parity.ODD),    letter: Parity.letter(Parity.ODD) },
        { value: "P",  digit: Parity.digit(Parity.EVEN),   letter: Parity.letter(Parity.EVEN) },
        { value: "IP", digit: Parity.digit(Parity.DOUBLE), letter: Parity.letter(Parity.ODD)
             + Parity.letter(Parity.EVEN), doubleParityAllowed: true },
        { value: "",   digit: "",                          letter: "" }
    ];

    assert.check(printTests, {
        category: "Parity",
        label: t => `print ${t.value}`,
        actual: (t: typeof printTests[number]) => {
            const p = Parity.from(t.value, { doubleParityAllowed: t.doubleParityAllowed });
            return [
                p.printDigit(),
                p.printLetter()
            ].join("|");
        },
        expected: (t: typeof printTests[number]) =>
            [t.digit, t.letter].join("|")
    });

    /* ==========================================================
       12. static factories
       ========================================================== */

    const staticFactoriesTests = [
        { label: "Parity.odd()", actual: () => Parity.odd().value,             expected: Parity.ODD },
        { label: "Parity.even()", actual: () => Parity.even().value,           expected: Parity.EVEN },
        { label: "Parity.double()", actual: () => Parity.double().value,       expected: Parity.DOUBLE },
        { label: "Parity.undefined()", actual: () => Parity.undefined().value, expected: Parity.UNDEFINED },
        { label: "Parity.double() autorise combineWith",
            actual: () => Parity.double().combineWith(Parity.odd()).value,     expected: Parity.DOUBLE }
    ];

    assert.check(staticFactoriesTests, {
        category: "Parity"
    });

    /* ==========================================================
       13. static containsParityLetter()
       ========================================================== */

    const containsTests = [
        { text: "Train I",  parity: Parity.ODD,     expected: true },
        { text: "Train I",  parity: Parity.EVEN,    expected: false },
        { text: "Train IP", parity: Parity.DOUBLE,  expected: true }
    ];

    assert.check(containsTests, {
        category: "Parity",
        label: t => `containsParityLetter "${t.text}"`,
        actual: (t: typeof containsTests[number]) =>
            Parity.containsParityLetter(t.text, t.parity)
    });

    /* ==========================================================
       14. static letter() / digit()
       ========================================================== */

    const staticTests = [
        { method: "letter", parity: Parity.ODD,  type: "string" },
        { method: "letter", parity: Parity.EVEN, type: "string" },
        { method: "digit",  parity: Parity.ODD,  type: "number" },
        { method: "digit",  parity: Parity.EVEN, type: "number" },
        { method: "digit",  parity: 999, expected: 0 }
    ];

    assert.check(staticTests, {
        category: "Parity",
        label: t =>
            `${t.method}(${t.parity})`,
        actual: (t: typeof staticTests[number]) => {
            const result = t.method === "letter"
                ? Parity.letter(t.parity)
                : Parity.digit(t.parity);
            return "expected" in t
                ? result
                : typeof result;
        },
        expected: (t: typeof staticTests[number]) =>
            "expected" in t
                ? t.expected
                : t.type
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */
       
    assert.printSummary("testParity");
}

function testTrainNumber(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    TrainNumber.load({ erase: true });

    /* ==========================================================
       1. CONSTRUCTEUR
       ========================================================== */

    const constructorTests = [
        { desc: "Nombre simple",            input: 146490,                     expected: "146490" },
        { desc: "Nombre simple",            input: 146490, doubleParity: true, expected: "146490/1" },
        { desc: "Chaîne avec slash",        input: "146490/1",                 expected: "146490/1" },
        { desc: "Chaîne avec slash",        input: "146490/91",                expected: "146490/1" },
        { desc: "Minuscules + parasites",   input: "w-14a6490",                expected: "W14A6490" }
    ];

    assert.check(constructorTests, {
        category: "TrainNumber",
        label: t => `TrainNumber.from(${JSON.stringify(t.input)}) (${t.desc})`,
        actual: (t: typeof constructorTests[number]) =>
            TrainNumber.from(t.input, { doubleParity: t.doubleParity })!.toString()
    });

    /* ==========================================================
       2. SETTER parity
       ========================================================== */

    const paritySetterTests = [
        { value: "146490", parity: Parity.ODD,  expected: "146491" },
        { value: "146491", parity: Parity.EVEN, expected: "146490" }
    ];
    
    assert.check(paritySetterTests, {
        category: "TrainNumber",
        label: t =>
            `${t.value} set parity(${t.parity})`,
        actual: (t: typeof paritySetterTests[number]) => {
            const trainNumber = TrainNumber.from(t.value)!;
            trainNumber.parity = t.parity;
            return trainNumber.value;
        }
    });

    /* ==========================================================
       3. SETTER doubleParity
       ========================================================== */

    const doubleParitySetterTests = [
        { value: "146490",   doubleParity: true,  expected: "146490/1" },
        { value: "146491",   doubleParity: true,  expected: "146491/0" },
        { value: "146490/1", doubleParity: false, expected: "146490" }
    ];

    assert.check(doubleParitySetterTests, {
        category: "TrainNumber",
        label: t => `${t.value} set doubleParity(${t.doubleParity})`,
        actual: (t: typeof doubleParitySetterTests[number]) => {
            const trainNumber = TrainNumber.from(t.value)!;
            trainNumber.doubleParity = t.doubleParity;
            return trainNumber.value;
        }
    });

    /* ==========================================================
       4. isCommercial
       ========================================================== */

    const isCommercialTests = [
        { value: 147490,   expected: true },
        { value: 146490,   expected: false },
        { value: "E46490", expected: false }
    ];

    assert.check(isCommercialTests, {
        category: "TrainNumber",
        label: t => `isCommercial(${t.value})`,
        actual: (t: typeof isCommercialTests[number]) =>
            TrainNumber.from(t.value)!.isCommercial
    });

    /* ==========================================================
       5. isEmptyPassenger
       ========================================================== */
    
    const isEmptyPassengerTests = [
        { value: 146490, expected: true },
        { value: 569907, expected: true },
        { value: 147490, expected: false }
    ];

    assert.check(isEmptyPassengerTests, {
        category: "TrainNumber",
        label: t => `isEmptyPassenger(${t.value})`,
        actual: (t: typeof isEmptyPassengerTests[number]) =>
            TrainNumber.from(t.value)!.isEmptyPassenger
    });

    /* ==========================================================
       6. isMouvement
       ========================================================== */

    const isMouvementTests = [
        { value: "E46490", expected: true },
        { value: "146490", expected: false }
    ];

    assert.check(isMouvementTests, {
        category: "TrainNumber",
        label: t => `isMouvement(${t.value})`,
        actual: (t: typeof isMouvementTests[number]) =>
            TrainNumber.from(t.value)!.isMouvement
    });

    /* ==========================================================
       7. zone
       ========================================================== */

    const zoneTests = [
        { value: 147490, expected: 4 },
        { value: 146490, expected: null }
    ];

    assert.check(zoneTests, {
        category: "TrainNumber",
        label: t => `zone(${t.value})`,
        actual: (t: typeof zoneTests[number]) =>
            TrainNumber.from(t.value)!.zone
    });

    /* ==========================================================
       8. battery
       ========================================================== */

    const batteryTests = [
        { value: 145824,     expected: 24 },
        { value: 147490,     expected: 90 },
        { value: "147490/1", expected: 91 },
        { value: 146490,     expected: null }
    ];

    assert.check(batteryTests, {
        category: "TrainNumber",
        label: t => `battery(${t.value})`,
        actual: (t: typeof batteryTests[number]) =>
            TrainNumber.from(t.value)!.battery
    });

    /* ==========================================================
       9. format()
       ========================================================== */

    const formatTests = [
        { value: "146490/1", abbreviate: true,  withDoubleParity: true,         expected: "6490/1" },
        { value: "146490/1", abbreviate: false, withDoubleParity: true,         expected: "146490/1" },
        { value: "146490/1", abbreviate: true,  withDoubleParity: false,        expected: "6490" },
        { value: "146490/1", abbreviate: false, withDoubleParity: false,        expected: "146490" },

        { value: "146490/1", forceParity: Parity.ODD,                           expected: "146491/0" },
        { value: "146490/1", forceParity: Parity.EVEN,                          expected: "146490/1" },
        { value: "146490/1", forceParity: Parity.ODD,  withDoubleParity: false, expected: "146491" },
        { value: "146490",   forceParity: Parity.ODD,                           expected: "" },
        { value: "146491",   forceParity: Parity.EVEN,                          expected: "" },
        { value: "146490",   forceParity: Parity.EVEN,                          expected: "146490" },
        { value: "146491",   forceParity: Parity.ODD,                           expected: "146491" },
        { value: "146490/1", forceParity: Parity.ODD,  abbreviate: true,        expected: "6491/0" },
        { value: "146490/1", forceParity: Parity.ODD,  abbreviate: true,
             withDoubleParity: false,                                           expected: "6491" }
    ];

    assert.check(formatTests, {
        category: "TrainNumber",
        label: t => `format(${t.value})`
            + (t.abbreviate !== undefined 
                ? ` (${t.abbreviate ? "abrégé" : "non abrégé"})` 
                : "")
            + (t.withDoubleParity !== undefined 
                ? ` (${t.withDoubleParity ? "avec double parité" : "sans double parité"})`
                : "")
            + (t.forceParity !== undefined
                ? ` (${t.forceParity ? `parité forcée ${t.forceParity}`: "sans parité forcée"})`
                : ""),
        actual: (t: typeof formatTests[number]) =>
            TrainNumber.from(t.value)!.format({
                abbreviate: t.abbreviate,
                withDoubleParity: t.withDoubleParity,
                forceParity: t.forceParity
        })
    }); 

    /* ==========================================================
       10. includes()
       ========================================================== */

       const tn = TrainNumber.from(146490);

       const includesTests = [
           { value: "146490",   expected: true },
           { value: "146491",   expected: true },
           { value: "146490/1", expected: true },
           { value: "146491/0", expected: true },
           { value: "146492",   expected: false }
       ];
   
       assert.check(includesTests, {
           category: "TrainNumber",
           label: t => `includes(${t.value})`,
           actual: (t: typeof includesTests[number]) =>
               tn!.includes(t.value)
       });

    /* ==========================================================
       11. toString()
       ========================================================== */

    const toStringTests = [
        { value: 146490,                     expected: "146490" },
        { value: 146491,                     expected: "146491" },
        { value: 146490, doubleParity: true, expected: "146490/1" },
        { value: 146491, doubleParity: true, expected: "146491/0" },
    ];

    assert.check(toStringTests, {
        category: "TrainNumber",
        label: t => `toString(${t.value})`,
        actual: (t: typeof toStringTests[number]) =>
            TrainNumber.from(t.value, { doubleParity: t.doubleParity })!.toString()
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testTrainNumber");
}

function testStation(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    Stations.load({ erase: true });

    /* ==========================================================
       1. load()
       ========================================================== */

    const sizeAfterFirstLoad = Stations.size;
    Stations.load({ erase: false });
    const sizeAfterSecondLoad = Stations.size;
    Stations.load({ erase: true });
    const sizeAfterReload = Stations.size;

    const loadTests = [
        { label: "Stations.load(true) - au moins une gare chargée",
            actual: Stations.size > 0,   expected: true },
        { label: "Stations.load(false) - pas de rechargement",
            actual: sizeAfterSecondLoad, expected: sizeAfterFirstLoad },
        { label: "Stations.load(true) - rechargement après erase",
            actual: sizeAfterReload,     expected: sizeAfterFirstLoad }
    ];

    assert.check(loadTests, {
        category: "Stations"
    });

    /* ==========================================================
       2. from()
       ========================================================== */

    StationsWithParity.load();

    const fromTests = [
        { label: "from(chaine)",                value: "PZB",                           expected: "PZB" },
        { label: "from(chaine avec parité)",    value: "PZB_1",                         expected: "PZB" },
        { label: "from(instance)",              value: Station.from("PZB"),             expected: "PZB" },
        { label: "from(StationWithParity)",     value: StationWithParity.from("PZB_1"), expected: "PZB" },
        { label: "from(null)",                  value: StationWithParity.from(null),    expected: undefined }
    ];

    assert.check(fromTests, {
        category: "Station",
        actual: (t: typeof fromTests[number]) => Station.from(t.value)?.abbreviation
    })

    /* ==========================================================
       3. get() / getById()
       ========================================================== */

    const firstStation = Stations.values()[0] as Station;

    const accessTests = [
        { label: "Stations contient au moins une Station",
            actual: firstStation instanceof Station,         expected: true         },
        { label: `Stations.get("${firstStation.abbreviation}") retourne la même instance`,
            actual: Stations.get(firstStation.abbreviation), expected: firstStation },
        { label: "Stations.getById(0) retourne une Station", 
            actual: Stations.getById(0) instanceof Station,  expected: true }
    ];

    assert.check(accessTests, {
        category: "Stations"
    });

    /* ==========================================================
       4. RATTACHEMENTS
       ========================================================== */

    const attachmentTests: AssertDDCheck[] = [];

    for (const station of Stations.values()) {

        if (station.referenceStation) {
            attachmentTests.push({
                label: `${station.abbreviation} référencée dans`
                    + ` ${station.referenceStation.abbreviation}.childStations`,
                actual: station.referenceStation.childStations.includes(station),
                expected: true
            });
        }

        for (const child of station.childStations) {
            attachmentTests.push({
                label: `${child.abbreviation}.referenceStation === ${station.abbreviation}`,
                actual: child.referenceStation,
                expected: station
            });
        }
    }

    assert.check(attachmentTests, {
        category: "Stations"
    });

    /* ==========================================================
       5. DONNÉES MÉTIERS
       ========================================================== */

    const dataTests = Array.from(Stations.values()).map(station => ({
        label: `Station ${station.abbreviation} - abréviation non vide`,
        actual: station.abbreviation !== "",
        expected: true
    }));

    assert.check(dataTests, {
        category: "Stations"
    });

    /* ==========================================================
       6. print()
       ========================================================== */

    let printSucceeded = true;
    const sheetAndTableName = "testGares";

    try {
        Stations.save({
            sheetName: sheetAndTableName,
            tableName: sheetAndTableName,
            startCell: "A1"
        });
    } catch {
        printSucceeded = false;
    }

    assert.check([{
        label: 'Stations.print() - impression dans "testGares"',
        actual: printSucceeded,
        expected: true
    }], {
        category: "Stations"
    });

    WorkbookServices.getSheet({ sheetName: sheetAndTableName })?.delete();

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testStation");
}

function testStationWithParity(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    StationsWithParity.load({ erase: true });

    /* ==========================================================
       1. CONSTRUCTION / hasDefinedParity()
       ========================================================== */

    const s = Station.from("JY");
    const sU = StationWithParity.from("JY");
    const sO = StationWithParity.from("JY_1");
    const sE = StationWithParity.from("JY_2");
    const childSwpU = StationWithParity.from("JY-146_1");
    const childSwpO = StationWithParity.from("JY-146_2");
    const childSwpE = StationWithParity.from("JY-146_3");

    const constructorTests = [
        { label: "Station sans parité", actual: sU!.hasDefinedParity(),     expected: false },
        { label: "Station avec parité", actual: sO!.hasDefinedParity(),     expected: true },
        { label: "Station parity odd",  actual: sO!.parity.is(Parity.ODD),  expected: true },
        { label: "Station parity even", actual: sE!.parity.is(Parity.EVEN), expected: true }
    ];

    assert.check(constructorTests, {
        category: "StationWithParity"
    });

    /* ==========================================================
       2. from()
       ========================================================== */

    const fromTests = [
        { label: "from(chaine sans parité)",    value: "JY",                    expected: "JY" },
        { label: "from(chaine avec parité)",    value: "JY_1",                  expected: "JY_1" },
        { label: "from(instance)",              value: sO,                      expected: "JY_1" },
        { label: "from(Station)",               value: Station.from("JY_1"),    expected: "JY" },
        { label: "from(null)",                  value: null,                    expected: undefined }
    ];

    assert.check(fromTests, {
        category: "StationWithParity",
        actual: (t: typeof fromTests[number]) =>
            StationWithParity.from(t.value)?.toString()
    })

    /* ==========================================================
       3. includes()
       ========================================================== */

    const includesTests = [
        { label: "includes undefined -> odd",          actual: sU!.includes(sO),        expected: true },
        { label: "includes odd -> undefined",          actual: sO!.includes(sU),        expected: false },
        { label: "includes même parité",               actual: sO!.includes(sO),        expected: true },
        { label: "includes station fille sans parité", actual: sU!.includes(childSwpU), expected: true },
        { label: "includes station fille avec parité", actual: sU!.includes(childSwpO), expected: true },
        { label: "includes station fille opposée",     actual: sO!.includes(childSwpE), expected: false }
    ];

    assert.check(includesTests, {
        category: "StationWithParity"
    });

    /* ==========================================================
       4. expandWithChildren()
       ========================================================== */

    const expandedU = sU!.expandWithChildren();
    const expandedO = sO!.expandWithChildren();
    const expandedAgain = sU!.expandWithChildren();
    const ids = expandedU.map(s => s.id);

    const expandTests = [
        { label: "expand undefined contient odd",
            actual: expandedU.some(s => s.parity.is(Parity.ODD)),         expected: true },
        { label: "expand undefined contient even",
            actual: expandedU.some(s => s.parity.is(Parity.EVEN)),        expected: true },
        { label: "expand avec parité définie",
            actual: expandedO.some(s => s.parity.is(Parity.ODD)),         expected: true },
        { label: "expand sans doublons",
            actual: ids.length,                                           expected: new Set(ids).size },
        { label: "cache utilisé",
            actual: expandedU,                                            expected: expandedAgain },
        { label: "expand stable",
            actual: expandedU.length,                                     expected: expandedAgain.length },
        { label: "expand avec visited externe",
            actual: sU!.expandWithChildren(new Set<number>()).length > 0, expected: true },
        { label: "key sans parité",
            actual: sU!.key,                                              expected: "JY" },
        { label: "key avec parité",
            actual: sO!.key,                                              expected: "JY_1" }
    ];

    assert.check(expandTests, {
        category: "StationWithParity"
    });

    /* ==========================================================
       5. stationAfterTurnaround()
       ========================================================== */

    const turned = sO!.stationAfterTurnaround();

    if (turned) {

        const turnaroundTests = [
            { label: "turnaround station identique",
                actual: turned.station,                expected: sO!.station },
            { label: "turnaround parité inversée",
                actual: turned.parity.is(Parity.EVEN), expected: true }
        ];

        assert.check(turnaroundTests, {
            category: "StationWithParity",
        });
    }

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testStationWithParity");
}

function testConnection(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    Connections.load({ erase: true });

    /* ==========================================================
       1. load()
       ========================================================== */

    const sizeAfterLoad = Connections.size;

    Connections.load({ erase: false });

    const firstConnection = Connections.values()[0] as Connection;

    const from = firstConnection.from;
    const to = firstConnection.to;

    const globalTests = [
        { label: "Connections.load(true)", 
            actual: Connections.size > 0,                                            expected: true },
        { label: "Connections.load(false)", 
            actual: Connections.size,                                                expected: sizeAfterLoad },
        { label: "Connections.values()", 
            actual: firstConnection instanceof Connection,                           expected: true },
        { label: "Connections.has(from, to)", 
            actual: Connections.has(from, to),                                       expected: true },
        { label: "Connections.get(from, to)", 
            actual: Connections.get(from, to),                                       expected: firstConnection },
        { label: "has(Station, Station)", 
            actual: Connections.has(from.station, to.station),                       expected: true },
        { label: "get(Station, Station)",
            actual: Connections.get(from.station, to.station) instanceof Connection, expected: true }
    ];

    assert.check(globalTests, {
        category: "Connections"
    });

    /* ==========================================================
       2. COHÉRENCE MÉTIER
       ========================================================== */

           const connectionTests: AssertDDEntry<AssertDDCheck>[] = [];

    for (const connection of Connections.values()) {
        const  tests: AssertDDEntry<AssertDDCheck>[] = [
            { label: `${connection} : from instanceof StationWithParity`, 
                actual: connection.from instanceof StationWithParity,                expected: true },
            { label: `${connection} : to instanceof StationWithParity`,
                actual: connection.to instanceof StationWithParity,                  expected: true },
            { label: `${connection} : from ≠ to`,     
                actual: !connection.from.equalsTo(connection.to),                    expected: true },
            { label: `${connection} : temps > 0 sauf retournement`,     
                actual: connection.withTurnaround || connection.time.excelValue > 0, expected: true },
            { label: `${connection} : temps relatif`,
                actual: connection.time.isRelative,                                  expected: true
            }
        ];
        connectionTests.push(
            ...tests
        );
    }

    assert.check(connectionTests, {
        category: "Connections"
    });

    /* ==========================================================
       3. print()
       ========================================================== */

    let printOk = true;
    const sheetAndTableName = "testConnexions";

    try {
        Connections.save({
            sheetName: sheetAndTableName,
            tableName: sheetAndTableName,
            startCell: "A1"
        });
    } catch {
        printOk = false;
    }

    assert.check([{
        label: "Connections.print()",
        actual: printOk,
        expected: true
    }], {
        category: "Connections"
    });

    WorkbookServices.getSheet({ sheetName: sheetAndTableName })?.delete();

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testConnection");
}

function testStop(
    options: Partial<AssertDDOptions> = {}
) {

    const assert = new AssertDD(options);

    const stop1 = new Stop({
        station: "PZB_1",
        stationAfterTurnaround: "PZB_2",
        arrivalTime: "08:00:00",
        departureTime: "08:02:00",
        areRelativeTimes: false,
        tracks: "A;B"
    });
    const stop2 = new Stop({
        station: "BFM_1",
        passageTime: "08:05:00",
        areRelativeTimes: false,
        tracks: "E"
    });
    const stop3 = new Stop({
        station: "SHL_1",
        departureTime: "07:58:00",
        areRelativeTimes: false,
        tracks: "E"
    });

    /* ==========================================================
       1. CONSTRUCTEUR
       ========================================================== */

    assert.check([{
        label: "instance créée",
        actual: stop1 instanceof Stop,
        expected: true
    }], {
        category: "Stop"
    });

    /* ==========================================================
       2. Station / key
       ========================================================== */

    const stationTests = [
        { label: "key",                 actual: stop1.key,                 expected: "PZB_1" },
        { label: "stationAbbreviation", actual: stop1.stationAbbreviation, expected: "PZB" }
    ];

    assert.check(stationTests, {
        category: "Stop"
    });

    /* ==========================================================
       3. withTurnaround() et stationAfterTurnaround()
       ========================================================== */

    const turnaroundTests = [
        { label: "withTurnaround (true)",  actual: stop1.withTurnaround,              expected: true },
        { label: "withTurnaround (false)", actual: stop2.withTurnaround,              expected: false },
        { label: "stationAfterTurnaround", actual: stop1.stationAfterTurnaround?.key, expected: "PZB_2" }
    ];

    assert.check(turnaroundTests, {
        category: "Stop"
    });

    /* ==========================================================
       4. withNonStopPassage
       ========================================================== */

    const withNonStopPassage = [
        { desc: "arrival+departure", value: stop1, expected: false },
        { desc: "passage",           value: stop2, expected: true },
        { desc: "departure",         value: stop3, expected: false }
    ];

    assert.check(withNonStopPassage, {
        category: "Stop",
        label: t => `withNonStopPassage (${t.desc})`,
        actual: (t: typeof withNonStopPassage[number]) => t.value.withNonStopPassage,
        expected: true
    });

    /* ==========================================================
       5. isIntermediateStop
       ========================================================== */

    const isIntermediateStop = [
        { desc: "arrival+departure", value: stop1, expected: true },
        { desc: "passage",           value: stop2, expected: true },
        { desc: "departure",         value: stop3, expected: false }
    ];

    assert.check(isIntermediateStop, {
        category: "Stop",
        label: t => `isIntermediateStop (${t.desc})`,
        actual: (t: typeof isIntermediateStop[number]) => t.value.isIntermediateStop,
        expected: true
    });

    /* ==========================================================
       6. HORAIRES
       ========================================================== */

    const timeTests = [
        { label: "arrivalTime défini",    actual: stop1.arrivalTime instanceof DateTime,   expected: true },
        { label: "departureTime défini",  actual: stop1.departureTime instanceof DateTime, expected: true },
        { label: "passageTime undefined", actual: stop1.passageTime,                       expected: undefined }
    ];

    assert.check(timeTests, { category: "Stop"
    });

    /* ==========================================================
       7. getTime()
       ========================================================== */

    const t1 = stop1.getTime();
    const t2 = stop1.getTime({ ignoreArrival: true });

    const getTimeTests = [
        { label: "getTime() retourne DateTime",
            time: (stop1.getTime()),                                             expected: true },
        { label: "getTime(true) retourne DateTime",
            time: stop1.getTime({ ignoreArrival: true }),                        expected: true },
        { label: "getTime(true) retourne DateTime",
            time: stop1.getTime({ ignoreArrival: true, ignoreDeparture: true }), expected: false }
    ];

    assert.check(getTimeTests, {
        category: "Stop",
        actual: (t: typeof getTimeTests[number]) => t.time instanceof DateTime
    });

    /* ==========================================================
       8. Tracks
       ========================================================== */

    stop1.addTrack("C");

    const trackTests = [
        { label: "tracks longueur", actual: stop1.tracks.length,        expected: 3 },
        { label: "addTrack",        actual: stop1.tracks.includes("C"), expected: true }
    ];

    assert.check(trackTests, {
        category: "Stop"
    });

    /* ==========================================================
       9. equalsTo() / includes()
       ========================================================== */

    const stopSame = new Stop({
        station: "PZB_1",
        stationAfterTurnaround: "PZB_2",
        arrivalTime: "08:00:00",
        departureTime: "08:02:00"
    });

    const stopWithoutParity = new Stop({
        station: "PZB",
        arrivalTime: "08:00:00",
        departureTime: "08:02:00"
    });

    const stopOther = new Stop({
        station: "SQY_1",
        arrivalTime: "08:00:00"
    });

    const compareTests = [
        { label: "equalsTo identique",                  actual: stop1.equalsTo(stopSame),
            expected: true },
        { label: "equalsTo différent",                  actual: stop1.equalsTo(stopOther),
            expected: false },
        { label: "includes sans parité -> avec parité", actual: stopWithoutParity.includes(stop1),
            expected: false },
        { label: "includes avec parité -> sans parité", actual: stop1.includes(stopWithoutParity),
            expected: false }
    ];

    assert.check(compareTests, {
        category: "Stop"
    });

    /* ==========================================================
       10. convertToRelativeTime()
       ========================================================== */

    const ref = DateTime.from("07:00:00");
    stop1.convertToRelativeTime(ref!);

    assert.check([{ 
        label: "convertToRelativeTime refTime",
        actual: stop1.arrivalTime!.compareTo(DateTime.from("01:00:00", { isRelative: true})! ),
        expected: 0
    }], {
        category: "Stop"
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testStop");
}

function testPath(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);

    /* ==========================================================
       1. CRÉATION DU Path
       ========================================================== */

    const path = Path.fromTerminals({
        origin: "PZB",
        departureTime: "08:00:00",
        destination: "SQY",
        arrivalTime: "09:00:00",
        signature: "PZB>SQY>MPU;VC"
    });

    const constructorTests = [
        { label: "Path instance créée", actual: path instanceof Path, expected: true },
        { label: "Path.signature",      actual: path.signature,       expected: "PZB>MPU;VC>SQY" }
    ];

    assert.check(constructorTests, {
        category: "Path"
    });

    /* ==========================================================
       2. findPath()
       ========================================================== */

    path.findPath();

    path.check();

    const findPathTests = [
        { label: "Path.findPath - FULL_PATH", actual: path.stopsChecked,     expected: Path.FULL_PATH },
        { label: "Path.stops non vide",       actual: path.stops.length > 1, expected: true }
    ];

    assert.check(findPathTests, {
        category: "Path"
    });

    /* ==========================================================
       3. PREMIER ET DERNIER ARRÊT
       ========================================================== */

    const first = path.stops[0];
    const last = path.stops[path.stops.length - 1];

    const firstAndLastStopsTests = [
        { label: "Premier arrêt PZB", actual: first.stationAbbreviation, expected: "PZB" },
        { label: "Dernier arrêt SQY", actual: last.stationAbbreviation,  expected: "SQY" }
    ];

    assert.check(firstAndLastStopsTests, {
        category: "Path"
    });

    /* ==========================================================
       4. PASSAGE PAR MPU OU VC
       ========================================================== */

    const viaTests = [
        { label: "Path passe par MPU ou VC",
            actual: path.stops.map(s => s.stationAbbreviation).includes("MPU") 
                || path.stops.map(s => s.stationAbbreviation).includes("VC"),
            expected: true }
    ];

    assert.check(viaTests, {
        category: "Path"
    });

    /* ==========================================================
       5. getStop()
       ========================================================== */

    const getStopTests = [
        { label: "getStop number",      actual: path.getStop(-1) instanceof Stop, 
            expected: true },
        { label: "getStop string",      actual: path.getStop("PZB") instanceof Stop, 
            expected: true },
        { label: "getStop Station",     actual: path.getStop(first.station)?.stationAbbreviation,
             expected: first.stationAbbreviation },
        { label: "getStop SWP key",     actual: path.getStop(first.station.key) instanceof Stop,
            expected: true },
        { label: "getStop sans parité", actual: path.getStop(first.station.station) instanceof Stop, 
            expected: true }
    ];

    assert.check(getStopTests, {
        category: "Path"
    });

    /* ==========================================================
       6. nextStop() / previousStop()
       ========================================================== */
    
    const next = path.nextStop(first);

    const nextStopTests = [
        { label: "nextStop retourne Stop", actual: next instanceof Stop,      expected: true },
        { label: "previousStop cohérent",  actual: path.previousStop(next!),  expected: first }
    ];

    assert.check(nextStopTests, {
        category: "Path"
    });

    /* ==========================================================
       7. INDEX ET POSITIONS
       ========================================================== */

    const indexTests = [
        { label: "Positions cohérentes", 
            actual: path.stops.every((s, i) => path["_stopPosition"].get(s.key) === i), expected: true },
        { label: "signatureIndex contient le Path", 
            actual: Paths.signatureIndex.get(path.signature)![0],                       expected: path }
    ];

    assert.check(indexTests, {
        category: "Path"
    });

    /* ==========================================================
       8. buildConnectionsFromStops()
       ========================================================== */

    const rebuiltConnections = path.buildConnectionsFromStops();
    const rebuilt = new Path();
    rebuilt.stops = Array.from(path.stops);
    rebuilt.stopsChecked = Path.FULL_PATH;

    const buildConnectionsTests = [
        { label: "buildConnectionsFromStops retourne connexions",
            actual: rebuiltConnections.length > 0,              expected: true },
        { label: "Reconstruction cohérente",
            actual: rebuilt.buildConnectionsFromStops().length, expected: rebuiltConnections.length },
        { label: "Temps ordonnés", 
            actual: path.stops.every((s, i, arr) => i === 0 || s.getTime()!
                .compareTo(arr[i - 1].getTime()!) >= 0 ),       expected: true }
    ];

    assert.check(buildConnectionsTests, {
        category: "Path"
    });

    /* ==========================================================
       SYNTHÈSE
       ========================================================== */

    assert.printSummary("testPath");
}