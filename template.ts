/**
 * Titre du fichier
 * 
 * Code Excel Automate pour .
 * 
 * @author Paul Guignier
 * @version 1.0
 * @package scr\FileName.ts
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

/**
 * Commentaire de fonction en français avec retour à la ligne avec alinéa
 *  si la ligne dépasse la partie visible, et avec un point à la fin de chaque phrase.
 * Une deuxième phrase est possible sans alinéa.
 * @param {string | undefined} param1 Paramètre 1 sans article devant,
 *  et avec un point à la fin.
 * @param {boolean} [param2=true] Si vrai (par défaut), donne l'action si vrai. Si faux,
 *  indique l'action si faux. La valeur par défaut est mise avec l'intitulé du paramètre
 *  entre crochets, puis dans le commentaire avec la mention par défaut entre parenthèses.
 * @returns {string} Valeur retournée avec un point à la fin.
 */
function myFunction(param1: string | undefined, param2: boolean = true): string {
    // Commentaire de ligne en français sans ponctuation

    /* Commentaire avec ponctuation.
     * Chaque ligne ne doit pas dépasser du cadre, un retour à la ligne est nécessaire
     *  avec alinéa (1 espace), si possible avant un opérateur. Dans ce cas un saut de ligne
     *  est nécessaire ensuite (sauf commentaires). */

    if (param2) {
        // Commentaire d'application de la condition
    }

    return param1!;
}

/**
 * Classe ClassName qui défini tel élément.
 */
class ClassName {

    /* Base de données des objets ClassName. */
    public static readonly DATAS = new Map<string, ClassName>();

    /* Cache pour l'extraction des éléments. */
    private static readonly CACHE = new Map<string, Map<string, number[]>>();

    property1: string;          // Propriété 1
    property2: string;          // Propriété 2
    property3?: ClassName;      // Propriété 3 décrivant l'objet d'une classe. Si texte long,
                                //  retour à la ligne avec un espace pour alinéa. Si l'objet ne
                                //  peut être défini dans le constructeur, mettre un ?

    /**
     * Constructeur de la classe ClassName.
     * @param {string} property1 Paramètre pour alimenter la propriété 1.
     * @param {string} property2 Paramètre pour alimenter la propriété 2.
     * @param {string} property3Name Abréviation du jour ou du groupe de jours de la semaine.
     */
    constructor(property1: string, property2: string, property3Name: string) {
        this.property1 = property1;
        this.property2 = property2;
        this.property3 = ClassName.DATAS.get(property3Name);
    }

    /**
     * Méthodes
     */
    
    
    /* Constantes de lecture du tableau Excel. */
    private static readonly SHEET = "Feuille";
    private static readonly TABLE = "Tableau";
    private static readonly COL_NOM1 = 0;
    private static readonly COL_NOM2 = 1;
    private static readonly COL_NOM3 = 2;

    /**
     * Charge les objets à partir du tableau "Tableau" de la feuille "Feuille".
     * Les objets sont stockés dans la base de données ClassName.DATAS.
     */
    public static loadFromExcel() { 

        const data = getDataFromTable(ClassName.SHEET, ClassName.TABLE);

        for (const  row of data.slice(1)) {
            // Vérifie si la ligne est vide (toutes les valeurs nulles ou vides)
            if (row.every(cell => !cell)) continue;

            // Extrait les valeurs
            let property1 = String(row[Day.COL_NOM1]);
            let fullName = String(row[Day.COL_FULL_NAME]);
            let abreviation = String(row[Day.COL_ABBREVIATION]);

            // Crée l'objet Day
            let day = new Day(numbersString, fullName, abreviation);
            PARAM.days.set(day.numbersString, day);
        }
    }
}