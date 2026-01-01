"use strict";
const SHEET_TB = "TB";
const TB_CELL_GARE = "B1";
const TB_CELL_JOUR = "B2";
const TB_CELL_HEURE = "B3";
const TB_CELL_ARRIVEE_DEPART = "B4";
const TB_COLS_TRAINPARGARE = "D:F";
var WORKBOOK;
function main(workbook) {
    WORKBOOK = workbook;
    let sheet = WORKBOOK.getWorksheet(SHEET_TB);
    let gare = sheet.getRange(TB_CELL_GARE).getValue();
    let jour = (typeof sheet.getRange(TB_CELL_JOUR).getValue() === 'number' ? sheet.getRange(TB_CELL_JOUR).getValue() : 1);
    let heure = (typeof sheet.getRange(TB_CELL_HEURE).getValue() === 'number' ? sheet.getRange(TB_CELL_HEURE).getValue() : 0);
    let arrivee_depart = (typeof sheet.getRange(TB_CELL_ARRIVEE_DEPART).getValue() === 'number' ? sheet.getRange(TB_CELL_ARRIVEE_DEPART).getValue() : 0);
    let range = sheet.getRange(TB_COLS_TRAINPARGARE); // Sélectionner les colonnes entières
    range.clear();
    range.getCell(0, 0).setValue("Train");
    range.getCell(0, 1).setValue("Heure");
    range.getCell(0, 2).setValue("Réut");
    let trains = chargerTrains();
    chargerArrets(trains);
    let trainsParGare = getTrainsByGare(gare, jour, heure, trains);
    writeTrains(range, gare, trainsParGare);
    sheet.getRange("E:E").setNumberFormat("hh:mm:ss");
    console.log("Fini !");
}
/**
 * Renvoie les trains passant par la gare donnée le jour donné
 * après l'heure donnée.
 * Les trains sont triés par ordre chronologique.
 * @param {string} gare - Gare
 * @param {number} jour - Jour
 * @param {number} heure - Heure
 * @param {Record<string, Train>} trains - Trains
 * @returns {Train[]} - Trains passant par la gare
 */
function getTrainsByGare(gare, jour, heure, trains) {
    let trainsByGare = [];
    Object.keys(trains).forEach((trainKey) => {
        let train = trains[trainKey];
        if (train.arrets && train.arrets[gare] && train.jour === jour && train.arrets[gare].getHeure() >= heure) { // Vérification si arrets est null
            let arret = Object.values(train.arrets).find((arret) => arret.gare === gare);
            if (arret) {
                trainsByGare.push(train);
            }
        }
    });
    // Trier les trains par ordre chronologique
    trainsByGare.sort((a, b) => {
        let heureA = a.arrets[gare].getHeure();
        let heureB = b.arrets[gare].getHeure();
        return heureA - heureB;
    });
    return trainsByGare;
}
/**
 * Ecrit les trains passant par la gare donnée dans le range
 * fourni. Les colonnes sont : Train, Heure, Réut.
 * @param {ExcelScript.Range} range - Range
 * @param {string} gare - Gare
 * @param {Train[]} trains - Trains
 */
function writeTrains(range, gare, trains) {
    let row = 1;
    Object.keys(trains).forEach((trainKey) => {
        let train = trains[trainKey];
        let arret = train.arrets[gare];
        if (arret) {
            range.getCell(row, 0).setValue(train.numero + arret.parite);
            range.getCell(row, 1).setValue(arret.getHeure());
            range.getCell(row, 2).setValue(train.reutilisation);
            row++;
        }
    });
}
class Train {
    /**
     * Constructeur d'un train.
     * @param {number} numero - Numéro du train
     * @param {number} sens - Sens du train
     * @param {number} jour - Jour du train
     * @param {string} codeMission - Code de mission du train
     * @param {number} heureDepart - Heure de départ du train
     * @param {string} gareDepart - Gare de départ du train
     * @param {number} heureArrivee - Heure d'arrivée du train
     * @param {string} gareArrivee - Gare d'arrivée du train
     * @param {string} garesVia - Gares via
     * @param {string} reutilisation - Reutilisation du train
     */
    constructor(numero, sens, jour, codeMission, heureDepart, gareDepart, heureArrivee, gareArrivee, garesVia, reutilisation) {
        this.numero = numero;
        this.sens = sens;
        this.jour = jour;
        this.codeMission = codeMission;
        this.heureDepart = heureDepart;
        this.gareDepart = gareDepart;
        this.heureArrivee = heureArrivee;
        this.gareArrivee = gareArrivee;
        this.garesVia = garesVia;
        this.reutilisation = reutilisation;
        this.arrets = {};
        let arret = new Arret(0, gareDepart, 0, heureDepart, "0");
        this.ajouterArret(arret);
        arret = new Arret(0, gareArrivee, heureArrivee, 0, "0");
        this.ajouterArret(arret);
    }
    /**
     * Ajoute un arret au train.
     * @param {Arret} arret - Arret à ajouter
     */
    ajouterArret(arret) {
        this.arrets[arret.gare] = arret;
    }
}
const SHEET_TRAINS = "Trains";
const TABLE_TRAINS = "Trains";
const TRAINS_COL_NUMERO = 0;
const TRAINS_COL_SENS = 1;
const TRAINS_COL_JOUR = 2;
const TRAINS_COL_CODE_MISSION = 3;
const TRAINS_COL_HEURE_DEPART = 4;
const TRAINS_COL_GARE_DEPART = 5;
const TRAINS_COL_HEURE_ARRIVEE = 6;
const TRAINS_COL_GARE_ARRIVEE = 7;
const TRAINS_COL_GARES_VIA = 8;
const TRAINS_COL_REUTILISATION = 9;
/**
 * Charge les trains à partir du tableau "Trains" de la feuille "Trains".
 * Les trains sont stockés dans un objet avec comme clés le numéro de train
 * suivi du jour et comme valeur l'objet Train.
 * @returns Un objet contenant les trains.
 */
function chargerTrains() {
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
        let numero = data[i][TRAINS_COL_NUMERO];
        let sens = data[i][TRAINS_COL_SENS];
        let jour = data[i][TRAINS_COL_JOUR];
        let codeMission = data[i][TRAINS_COL_CODE_MISSION];
        let heureDepart = data[i][TRAINS_COL_HEURE_DEPART];
        let gareDepart = data[i][TRAINS_COL_GARE_DEPART];
        let heureArrivee = data[i][TRAINS_COL_HEURE_ARRIVEE];
        let gareArrivee = data[i][TRAINS_COL_GARE_ARRIVEE];
        let garesVia = data[i][TRAINS_COL_GARES_VIA];
        let reutilisation = data[i][TRAINS_COL_REUTILISATION];
        let train = new Train(numero, sens, jour, codeMission, heureDepart, gareDepart, heureArrivee, gareArrivee, garesVia, reutilisation);
        let key = numero + "_" + jour;
        trains[key] = train;
    }
    return trains;
}
class Arret {
    /**
     * Constructeur d'un arret.
     * @param {number} parite - Parit  de l'arret
     * @param {string} gare - Gare de l'arret
     * @param {number} heureArrivee - Heure d'arriv e  l'arret
     * @param {number} heureDepart - Heure de d part  l'arret
     * @param {string} voie - Voie de l'arret
     */
    constructor(parite, gare, heureArrivee, heureDepart, voie) {
        this.parite = parite;
        this.gare = gare;
        this.heureArrivee = heureArrivee;
        this.heureDepart = heureDepart;
        this.voie = voie;
    }
    /**
     * Renvoie l'heure d'arriv e  ou de d part  l'arret.
     * Si l'heure d'arriv e  est d finie, renvoie cette heure.
     * Sinon, renvoie l'heure de d part.
     * @returns {number} L'heure d'arriv e  ou de d part.
     */
    getHeure() {
        return this.heureArrivee ? this.heureArrivee : this.heureDepart;
    }
}
const SHEET_ARRETS = "Arrêts";
const TABLE_ARRETS = "Arrêts";
const ARRETS_COL_NUMERO_TRAIN = 0;
const ARRETS_COL_PARITE = 1;
const ARRETS_COL_JOUR = 2;
const ARRETS_COL_GARE = 3;
const ARRETS_COL_HEURE_ARRIVEE_ARRET = 4;
const ARRETS_COL_HEURE_DEPART_ARRET = 5;
const ARRETS_COL_VOIE = 6;
function chargerArrets(trains) {
    let sheet = WORKBOOK.getWorksheet(SHEET_ARRETS);
    if (!sheet) {
        console.log("La feuille " + SHEET_ARRETS + " n'existe pas !");
        return {};
    }
    const table = sheet.getTable(TABLE_ARRETS);
    if (!table) {
        console.log("Le tableau " + TABLE_ARRETS + " n'existe pas !");
        return {};
    }
    const data = table.getRange().getValues();
    for (let i = 0; i < data.length; i++) {
        let numeroTrain = data[i][ARRETS_COL_NUMERO_TRAIN];
        let parite = data[i][ARRETS_COL_PARITE];
        let jour = data[i][ARRETS_COL_JOUR];
        let gare = data[i][ARRETS_COL_GARE];
        let heureArrivee = data[i][ARRETS_COL_HEURE_ARRIVEE_ARRET];
        let heureDepart = data[i][ARRETS_COL_HEURE_DEPART_ARRET];
        let voie = data[i][ARRETS_COL_VOIE];
        let arret = new Arret(parite, gare, heureArrivee, heureDepart, voie);
        let key = numeroTrain + "_" + jour;
        if (trains[key]) {
            trains[key].ajouterArret(arret);
        }
    }
    console.log("Arrêts ajoutés !");
}
class Station {
    /**
     * Constructeur d'une gare.
     * @param {string} abbreviation - Abréviation de la gare
     * @param {string} name - Nom de la gare
     * @param {boolean} reversedParity - Parité de la gare (true si inversée, false sinon)
     * @param {string[]} variants - Variantes de la gare
     */
    constructor(abbreviation, name, reversedParity, variants = []) {
        this.abbreviation = abbreviation;
        this.name = name;
        this.reversedParity = reversedParity;
        this.variants = variants;
    }
}
const SHEET_GARES = "Param";
const TABLE_GARES = "Gares";
const GARES_COL_ABRV = 0;
const GARES_COL_NOM = 1;
const GARES_COL_PARITE_INVERSEE = 2;
/**
 * Charge les gares à partir du tableau "Gares" de la feuille "Param".
 * @returns Un objet contenant les gares sous forme de clés (abréviation) et de valeurs (objets Gare).
 */
function chargerGares() {
    let gares = {};
    let sheet = WORKBOOK.getWorksheet(SHEET_GARES);
    if (!sheet) {
        console.log("La feuille " + SHEET_GARES + " n'existe pas !");
        return gares;
    }
    const table = sheet.getTable(TABLE_GARES);
    if (!table) {
        console.log("Le tableau " + TABLE_GARES + " n'existe pas !");
        return gares;
    }
    let data = table.getRange().getValues();
    for (let i = 1; i < data.length; i++) {
        let abreviation = data[i][GARES_COL_ABRV];
        let nom = data[i][GARES_COL_NOM];
        ;
        let pariteInversee = data[i][GARES_COL_PARITE_INVERSEE];
        let gare = new Gare(abreviation, nom, pariteInversee);
        gares[abreviation] = gare;
    }
    return gares;
}
