const SHEET_TB = "TB";
const TB_CELL_GARE = "B1";
const TB_CELL_JOUR = "B2";
const TB_CELL_HEURE = "B3";
const TB_CELL_ARRIVEE_DEPART = "B4";
const TB_COLS_TRAINPARGARE = "D:F";

function main(workbook: ExcelScript.Workbook) {

  const sheet = workbook.getWorksheet("TB");
  const gare = sheet.getRange(TB_CELL_GARE).getValue();
  const jour = typeof sheet.getRange(TB_CELL_JOUR).getValue() === 'number' ? sheet.getRange(TB_CELL_JOUR).getValue() : 1;
  const heure = typeof sheet.getRange(TB_CELL_HEURE).getValue() === 'number' ? sheet.getRange(TB_CELL_HEURE).getValue() : 0;
  const arrivee_depart = typeof sheet.getRange(TB_CELL_ARRIVEE_DEPART).getValue() === 'number' ? sheet.getRange(TB_CELL_ARRIVEE_DEPART).getValue() : 0;

  const range = sheet.getRange(TB_COLS_TRAINPARGARE); // Sélectionner les colonnes entières
  range.clear();
  range.getCell(0, 0).setValue("Train");
  range.getCell(0, 1).setValue("Heure");
  range.getCell(0, 2).setValue("Réut");

  let trains = chargerTrains(workbook);
  chargerArrets(workbook, trains);
  let trainsParGare = getTrainsByGare(gare, jour, heure, trains);
  writeTrains(range, gare, trainsParGare);

  sheet.getRange("E:E").setNumberFormat("hh:mm:ss");
  console.log("Fini !");
}

function getTrainsByGare(gare: string, jour: number, heure: number, trains: Train[]): Train[] {
  const trainsByGare: Train[] = [];

  Object.keys(trains).forEach((trainKey) => {
    const train: Train = trains[trainKey];
    if (train.arrets && train.arrets[gare] && train.jour === jour && train.arrets[gare].getHeure() >= heure) { // Vérification si arrets est null
      const arret = Object.values(train.arrets).find((arret) => arret.gare === gare);
      if (arret) {
        trainsByGare.push(train);
      }
    }
  });

  // Trier les trains par ordre chronologique
  trainsByGare.sort((a, b) => {
    const heureA: number = a.arrets[gare].getHeure();
    const heureB: number = b.arrets[gare].getHeure();
    return heureA - heureB;
  });

  return trainsByGare;
}

function writeTrains(range: ExcelScript.Range, gare: string, trains: Train[]) {
  let row = 1;
  Object.keys(trains).forEach((trainKey) => {
    const train: Train = trains[trainKey];
    const arret: Arret = train.arrets[gare];
    if (arret) {
      range.getCell(row, 0).setValue(train.numero + arret.parite);
      range.getCell(row, 1).setValue(arret.getHeure());
      range.getCell(row, 2).setValue(train.reutilisation);
      row++;
    }
  });
}

class Train {
  numero: number;
  sens: number;
  jour: number;
  codeMission: string;
  heureDepart: number;
  gareDepart: string;
  heureArrivee: number;
  gareArrivee: string;
  reutilisation: string;
  arrets: { [abreviation: string]: Arret };

  constructor(numero: number, sens: number, jour: number, codeMission: string, heureDepart: string, gareDepart: string, heureArrivee: string, gareArrivee: string, reutilisation: string) {
    this.numero = numero;
    this.sens = sens;
    this.jour = jour;
    this.codeMission = codeMission;
    this.heureDepart = heureDepart;
    this.gareDepart = gareDepart;
    this.heureArrivee = heureArrivee;
    this.gareArrivee = gareArrivee;
    this.reutilisation = reutilisation;
    this.arrets = {};

  }

  ajouterArret(arret: Arret) {
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
const TRAINS_COL_REUTILISATION = 8;

function chargerTrains(workbook: ExcelScript.Workbook): Record<string, Train> {
  let sheet = workbook.getWorksheet(SHEET_TRAINS);
  if (!sheet) {
    console.log("La feuille " + SHEET_TRAINS + " n'existe pas !");
    return {};
  }
  const table = sheet.getTable(TABLE_TRAINS);
  if (!table) {
    console.log("Le tableau " + TABLE_TRAINS + " n'existe pas !");
    return {};
  }
  const data = table.getRange().getValues();

  let trains: Record<string, Train> = {};
  for (let i = 0; i < data.length; i++) {
    let numero = data[i][TRAINS_COL_NUMERO] as number;
    let sens = data[i][TRAINS_COL_SENS] as number;
    let jour = data[i][TRAINS_COL_JOUR] as number;
    let codeMission = data[i][TRAINS_COL_CODE_MISSION] as string;
    let heureDepart = data[i][TRAINS_COL_HEURE_DEPART] as string;
    let gareDepart = data[i][TRAINS_COL_GARE_DEPART] as string;
    let heureArrivee = data[i][TRAINS_COL_HEURE_ARRIVEE] as number;
    let gareArrivee = data[i][TRAINS_COL_GARE_ARRIVEE] as string;
    let reutilisation = data[i][TRAINS_COL_REUTILISATION] as number;

    let train = new Train(numero, sens, jour, codeMission, heureDepart, gareDepart, heureArrivee, gareArrivee, reutilisation);
    let key = numero + "_" + jour;
    trains[key] = train;
  }

  return trains;
}

class Arret {
  parite: number;
  gare: string;
  heureArrivee: number;
  heureDepart: number;
  voie: string;

  constructor(parite: number, gare: string, heureArrivee: string, heureDepart: string, voie: string) {
    this.parite = parite;
    this.gare = gare;
    this.heureArrivee = heureArrivee;
    this.heureDepart = heureDepart;
    this.voie = voie;
  }

  getHeure(): number {
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

function chargerArrets(workbook: ExcelScript.Workbook, trains: Record<string, Train>) {
  let sheet = workbook.getWorksheet(SHEET_ARRETS);
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
    let numeroTrain = data[i][ARRETS_COL_NUMERO_TRAIN] as string;
    let parite = data[i][ARRETS_COL_PARITE] as number;
    let jour = data[i][ARRETS_COL_JOUR] as number;
    let gare = data[i][ARRETS_COL_GARE] as string;
    let heureArrivee = data[i][ARRETS_COL_HEURE_ARRIVEE_ARRET] as number;
    let heureDepart = data[i][ARRETS_COL_HEURE_DEPART_ARRET] as number;
    let voie = data[i][ARRETS_COL_VOIE] as string;

    let arret = new Arret(parite, gare, heureArrivee, heureDepart, voie);
    let key = numeroTrain + "_" + jour;

    if (trains[key]) {
      trains[key].ajouterArret(arret);
    }
  }

  console.log("Arrêts ajoutés !");
}

class Gare {
  abreviation: string;
  pr: string;
  nom: string;
  pariteInversee: string;

  constructor(abreviation: string, pr: string, nom: string, pariteInversee: string) {
    this.abreviation = abreviation;
    this.pr = pr;
    this.nom = nom;
    this.pariteInversee = pariteInversee;
  }
}

const SHEET_GARES = "Param";
const TABLE_GARES = "Gares";
const GARES_COL_ABRV = 0;
const GARES_COL_NOM = 1;
const GARES_COL_PARITE_INVERSEE = 2;

function chargerGares(workbook: ExcelScript.Workbook): Gare[] {
  let sheet = workbook.getWorksheet(SHEET_GARES);
  if (!sheet) {
    console.log("La feuille " + SHEET_GARES + " n'existe pas !");
    return {};
  }
  const table = sheet.getTable(TABLE_GARES);
  if (!table) {
    console.log("Le tableau " + TABLE_GARES + " n'existe pas !");
    return {};
  }
  const data = table.getRange().getValues();

  let gares: Record<string, Gare> = {};
  for (let i = 1; i < data.length; i++) {
    const abreviation = data[i][GARES_COL_ABRV];
    const nom = data[i][GARES_COL_NOM];
    const pariteInversee = data[i][GARES_COL_PARITE_INVERSEE];

    let gare = new Gare(abreviation, nom, pariteInversee);
    gares[abreviation] = gare;
  }

  return gares;
}