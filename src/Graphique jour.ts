function main(workbook: ExcelScript.Workbook) {
    const wsTrains = workbook.getWorksheet("Trains");
    const wsLundi = workbook.getWorksheet("Lundi");
    wsLundi.getUsedRange()?.clear(); // vider la feuille avant de remplir
  
    // Récupérer le tableau "Trains"
    const table = wsTrains.getTable("Trains");
    const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  
    // Colonnes importantes (index 0-based)
    const COL_NUM_TRAIN = 2;
    const COL_DEP_HOUR = 8;
    const COL_DEP_STATION = 9;
    const COL_ARR_HOUR = 10;
    const COL_ARR_STATION = 11;
    const COL_NEXT = 30;
    const COL_PREV = 31;
  
    // Fonction pour convertir une heure Excel → index de colonne (0-based pour tableau JS)
    function timeToColumn(time: number): number {
      let totalMinutes = Math.floor(time * 24 * 60);
      let hour = Math.floor(totalMinutes / 60);
      let minute = totalMinutes % 60;
      return (hour - 3.5) * 4 + Math.floor(minute / 15); // 0-based
    }
  
    let allRows: (string | null)[][] = [];
    let maxCols = 0;
  
    // Parcourir les trains sans précédent
    rows.forEach((row, rowIndex) => {
      if (!row[COL_PREV]) {
        let outputRow: (string | null)[] = [];
        let currentIndex: number | null = rowIndex;
        outputRow[0] = currentIndex + 2;
        while (currentIndex !== null) {
          let train = rows[currentIndex];
          let numTrain = (train[COL_NUM_TRAIN] as string);
          let depTime = train[COL_DEP_HOUR] as number;
          let arrTime = train[COL_ARR_HOUR] as number;
          let depStation = train[COL_DEP_STATION] as string;
          let arrStation = train[COL_ARR_STATION] as string;
          let nextTrain = train[COL_NEXT] as number;
  
          let colDep = timeToColumn(depTime);
  
          // Départ : gare dans colonne précédente, train dans colonne de départ
          if (colDep > 0 && (outputRow[colDep - 1] === undefined || outputRow[colDep - 1] === null)) {
            outputRow[colDep - 1] = depStation;
          }
          outputRow[colDep] = '=' + numTrain + '&""';
  
          // Si pas de suivant → arrivée
          if (!nextTrain) {
            let colArr = timeToColumn(arrTime);
            outputRow[colArr] = arrStation;
          }
  
          // Avancer au train suivant
          currentIndex = nextTrain ? (nextTrain - 2) : null; // si colonne contient numéro de ligne (1-based)
        }
  
        allRows.push(outputRow);
        maxCols = Math.max(maxCols, outputRow.length);
      }
    });
  
    // Normaliser toutes les lignes à la même longueur
    let normalized: (string | null)[][] = allRows.map(row => {
      let newRow = new Array(maxCols).fill(null) as (string | number | boolean)[][];
      row.forEach((val, idx) => { if (val !== undefined) newRow[idx] = val; });
      return newRow;
    });
  
    // Écrire dans la feuille "Lundi"
    if (normalized.length > 0) {
      wsLundi.getRangeByIndexes(0, 0, normalized.length, maxCols).setValues(normalized);
    }
  }
  