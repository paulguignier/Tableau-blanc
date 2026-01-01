function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();

    // Nettoyer les anciennes données dans les colonnes E
    sheet.getRange("E:Z").clear();

    // Récupérer les valeurs de départ, arrivée et via
    let start = sheet.getRange("B1").getValue() as string;
    let end = sheet.getRange("B2").getValue() as string;
    let via = (sheet.getRange("B3").getValue() as string).split(";");
    let changeTime = sheet.getRange("B4").getValue() as number;  // Temps de changement de sens
    let rawConnections = sheet.getRange("A5:D400").getValues();

    // Créer les connexions et les variantes
    let { connections, variants } = createConnectionsAndVariants(rawConnections);

    // Générer toutes les combinaisons possibles de parcours
    let allCombinations = generateCombinations(start, end, via, variants);

    // Trouver le chemin le plus court parmi toutes les combinaisons
    let shortestPath = findShortestPath(connections, allCombinations, changeTime);

    // Afficher le chemin le plus court dans la colonne F
    if (shortestPath) {
        sheet.getRange("F1").setValue("Chemin le plus court");
        sheet.getRange("F2").setValue(`Distance totale : ${shortestPath.totalDistance} min`);
        for (let i = 0; i < shortestPath.path.length; i++) {
            sheet.getRange(`F${i + 3}`).setValue(shortestPath.path[i]);
        }
    } else {
        sheet.getRange("F1").setValue("Aucun chemin trouvé");
    }
}

function createConnectionsAndVariants(rawConnections: (string | number | boolean)[][]): {connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, variants: Map<string, string[]>} {
    let connections = new Map<string, Map<string, { time: number, needsTurnaround: boolean }>>();
    let variants = new Map<string, string[]>();

    for (let row of rawConnections) {
        let from = row[0] as string;
        let to = row[1] as string;
        let time = row[2] as number;
        let needsTurnaround = (row[3] === 'RBT') as boolean;

        if (!connections.has(from)) {
            connections.set(from, new Map<string, { time: number, needsTurnaround: boolean }>());
        }
        connections.get(from).set(to, { time, needsTurnaround });

        // Créer les variantes pour la gare 'from'
        let baseFrom = from.split('_')[0];
        if (!variants.has(baseFrom)) {
            variants.set(baseFrom, []);
        }
        if (!variants.get(baseFrom)!.includes(from)) {
            variants.get(baseFrom)!.push(from);
        }

        // Créer les variantes pour la gare 'to'
        let baseTo = to.split('_')[0];
        if (!variants.has(baseTo)) {
            variants.set(baseTo, []);
        }
        if (!variants.get(baseTo)!.includes(to)) {
            variants.get(baseTo)!.push(to);
        }
    }

    return { connections, variants };
}

/**
 * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
 * @param connections - La carte des connexions entre les gares.
 * @param allCombinations - La liste de toutes les combinaisons de parcours à évaluer.
 * @param changeTime - Temps de changement de sens.
 * @returns Un objet contenant le chemin le plus court et sa distance totale, ou null si aucun chemin n'est trouvé.
 */
function findShortestPath(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, allCombinations: string[][], changeTime: number): { path: string[], totalDistance: number } | null {
    let shortestPath: { path: string[], totalDistance: number } | null = null;

    for (let combination of allCombinations) {
        // Calculer le chemin complet et la distance totale pour la combinaison actuelle
        let { path, totalDistance } = calculateCompletePath(connections, combination, changeTime);

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
 * @param connections - La carte des connexions entre les gares.
 * @param combination - La liste ordonnée des gares à parcourir.
 * @returns Un objet contenant le chemin complet et la distance totale.
 */
function calculateCompletePath(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, combination: string[], changeTime: number): { path: string[], totalDistance: number } {
    let completePath: string[] = [];
    let totalDistance = 0;

    for (let i = 0; i < combination.length - 1; i++) {
        let segmentStart = combination[i];
        let segmentEnd = combination[i + 1];

        // Trouver le chemin le plus court pour le tronçon actuel
        let segmentPath = dijkstra(connections, segmentStart, segmentEnd, changeTime);

        if (segmentPath.length === 0) {
            // Si aucun chemin n'est trouvé pour ce tronçon, retourner un chemin vide
            return { path: [], totalDistance: 0 };
        }

        // Calculer la distance pour ce tronçon
        let segmentDistance = calculatePathTime(connections, segmentPath, changeTime);

        // Ajouter la distance du tronçon à la distance totale
        totalDistance += segmentDistance;

        // Ajouter le chemin du tronçon au chemin complet
        // Éviter de dupliquer les gares intermédiaires
        if (completePath.length > 0) {
            segmentPath.shift(); // Retirer la première gare pour éviter la duplication
        }
        completePath.push(...segmentPath);
    }

    return { path: completePath, totalDistance };
}

/**
 * Calcule le temps total pour un chemin donné en tenant compte des temps de trajet
 * et des éventuels temps de changement de sens.
 * @param connections - La carte des connexions entre les gares, incluant le temps de trajet et l'information sur le besoin de changement de sens.
 * @param path - La liste ordonnée des gares constituant le chemin.
 * @param changeTime - Le temps de changement de sens à ajouter lorsque nécessaire.
 * @returns Le temps total du chemin, incluant les temps de trajet et de changement de sens.
 */
function calculatePathTime(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean}>>, path: string[], changeTime: number): number {
    let totalTime  = 0;

    for (let i = 0; i < path.length - 1; i++) {
        let from = path[i];
        let to = path[i + 1];
        let connection = connections.get(from)?.get(to);
        if (connection) {
            totalTime += connection.time;
            // Ajouter le temps de changement de sens sauf pour le premier segment
            if (connection.needsTurnaround) {
                totalTime += changeTime;
            }
        }
    }

    return totalTime ;
}

/**
 * Cherche le chemin le plus court entre le départ et l'arrivée
 * en appliquant Dijkstra.
 * @param connections - La carte des connexions entre les gares.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @returns Le chemin le plus court.
 */
function dijkstra(connections: Map<string, Map<string, { time: number, needsTurnaround: boolean }>>, start: string, end: string, changeTime: number): string[] {
    let distances = new Map<string, number>();
    let previousNodes = new Map<string, string | null>();
    let unvisited = new Set<string>(connections.keys());
    let path: string[] = [];

    // Initialisation des distances
    for (let node of unvisited) {
        distances.set(node, Infinity);
        previousNodes.set(node, null);
    }
    distances.set(start, 0);

    while (unvisited.size > 0) {
        let currentNode = Array.from(unvisited).reduce((minNode, node) =>
            distances.get(node) < distances.get(minNode) ? node : minNode
        );

        if (distances.get(currentNode) === Infinity) break; // Aucun chemin

        unvisited.delete(currentNode);

        // Examiner les voisins avec les nouveaux attributs
        for (let [neighbor, { time, needsTurnaround }] of connections.get(currentNode) || []) {
            let additionalTime = time;
            if (needsTurnaround && currentNode !== start) {  // Si un changement de sens est nécessaire, ajouter du temps
                additionalTime += changeTime;
            }
            
            let newDist = distances.get(currentNode) + additionalTime;
            if (newDist < distances.get(neighbor)) {
                distances.set(neighbor, newDist);
                previousNodes.set(neighbor, currentNode);
            }
        }
    }

    // Retracer le chemin
    let step = end;
    while (step) {
        path.unshift(step);
        step = previousNodes.get(step);
    }

    // Si le chemin est valide
    return path[0] === start ? path : [];
}

/**
 * Génère toutes les combinaisons de routes possibles pour aller de start à end en passant par les gares intermédiaires via.
 * Les variantes de chaque gare sont incluses en utilisant la Map variants.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @param via - Les gares intermédiaires à passer par.
 * @param variants - La Map qui permet de récupérer les variantes pour chaque gare.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une combinaison de route possible.
 */
function generateCombinations(start: string, end: string, via: string[], variants: Map<string, string[]>): string[][] {
    // Filtrer les gares intermédiaires pour éliminer les chaînes vides
    let filteredVia = via.filter(v => v.trim() !== "");

    // Générer les permutations des gares intermédiaires
    let viaPermutations = permute(filteredVia);

    // Ajouter start au début et end à la fin de chaque permutation
    let routes = viaPermutations.map(permutation => [start, ...permutation, end]);

    // Étendre chaque route pour inclure toutes les variantes possibles
    let allCombinations: string[][] = routes.flatMap(route => expandPermutations(route, variants));

    return allCombinations;
}   
  
/**
 * Renvoie toutes les variantes possibles pour une gare.
 * Une variante correspond au sens de passage dans la gare : GARE_1 en impair, GARE_2 en pair.
 * Seules les gares de retournement permettent de passer d'une gare à l'autre
 * Si la gare a déjà un suffixe imposé (_), renvoie uniquement cette gare avec suffixe [gare].
 * Sinon, renvoie toutes les variantes associées.
 * @param gare - La gare dont on cherche les variantes.
 * @param variants - La Map qui permet de récupérer les variantes pour chaque gare.
 * @returns Un tableau contenant toutes les variantes possibles pour la gare.
 */
function getAllVariants(gare: string, variants: Map<string, string[]>): string[] {
    // Si la gare a un suffixe (_), renvoyer uniquement [gare]
    if (gare.includes('_')) {
        return [gare];
    }
    // Sinon, renvoyer toutes les variantes associées
    return variants.get(gare) || [];
}
  
/**
 * Génère toutes les permutations possibles d'un tableau de chaînes.
 * @param arr - Le tableau de chaînes à permuter.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une permutation possible.
 */
function permute(arr: string[]): string[][] {
    if (arr.length === 0) return [[]];
    if (arr.length === 1) return [[arr[0]]];

    let result: string[][] = [];

    for (let i = 0; i < arr.length; i++) {
        let rest = [...arr.slice(0, i), ...arr.slice(i + 1)];
        let restPermutations = permute(rest);
        
        for (let perm of restPermutations) {
            result.push([arr[i], ...perm]);
        }
    }

    return result;
}

/**
 * Étend une permutation de gares pour inclure toutes les variantes possibles.
 * @param permutation - La permutation de gares à étendre.
 * @param variants - La Map qui permet de récupérer les variantes pour chaque gare.
 * @returns Un tableau de tableaux, chaque sous-tableau représentant une permutation possible avec toutes les variantes.
 */
function expandPermutations(permutation: string[], variants: Map<string, string[]>): string[][] {
    if (permutation.length === 0) return [[]];

    let first = getAllVariants(permutation[0], variants);
    let restExpanded = expandPermutations(permutation.slice(1), variants);

    let result: string[][] = [];
    for (let f of first) {
        for (let r of restExpanded) {
            result.push([f, ...r]);
        }
    }

    return result;
}