function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();

    // Nettoyer les anciennes données dans les colonnes E
    sheet.getRange("E:Z").clear();

    // Récupérer les valeurs de départ, arrivée et via
    let start = sheet.getRange("B1").getValue() as string;
    let end = sheet.getRange("B2").getValue() as string;
    let via = (sheet.getRange("B3").getValue() as string).split(";");
    let rawConnections = sheet.getRange("A5:C400").getValues();

    // Créer les connexions et les variantes
    let { connections, variants } = createConnectionsAndVariants(rawConnections);

    // Générer toutes les combinaisons possibles de parcours
    let allCombinations = generateCombinations(start, end, via, variants);

    // Trouver le chemin le plus court parmi toutes les combinaisons
    let shortestPath = findShortestPath(connections, allCombinations);

    // Afficher le chemin le plus court dans la colonne E
    if (shortestPath) {
        sheet.getRange("E1").setValue("Chemin le plus court");
        sheet.getRange("E2").setValue(`Distance totale : ${shortestPath.totalDistance} min`);
        for (let i = 0; i < shortestPath.path.length; i++) {
            sheet.getRange(`E${i + 3}`).setValue(shortestPath.path[i]);
        }
    } else {
        sheet.getRange("E1").setValue("Aucun chemin trouvé");
    }
}

/**
 * Trouve le chemin le plus court parmi toutes les combinaisons possibles.
 * @param connections - La carte des connexions entre les gares.
 * @param allCombinations - La liste de toutes les combinaisons de parcours à évaluer.
 * @returns Un objet contenant le chemin le plus court et sa distance totale, ou null si aucun chemin n'est trouvé.
 */
function findShortestPath(connections: Map<string, Map<string, number>>, allCombinations: string[][]): { path: string[], totalDistance: number } | null {
    let shortestPath: { path: string[], totalDistance: number } | null = null;

    for (let combination of allCombinations) {
        // Calculer le chemin complet et la distance totale pour la combinaison actuelle
        let { path, totalDistance } = calculateCompletePath(connections, combination);

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
function calculateCompletePath(connections: Map<string, Map<string, number>>, combination: string[]): { path: string[], totalDistance: number } {
    let completePath: string[] = [];
    let totalDistance = 0;

    for (let i = 0; i < combination.length - 1; i++) {
        let segmentStart = combination[i];
        let segmentEnd = combination[i + 1];

        // Trouver le chemin le plus court pour le tronçon actuel
        let segmentPath = dijkstra(connections, segmentStart, segmentEnd);

        if (segmentPath.length === 0) {
            // Si aucun chemin n'est trouvé pour ce tronçon, retourner un chemin vide
            return { path: [], totalDistance: 0 };
        }

        // Calculer la distance pour ce tronçon
        let segmentDistance = calculatePathDistance(connections, segmentPath);

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
 * Calcule la distance totale d'un chemin donné.
 * @param connections - La carte des connexions entre les gares.
 * @param path - La liste ordonnée des gares constituant le chemin.
 * @returns La distance totale du chemin.
 */
function calculatePathDistance(connections: Map<string, Map<string, number>>, path: string[]): number {
    let distance = 0;

    for (let i = 0; i < path.length - 1; i++) {
        let from = path[i];
        let to = path[i + 1];
        distance += connections.get(from)?.get(to) ?? 0;
    }

    return distance;
}

/**
 * Cherche le chemin le plus court entre le départ et l'arrivée
 * en appliquant Dijkstra.
 * @param connections - La carte des connexions entre les gares.
 * @param start - La gare de départ.
 * @param end - La gare d'arrivée.
 * @returns Le chemin le plus court.
 */
function dijkstra(connections: Map<string, Map<string, number>>, start: string, end: string): string[] {
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

        for (let [neighbor, time] of connections.get(currentNode) || []) {
            let newDist = distances.get(currentNode) + time;
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
 * Crée les connexions et les variantes pour chaque gare.
 * @param rawConnections - La liste des connexions brutes.
 * @returns Un objet contenant les connexions et les variantes.
 */
function createConnectionsAndVariants(rawConnections: (string | number)[][]): {connections: Map<string, Map<string, number>>, variants: Map<string, string[]>} {
    let connections = new Map<string, Map<string, number>>();
    let variants = new Map<string, string[]>();

    for (let row of rawConnections) {
        let from = row[0] as string;
        let to = row[1] as string;
        let time = row[2] as number;

        if (!connections.has(from)) {
            connections.set(from, new Map<string, number>());
        }
        connections.get(from).set(to, time);

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