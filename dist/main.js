"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const ChargementTrains_1 = require("./ChargementTrains");
// Simule des données de test ici
(0, ChargementTrains_1.loadParams)();
(0, ChargementTrains_1.loadConnections)();
const combinaisons = (0, ChargementTrains_1.generateCombinations)("MPU", "ETP", "PRU");
console.log((0, ChargementTrains_1.findShortestPath)(combinaisons));
