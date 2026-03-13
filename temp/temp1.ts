public expandRoutes(): string[][] {

    let result: string[][] = [[]];

    for (const group of this.routeStations) {

        const perms = Path.permutations(group);
        const newResult: string[][] = [];

        for (const base of result) {
            for (const perm of perms) {
                newResult.push([...base, ...perm]);
            }
        }

        result = newResult;
    }

    return result;
}