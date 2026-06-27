public static loadRow(
    row: unknown[],
    databaseColumns: Record<string, number>,
    columns: Record<string, any>
): Record<string, unknown> {

    const result: Record<string, unknown> = {};

    for (const [columnName, excelColumn] of Object.entries(databaseColumns)) {

        const definition = columns[columnName];

        if (!definition?.load) continue;

        const property =
            definition.property
            ?? columnName;

        let value: unknown;

        if (definition.load.required) {

            value = WorkbookService.getRequiredValue({
                type: definition.load.type,
                row,
                col: excelColumn
            });

        } else {

            value = WorkbookService.getValueOrUndefined({
                type: definition.load.type,
                row,
                col: excelColumn
            });
        }

        if (
            value !== undefined
            && definition.load.parse
        ) {
            value = definition.load.parse(value);
        }

        result[property] = value;
    }

    return result;
}