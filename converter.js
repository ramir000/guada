const XLSX = require("xlsx");
const fs = require("fs");

const dbFile = "Base de Datos GDO.xlsx";
const otherFile = "TAREAS_PLANES_GDO_PLANTAS - Copia.xlsx";

const args = process.argv.slice(2);

function getArgument(name, defaultValue = undefined) {
    const index = args.indexOf(`--${name}`);
    return index >= 0 ? args[index + 1] : defaultValue;
}

function normalize(value) {
    return String(value ?? "")
        .trim()
        .toLowerCase();
}

function loadSheet(filePath, sheetName) {
    const workbook = XLSX.readFile(filePath);

    const selectedSheet =
        sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(selectedSheet)) {
        throw new Error(
            `La hoja "${selectedSheet}" no existe en ${filePath}. Hojas disponibles: ${workbook.SheetNames.join(", ")}`
        );
    }

    const worksheet = workbook.Sheets[selectedSheet];

    return XLSX.utils.sheet_to_json(worksheet, {
        defval: ""
    });
}

function filterRows(rows, column, value) {
    if (!rows.length) {
        return [];
    }

    if (!(column in rows[0])) {
        throw new Error(
            `La columna "${column}" no existe. Columnas disponibles: ${Object.keys(rows[0]).join(", ")}`
        );
    }

    return rows.filter((row) => {
        return normalize(row[column]) === normalize(value);
    });
}

try {
    const dbSheet = getArgument("db-sheet");
    const otherSheet = getArgument("other-sheet");

    const dbColumn = getArgument("db-column");
    const dbValue = getArgument("db-value");

    const otherColumn = getArgument("other-column");
    const otherValue = getArgument("other-value");

    if (!dbColumn || dbValue === undefined) {
        throw new Error(
            "Debes indicar --db-column y --db-value."
        );
    }

    if (!otherColumn || otherValue === undefined) {
        throw new Error(
            "Debes indicar --other-column y --other-value."
        );
    }

    const databaseRows = loadSheet(dbFile, dbSheet, 11);
    const otherRows = loadSheet(otherFile, otherSheet, 0);

    const filteredDatabaseRows = filterRows(
        databaseRows,
        dbColumn,
        dbValue
    );

    const filteredOtherRows = filterRows(
        otherRows,
        otherColumn,
        otherValue
    );

    const result = {
        database: {
            file: dbFile,
            sheet: dbSheet || "primera hoja",
            filter: {
                column: dbColumn,
                value: dbValue
            },
            total: filteredDatabaseRows.length,
            rows: filteredDatabaseRows
        },
        otherFile: {
            file: otherFile,
            sheet: otherSheet || "primera hoja",
            filter: {
                column: otherColumn,
                value: otherValue
            },
            total: filteredOtherRows.length,
            rows: filteredOtherRows
        }
    };

    fs.writeFileSync(
        "resultado.json",
        JSON.stringify(result, null, 2),
        "utf8"
    );

    console.log("Archivo resultado.json creado correctamente.");
    console.log(`Resultados de base de datos: ${filteredDatabaseRows.length}`);
    console.log(`Resultados del otro archivo: ${filteredOtherRows.length}`);
} catch (error) {
    console.error(`Error: ${error.message}`);
    process.exit(1);
}