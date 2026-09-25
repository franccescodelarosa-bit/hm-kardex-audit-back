
import { Rule003Exporter } from "./Rule003Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

function baseMetadata(overrides: Partial<any> = {}) {
    return {
        fromIndex: 11,
        toIndex: 12,
        finalBalance: { quantity: 100, unitCost: 8, totalCost: 904 },
        initialBalance: { quantity: 100, unitCost: 7, totalCost: 791 },
        differences: ["Costo Unitario", "Costo Total"],
        ...overrides
    };
}

describe("Rule003Exporter", () => {
    it("muestra SOLO la fila del campo con diferencia (decisión del equipo): si solo difiere el Costo Total, no aparece la fila de Costo Unitario", async () => {
        const exporter = new Rule003Exporter();
        const results = [
            {
                product_code: "006749",
                product_name: "PEGAMENTO E/LATA DE 1/32 T/MED.M/AFRICANO",
                risk_level: "ALTO",
                metadata: baseMetadata({
                    // Caso real: Costo Unitario identico (4.04 = 4.04), solo Costo Total difirio
                    finalBalance: { quantity: 247, unitCost: 4.04, totalCost: 997.88 },
                    initialBalance: { quantity: 251, unitCost: 4.04, totalCost: 1014.04 },
                    differences: ["Costo Total"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 de header + 1 sola fila
        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Total");
        expect(sheet.getRow(5).getCell(5).value).toBe(997.88);
        expect(sheet.getRow(5).getCell(6).value).toBe(1014.04);
        const trace = String(sheet.getRow(5).getCell(10).value);
        expect(trace).toContain("Campos con diferencia: Costo Total");
        expect(trace).not.toContain("Costo Unitario");
    });

    it("si solo difiere el Costo Unitario, muestra solo esa fila", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: baseMetadata({
                finalBalance: { quantity: 10, unitCost: 5, totalCost: 50 },
                initialBalance: { quantity: 10, unitCost: 6, totalCost: 50 },
                differences: ["Costo Unitario"]
            })
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Unitario");
    });

    it("registros viejos sin 'differences': muestra solo los campos que realmente difieren", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: baseMetadata({
                finalBalance: { quantity: 10, unitCost: 0, totalCost: 0.05 },
                initialBalance: { quantity: 10, unitCost: 0, totalCost: 0 },
                differences: undefined
            })
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Total");
    });

    it("si las dos difieren, muestra las 2 filas, cada trazabilidad habla SOLO de su propio campo (sin mezclar)", async () => {
        const exporter = new Rule003Exporter();
        const results = [
            {
                product_code: "019706",
                product_name: "SHORT B/VARON TELA WING JASPEADO C/BOLS",
                risk_level: "ALTO",
                metadata: baseMetadata()
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6);

        const unitRowTrace = String(sheet.getRow(5).getCell(10).value);
        expect(unitRowTrace).toContain("Campos con diferencia: Costo Unitario");
        expect(unitRowTrace).not.toContain("Costo Total");

        const totalRowTrace = String(sheet.getRow(6).getCell(10).value);
        expect(totalRowTrace).toContain("Campos con diferencia: Costo Total");
        expect(totalRowTrace).not.toContain("Costo Unitario");
    });

    it("si el producto no existe en el mes siguiente (initialBalance null), no revienta y muestra una fila clara", async () => {
        const exporter = new Rule003Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "ALTO",
                metadata: {
                    fromIndex: 3,
                    toIndex: 4,
                    finalBalance: { quantity: 50, unitCost: 5, totalCost: 250 },
                    initialBalance: null,
                    differences: undefined
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 de header + 1 sola fila
        expect(sheet.getRow(5).getCell(5).value).toBe(250);
        expect(sheet.getRow(5).getCell(6).value).toBe("Producto no encontrado");
        expect(sheet.getRow(5).getCell(7).value).toBe("No aplicable");
        const trace = String(sheet.getRow(5).getCell(10).value);
        expect(trace).not.toContain("null");
        expect(trace).toContain("no tiene Kardex registrado");
    });

    it("producto no encontrado que cerró en 0: 'Encontrado' dice 'Producto no encontrado', no un 0 que parezca sin diferencia", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            product_code: "000144",
            product_name: "PRODUCTO AGOTADO",
            risk_level: "ALTO",
            metadata: {
                fromIndex: 3,
                toIndex: 4,
                finalBalance: { quantity: 0, unitCost: 0, totalCost: 0 },
                initialBalance: null
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(4).value).toBe("Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(0);
        expect(sheet.getRow(5).getCell(6).value).toBe("Producto no encontrado");
        expect(sheet.getRow(5).getCell(7).value).toBe("No aplicable");
    });

    it("INITIAL_BALANCE_NOT_FOUND_NEXT_MONTH: 'Sin Saldo Inicial en el mes siguiente', encontrado dice que falta la op 16", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            error_type: "INITIAL_BALANCE_NOT_FOUND_NEXT_MONTH",
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: {
                fromIndex: 3,
                toIndex: 4,
                finalBalance: { quantity: 10, unitCost: 5, totalCost: 50 },
                initialBalance: null
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(5).getCell(1).value).toBe("Marzo → Abril");
        expect(sheet.getRow(5).getCell(4).value).toBe("Sin Saldo Inicial en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(50);
        expect(sheet.getRow(5).getCell(6).value).toBe("Sin Saldo Inicial (TipoOp 16)");
        expect(sheet.getRow(5).getCell(7).value).toBe("No aplicable");
        const trace = String(sheet.getRow(5).getCell(10).value);
        expect(trace).toContain("no tiene Saldo Inicial (TipoOp 16) en Abril");
        expect(trace).not.toContain("null");
    });

    it("PRODUCT_NOT_FOUND_NEXT_MONTH con missingFields: una fila por campo distinto de 0, esperado = valor de ese campo", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            error_type: "PRODUCT_NOT_FOUND_NEXT_MONTH",
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: {
                fromIndex: 1,
                toIndex: 2,
                finalBalance: { quantity: 10, unitCost: 0, totalCost: 50 },
                initialBalance: null,
                missingFields: ["Cantidad", "Costo Total"]
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6); // 4 de header + 2 campos
        expect(sheet.getRow(5).getCell(4).value).toBe("Cantidad - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(10);
        expect(sheet.getRow(5).getCell(6).value).toBe("Producto no encontrado");
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Total - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(6).getCell(5).value).toBe(50);
    });

    it("INITIAL_BALANCE_NOT_FOUND_NEXT_MONTH con missingFields: una fila por campo distinto de 0", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            error_type: "INITIAL_BALANCE_NOT_FOUND_NEXT_MONTH",
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: {
                fromIndex: 1,
                toIndex: 2,
                finalBalance: { quantity: 0, unitCost: 142.06, totalCost: 0 },
                initialBalance: null,
                missingFields: ["Costo Unitario"]
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario - Sin Saldo Inicial en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(142.06);
        expect(sheet.getRow(5).getCell(6).value).toBe("Sin Saldo Inicial (TipoOp 16)");
    });
});

