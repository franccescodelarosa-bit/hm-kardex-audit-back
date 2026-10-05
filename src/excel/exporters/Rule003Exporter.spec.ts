
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
    it("aunque solo difiera el Costo Total, muestra las DOS filas (Costo Unitario sin diferencia + Costo Total)", async () => {
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

        expect(sheet.rowCount).toBe(5); // 3 de header + las 2 filas de costo

        expect(sheet.getRow(4).getCell(4).value).toBe("Continuidad de Costo Unitario");
        expect(sheet.getRow(4).getCell(5).value).toBe(4.04);
        expect(sheet.getRow(4).getCell(6).value).toBe(4.04);
        expect(sheet.getRow(4).getCell(7).value).toBe(0);
        const unitTrace = String(sheet.getRow(4).getCell(10).value);
        // Primero la información del campo, y al final la aclaración.
        expect(unitTrace).toContain("Costo Unitario Final: 4.04");
        expect(unitTrace).toContain("Costo Unitario Inicial: 4.04");
        expect(unitTrace.split("\n").pop()).toBe("Sin diferencia en Costo Unitario");
        expect(unitTrace).not.toContain("Campos con diferencia");

        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Total");
        expect(sheet.getRow(5).getCell(5).value).toBe(997.88);
        expect(sheet.getRow(5).getCell(6).value).toBe(1014.04);
        const totalTrace = String(sheet.getRow(5).getCell(10).value);
        expect(totalTrace).toContain("Campos con diferencia: Costo Total");
        expect(totalTrace).not.toContain("Costo Unitario");
    });

    it("aunque solo difiera el Costo Unitario, muestra las DOS filas", async () => {
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
        expect(sheet.getRow(4).getCell(4).value).toBe("Continuidad de Costo Unitario");
        expect(String(sheet.getRow(4).getCell(10).value)).toContain("Campos con diferencia: Costo Unitario");
        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Total");
        expect(sheet.getRow(5).getCell(7).value).toBe(0);
        expect(String(sheet.getRow(5).getCell(10).value)).toContain("Sin diferencia en Costo Total");
    });

    it("registros viejos sin 'differences': si algún campo realmente difiere, muestra las dos filas", async () => {
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
        expect(sheet.getRow(4).getCell(4).value).toBe("Continuidad de Costo Unitario");
        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Total");
    });

    it("si ningún campo difiere no genera filas (las dos filas aparecen solo cuando hay un hallazgo)", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: baseMetadata({
                finalBalance: { quantity: 10, unitCost: 5, totalCost: 50 },
                initialBalance: { quantity: 10, unitCost: 5, totalCost: 50 },
                differences: []
            })
        }], header);

        expect(workbook.worksheets[0].rowCount).toBe(3);
    });

    it("sin unidades en ambos saldos el Costo Unitario no se compara, pero su fila se muestra con esa aclaración", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: baseMetadata({
                // El Kardex deja el último costo al cerrar en 0 y abre con costo 0
                finalBalance: { quantity: 0, unitCost: 90.7, totalCost: 5 },
                initialBalance: { quantity: 0, unitCost: 0, totalCost: 0 },
                differences: ["Costo Total"]
            })
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(4).getCell(4).value).toBe("Continuidad de Costo Unitario");
        const unitTrace = String(sheet.getRow(4).getCell(10).value);
        expect(unitTrace).toContain("sin unidades en ambos saldos");
        expect(unitTrace).not.toContain("Campos con diferencia");
        expect(String(sheet.getRow(5).getCell(10).value)).toContain("Campos con diferencia: Costo Total");
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

        expect(sheet.rowCount).toBe(5);

        const unitRowTrace = String(sheet.getRow(4).getCell(10).value);
        expect(unitRowTrace).toContain("Campos con diferencia: Costo Unitario");
        expect(unitRowTrace).not.toContain("Costo Total");

        const totalRowTrace = String(sheet.getRow(5).getCell(10).value);
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

        expect(sheet.rowCount).toBe(4); // 3 de header + 1 sola fila
        expect(sheet.getRow(4).getCell(5).value).toBe(250);
        expect(sheet.getRow(4).getCell(6).value).toBe("Producto no encontrado");
        expect(sheet.getRow(4).getCell(7).value).toBe("No aplicable");
        const trace = String(sheet.getRow(4).getCell(10).value);
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

        expect(sheet.getRow(4).getCell(4).value).toBe("Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(4).getCell(5).value).toBe(0);
        expect(sheet.getRow(4).getCell(6).value).toBe("Producto no encontrado");
        expect(sheet.getRow(4).getCell(7).value).toBe("No aplicable");
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

        expect(sheet.rowCount).toBe(4);
        expect(sheet.getRow(4).getCell(1).value).toBe("Marzo → Abril");
        expect(sheet.getRow(4).getCell(4).value).toBe("Sin Saldo Inicial en el mes siguiente");
        expect(sheet.getRow(4).getCell(5).value).toBe(50);
        expect(sheet.getRow(4).getCell(6).value).toBe("Sin Saldo Inicial (TipoOp 16)");
        expect(sheet.getRow(4).getCell(7).value).toBe("No aplicable");
        const trace = String(sheet.getRow(4).getCell(10).value);
        expect(trace).toContain("no tiene Saldo Inicial (TipoOp 16) en Abril");
        expect(trace).not.toContain("null");
    });

    it("PRODUCT_NOT_FOUND_NEXT_MONTH con missingFields: Cantidad solo si es distinta de 0, y SIEMPRE las dos filas de costo", async () => {
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

        expect(sheet.rowCount).toBe(6); // 3 de header + Cantidad + Costo Unitario + Costo Total
        expect(sheet.getRow(4).getCell(4).value).toBe("Cantidad - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(4).getCell(5).value).toBe(10);
        expect(sheet.getRow(4).getCell(6).value).toBe("Producto no encontrado");
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(0);
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Total - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(6).getCell(5).value).toBe(50);
    });

    it("caso real 46045: cierra con cantidad 0 y solo el Costo Unitario distinto de 0 -> muestra Costo Unitario y Costo Total (sin fila de Cantidad)", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            error_type: "PRODUCT_NOT_FOUND_NEXT_MONTH",
            product_code: "46045",
            product_name: "PANTALON FILIPPO ALPI PIMA COMFORT BOCELLI SILVER 36",
            risk_level: "ALTO",
            metadata: {
                fromIndex: 1,
                toIndex: 2,
                finalBalance: { quantity: 0, unitCost: 90.7, totalCost: 0 },
                initialBalance: null,
                missingFields: ["Costo Unitario"]
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(4).getCell(1).value).toBe("Enero → Febrero");
        expect(sheet.getRow(4).getCell(4).value).toBe("Costo Unitario - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(4).getCell(5).value).toBe(90.7);
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(0);

        // La fila que es hallazgo no lleva la aclaración; la que no lo es, la lleva al final.
        expect(String(sheet.getRow(4).getCell(10).value)).not.toContain("Sin diferencia");
        const totalTrace = String(sheet.getRow(5).getCell(10).value);
        expect(totalTrace).toContain("no se puede validar la continuidad");
        expect(totalTrace.split("\n").pop()).toBe("Sin diferencia en Costo Total");
    });

    it("INITIAL_BALANCE_NOT_FOUND_NEXT_MONTH con missingFields: también las dos filas de costo", async () => {
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
        expect(sheet.getRow(4).getCell(4).value).toBe("Costo Unitario - Sin Saldo Inicial en el mes siguiente");
        expect(sheet.getRow(4).getCell(5).value).toBe(142.06);
        expect(sheet.getRow(4).getCell(6).value).toBe("Sin Saldo Inicial (TipoOp 16)");
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total - Sin Saldo Inicial en el mes siguiente");
        expect(sheet.getRow(5).getCell(5).value).toBe(0);
    });

    it("missingFields solo con Cantidad: Cantidad + las dos filas de costo", async () => {
        const exporter = new Rule003Exporter();
        const workbook = await exporter.export([{
            error_type: "PRODUCT_NOT_FOUND_NEXT_MONTH",
            product_code: "000001",
            product_name: "PRODUCTO A",
            risk_level: "ALTO",
            metadata: {
                fromIndex: 1,
                toIndex: 2,
                finalBalance: { quantity: 3, unitCost: 0, totalCost: 0 },
                initialBalance: null,
                missingFields: ["Cantidad"]
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6);
        expect(sheet.getRow(4).getCell(4).value).toBe("Cantidad - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario - Producto no encontrado en el mes siguiente");
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Total - Producto no encontrado en el mes siguiente");
    });
});

