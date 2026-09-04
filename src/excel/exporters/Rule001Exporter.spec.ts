import { Rule001Exporter } from "./Rule001Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule001Exporter", () => {
    it("encabezados propios: 'Valor esperado del inventario final' y 'Valor encontrado del Kardex'", async () => {
        const exporter = new Rule001Exporter();
        const workbook = await exporter.export([], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(4).getCell(5).value).toBe("Valor esperado del inventario final");
        expect(sheet.getRow(4).getCell(6).value).toBe("Valor encontrado del Kardex");
    });

    it("si SOLO la Cantidad difiere, muestra una sola fila (no las 3)", async () => {
        const exporter = new Rule001Exporter();
        const results = [
            {
                error_type: "INVENTORY_MISMATCH",
                product_code: "000123",
                product_name: "PRODUCTO A",
                risk_level: "CRITICO",
                description: "El inventario final no coincide con el Saldo Inicial del Kardex (Cantidad).",
                metadata: {
                    month: 1,
                    inventoryCode: "000123",
                    normalizedCode: "123",
                    inventoryStock: 100,
                    kardexStock: 90, // difiere
                    inventoryUnitCost: 10,
                    kardexUnitCost: 10, // igual
                    inventoryTotalCost: 1000,
                    kardexTotalCost: 1000,
                    kardexMovements: 5
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        // fila 5 = unica fila de datos
        expect(sheet.getRow(5).getCell(4).value).toBe("Cantidad");
        // no debe haber una fila 6 (Costo Unitario) ni 7 (Costo Total)
        expect(sheet.getRow(6).getCell(4).value).toBeFalsy();
        expect(sheet.getRow(7).getCell(4).value).toBeFalsy();
    });

    it("si difieren Cantidad Y Costo Total (pero no Costo Unitario), muestra 2 filas", async () => {
        const exporter = new Rule001Exporter();
        const results = [
            {
                error_type: "INVENTORY_MISMATCH",
                product_code: "000456",
                product_name: "PRODUCTO B",
                risk_level: "CRITICO",
                description: "...",
                metadata: {
                    month: 1,
                    inventoryStock: 100,
                    kardexStock: 80, // difiere
                    inventoryUnitCost: 10,
                    kardexUnitCost: 10, // igual
                    inventoryTotalCost: 1000,
                    kardexTotalCost: 850, // difiere
                    kardexMovements: 3
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const tipos = [
            sheet.getRow(5).getCell(4).value,
            sheet.getRow(6).getCell(4).value
        ];
        expect(tipos).toContain("Cantidad");
        expect(tipos).toContain("Costo Total");
        expect(tipos).not.toContain("Costo Unitario");
        expect(sheet.getRow(7).getCell(4).value).toBeFalsy();
    });

    it("(reproduce el caso real reportado): la Descripcion de CADA fila menciona SOLO su propio campo, no los 3 juntos", async () => {
        const exporter = new Rule001Exporter();
        const results = [
            {
                error_type: "INVENTORY_MISMATCH",
                product_code: "035-F",
                product_name: "BOLIGRAFO 50-1 COLOR AZUL/ROJO/NEGRO",
                risk_level: "CRITICO",
                description: "El inventario final no coincide con el Saldo Inicial del Kardex (Cantidad, Costo Unitario, Costo Total).",
                metadata: {
                    month: 1,
                    inventoryStock: 35,
                    kardexStock: 40,
                    inventoryUnitCost: 20,
                    kardexUnitCost: 21.2,
                    inventoryTotalCost: 700,
                    kardexTotalCost: 848,
                    kardexMovements: 22
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const traceCantidad = String(sheet.getRow(5).getCell(10).value);
        const traceCostoUnitario = String(sheet.getRow(6).getCell(10).value);
        const traceCostoTotal = String(sheet.getRow(7).getCell(10).value);

        expect(traceCantidad).toContain("Descripción: El inventario final no coincide con el Saldo Inicial del Kardex (Cantidad).");
        expect(traceCantidad).not.toContain("Costo Unitario");
        expect(traceCantidad).not.toContain("Costo Total");

        expect(traceCostoUnitario).toContain("Descripción: El inventario final no coincide con el Saldo Inicial del Kardex (Costo Unitario).");
        expect(traceCostoUnitario).not.toContain("Cantidad,");
        expect(traceCostoUnitario).not.toContain("Costo Total");

        expect(traceCostoTotal).toContain("Descripción: El inventario final no coincide con el Saldo Inicial del Kardex (Costo Total).");
        expect(traceCostoTotal).not.toContain("Cantidad,");
        expect(traceCostoTotal).not.toContain("Costo Unitario");
    });

    it("si los 3 campos difieren, sigue mostrando las 3 filas (comportamiento real de un mismatch total)", async () => {
        const exporter = new Rule001Exporter();
        const results = [
            {
                error_type: "INVENTORY_MISMATCH",
                product_code: "000789",
                product_name: "PRODUCTO C",
                risk_level: "CRITICO",
                description: "...",
                metadata: {
                    month: 1,
                    inventoryStock: 100,
                    kardexStock: 80,
                    inventoryUnitCost: 10,
                    kardexUnitCost: 12,
                    inventoryTotalCost: 1000,
                    kardexTotalCost: 960,
                    kardexMovements: 3
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const tipos = [
            sheet.getRow(5).getCell(4).value,
            sheet.getRow(6).getCell(4).value,
            sheet.getRow(7).getCell(4).value
        ];
        expect(tipos).toEqual(["Cantidad", "Costo Unitario", "Costo Total"]);
    });

    it("la trazabilidad incluye el punto de partida (cierre del ejercicio anterior), con el año calculado dinamico segun el header", async () => {
        const exporter = new Rule001Exporter();
        const results = [
            {
                error_type: "INVENTORY_MISMATCH",
                product_code: "000123",
                product_name: "PRODUCTO A",
                risk_level: "CRITICO",
                description: "...",
                metadata: {
                    month: 1,
                    inventoryStock: 100,
                    kardexStock: 90,
                    inventoryUnitCost: 10,
                    kardexUnitCost: 10,
                    inventoryTotalCost: 1000,
                    kardexTotalCost: 1000,
                    kardexMovements: 5
                }
            }
        ];

        // header.year = 2024 -> el punto de partida tiene que decir "dic 2023"
        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(trace).toContain(
            "Punto de partida: Inventario final del cierre del ejercicio anterior (dic 2023)"
        );
    });

    it("el punto de partida usa el año del header, no un año fijo -- si la auditoria es de 2025, dice dic 2024", async () => {
        const exporter = new Rule001Exporter();
        const header2025: ReportHeader = { ...header, year: 2025 };
        const results = [
            {
                error_type: "PRODUCT_NOT_FOUND",
                product_code: "000999",
                product_name: "PRODUCTO SIN KARDEX",
                risk_level: "CRITICO",
                description: "El producto no existe en el Kardex.",
                recommendation: "Verifique que el producto exista en ambos archivos.",
                metadata: { month: 0 }
            }
        ];

        const workbook = await exporter.export(results, header2025);
        const sheet = workbook.worksheets[0];
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(trace).toContain(
            "Punto de partida: Inventario final del cierre del ejercicio anterior (dic 2024)"
        );
    });

    it("PRODUCT_NOT_FOUND sigue mostrando su propia fila, sin cambios (no pasa por el filtro de campos)", async () => {
        const exporter = new Rule001Exporter();
        const results = [
            {
                error_type: "PRODUCT_NOT_FOUND",
                product_code: "000999",
                product_name: "PRODUCTO SIN KARDEX",
                risk_level: "CRITICO",
                description: "El producto no existe en el Kardex.",
                recommendation: "Verifique que el producto exista en ambos archivos.",
                metadata: { month: 0 }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(4).value).toBe("Producto no encontrado en Kardex");
    });
});
