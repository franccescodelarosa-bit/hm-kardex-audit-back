
import { Rule004Exporter } from "./Rule004Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

/**
 * Columnas propias de RULE_004 (no usa las 10 genéricas de BaseExcelExporter
 * -- confirmado con la usuaria: Fecha Emisión, Fecha Almacén, RUC, Proveedor,
 * Documento y Documento Normalizado pasan a ser columnas visibles, no texto
 * adentro de Trazabilidad):
 *   1 Periodo | 2 Fecha Emisión | 3 Fecha Almacén | 4 RUC Proveedor |
 *   5 Proveedor | 6 Documento | 7 Documento Normalizado |
 *   8 Código del producto | 9 Descripción | 10 Tipo de inconsistencia |
 *   11 Valor esperado | 12 Valor encontrado | 13 Diferencia |
 *   14 % Diferencia | 15 Nivel de riesgo | 16 Trazabilidad
 */
describe("Rule004Exporter", () => {
    it("Fecha Emisión, Fecha Almacén, RUC, Proveedor, Documento y Documento Normalizado ahora son columnas propias", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-01-06",
                    issueDate: "2023-12-06",
                    warehouseDate: "2024-01-04",
                    supplierRuc: "20136836545",
                    supplier: "ARDILES SAC",
                    document: "Fac-F001-501064",
                    normalizedDocument: "F00100501064",
                    month: 1,
                    expectedCost: 850,
                    foundCost: 820,
                    evaluatedProducts: [
                        { code: "000123", description: "PEGAMENTO X", cost: 500 },
                        { code: "000456", description: "TORNILLO Y", cost: 320 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(2).value).toBe("2023-12-06");
        expect(sheet.getRow(5).getCell(3).value).toBe("2024-01-04");
        expect(sheet.getRow(5).getCell(4).value).toBe("20136836545");
        expect(sheet.getRow(5).getCell(5).value).toBe("ARDILES SAC");
        expect(sheet.getRow(5).getCell(6).value).toBe("Fac-F001-501064");
        expect(sheet.getRow(5).getCell(7).value).toBe("F00100501064");
    });

    it("codigo/descripcion (con comas), esperado/encontrado y tipo de inconsistencia siguen en su lugar, corridos por las columnas nuevas", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-01-06",
                    issueDate: "2023-12-06",
                    warehouseDate: "2024-01-04",
                    supplierRuc: "20136836545",
                    supplier: "ARDILES SAC",
                    document: "Fac-F001-501064",
                    normalizedDocument: "F00100501064",
                    month: 1,
                    expectedCost: 850,
                    foundCost: 820,
                    evaluatedProducts: [
                        { code: "000123", description: "PEGAMENTO X", cost: 500 },
                        { code: "000456", description: "TORNILLO Y", cost: 320 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(8).value).toBe("000123, 000456");
        expect(sheet.getRow(5).getCell(9).value).toBe("PEGAMENTO X, TORNILLO Y");
        expect(sheet.getRow(5).getCell(10).value).toBe("Mercadería en tránsito no registrada");
        expect(sheet.getRow(5).getCell(11).value).toBe(850); // esperado
        expect(sheet.getRow(5).getCell(12).value).toBe(820); // encontrado

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).not.toContain("undefined");
        expect(trace).toContain("Productos Encontrados: 000123 - PEGAMENTO X, 000456 - TORNILLO Y");
        // Ya no se repiten en la trazabilidad -- ahora son columnas propias
        expect(trace).not.toContain("RUC:");
        expect(trace).not.toContain("Proveedor:");
    });

    it("factura no registrada en el Kardex: 'Encontrado' dice 'Documento no encontrado', no un 0 (cambio 1)", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-03-10",
                    issueDate: "2024-02-01",
                    warehouseDate: "2024-03-05",
                    supplierRuc: "20999999999",
                    supplier: "PROVEEDOR SIN MATCH",
                    document: "Fac-F001-9999",
                    normalizedDocument: "F00100009999",
                    month: 3,
                    expectedCost: 500,
                    foundCost: 0,
                    evaluatedProducts: []
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(String(sheet.getRow(5).getCell(8).value)).toBe("");
        expect(String(sheet.getRow(5).getCell(9).value)).toBe("");
        expect(sheet.getRow(5).getCell(10).value).toBe("Mercadería en tránsito no registrada");
        expect(sheet.getRow(5).getCell(11).value).toBe(500);
        expect(sheet.getRow(5).getCell(12).value).toBe("Documento no encontrado");
        expect(sheet.getRow(5).getCell(13).value).toBe("No aplicable");
        expect(sheet.getRow(5).getCell(14).value).toBe("No aplicable");
    });

    it("registro viejo con noEvaluable=true se muestra como 'Registrada en otro mes' (decisión del equipo)", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-01-05",
                    issueDate: "2023-12-20",
                    warehouseDate: "2024-01-05",
                    supplierRuc: "20999999999",
                    supplier: "PROVEEDOR SIN DATOS",
                    document: "Fac-E001-9000",
                    normalizedDocument: "E00100009000",
                    month: 1,
                    expectedCost: 500,
                    foundCost: 0,
                    isIncident: false,
                    noEvaluable: true,
                    usedFallback: true,
                    thresholdPercent: 5,
                    evaluatedProducts: []
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(10).value).toBe("Registrada en otro mes");
        expect(sheet.getRow(5).getCell(12).value).toBe("Registrada en otro mes");
        expect(sheet.getRow(5).getCell(13).value).toBe("No aplicable");
        expect(sheet.getRow(5).getCell(14).value).toBe("No aplicable");
    });

    it("usedFallback=true (encontro por documento, no por codigo) -- la trazabilidad lo dice explicitamente", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-01-13",
                    issueDate: "2023-12-24",
                    warehouseDate: "2024-01-13",
                    supplierRuc: "10442286260",
                    supplier: "ESTELA VILCHEZ ELISA",
                    document: "Fac-E001-1007",
                    normalizedDocument: "E00100001007",
                    month: 1,
                    expectedCost: 650,
                    foundCost: 400,
                    isIncident: true,
                    usedFallback: true,
                    thresholdPercent: 5,
                    evaluatedProducts: [
                        { code: "000123", description: "PRODUCTO REAL", cost: 650 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(10).value).toBe("INCIDENCIA");

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).toContain("Fuente de búsqueda: Factura (todas sus líneas)");
    });

    it("usedFallback=false (encontro directo por codigo adquirido) -- la trazabilidad tambien lo dice", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-01-08",
                    issueDate: "2023-12-01",
                    warehouseDate: "2024-01-08",
                    supplierRuc: "20136836545",
                    supplier: "ARDILES SAC",
                    document: "Fac-F001-7000",
                    normalizedDocument: "F00100007000",
                    month: 1,
                    expectedCost: 200,
                    foundCost: 100,
                    isIncident: true,
                    usedFallback: false,
                    thresholdPercent: 5,
                    evaluatedProducts: [
                        { code: "000500", description: "PRODUCTO POR CODIGO", cost: 200 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).toContain("Fuente de búsqueda: Factura, acotada a los Códigos Adquiridos");
    });

    it("documento SI encontrado, con validacion de costo -- muestra INCIDENCIA/ACEPTADA, no la etiqueta de 'no registrada'", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "MEDIO",
                metadata: {
                    transitItem: "2024-01-06",
                    issueDate: "2023-12-06",
                    warehouseDate: "2024-01-04",
                    supplierRuc: "20136836545",
                    supplier: "ARDILES SAC",
                    document: "Fac-F001-501064",
                    normalizedDocument: "F00100501064",
                    month: 1,
                    expectedCost: 850,
                    foundCost: 500,
                    isIncident: true,
                    thresholdPercent: 5,
                    evaluatedProducts: [
                        { code: "000123", description: "PEGAMENTO X", cost: 500 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(10).value).toBe("INCIDENCIA");
        expect(sheet.getRow(5).getCell(10).value).not.toBe("Mercadería en tránsito no registrada");

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).toContain("Umbral permitido: 5");
    });

    it("(decisión del equipo) las facturas ACEPTADAS SÍ se muestran, con el tipo 'ACEPTADA'", async () => {
        const exporter = new Rule004Exporter();
        const workbook = await exporter.export([
            {
                error_type: "ACCEPTED",
                risk_level: "BAJO",
                metadata: {
                    document: "Fac-F001-1", normalizedDocument: "F0010001", month: 1,
                    expectedCost: 500, foundCost: 500, isIncident: false, usedFallback: true,
                    thresholdPercent: 5, evaluatedProducts: [{ code: "1", description: "A", cost: 500 }]
                }
            },
            {
                error_type: "TRANSIT_COST_MISMATCH",
                risk_level: "MEDIO",
                metadata: {
                    document: "Fac-F001-2", normalizedDocument: "F0010002", month: 1,
                    expectedCost: 500, foundCost: 100, isIncident: true, usedFallback: true,
                    thresholdPercent: 5, evaluatedProducts: [{ code: "2", description: "B", cost: 100 }]
                }
            }
        ], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6); // 4 de header + ACEPTADA + INCIDENCIA
        expect(sheet.getRow(5).getCell(6).value).toBe("Fac-F001-1");
        expect(sheet.getRow(5).getCell(10).value).toBe("ACEPTADA");
        expect(sheet.getRow(5).getCell(12).value).toBe(500);
        expect(String(sheet.getRow(5).getCell(16).value)).toContain("Resultado: ACEPTADA");
        expect(sheet.getRow(6).getCell(6).value).toBe("Fac-F001-2");
        expect(sheet.getRow(6).getCell(10).value).toBe("INCIDENCIA");
        expect(String(sheet.getRow(6).getCell(16).value)).toContain("Resultado: CONTINGENCIA");
    });

    it("TRANSIT_REGISTERED_OTHER_MONTH: 'Registrada en otro mes', encontrado dice en qué mes se registró", async () => {
        const exporter = new Rule004Exporter();
        const workbook = await exporter.export([{
            error_type: "TRANSIT_REGISTERED_OTHER_MONTH",
            risk_level: "MEDIO",
            metadata: {
                issueDate: "2024-02-20",
                warehouseDate: "2024-03-10",
                document: "Fac-F001-600",
                normalizedDocument: "F00100000600",
                month: 3,
                expectedCost: 1000,
                foundCost: 0,
                registeredMonths: [4],
                evaluatedProducts: []
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(1).value).toBe("Marzo");
        expect(sheet.getRow(5).getCell(10).value).toBe("Registrada en otro mes");
        expect(sheet.getRow(5).getCell(11).value).toBe(1000);
        expect(sheet.getRow(5).getCell(12).value).toBe("Registrada en Abril");
        expect(sheet.getRow(5).getCell(13).value).toBe("No aplicable");
        expect(sheet.getRow(5).getCell(14).value).toBe("No aplicable");

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).toContain("Mes de ingreso al almacén: Marzo");
        expect(trace).toContain("Mes registrado en Kardex: Abril");
    });

    it("el % de diferencia es el que calculó la regla: esperado 0 y encontrado 500 -> 100 %, no 0 % (cambio 3)", async () => {
        const exporter = new Rule004Exporter();
        const workbook = await exporter.export([{
            error_type: "TRANSIT_COST_MISMATCH",
            risk_level: "MEDIO",
            metadata: {
                document: "Fac-F001-1234", normalizedDocument: "F00100001234", month: 1,
                expectedCost: 0, foundCost: 500, difference: -500, differencePercent: 100,
                isIncident: true, usedFallback: true, thresholdPercent: 5,
                evaluatedProducts: [{ code: "1", description: "A", cost: 500 }]
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(10).value).toBe("INCIDENCIA");
        expect(sheet.getRow(5).getCell(14).value).toBe("100.00 %");
    });

    it("TRANSIT_INVALID_WAREHOUSE_DATE: 'Fecha de ingreso inválida', muestra el texto que vino en la celda (cambio 4)", async () => {
        const exporter = new Rule004Exporter();
        const workbook = await exporter.export([{
            error_type: "TRANSIT_INVALID_WAREHOUSE_DATE",
            risk_level: "MEDIO",
            metadata: {
                month: null,
                issueDate: "2024-02-20",
                warehouseDate: "15/03/2024",
                rawWarehouseDate: "15/03/2024",
                document: "Fac-F001-8888",
                normalizedDocument: "F00100008888",
                expectedCost: 400,
                foundCost: 0,
                evaluatedProducts: []
            }
        }], header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(1).value).toBe("Sin período");
        expect(sheet.getRow(5).getCell(10).value).toBe("Fecha de ingreso inválida");
        expect(sheet.getRow(5).getCell(11).value).toBe(400);
        expect(sheet.getRow(5).getCell(12).value).toBe("No evaluado");
        expect(sheet.getRow(5).getCell(13).value).toBe("No aplicable");
        expect(sheet.getRow(5).getCell(14).value).toBe("No aplicable");
        expect(String(sheet.getRow(5).getCell(16).value)).toContain("Fecha de ingreso recibida: 15/03/2024");
    });
});

