
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

    it("si no se encontro ningun producto por los codigos adquiridos, la lista queda vacia sin romper nada", async () => {
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
        expect(sheet.getRow(5).getCell(12).value).toBe(0);
    });

    it("noEvaluable=true (ni por codigo ni por documento se encontro nada) -- muestra etiqueta propia, no INCIDENCIA", async () => {
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

        const tipo = String(sheet.getRow(5).getCell(10).value);
        expect(tipo).not.toBe("INCIDENCIA");
        expect(tipo).not.toBe("Mercadería en tránsito no registrada");
        expect(tipo).toBe("Sin datos suficientes para evaluar el costo");

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).toContain("Fuente de búsqueda: Documento (fallback -- no había Códigos Adquiridos)");
        expect(trace).toContain("SIN DATOS PARA EVALUAR");
    });

    it("usedFallback=true (encontro por documento, no por codigo) -- la trazabilidad lo dice explicitamente", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "BAJO",
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
                    foundCost: 650,
                    isIncident: false,
                    noEvaluable: false,
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

        expect(sheet.getRow(5).getCell(10).value).toBe("ACEPTADA");

        const trace = String(sheet.getRow(5).getCell(16).value);
        expect(trace).toContain("Fuente de búsqueda: Documento (fallback -- no había Códigos Adquiridos)");
    });

    it("usedFallback=false (encontro directo por codigo adquirido) -- la trazabilidad tambien lo dice", async () => {
        const exporter = new Rule004Exporter();
        const results = [
            {
                risk_level: "BAJO",
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
                    foundCost: 200,
                    isIncident: false,
                    noEvaluable: false,
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
        expect(trace).toContain("Fuente de búsqueda: Códigos Adquiridos");
        expect(trace).not.toContain("fallback");
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
});
