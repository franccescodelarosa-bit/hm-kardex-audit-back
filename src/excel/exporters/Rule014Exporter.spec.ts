
import { Rule014Exporter } from "./Rule014Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

function baseMetadata(overrides: Partial<any> = {}) {
    return {
        month: 1,
        initialBalance: { quantity: 217963, totalCost: 945647.12 },
        totals: {
            entry: { quantity: 88017, totalCost: 310438.92 },
            exit: { quantity: 74947, totalCost: 274389.11 }
        },
        // Resultado de la FORMULA (Inicio + Entrada - Salida) -- el diagrama lo llama "valor encontrado"
        expectedFinalBalance: { quantity: 231033, totalCost: 981696.93 },
        costTolerance: { percentage: 0, lowerLimit: 981696.93, upperLimit: 981696.93 },
        // Inventario Valorizado de CIERRE real -- el diagrama lo llama "valor esperado"
        actualFinalBalance: { quantity: 231100, totalCost: 990339.13 },
        difference: { quantity: -67, totalCost: -8642.2 },
        productCount: 40985,
        movementCount: 52422,
        differences: ["Costo valorizado fuera del rango permitido"],
        ...overrides
    };
}

const GREEN = "FF00B050";
const BLUE = "FF0070C0";
const LIGHT_BLUE = "FF00B0F0";
const RED = "FFFF0000";

function traceText(value: any): string {
    return value.richText.map((t: any) => t.text).join("");
}

describe("Rule014Exporter", () => {
    it("'Valor esperado' muestra el Inventario de Cierre y 'Valor encontrado' muestra la fórmula, segun el diagrama oficial", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(4).getCell(5).value).toBe(990339.13); // Valor esperado = Cierre real
        expect(sheet.getRow(4).getCell(6).value).toBe(981696.93); // Valor encontrado = Formula
    });

    it("YA NO existe la fila de Cantidad -- RULE_014 solo valida costo (confirmado contra el diagrama y el Anexo 03)", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(4); // 3 de header + 1 sola fila (costo)
        expect(sheet.getRow(4).getCell(4).value).not.toBe("Sumatoria Consolidada - Cantidad");
    });

    it("la trazabilidad sigue el formato del cliente: Esperado -> definicion de Encontrado -> formula con montos -> Diferencia", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const tipo = sheet.getRow(4).getCell(4).value;
        const trace = traceText(sheet.getRow(4).getCell(10).value);

        expect(tipo).toBe("Costo valorizado fuera del rango permitido");

        expect(trace).toBe([
            "VALOR ESPERADO = SUMATORIA DE LOS SALDOS FINALES DE CADA ITEM DEL MES SEGUN EL KARDEX : S/ 990,339.13",
            "VALOR ENCONTRADO = SUMATORIA DE LOS SALDOS INICIALES DE CADA ITEM DEL MES + TOTAL ENTRADAS (Todos los ingresos) - TOTAL SALIDAS (todos los egresos)",
            "= RESULTADO DE LA FORMULA (VALOR ENCONTRADO):",
            "S/ 945,647.12 + S/ 310,438.92 - S/ 274,389.11",
            "= S/ 981,696.93",
            "",
            "Diferencia = S/ 990,339.13 - S/ 981,696.93 = S/ 8,642.20",
            "Tolerancia: Sin tolerancia (debe coincidir exacto)",
            "Productos consolidados: 40,985.00",
            "Movimientos consolidados: 52,422.00"
        ].join("\n"));
    });

    it("los montos de millones usan el apostrofo como separador (S/ 2´446,924.07), igual que la captura del cliente", async () => {
        const exporter = new Rule014Exporter();
        const results = [{
            risk_level: "CRITICO",
            metadata: baseMetadata({
                initialBalance: { quantity: 0, totalCost: 2475920.67 },
                totals: {
                    entry: { quantity: 0, totalCost: 94778.22 },
                    exit: { quantity: 0, totalCost: 125977.9 }
                },
                expectedFinalBalance: { quantity: 0, totalCost: 2444720.99 },
                actualFinalBalance: { quantity: 0, totalCost: 2446924.07 },
                productCount: 13808,
                movementCount: 15272
            })
        }];

        const workbook = await exporter.export(results, header);
        const trace = traceText(workbook.worksheets[0].getRow(4).getCell(10).value);

        expect(trace).toContain("SEGUN EL KARDEX : S/ 2´446,924.07");
        expect(trace).toContain("S/ 2´475,920.67 + S/ 94,778.22 - S/ 125,977.90");
        expect(trace).toContain("= S/ 2´444,720.99");
        expect(trace).toContain("Diferencia = S/ 2´446,924.07 - S/ 2´444,720.99 = S/ 2,203.08");
        expect(trace).toContain("Productos consolidados: 13,808.00");
        expect(trace).toContain("Movimientos consolidados: 15,272.00");
    });

    it("la trazabilidad usa colores: esperado en verde, encontrado en azul y operadores en rojo, con los totales subrayados", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const value = workbook.worksheets[0].getRow(4).getCell(10).value as any;
        const fragment = (text: string) => value.richText.find((t: any) => t.text === text);

        expect(fragment("S/ 990,339.13").font).toMatchObject({ color: { argb: GREEN }, underline: true });
        expect(fragment("S/ 981,696.93").font).toMatchObject({ color: { argb: BLUE }, underline: true });
        const fragments = (text: string) => value.richText.filter((t: any) => t.text === text);
        // En la linea de Diferencia van los montos (no las etiquetas), sin subrayar
        expect(fragments("S/ 990,339.13")[1].font).toMatchObject({ color: { argb: GREEN }, underline: false });
        expect(fragments("S/ 981,696.93")[1].font).toMatchObject({ color: { argb: LIGHT_BLUE }, underline: false });
        expect(fragment(" + ").font.color.argb).toBe(RED);
        expect(fragment(" - ").font.color.argb).toBe(RED);
    });

    it("la trazabilidad YA NO incluye el mensaje del Anexo 03 -- el cliente lo pidio asi (sigue en la descripcion del finding)", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const trace = traceText(workbook.worksheets[0].getRow(4).getCell(10).value);

        expect(trace).not.toContain("ERROR DE CONSOLIDACIÓN DEL KARDEX Y LA DIFERENCIA");
    });
});
