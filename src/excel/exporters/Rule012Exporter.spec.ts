
import { Rule012Exporter } from "./Rule012Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule012Exporter", () => {
    it("usa el numero de documento (result.product_code) en 'Codigo del producto', no una fecha", async () => {
        const exporter = new Rule012Exporter();
        const results = [
            {
                product_code: "Fac-F001-503371",
                product_name: "",
                risk_level: "MEDIO",
                metadata: {
                    issueDate: "2023-12-13T00:00:00.000Z",
                    warehouseDate: "2024-01-04T00:00:00.000Z",
                    supplierRuc: "20136836545",
                    supplier: "ARDILES SAC",
                    document: "Fac-F001-503371",
                    normalizedDocument: "F00100503371",
                    expectedCost: 837.32,
                    kardexCost: 849.42,
                    difference: -12.1,
                    differencePercent: 1.45,
                    thresholdPercent: 5,
                    isIncident: false,
                    movements: 1,
                    month: 1,
                    transitItem: "2023-12-13"
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const row = sheet.getRow(5);

        expect(row.getCell(2).value).toBe("Fac-F001-503371"); // Codigo del producto -- NO "2023-12-13"
        expect(String(row.getCell(10).value)).not.toContain("undefined");
        expect(String(row.getCell(10).value)).toContain("Documento Normalizado: F00100503371");
    });
});
