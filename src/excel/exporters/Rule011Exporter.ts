import ExcelJS from "exceljs";

export interface Rule011Metadata {
    date: string;
    month: number;
    document: string;
    operation: string;
    previousCost: number;
    currentCost: number;
    variationPercent: number;
}

export class Rule011Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_011");

        worksheet.mergeCells("A1:K1");

        worksheet.getCell("A1").value =
            "RULE_011 - Variación Inusual del Costo Unitario";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Código",
            "Producto",
            "Mes",
            "Fecha",
            "Documento",
            "Operación",
            "Costo Unitario Anterior",
            "Costo Unitario Actual",
            "Variación (%)",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 40;
        worksheet.getColumn(3).width = 10;
        worksheet.getColumn(4).width = 15;
        worksheet.getColumn(5).width = 20;
        worksheet.getColumn(6).width = 25;
        worksheet.getColumn(7).width = 22;
        worksheet.getColumn(8).width = 22;
        worksheet.getColumn(9).width = 18;
        worksheet.getColumn(10).width = 55;
        worksheet.getColumn(11).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule011Metadata;

            worksheet.addRow([
                result.product_code,
                result.product_name,
                metadata.month,
                metadata.date,
                metadata.document,
                metadata.operation,
                metadata.previousCost,
                metadata.currentCost,
                metadata.variationPercent,
                result.description,
                result.recommendation
            ]);

        }

        worksheet.views = [
            {
                state: "frozen",
                ySplit: 3
            }
        ];

        worksheet.autoFilter = {
            from: "A3",
            to: "K3"
        };

        return workbook;

    }

}