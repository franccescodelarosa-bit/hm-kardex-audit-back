import ExcelJS from "exceljs";

export interface Rule008Metadata {
    source: string;
}

export class Rule008Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_008");

        worksheet.mergeCells("A1:E1");

        worksheet.getCell("A1").value =
            "RULE_008 - Productos No Encontrados en el Maestro";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Origen",
            "Código",
            "Producto",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 20;
        worksheet.getColumn(3).width = 45;
        worksheet.getColumn(4).width = 60;
        worksheet.getColumn(5).width = 60;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule008Metadata;

            worksheet.addRow([
                metadata.source,
                result.product_code,
                result.product_name,
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
            to: "E3"
        };

        return workbook;

    }

}