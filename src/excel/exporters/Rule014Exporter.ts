import ExcelJS from "exceljs";

export interface Rule014Metadata {
    month: number;

    initialBalance: {
        quantity: number;
        totalCost: number;
    };

    totals: {
        entry: {
            quantity: number;
            totalCost: number;
        };

        exit: {
            quantity: number;
            totalCost: number;
        };
    };

    expectedFinalBalance: {
        quantity: number;
        totalCost: number;
    };

    costTolerance: {
        percentage: number;
        lowerLimit: number;
        upperLimit: number;
    };

    actualFinalBalance: {
        quantity: number;
        totalCost: number;
    };

    difference: {
        quantity: number;
        totalCost: number;
    };

    productCount: number;
    movementCount: number;

    differences: string[];
}

export class Rule014Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet =
            workbook.addWorksheet("RULE_014");

        /*
         * ============================================================
         * TÍTULO
         * ============================================================
         */

        worksheet.mergeCells("A1:R1");

        worksheet.getCell("A1").value =
            "RULE_014 - Validación Consolidada de Sumatorias Mensuales";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        /*
         * ============================================================
         * CABECERAS
         * ============================================================
         */

        worksheet.addRow([
            "Mes",                                         // A

            "Saldo Inicial - Unidades",                    // B
            "Saldo Inicial - Costo Valorizado",            // C

            "Suma Entradas - Unidades",                    // D
            "Suma Entradas - Costo Valorizado",            // E

            "Suma Salidas - Unidades",                     // F
            "Suma Salidas - Costo Valorizado",             // G

            "Saldo Final Esperado - Unidades",             // H
            "Costo Valorizado Calculado",                  // I

            "Límite Inferior Costo (97.5%)",               // J
            "Límite Superior Costo (102.5%)",              // K

            "Saldo Final Real - Unidades",                 // L
            "Saldo Final Real - Costo Valorizado",         // M

            "Cantidad de Productos",                       // N
            "Cantidad de Movimientos",                     // O

            "Diferencias",                                 // P
            "Descripción",                                 // Q
            "Recomendación"                                // R
        ]);

        /*
         * ============================================================
         * ANCHOS
         * ============================================================
         */

        worksheet.getColumn(1).width = 12;  // Mes

        worksheet.getColumn(2).width = 25;  // Inicial unidades
        worksheet.getColumn(3).width = 32;  // Inicial costo

        worksheet.getColumn(4).width = 25;  // Entradas unidades
        worksheet.getColumn(5).width = 32;  // Entradas costo

        worksheet.getColumn(6).width = 25;  // Salidas unidades
        worksheet.getColumn(7).width = 32;  // Salidas costo

        worksheet.getColumn(8).width = 32;  // Esperado unidades
        worksheet.getColumn(9).width = 30;  // Costo calculado

        worksheet.getColumn(10).width = 32; // Límite inferior
        worksheet.getColumn(11).width = 32; // Límite superior

        worksheet.getColumn(12).width = 28; // Real unidades
        worksheet.getColumn(13).width = 35; // Real costo

        worksheet.getColumn(14).width = 24; // Productos
        worksheet.getColumn(15).width = 26; // Movimientos

        worksheet.getColumn(16).width = 45; // Diferencias
        worksheet.getColumn(17).width = 70; // Descripción
        worksheet.getColumn(18).width = 70; // Recomendación

        /*
         * ============================================================
         * DATOS
         * ============================================================
         */

        for (const result of results) {

            const metadata =
                result.metadata as Rule014Metadata;

            worksheet.addRow([

                /*
                 * MES CONSOLIDADO
                 */
                metadata.month,

                /*
                 * SALDO INICIAL CONSOLIDADO
                 */
                metadata.initialBalance.quantity,
                metadata.initialBalance.totalCost,

                /*
                 * SUMA DE TODAS LAS ENTRADAS DEL MES
                 */
                metadata.totals.entry.quantity,
                metadata.totals.entry.totalCost,

                /*
                 * SUMA DE TODAS LAS SALIDAS DEL MES
                 */
                metadata.totals.exit.quantity,
                metadata.totals.exit.totalCost,

                /*
                 * RESULTADO CALCULADO
                 */
                metadata.expectedFinalBalance.quantity,
                metadata.expectedFinalBalance.totalCost,

                /*
                 * RANGO PERMITIDO DEL COSTO
                 */
                metadata.costTolerance.lowerLimit,
                metadata.costTolerance.upperLimit,

                /*
                 * SALDO FINAL REAL CONSOLIDADO
                 */
                metadata.actualFinalBalance.quantity,
                metadata.actualFinalBalance.totalCost,

                /*
                 * INFORMACIÓN DEL CONSOLIDADO
                 */
                metadata.productCount,
                metadata.movementCount,

                /*
                 * RESULTADO DE LA VALIDACIÓN
                 */
                metadata.differences.join(", "),

                result.description,
                result.recommendation
            ]);
        }

        /*
         * ============================================================
         * FORMATO NUMÉRICO
         * ============================================================
         *
         * B:M son cantidades y costos.
         */

        for (let column = 2; column <= 13; column++) {

            worksheet
                .getColumn(column)
                .numFmt = "#,##0.00";
        }

        /*
         * Productos y movimientos son enteros.
         */
        worksheet.getColumn(14).numFmt = "#,##0";
        worksheet.getColumn(15).numFmt = "#,##0";

        /*
         * ============================================================
         * CABECERA
         * ============================================================
         */

        const headerRow =
            worksheet.getRow(3);

        headerRow.font = {
            bold: true
        };

        headerRow.alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true
        };

        headerRow.height = 50;

        /*
         * ============================================================
         * AJUSTE DE TEXTO
         * ============================================================
         */

        worksheet.getColumn(16).alignment = {
            vertical: "top",
            wrapText: true
        };

        worksheet.getColumn(17).alignment = {
            vertical: "top",
            wrapText: true
        };

        worksheet.getColumn(18).alignment = {
            vertical: "top",
            wrapText: true
        };

        /*
         * ============================================================
         * CONGELAR CABECERA
         * ============================================================
         */

        worksheet.views = [
            {
                state: "frozen",
                ySplit: 3
            }
        ];

        /*
         * ============================================================
         * FILTRO
         * ============================================================
         */

        worksheet.autoFilter = {
            from: "A3",
            to: "R3"
        };

        return workbook;
    }
}