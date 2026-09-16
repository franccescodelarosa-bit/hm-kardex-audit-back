export class DateUtils {

    private static readonly MONTHS = [
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre"
    ];

    static monthName(month: number | null | undefined): string {
        if (month === null || month === undefined) {
            return "Sin período";
        }
        return this.MONTHS[month - 1] ?? month.toString();
    }

}