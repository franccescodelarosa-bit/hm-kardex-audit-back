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

    static monthName(month: number): string {
        return this.MONTHS[month - 1] ?? month.toString();
    }

}