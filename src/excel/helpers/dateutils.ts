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

    static dottedDate(date: Date): string {
        const [year, month, day] = new Intl.DateTimeFormat("en-CA", {
            timeZone: "America/Lima",
            year: "numeric",
            month: "2-digit",
            day: "2-digit"
        }).format(date).split("-");
        return `${day}.${month}.${year}`;
    }

}