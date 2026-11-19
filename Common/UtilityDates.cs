using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public class UtilityDates
    {

        public static DateTime GetLastDayOfMonth(int Month, int Year)
        {
            // set return value to the last day of the month
            // for any date passed in to the method

            // create a datetime variable set to the passed in date
            DateTime dtTo = new DateTime(Year, Month, 1);

            // overshoot the date by a month
            dtTo = dtTo.AddMonths(1);

            // remove all of the days in the next month
            // to get bumped down to the last day of the
            // previous month
            dtTo = dtTo.AddDays(-(dtTo.Day));

            // return the last day of the month
            return dtTo;
        }
        public static string f_RecortaFecha(StringBuilder Cadena)
        {
            string functionReturnValue = null;
            string aCadena = Cadena.ToString();
            int lPosCaracter = 0;
            int lAcumulador = 0;

            lPosCaracter = aCadena.IndexOf('/', 0);
            lAcumulador += lPosCaracter;
            lPosCaracter = aCadena.IndexOf('/', lPosCaracter + 1);
            lAcumulador += lPosCaracter;

            if (lAcumulador > 0)
            {
                if (lAcumulador == 4)
                    functionReturnValue = aCadena.Substring(0, 8);
                else
                    functionReturnValue = aCadena.Substring(0, (aCadena.Length >= 10) ? 10 : aCadena.Length);
            }
            else
            {
                functionReturnValue = new DateTime(1989, 1, 1).ToString("MM/dd/yyyy");
            }
            return functionReturnValue;
        }

        public static DateTime f_RecortaFecha(string fecha)
        {
            try
            {
                fecha = f_RecortaFecha(new StringBuilder(fecha));
                string[] aux = fecha.Split('/');
                //DD/MM/YY
                return new DateTime(int.Parse(aux[2]), int.Parse(aux[1]), int.Parse(aux[0]));
            }
            catch
            {
                return new DateTime(1989, 1, 1);
            }
        }

        public static DateTime ObtieneFechaInicio(string periodo)
        {
            try
            {
                //Llega la informaicon como 2019\09-SEPTIEMBRE

                string[] split = periodo.Split('\\');

                int.TryParse(split[0], out int year);
                split = split[1].Split('-');

                var month = DateTime.ParseExact(split[1], "MMMM", CultureInfo.CurrentCulture).Month;


                return new DateTime(year, month, 1);
            }
            catch
            {
                //No tiene el Path el formato esperado.
                return new DateTime ( DateTime.Now.Year,DateTime.Now.Month,1);
            }

        }
        public static DateTime f_RecortaFechaMX(string fecha)
        {
            try
            {
                fecha = f_RecortaFecha(new StringBuilder(fecha));
                string[] aux = fecha.Split('/');
                //DD/MM/YY
                return new DateTime(int.Parse(aux[2]), int.Parse(aux[1]), int.Parse(aux[0]));
            }
            catch
            {
                return new DateTime();
            }
        }

        public static DateTime f_RecortaFecha2(string fecha)
        {
            try
            {
                fecha = f_RecortaFecha(new StringBuilder(fecha));
                string[] aux = fecha.Split('/');
                //MM/DD/YY
                return new DateTime(int.Parse(aux[2]), int.Parse(aux[0]), int.Parse(aux[1]));
            }
            catch
            {
                return new DateTime();
            }
        }

        public static DateTime GetLastDayOfMonth()
        {
            return GetLastDayOfMonth(DateTime.Now.Month, DateTime.Now.Year);
        }

        public static string ObtienePeriodo(DateTime fecha)
        {
            string yearString = fecha.Year.ToString();
            string fullMonthName = fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-MX"));
            return (fullMonthName + yearString).ToUpper();
        }
        public static string ObtienePeriodoCorto(DateTime fecha)
        {
            string yearString = fecha.ToString("yy");
            string fullMonthName = fecha.ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-MX"));
            return (yearString+fullMonthName.Replace(".","")  ).ToUpper();
        }
        public static string ObtieneFecha(int periodo)
        {
            var year = DateTime.Now.Month == 1 ? DateTime.Now.Year - 1 : DateTime.Now.Year;
            var dateTime = new DateTime(year, periodo + 1, periodo + 1);
            return dateTime.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es"));
        }
        public static DateTime FromExcelSerialDate(double SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }
    }
}
