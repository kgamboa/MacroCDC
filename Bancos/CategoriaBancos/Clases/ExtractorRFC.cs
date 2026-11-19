using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
namespace TesoreriaCobranza
{
    public class RFCExtractor
    {
        public static List<string> ExtraerRFCs(string texto)
        {
            List<string> rfcList = new List<string>();

            if (string.IsNullOrEmpty(texto))
            {
                return rfcList;
            }

            // Expresión regular ajustada con delimitadores para evitar capturas parciales:
            // - (?<![A-Za-z0-9]) asegura que no haya letras o números antes del RFC
            // - [A-Za-z]{3,4}: Exactamente 3 o 4 letras al inicio
            // - \d{6}: Exactamente 6 dígitos
            // - [A-Za-z]{3}: Exactamente 3 letras al final
            // - (?! [A-Za-z0-9]) asegura que no haya letras o números inmediatamente después
            string pattern = @"(?<![A-Za-z0-9])[A-Za-z]{3,4}\d{6}[A-Za-z]{3}(?![A-Za-z0-9])";

            // Buscar todas las coincidencias en el texto
            MatchCollection matches = Regex.Matches(texto, pattern);

            // Procesar cada coincidencia para corregir errores comunes y filtrar
            foreach (Match match in matches)
            {
                if (match.Success && match.Value.Length >= 12 && match.Value.Length <= 13)
                {
                    string potentialRFC = match.Value.ToUpper();
                    string correctedRFC = CorregirErroresTipograficos(potentialRFC);

                    if (!string.IsNullOrEmpty(correctedRFC) && EsRFCValido(correctedRFC))
                    {
                        rfcList.Add(correctedRFC);
                    }
                }
            }

            return rfcList.Distinct().ToList(); // Eliminar duplicados
        }

        private static string CorregirErroresTipograficos(string rfc)
        {
            // Corrección de errores comunes: 0 (cero) por O (letra), 1 por I, 5 por S, etc.
            string corrected = rfc;

            // Reemplazar 0 por O, 1 por I, 5 por S solo en las primeras 3-4 posiciones (donde suelen estar las letras)
            if (corrected.Length >= 3)
            {
                string prefix = corrected.Substring(0, Math.Min(4, corrected.Length));
                string suffix = corrected.Substring(Math.Min(4, corrected.Length));

                // Reemplazar caracteres problemáticos en el prefijo
                prefix = Regex.Replace(prefix, @"0", "O");
                prefix = Regex.Replace(prefix, @"1", "I");
                prefix = Regex.Replace(prefix, @"5", "S");
                prefix = Regex.Replace(prefix, @"4", "A");
                prefix = Regex.Replace(prefix, @"8", "B");

                corrected = prefix + suffix;
            }

            return corrected;
        }

        private static bool EsRFCValido(string rfc)
        {
            // Validar que el RFC tiene el formato esperado: 3-4 letras, 6 dígitos, 3 letras
            if (rfc.Length < 12 || rfc.Length > 13)
                return false;

            // Verificar que las primeras 3-4 posiciones sean letras (después de corrección)
            string prefix = rfc.Substring(0, Math.Min(4, rfc.Length));
            if (!Regex.IsMatch(prefix, @"^[A-Za-z]+$"))
                return false;

            // Verificar que los siguientes 6 caracteres sean dígitos
            if (rfc.Length >= 10)
            {
                string digits = rfc.Substring(Math.Min(4, rfc.Length), 6);
                if (!Regex.IsMatch(digits, @"^\d{6}$"))
                    return false;
            }

            // Verificar que los últimos 3 caracteres sean letras
            if (rfc.Length >= 12)
            {
                string suffix = rfc.Substring(rfc.Length - 3);
                if (!Regex.IsMatch(suffix, @"^[A-Za-z]{3}$"))
                    return false;
            }

            return true;
        }

    }
}