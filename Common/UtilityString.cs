using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;

namespace Common
{
    public static class StringExtensions
    {
        public static string RemoveAccents(this string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            // Normaliza la cadena a la forma D (FormD), que descompone
            // los caracteres acentuados en el carácter base y el signo diacrítico.
            string normalizedString = text.Normalize(NormalizationForm.FormD);
            StringBuilder stringBuilder = new StringBuilder();

            foreach (char c in normalizedString)
            {
                // Obtiene la categoría Unicode del carácter.
                // Si no es un NonSpacingMark (un signo diacrítico), lo añade al StringBuilder.
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            // Opcional: Normaliza la cadena resultante a la forma C (FormC)
            // para recomponer los caracteres que puedan haberse dividido
            // pero que no eran acentos (esto es menos común pero asegura
            // una representación canónica si es importante).
            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
        public static string UpperRemoveAccents(this string text)
        {
            return text.RemoveAccents().ToUpper();
        }
    }
    public class UtilityString
    {
        public static int getEditDistance(string X, string Y)
        {
            int m = X.Length;
            int n = Y.Length;

            int[][] T = new int[m + 1][];
            for (int i = 0; i < m + 1; ++i)
            {
                T[i] = new int[n + 1];
            }

            for (int i = 1; i <= m; i++)
            {
                T[i][0] = i;
            }
            for (int j = 1; j <= n; j++)
            {
                T[0][j] = j;
            }

            int cost;
            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    cost = X[i - 1] == Y[j - 1] ? 0 : 1;
                    T[i][j] = Math.Min(Math.Min(T[i - 1][j] + 1, T[i][j - 1] + 1),
                            T[i - 1][j - 1] + cost);
                }
            }

            return T[m][n];
        }

        public static double findSimilarity(string x, string y)
        {
            if (x == null || y == null)
            {
                throw new ArgumentException("Strings must not be null");
            }
            x = x.ToLower();
            y = y.ToLower();
            double maxLength = Math.Max(x.Length, y.Length);
            if (maxLength > 0)
            {
                // optionally ignore case if needed
                return (maxLength - getEditDistance(x, y)) / maxLength;
            }
            return 1.0;
        }


        public static void Main()
        {
            string s1 = "Techie Delight";
            string s2 = "Tech Delight";

            double similarity = findSimilarity(s1, s2);

            Console.WriteLine(similarity);
        }
    }
}




