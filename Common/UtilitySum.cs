using System.Collections.Generic;
using System.Linq;

public static class UtilitySum
{
    public static List<int> Search(decimal suma, int numero, List<IElement> lista)
    {
        // Ordenar la lista en orden descendente
        var sortedList = lista.OrderByDescending(x => x.Total).ToList();

        // Matriz de programación dinámica: dp[i, j] indica si una suma j es posible con i elementos
        var dp = new bool[sortedList.Count + 1, (int)suma + 1];
        dp[0, 0] = true;

        // Llenar la matriz de programación dinámica
        for (int i = 1; i <= sortedList.Count; i++)
        {
            for (int j = 0; j <= (int)suma; j++)
            {
                dp[i, j] = dp[i - 1, j];
                if (j >= sortedList[i - 1].Total)
                {
                    dp[i, j] |= dp[i - 1, j - (int)sortedList[i - 1].Total];
                }
            }
        }

        // Buscar combinación de tamaño válido
        for (int size = 1; size <= numero; size++)
        {
            if (dp[size, (int)suma])
            {
                var result = new List<int>();
                int currentSum = (int)suma;
                int currentSize = size;

                // Reconstruir la combinación
                for (int i = sortedList.Count - 1; i >= 0; i--)
                {
                    if (currentSum >= sortedList[i].Total && dp[currentSize - 1, currentSum - (int)sortedList[i].Total])
                    {
                        result.Add(sortedList[i].RowIndex);
                        currentSum -= (int)sortedList[i].Total;
                        currentSize--;
                    }
                }
                return result;
            }
        }

        return new List<int>();
    }

    public class IElement
    {
        public int RowIndex { get; set; }
        public decimal Total { get; set; }
    }
}