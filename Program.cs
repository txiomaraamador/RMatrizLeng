using System;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // Establecer el contexto de la licencia para EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        string filePath = "C:/Users/xioma/OneDrive/Escritorio/Automatas I/(P7) Analizador Léxico programa primer entrega/MatrizLeng.xlsx";

        try
        {
            
            string[,] matriz = LeerMatrizDesdeExcel(filePath);

            while (true)
            {
                Console.Write("Ingrese la palabra que desea buscar (o escriba 'salir' para salir): ");
                string palabra = Console.ReadLine();

                if (palabra.ToLower() == "salir")
                {
                    break;
                }

                
                string resultado = BuscarPalabra(matriz, palabra);
                Console.WriteLine(resultado);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }

    }

    static string[,] LeerMatrizDesdeExcel(string filePath)
    {
        using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
        {
            if (package.Workbook.Worksheets.Count == 0)
            {
                throw new Exception("No hay hojas de cálculo en el archivo Excel.");
            }

            var worksheet = package.Workbook.Worksheets[0];

            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            if (rows == 0 || columns == 0)
            {
                throw new Exception("La hoja de cálculo está vacía.");
            }

            string[,] matriz = new string[rows, columns];

            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= columns; col++)
                {
                    matriz[row - 1, col - 1] = worksheet.Cells[row, col].Text;
                }
            }

            return matriz;
        }
    }


static string BuscarPalabra(string[,] matriz, string palabra)
    {
        string posActual = "q0";
        Console.WriteLine($"El estado inicial: {posActual}");

        for (int i = 0; i < palabra.Length; i++)
        {
            // Encuentra la columna de la letra actual
            int columna = Array.IndexOf(matriz.GetRow(0), palabra[i].ToString());

            // Busca la posición de la letra en la fila actual
            for (int fila = 1; fila < matriz.GetLength(0); fila++)
            {
                if (matriz[fila, 0] == posActual)
                {
                    Console.Write($"la letra {palabra[i]}, va del estado: {posActual},");
                    // Obtiene la nueva posición
                    posActual = matriz[fila, columna];
                    Console.WriteLine($" al estado: {posActual}");
                    break;
                }
                else if (fila == matriz.GetLength(0) - 1)
                {
                    // Si no se encuentra la letra en la posición actual, la palabra no existe
                    return "No existe la palabra\n";
                }
            }
        }

        // Verifica si la última posición es q1000*
        if (posActual == "q1000*")
        {
            return "La palabra existe en la matriz\n";
        }
        else
        {
            return "No existe la palabra\n";
        }
    }
}

static class Extensions
{
    public static T[] GetRow<T>(this T[,] array, int row)
    {
        int length = array.GetLength(1);
        T[] result = new T[length];
        for (int i = 0; i < length; i++)
        {
            result[i] = array[row, i];
        }
        return result;
    }
}
