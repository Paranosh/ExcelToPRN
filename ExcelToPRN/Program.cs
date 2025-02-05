using System;
using System.IO;
using ClosedXML.Excel;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Arrastra un archivo Excel sobre el ejecutable.");
            return;
        }

        string excelPath = args[0];

        if (!File.Exists(excelPath) || !excelPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("El archivo no es válido. Asegúrate de que sea un Excel (.xlsx).");
            return;
        }

        string directory = Path.GetDirectoryName(excelPath);
        string baseName = Path.GetFileNameWithoutExtension(excelPath);
        string outputFilePath = Path.Combine(directory, baseName + ".prn");
        string infoFilePath = Path.Combine(directory, baseName + "_info.txt");

        try
        {
            int[] columnWidths = ConvertExcelToPrn(excelPath, outputFilePath);
            GenerateInfoFile(excelPath, infoFilePath, columnWidths);
            Console.WriteLine($"Conversión exitosa. Archivos guardados en:\n{outputFilePath}\n{infoFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static int[] ConvertExcelToPrn(string excelPath, string outputFilePath)
    {
        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1); // Primera hoja
            var range = worksheet.RangeUsed();
            int rowCount = range.RowCount();
            int colCount = range.ColumnCount();

            // Determinar el ancho máximo de cada columna
            int[] columnWidths = new int[colCount];

            for (int col = 1; col <= colCount; col++)
            {
                int maxLength = 0;

                for (int row = 1; row <= rowCount; row++)
                {
                    string cellValue = worksheet.Cell(row, col).GetString().Trim();
                    maxLength = Math.Max(maxLength, cellValue.Length);
                }

                columnWidths[col - 1] = maxLength + 1; // Agregar un espacio extra para separación
            }

            using (var writer = new StreamWriter(outputFilePath))
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        string value = worksheet.Cell(row, col).GetString().Trim();
                        writer.Write(value.PadRight(columnWidths[col - 1])); // Rellenar con espacios
                    }
                    writer.WriteLine();
                }
            }

            return columnWidths; // Retornar los anchos de las columnas
        }
    }

    static void GenerateInfoFile(string excelPath, string infoFilePath, int[] columnWidths)
    {
        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1); // Primera hoja

            using (var writer = new StreamWriter(infoFilePath))
            {
                int colCount = columnWidths.Length;
                int position = 1; // Comienza en 1

                for (int col = 1; col <= colCount; col++)
                {
                    string value = worksheet.Cell(1, col).GetString().Trim(); // Primera fila (encabezado)
                    int length = columnWidths[col - 1]; // Longitud de la columna incluyendo espacios

                    writer.WriteLine($"{value}  Inicio: {position}  Longitud: {length}");

                    position += length; // Mover la posición para la siguiente columna
                }
            }
        }
    }
}
