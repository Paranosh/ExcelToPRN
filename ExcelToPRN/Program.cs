using System;
using System.IO;
using OfficeOpenXml;

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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // Primera hoja

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // Determinar el ancho máximo de cada columna
            int[] columnWidths = new int[colCount];

            for (int col = 1; col <= colCount; col++)
            {
                int maxLength = 0;

                for (int row = 1; row <= rowCount; row++)
                {
                    string cellValue = worksheet.Cells[row, col].Text;
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
                        string value = worksheet.Cells[row, col].Text;
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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // Primera hoja

            using (var writer = new StreamWriter(infoFilePath))
            {
                int colCount = columnWidths.Length;
                int position = 1; // Ahora comienza en 1

                for (int col = 1; col <= colCount; col++)
                {
                    string value = worksheet.Cells[1, col].Text; // Primera fila (encabezado)
                    int length = columnWidths[col - 1]; // Longitud de la columna incluyendo espacios

                    writer.WriteLine($"{value}  Inicio: {position}  Longitud: {length}");

                    position += length; // Mover la posición para la siguiente columna
                }
            }
        }
    }
}
