using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Rishith Rao Nooli");

        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        var totals = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in Directory.GetFiles(folderPath, "*.xlsx"))
        {
            Console.WriteLine($"\nFile: {Path.GetFileName(file)}");

            using var package = new ExcelPackage(new FileInfo(file));
            var worksheet = package.Workbook.Worksheets[0];

            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                string columnB = worksheet.Cells[row, 2].Text.Trim();
                string columnC = worksheet.Cells[row, 3].Text.Trim();

                if (columnB.StartsWith("Subaward:", StringComparison.OrdinalIgnoreCase))
                {
                    string name = columnB.Replace("Subaward:", "").Trim();

                    if (string.IsNullOrWhiteSpace(name))
                        name = columnC;

                    if (string.IsNullOrWhiteSpace(name))
                        name = "Unnamed Subrecipient";

                    decimal amount = 0;

                    for (int col = 5; col <= worksheet.Dimension.Columns; col++)
                    {
                        string value = worksheet.Cells[row, col].Text
                            .Replace("$", "")
                            .Replace(",", "")
                            .Trim();

                        if (decimal.TryParse(value, out decimal number))
                            amount += number;
                    }

                    Console.WriteLine($"  - {name}");

                    if (!totals.ContainsKey(name))
                        totals[name] = 0;

                    totals[name] += amount;
                }
            }
        }

        Console.WriteLine("\n==============================");
        Console.WriteLine("Subrecipient Totals");
        Console.WriteLine("==============================");

        foreach (var item in totals)
        {
            Console.WriteLine($"{item.Key,-25} ${item.Value:N2}");
        }
    }
}