using Xunit;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;

public class UnitTest1
{
    [Fact]
    public void Check_Subrecipients_In_File1()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Test");

        string path = Path.Combine(
            Directory.GetCurrentDirectory(),
            "..", "..", "..", "..",
            "SubawardApp", "Data", "SubawardBudgetExample1.xlsx"
        );

        using var package = new ExcelPackage(new FileInfo(path));
        var sheet = package.Workbook.Worksheets[0];

        var names = new List<string>();

        for (int row = 1; row <= sheet.Dimension.Rows; row++)
        {
            string columnB = sheet.Cells[row, 2].Text.Trim();
            string columnC = sheet.Cells[row, 3].Text.Trim();

            if (columnB.StartsWith("Subaward:", System.StringComparison.OrdinalIgnoreCase))
            {
                string name = columnB.Replace("Subaward:", "").Trim();

                if (string.IsNullOrWhiteSpace(name))
                    name = columnC;

                names.Add(name);
            }
        }

        Assert.Contains("Indiana", names);
        Assert.Contains("Mayo", names);
        Assert.Contains("Purdue", names);
        Assert.Contains("Florida", names);
        Assert.Equal(4, names.Count);
    }
}