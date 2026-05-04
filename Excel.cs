using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using Aspose.Cells;
using System.Xml;
using System.Collections;
using System.Configuration;

namespace Excel
{
    public class Test
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();

            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            cells["A1"].PutValue("Columna 1");
            cells["B1"].PutValue("Columna 2");
            cells["C1"].PutValue("Columna 3");
            cells["A2"].PutValue("12121212");
            cells["B2"].PutValue("");
            cells["C2"].PutValue("");
            cells["A3"].PutValue("04/05/2026");
            cells["B3"].PutValue(4);

            DataTable dataTable = cells.ExportDataTableAsString(0, 0, 10, 3, true);

            Console.WriteLine("Rows returned: " + dataTable.Rows.Count);

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                Console.WriteLine(
                    $"Row {i}: " + string.Join("|", dataTable.Rows[i].ItemArray)
                );
            }
        }
    }
}