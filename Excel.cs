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
            cells["C2"].PutValue(null);
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

        public static void CleanupRun()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Simulate headers (14 columns like your real file)
            for (int col = 0; col < 14; col++)
            {
                cells[0, col].PutValue("Col " + col);
            }

            // Row 1 - FULL data
            for (int col = 0; col < 14; col++)
            {
                cells[1, col].PutValue("R1C" + col);
            }

            // Row 2 - LAST COLUMN EMPTY (like your problematic case)
            for (int col = 0; col < 13; col++)
            {
                cells[2, col].PutValue("R2C" + col);
            }
            cells["C13"].PutValue(null);
            // cells[2,13] intentionally left empty

            // Row 3 - VALID DATA AFTER empty-last-column row
            for (int col = 0; col < 14; col++)
            {
                cells[3, col].PutValue("R3C" + col);
            }
            
            // Export like Bizagi does
            DataTable dt = cells.ExportDataTableAsString(0, 0, 10, 14, true);

            Console.WriteLine("=== BEFORE CLEANUP ===");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Console.WriteLine(
                    $"Row {i}: " + string.Join("|", dt.Rows[i].ItemArray)
                );
            }

            // Apply YOUR EXACT cleanup logic
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool isEmpty = true;

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                    {
                        isEmpty = false;
                        break;
                    }
                }

                if (isEmpty == true)
                {
                    dt.Rows.RemoveAt(i);
                    i--;
                }
            }

            Console.WriteLine("\n=== AFTER CLEANUP ===");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Console.WriteLine(
                    $"Row {i}: " + string.Join("|", dt.Rows[i].ItemArray)
                );
            }
        }
    }
}