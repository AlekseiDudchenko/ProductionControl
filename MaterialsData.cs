using System;
using System.Data;
using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;

namespace ProductionControl
{
    class MaterialsData
    {
        public DataView MaterialDataView
        {
            get
            {
                Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Excel.Range range;

                // указываем файл для работы и лист
                Excel.Workbook workbook = excelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Material.xlsx");
                Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Mat"]; 
                           
                range = worksheet.UsedRange;
                DataTable materialDataTable = new DataTable();

                materialDataTable.Columns.Add("Код");
                materialDataTable.Columns.Add("Наименование");

                int column = 0;

                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    DataRow dr = materialDataTable.NewRow();

                    dr[0] = (range.Cells[row, 1] as Excel.Range).Value2.ToString();
                    dr[1] = (range.Cells[row, 2] as Excel.Range).Value2.ToString();

                    materialDataTable.Rows.Add(dr);
                    materialDataTable.AcceptChanges();
                }

                workbook.Close(true, Missing.Value, Missing.Value);
                excelApp.Quit();

                return materialDataTable.DefaultView;
            }

        }
    }
}
