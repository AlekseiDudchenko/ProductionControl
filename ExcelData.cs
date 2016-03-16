using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductionControl
{
    public class ExcelData
    {
        /// <summary>
        /// Имя файла для работы с Движение средств цеха
        /// </summary>
        public string FileName = "\\11.xlsx";
        
        public DataView MoveData
        {       
            get
            {
                Excel.Application excelApp = new Excel.Application();               

                // указываем файл и лист для работы                
                Excel.Workbook workbook = excelApp.Workbooks.Open(Environment.CurrentDirectory + FileName);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets["Test Sheet"]; 
                       
                Excel.Range range = worksheet.UsedRange;
                DataTable dataTable = new DataTable();

                dataTable.Columns.Add("№");
                dataTable.Columns.Add("Дата");
                dataTable.Columns.Add("Приход");
                dataTable.Columns.Add("Расход");
                dataTable.Columns.Add("Итого");

                int column = 0;
                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    DataRow dataRow = dataTable.NewRow();

                    dataRow[0] = (range.Cells[row, 1] as Excel.Range).Value2.ToString();
                    dataRow[1] = (range.Cells[row, 2] as Excel.Range).Value2.ToString();
                    dataRow[2] = (range.Cells[row, 3] as Excel.Range).Value2.ToString();
                    dataRow[3] = (range.Cells[row, 6] as Excel.Range).Value2.ToString();
                    dataRow[4] = (range.Cells[row, 9] as Excel.Range).Value2.ToString();

                    dataTable.Rows.Add(dataRow);
                    dataTable.AcceptChanges();
                    
                }

                workbook.Close(true, Missing.Value, Missing.Value);
                excelApp.Quit();

                return dataTable.DefaultView;
            }
        }

       
    }
}