using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CreditApp
{
    class ExcelClass
    {
        // устанавливаем адрес файла
        private string filename = Environment.CurrentDirectory + "\\7.xlsx";
        //private string filename = "C:\\12.xlsx";
     
        /// <summary>
        /// Единичи измерения материалов
        /// </summary>
        public string[] EdiniciIzmerenia = new string[200];

        /// <summary>
        /// Список наименований материалов
        /// </summary>
        public string[] MaterialsNames = new string[200];

        /// <summary>
        /// Количество материалов в файле
        /// </summary>
        public int NamberMaterials;
 
        /// <summary>
        /// Содержит адрес и имя файла
        /// </summary>
        public string Filename
        {
            get { return filename; }
        }


        public void GetMaterials()
        {
            // открываем документ и лист для считывания данных для comboBox
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(Filename);
            Worksheet MatWorksheet = (Worksheet)workbook.Sheets["Mat"];
            Range MatRange = MatWorksheet.UsedRange;

            // проход по всем строкам листа Mat
            for (int i = 3; i <= MatRange.Rows.Count; i++)
            {
                // заполняем comboBox значениями
                MaterialsNames[i - 3] = MatWorksheet.Cells[i, 2].Value;
                // запоминаем единици измерения
                EdiniciIzmerenia[i - 3] = MatWorksheet.Cells[i, 3].Value;
            }

            NamberMaterials = MatRange.Rows.Count;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();       
        }
    }
}
