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
        private string filename = Environment.CurrentDirectory + "\\11.xlsx";
        //private string filename = "C:\\12.xlsx";
     
        /// <summary>
        /// Единичи измерения материалов
        /// </summary>
        public string[] Units = new string[200];

        /// <summary>
        /// Список наименований материалов
        /// </summary>
        public string[] MaterialsNames = new string[200];

        /// <summary>
        /// Количество материалов в файле
        /// </summary>
        public int NamberMaterials;

        public List<string> Providers = new List<string>(); 
 
        /// <summary>
        /// Содержит адрес и имя файла
        /// </summary>
        public string Filename
        {
            get { return filename; }
        }

        public List<string> GetProviders()
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(Filename);
            Worksheet providerWorksheet = (Worksheet)workbook.Sheets["Поставщики"];
            Range providerRange = providerWorksheet.UsedRange;

            for (int i = 2; i <= providerRange.Rows.Count; i++)
            {
                if (providerWorksheet.Cells[i, 2] != null)
                    Providers.Add(providerWorksheet.Cells[i, 2].Value.ToString());
            }
            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit(); 

            return Providers;
        }


        public void GetMaterials()
        {
            // открываем документ и лист для считывания данных для comboBox
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(Filename);
            Worksheet MatWorksheet = (Worksheet)workbook.Sheets["Mat"];
            Range MatRange = MatWorksheet.UsedRange;

            // проход по всем строкам листа Mat
            for (int i = 3; i <= MatRange.Rows.Count; i++)
            {
                // заполняем comboBox значениями
                MaterialsNames[i - 3] = MatWorksheet.Cells[i, 2].Value.ToString();
                // запоминаем единици измерения
                Units[i - 3] = MatWorksheet.Cells[i, 3].ToString();
            }
            

            NamberMaterials = MatRange.Rows.Count;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();       
        }
    }
}
