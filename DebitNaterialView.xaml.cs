using System;
using System.Windows;

   //using ExcelLibrary.SpreadSheet;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductionControl
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class DebitMaterialView : Window
    {
        public DebitMaterialView()
        {
            InitializeComponent();

            // открываем документ и лист для считывания данных для comboBox
            Excel.Application excelApp = new Excel.Application();                          
            Excel.Workbook workbook = excelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Material.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets["Mat"];

            // заполняем comboBox значениями
            for (int i = 3; i < 76; i++)
            {
                MaterialComboBox.Items.Add(worksheet.Cells[i, 2].Value);
            }


        }

    }
}


