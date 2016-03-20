using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для CreditMoneyWindow.xaml
    /// </summary>
    public partial class CreditMoneyWindow : Window
    {
        ExcelClass excel = new ExcelClass();

        public CreditMoneyWindow()
        {
            InitializeComponent();
            
            // текущая дата
            DatePicker.Text = DateTime.Today.ToString();
        }



        /// <summary>
        /// Кнопка добавить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);

            Worksheet creditMoneyWorksheet = (Worksheet)workbook.Sheets["Расход ДС"];
            Range debitRange = creditMoneyWorksheet.UsedRange;

            // номер последней заполненой строки
            int lastRow = creditMoneyWorksheet.UsedRange.Rows.Count;

            creditMoneyWorksheet.Cells[lastRow + 1, 1] = lastRow;
            creditMoneyWorksheet.Cells[lastRow + 1, 2] = DatePicker.Text;
            creditMoneyWorksheet.Cells[lastRow + 1, 3] = DocNamberTexBox.Text;
            creditMoneyWorksheet.Cells[lastRow + 1, 4] = CreditComboBox.Text;
            creditMoneyWorksheet.Cells[lastRow + 1, 5] = DiskriptonsCreditMoneyTextBox.Text;
            creditMoneyWorksheet.Cells[lastRow + 1, 6] = CreditMoneyTextBox.Text;

            //(creditMoneyWorksheet.Cells[lastRow + 1, 6]) as Microsoft.Office.Interop.Excel.Range) ///.NumberFormat = "Денежный";


            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            MessageBox.Show("OK!");

        }


        /// <summary>
        /// Кнопака закрыть
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
