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
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для AddNewProviderWindow.xaml
    /// </summary>
    public partial class AddNewProviderWindow : Window
    {
        ExcelClass excel = new ExcelClass();

        public AddNewProviderWindow()
        {
            InitializeComponent();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);

            Worksheet creditWorksheet = workbook.Sheets["Поставщики"];
            Range myRange = creditWorksheet.UsedRange;

            creditWorksheet.Cells[myRange.Rows.Count + 1, 1] = myRange.Rows.Count;
            creditWorksheet.Cells[myRange.Rows.Count + 1, 2] = NameTextBox.Text;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            Close();
            MessageBox.Show("Поставщик добавлен");
        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            AddButton.IsEnabled = NameTextBox.Text != String.Empty;
        }
    }
}
