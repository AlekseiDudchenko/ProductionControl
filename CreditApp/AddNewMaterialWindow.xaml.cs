using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для AddNewMaterialWindow.xaml
    /// </summary>
    public partial class AddNewMaterialWindow : Window
    {
        public AddNewMaterialWindow()
        {
            InitializeComponent();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //класс новый материал
            Material newMaterial = new Material();

            newMaterial.Name = NameTextMox.Text;
            newMaterial.Cod = CodTextBox.Text;
            newMaterial.Units = UnitsTextBox.Text;

            ExcelClass excel = new ExcelClass(); 

            //записать в эксель

            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);

            Worksheet matWorksheet = workbook.Sheets["Mat"];
            Range myRange = matWorksheet.UsedRange;

            //Получаем номер последней заполненной строки
            int lastrow = matWorksheet.UsedRange.Rows.Count;

            // дата 
            matWorksheet.Cells[lastrow + 1, 1] = newMaterial.Cod;
            // номер документа
            matWorksheet.Cells[lastrow+1, 2] = newMaterial.Name;
            //
            matWorksheet.Cells[lastrow + 1, 3] = newMaterial.Units;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            MessageBox.Show("Добавлен новый материал: " + newMaterial.Name + "\nКод: " + newMaterial.Cod + "\nЕдиницы измерения: " + newMaterial.Units);
            Close();
        }

        private void NameTextMox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(NameTextMox, UnitsTextBox, CodTextBox);
        }

        private void CodTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(NameTextMox, UnitsTextBox, CodTextBox);
        }

        private void UnitsTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(NameTextMox, UnitsTextBox, CodTextBox);
        }


    }
}
