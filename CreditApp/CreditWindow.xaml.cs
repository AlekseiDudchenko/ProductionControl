using System;
using System.Collections.Generic;
using System.IO;
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
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class CreditWindow : Window
    {
        ExcelClass excel = new ExcelClass();

        public CreditWindow()
        {
            InitializeComponent();


            // открываем документ и лист для считывания данных для comboBox
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook =
                excelApp.Workbooks.Open(excel.Filename);
            Microsoft.Office.Interop.Excel.Worksheet worksheet =
                (Microsoft.Office.Interop.Excel.Worksheet) workbook.Sheets["Mat"];

            // заполняем comboBox значениями
            for (int i = 3; i < 76; i++)
            {
                MaterialComboBox.Items.Add(worksheet.Cells[i, 2].Value);
            }


            /*
            // получем номер последнего документа
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet) workbook.Sheets["Расход материалов"];
            Range range = worksheet.UsedRange;
            DocNamberTexBox.Text = worksheet.Cells[2, 2].Value;
            */


            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // заполняем текущее время
            //DateTexBox.Text = DateTime.Now.ToString("dd.MM.yyyy");


        }

        private void Button_Click_Add(object sender, RoutedEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet worksheet = (Worksheet)workbook.Sheets["Расход материалов"];
            Range range = worksheet.UsedRange;
            Worksheet debitWorksheet = (Worksheet) workbook.Sheets["Приход материалов"];
            Range debitRange = debitWorksheet.UsedRange;

            int column = 0;

            //Получаем номер последней заполненной строки
            int lastrow = worksheet.UsedRange.Rows.Count;

            // проверям совпадение номера документа
            Range cellRange = (Range)worksheet.Cells[lastrow - 1, 3];
            string cellValue = cellRange.Value.ToString();

            Worksheet worksheetAnalitica = workbook.Sheets["Аналитика"];


            if (cellValue == DocNamberTexBox.Text)
            {
                worksheet.Cells[lastrow - 1, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;
            }
            // добавляем новую строку ели номер не совпал
            else
            {
                for (column = 3; column <= range.Columns.Count; column++)
                {
                    string letter = "";

                    if (column < 26 + 1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 1);
                        letter += letter1;
                    }
                    if (26 + 1 <= column  & column < 52+1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 26 - 1);
                        letter = "A" + letter1;
                    }
                    if (52+1 <= column  & column < 78+1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 52 - 1);
                        letter = "B" + letter1;
                    }
                    if (78+1 <= column & column < 103+1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 78 - 1);
                        letter = "C" + letter1;
                    }

                    // стираем старые формулы
                    worksheet.Cells[lastrow, column] = "";
                    string formula = "=СУММ(" + letter + "3:" + letter + lastrow + ")";

                    // записываем формулу суммы по столбцам
                    worksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                    // меняем формулу в аналитике
                    string analitica = "='Приход материалов'!" + letter + debitRange.Rows.Count + "-'Расход материалов'!" +
                                       letter + lastrow;

                    Range analiticaRange = (Range)worksheetAnalitica.Cells[5, column];
                    analiticaRange.FormulaLocal = analitica; 
                }

                // заполняем заполняем ячейки файла данными
                worksheet.Cells[lastrow, 1] = lastrow - 2;
                worksheet.Cells[lastrow, 2] = DateTexBox.Text;
                worksheet.Cells[lastrow, 3] = DocNamberTexBox.Text;
                worksheet.Cells[lastrow, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;
            }

                      

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // выводим информационно сообщение
            MessageBox.Show("Добавлен расход материала " + MaterialComboBox.Text + " в размере " + CreditMaterialTextBox.Text);

            MaterialComboBox.Text = "выберите материал"; // так не работает ))
            CreditMaterialTextBox.Text = String.Empty;
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
