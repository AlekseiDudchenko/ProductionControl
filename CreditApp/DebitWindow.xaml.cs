using System;
using System.IO;
using System.Net.Configuration;
using System.Reflection;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class DebitWindow : Window
    {
        private int lastrow = 0;
        ExcelClass excel = new ExcelClass();

        public DebitWindow()
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

            // получем номер последнего документа
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet) workbook.Sheets["Приход материалов"];
            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

            DocNamberTexBox.Text = worksheet.Cells[2, 2].Value;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // заполняем текущее время
            DateTexBox.Text = DateTime.Now.ToString("dd.MM.yyyy");


        }

        private void Button_Click_Add(object sender, RoutedEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook =
                excelApp.Workbooks.Open(excel.Filename);
            Worksheet debitWorksheet = (Worksheet) workbook.Sheets["Приход материалов"];
            Range debitRange = debitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
            Worksheet creditWorksheet = workbook.Sheets["Аналитика"];
            Range creditRange = creditWorksheet.UsedRange;

            int column = 0;

            //Получаем номер последней заполненной строки
            //Получаем номер последней заполненной строки
            int lastrow = debitWorksheet.UsedRange.Rows.Count;

            // проверям совпадение номера документа
            Range cellRange = (Range)debitWorksheet.Cells[lastrow - 1, 3];
            string cellValue = cellRange.Value.ToString();

           

            if (cellValue == DocNamberTexBox.Text)
            {
                debitWorksheet.Cells[lastrow - 1, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;
            }
            // добавляем новую строку ели номер не совпал
            else
            {
                for (column = 3; column <= debitRange.Columns.Count; column++)
                {
                    string letter = "";

                    {
                        char letter1 = Convert.ToChar(65 + column - 1);
                        letter += letter1;
                    }
                    if (26 + 1 <= column & column < 52 + 1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 26 - 1);
                        letter = "A" + letter1;
                    }
                    if (52 + 1 <= column & column < 78 + 1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 52 - 1);
                        letter = "B" + letter1;
                    }
                    if (78 + 1 <= column & column < 103 + 1)
                    {
                        char letter1 = Convert.ToChar(65 + column - 78 - 1);
                        letter = "C" + letter1;
                    }

                    // стираем старые формулы
                    debitWorksheet.Cells[lastrow, column] = "";
                    string formula = "=СУММ(" + letter + "3:" + letter + lastrow + ")";

                    // записываем формулу суммы по столбцам
                    debitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                    // меняем формулу в аналитике
                    string analitica = "='Приход материалов'!" + letter + lastrow + "-'Расход материалов'!" +
                                       letter + creditRange.Rows.Count;

                    Range analiticaRange = (Range)analiticaWorksheet.Cells[5, column];
                    analiticaRange.FormulaLocal = analitica; 
                }

                // заполняем заполняем ячейки файла данными
                debitWorksheet.Cells[lastrow, 1] = lastrow - 2;
                debitWorksheet.Cells[lastrow, 2] = DateTexBox.Text;
                debitWorksheet.Cells[lastrow, 3] = DocNamberTexBox.Text;
                debitWorksheet.Cells[lastrow, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;


                // меняем формулу в аналитике
                debitWorksheet = workbook.Sheets["Аналитика"];
                Range cellsRange = (Range)debitWorksheet.Cells[5, 2];
                cellsRange.Formula = "='Приход материалов'!D" + (lastrow + 1) + "-'Расход материалов'!D32";



                // закрываем Excel
                workbook.Close(true, Missing.Value, Missing.Value);
                excelApp.Quit();

                // выводим информационно сообщение
                MessageBox.Show("Добавлен приход материала " + MaterialComboBox.Text + " в размере " +
                                CreditMaterialTextBox.Text);

                MaterialComboBox.Text = "выберите материал"; // так не работает ))
                CreditMaterialTextBox.Text = String.Empty;
            }
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
