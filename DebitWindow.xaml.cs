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
        private int lastrow = 0; //???
        ExcelClass excel = new ExcelClass();

        public DebitWindow()
        {
            InitializeComponent();

            // открываем документ и лист для считывания данных для comboBox
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet worksheet = (Worksheet) workbook.Sheets["Mat"];

            // заполняем comboBox значениями
            for (int i = 3; i < 76; i++)
            {
                MaterialComboBox.Items.Add(worksheet.Cells[i, 2].Value);
            }

            /*
            // получем номер последнего документа в excele
            worksheet = (Worksheet) workbook.Sheets["Приход материалов"];
            Range range = worksheet.UsedRange;
            DocNamberTexBox.Text = worksheet.Cells[2, 2].Value;
            */

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_Add(object sender, RoutedEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet debitWorksheet = (Worksheet) workbook.Sheets["Приход материалов"];
            Range debitRange = debitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];

            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;
          
            int column = 0;

            //Получаем номер последней заполненной строки
            int lastrow = debitWorksheet.UsedRange.Rows.Count;

            // проверям совпадение номера документа
            Range cellRange = (Range)debitWorksheet.Cells[lastrow - 1, 3];
            // если номер документа совпал с последней записью
            if (cellRange.Value.ToString() == DocNamberTexBox.Text)
            {
                // добавляем в существующую записьб еще материал
                debitWorksheet.Cells[lastrow - 1, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;
            }
            // если номер последнего документа в excele не совпал с введенным в форме
            else
            {
                for (column = 4; column <= debitRange.Columns.Count; column++)
                {
                    // формируем адреса ячеек для формулы
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
                    // формируем новую формулу
                    string formula = "=СУММ(" + letter + "3:" + letter + lastrow + ")";
                    // записываем новую формулу суммы по столбцам в соответствующие ячейки
                    debitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                    // формируем формулу для Аналитики
                    string analitica = "='Приход материалов'!" + letter + (lastrow + 1) + "-'Расход материалов'!" +
                                       letter + creditRange.Rows.Count;
                    // записываем новую формулу в аналитику
                    Range analiticaRange = (Range)analiticaWorksheet.Cells[5, column];
                    analiticaRange.FormulaLocal = analitica; 
                }

                // заполняем заполняем ячейки файла данными
                debitWorksheet.Cells[lastrow, 1] = lastrow - 2;
                debitWorksheet.Cells[lastrow, 2] = DatePicker.Text;
                debitWorksheet.Cells[lastrow, 3] = DocNamberTexBox.Text;
                debitWorksheet.Cells[lastrow, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;

                /*
                // меняем формулу в аналитике
                debitWorksheet = workbook.Sheets["Аналитика"];
                Range cellsRange = (Range)debitWorksheet.Cells[5, 2];
                cellsRange.Formula = "='Приход материалов'!D" + (lastrow + 1) + "-'Расход материалов'!D32";
                */


                // закрываем Excel
                workbook.Close(true, Missing.Value, Missing.Value);
                excelApp.Quit();

                // выводим информационно сообщение
                MessageBox.Show("Добавлен приход материала " + MaterialComboBox.Text + " в размере " +
                                CreditMaterialTextBox.Text);

                // обнуляем значения элементов формы
                MaterialComboBox.SelectedIndex = -1;
                CreditMaterialTextBox.Text = String.Empty;
            }
        }


        /// <summary>
        /// Кнопка закрыть
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
