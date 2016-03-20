using System;
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
        //TODO что за переменная
        private int lastrow = 0; //???
        string[] ediniciIzmerenia = new string[200];  //

        // для получения адреса файла
        ExcelClass excel = new ExcelClass();

        public DebitWindow()
        {
            InitializeComponent();
            // устанавливаем значение поля суммы
            SummTextBox.Content = 0;
            AddButton.IsEnabled = false;
            EdiniciIzmereniaLabel.Content = "ед.";

            // открываем документ и лист для считывания данных для comboBox
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet MatWorksheet = (Worksheet) workbook.Sheets["Mat"];
            Range MatRange = MatWorksheet.UsedRange;
          
            // проход по всем строкам листа Mat
            for (int i = 3; i <= MatRange.Rows.Count; i++)
            {
                // заполняем comboBox значениями
                MaterialComboBox.Items.Add(MatWorksheet.Cells[i, 2].Value);
                // запоминаем единици измерения
                ediniciIzmerenia[i - 3] = MatWorksheet.Cells[i, 3].Value;
            }

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // заполняем текущее время
            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");
        }

        /// <summary>
        /// Кнопка добавить
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
            Worksheet priceDebitWorksheet = (Worksheet) workbook.Sheets["Цена прихода"];
            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;

            int column = 0;

            //Получаем номер последней заполненной строки
            int lastrow = debitWorksheet.UsedRange.Rows.Count;

            // проверям совпадение номера документа c последней записью в строке
            Range cellRange = (Range) debitWorksheet.Cells[lastrow - 1, 3];
            // если номер документа совпал с последней записью
            if (cellRange.Value.ToString() == DocNamberTexBox.Text)
            {
                // добавляем в существующую записьб еще материал
                debitWorksheet.Cells[lastrow - 1, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;
                priceDebitWorksheet.Cells[lastrow - 1, MaterialComboBox.SelectedIndex + 4] = PriceTextBox.Text;
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
                    priceDebitWorksheet.Cells[lastrow, column] = "";

                    // формируем новую формулу
                    string formula = "=СУММ(" + letter + "3:" + letter + lastrow + ")";

                    // записываем новую формулу суммы по столбцам в соответствующие ячейки 
                    // ...для Прихода материалов
                    debitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;
                    // ... для Цены прихода
                    priceDebitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                    // формируем формулу для Аналитики
                    string analitica = "='Приход материалов'!" + letter + (lastrow + 1) + "-'Расход материалов'!" +
                                       letter + creditRange.Rows.Count;
                    // записываем новую формулу в аналитику
                    Range analiticaRange = (Range) analiticaWorksheet.Cells[5, column];
                    analiticaRange.FormulaLocal = analitica;
                }


                // заполняем ячейки файла данными из формы о ПРИХОДЕ материала 
                debitWorksheet.Cells[lastrow, 1] = lastrow - 2;
                debitWorksheet.Cells[lastrow, 2] = DatePicker.Text;
                debitWorksheet.Cells[lastrow, 3] = DocNamberTexBox.Text;
                debitWorksheet.Cells[lastrow, MaterialComboBox.SelectedIndex + 4] = CreditMaterialTextBox.Text;

                // заполняем ячейки файла данными из формы о ЦЕНЕ материала               
                priceDebitWorksheet.Cells[lastrow, 1] = lastrow - 2;
                priceDebitWorksheet.Cells[lastrow, 2] = DatePicker.Text;
                priceDebitWorksheet.Cells[lastrow, 3] = DocNamberTexBox.Text;
                priceDebitWorksheet.Cells[lastrow, MaterialComboBox.SelectedIndex + 4] = PriceTextBox.Text;

            }
            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // выводим информационно сообщение
            MessageBox.Show("Добавлен приход материала " + MaterialComboBox.Text + " в размере " +
                            CreditMaterialTextBox.Text);


            //TODO сменить имена переменных на человеческие
            // отрезаем символы "руб" от LocalSumm.Content
            string localSumm = Convert.ToString(LocalSumm.Content);
            int position = localSumm.IndexOf("руб");
            string localsumm2 = "";

            for (int i = 0; i < position; i++)
            {
                localsumm2 += localSumm[i];
            }

            // обнуляем значения элементов формы
            // сбросили индекс выбранного материала
            MaterialComboBox.SelectedIndex = -1;
            // подсчитываем введенную за все шаги сумму и показываем
            SummTextBox.Content = Convert.ToString(Convert.ToInt32(SummTextBox.Content) + Convert.ToInt32(localsumm2));
            // сбрасываем локальную сумму
            LocalSumm.Content = "";
            // сбрасываем поля для ввода количества и цены
            CreditMaterialTextBox.Text = String.Empty;
            PriceTextBox.Text = String.Empty;
            AddButton.IsEnabled = false;
        }

        /// <summary>
        /// Изменения в поле ввода цены материала
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PriceTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            CreditMaterialTextBox_TextChanged(sender, e);
        }

        /// <summary>
        /// Изменение в поле ввода количества материала
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CreditMaterialTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            bool flag;

            // проверяем на ввод цифр
            try
            {
                Convert.ToInt32(CreditMaterialTextBox.Text);
                Convert.ToInt32(PriceTextBox.Text);
                flag = true;
            }
            catch (FormatException)
            {
                flag = false;
            }

            if (CreditMaterialTextBox.Text != "" & (PriceTextBox.Text != "") & flag)
            {
                if (Convert.ToInt32(CreditMaterialTextBox.Text) != 0 & 
                    Convert.ToInt32(PriceTextBox.Text) != 0 &
                    MaterialComboBox.SelectedIndex != -1)
                {
                    AddButton.IsEnabled = true;
                }
                else // нельзя записать количество 0 или по цене 0 или не выбрав материал
                {
                    AddButton.IsEnabled = false;
                }

                LocalSumm.Content = Convert.ToString(Convert.ToInt32(CreditMaterialTextBox.Text) *
                                     Convert.ToInt32(PriceTextBox.Text)) + "руб";
            }
            else  // нельзя записывать если не конвертируется количество и цена в числа
            {
                AddButton.IsEnabled = false;             
            }
        }


        private void MaterialComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            EdiniciIzmereniaLabel.Content = ediniciIzmerenia[MaterialComboBox.SelectedIndex + 3];

            bool flag;

            // проверяем на ввод цифр
            try
            {
                Convert.ToInt32(CreditMaterialTextBox.Text);
                Convert.ToInt32(PriceTextBox.Text);
                flag = true;
            }
            catch (FormatException)
            {
                flag = false;
            }

            if (CreditMaterialTextBox.Text != "" & (PriceTextBox.Text != "") & flag)
            {
                if (Convert.ToInt32(CreditMaterialTextBox.Text) != 0 &
                    Convert.ToInt32(PriceTextBox.Text) != 0 &
                    MaterialComboBox.SelectedIndex != -1)
                {
                    AddButton.IsEnabled = true;
                }
                else // нельзя записать количество 0 или по цене 0 или не выбрав материал
                {
                    AddButton.IsEnabled = false;
                }

                LocalSumm.Content = Convert.ToString(Convert.ToInt32(CreditMaterialTextBox.Text) *
                                     Convert.ToInt32(PriceTextBox.Text)) + "руб";
            }
            else  // нельзя записывать если не конвертируется количество и цена в числа
            {
                AddButton.IsEnabled = false;
            }
        }

        private void NewBillButton_OnClick(object sender, RoutedEventArgs e)
        {
            Window newBillWindow = new NewBillWindow();
            newBillWindow.Show();
            this.Close();

            //throw new NotImplementedException();
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
