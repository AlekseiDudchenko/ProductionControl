using System;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
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
        // массив для хранения индексов материалов
        string[] ediniciIzmerenia = new string[200];  //

        // для получения адреса файла
        ExcelClass excel = new ExcelClass();

        /// <summary>
        /// Переменная для подсчета количества записей в массиве = количество нажатий на кнопку Добавить
        /// </summary>
        private int NumberClickButton;

        public DebitWindow()
        {
            InitializeComponent();
            // устанавливаем значение поля суммы
            SummTextBox.Content = 0;
            AddButton.IsEnabled = false;
            EdiniciIzmereniaLabel.Content = "ед.";
            AddButton.IsEnabled = false;

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

            // обнуляем счетчик количества нажатй на кнопку дабавть
            // TODO нужно ли указывать значение при загрузке страници? Всегда ли при объявлении переменной будет присваиваться 0 по умолчанию?
            NumberClickButton = 0;

        }

        //создаем массив экземпляров класса ПРИХОД МАТЕРИАЛА. Размер массива 25 по условию заказчика. Объясняется тем, что больше 25 записией в одном документе о приходе материала не бывает.
        DebitMaterial[] ArrayDebitMaterial = new DebitMaterial[25];
        ObservableCollection<DebitMaterial> Collection = new ObservableCollection<DebitMaterial>();

        /// <summary>
        /// Кнопка добавить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_Add(object sender, RoutedEventArgs e)
        {
            DebitMaterial newDebitMaterial = new DebitMaterial();

            // заполняем экземпляр ЗАПИСЬ ПРИХОДА МАТЕРИАЛА  
            newDebitMaterial.DocumentNumber = DocNamberTexBox.Text;
            newDebitMaterial.Data = DatePicker.Text;
            newDebitMaterial.Summ = Convert.ToInt32(BillSummLabel.Content);
            newDebitMaterial.Material = MaterialComboBox.Text;
            newDebitMaterial.MaterialIndex = MaterialComboBox.SelectedIndex;
            newDebitMaterial.Debit = Convert.ToDouble(DebitMaterialTextBox.Text);
            newDebitMaterial.Price = Convert.ToDouble(PriceTextBox.Text);
            newDebitMaterial.LocalSumm = newDebitMaterial.Debit * newDebitMaterial.Price;

            ArrayDebitMaterial[NumberClickButton] = newDebitMaterial;

            // считаем очередное нажатие кнопки
            NumberClickButton += 1;

            // добавляем в коллекцию и отображаем в DataGrid
            Collection.Add(newDebitMaterial);
            DataGrid.ItemsSource = Collection;
            DataGrid.Items.Refresh();

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
            DebitMaterialTextBox.Text = String.Empty;
            PriceTextBox.Text = String.Empty;
            AddButton.IsEnabled = false;

            // подкращиваем введенную сумму
            Brush newBrush = Brushes.Yellow;

            if (SummTextBox.Content.ToString() == BillSummLabel.Content.ToString())
                newBrush = Brushes.LawnGreen;
            if (Convert.ToInt32(SummTextBox.Content) > Convert.ToInt32(BillSummLabel.Content))
                newBrush = Brushes.Red;

            SummTextBox.Background = newBrush;

        }

        /// <summary>
        /// сохранение Прихода материала в Excel
        /// </summary>
        private void SaveData()
        {
            // РАБОТА С ФАЙЛОМ
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet debitWorksheet = (Worksheet)workbook.Sheets["Приход материалов"];
            Range debitRange = debitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
            Worksheet priceDebitWorksheet = (Worksheet)workbook.Sheets["Цена прихода"];
            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;

            // TODO где используется?
            int column = 0;

            //Получаем номер последней заполненной строки TODO переместить
            int lastrow = debitWorksheet.UsedRange.Rows.Count;

            for (column = 4; column <= debitRange.Columns.Count; column++)
            {
                #region Формирование и перезапись формул

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
                Range analiticaRange = (Range)analiticaWorksheet.Cells[5, column];
                analiticaRange.FormulaLocal = analitica;

                #endregion
            }

            // заполняем ячейки файла данными из формы о ПРИХОДЕ материала и ЦЕНЕ
            // порядковый номер
            debitWorksheet.Cells[lastrow, 1] = lastrow - 2;
            priceDebitWorksheet.Cells[lastrow, 1] = lastrow - 2;
            // дата 
            debitWorksheet.Cells[lastrow, 2] = ArrayDebitMaterial[0].Data;
            priceDebitWorksheet.Cells[lastrow, 2] = ArrayDebitMaterial[0].Data;
            // номер счёта-фактуры
            debitWorksheet.Cells[lastrow, 3] = ArrayDebitMaterial[0].DocumentNumber;
            priceDebitWorksheet.Cells[lastrow, 3] = ArrayDebitMaterial[0].DocumentNumber;
            // количество материала и цена
            for (int i = 0; i < NumberClickButton; i++)
            {
                debitWorksheet.Cells[lastrow, ArrayDebitMaterial[i].MaterialIndex + 4] = ArrayDebitMaterial[i].Debit;
                priceDebitWorksheet.Cells[lastrow, MaterialComboBox.SelectedIndex + 4] = ArrayDebitMaterial[i].Price;
            }

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // выводим сообщение и закрываем текущее окно
            MessageBox.Show("Данные успешно внесены в базу!");
            this.Close();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (Convert.ToInt32(SummTextBox.Content) != Convert.ToInt32(BillSummLabel.Content))
            {
                if (MessageBox.Show("Суммы не совпадают!\n Все равно продолжить?", "Все равно продолжить?",
                    MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    SaveData();  
            }
            else
            {
                SaveData();
            }
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
                Convert.ToInt32(DebitMaterialTextBox.Text);
                Convert.ToInt32(PriceTextBox.Text);
                flag = true;
            }
            catch (FormatException)
            {
                flag = false;
            }

            if (DebitMaterialTextBox.Text != "" & (PriceTextBox.Text != "") & flag)
            {
                if (Convert.ToInt32(DebitMaterialTextBox.Text) != 0 & 
                    Convert.ToInt32(PriceTextBox.Text) != 0 &
                    MaterialComboBox.SelectedIndex != -1)
                {
                    AddButton.IsEnabled = true;
                }
                else // нельзя записать количество 0 или по цене 0 или не выбрав материал
                {
                    AddButton.IsEnabled = false;
                }

                LocalSumm.Content = Convert.ToString(Convert.ToInt32(DebitMaterialTextBox.Text) *
                                     Convert.ToInt32(PriceTextBox.Text)) + "руб";
            }
            else  // нельзя записывать если не конвертируется количество и цена в числа
            {
                AddButton.IsEnabled = false;             
            }
        }

        /// <summary>
        /// Изменения в выборе материала из списка
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MaterialComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

            // если не равно -1 то присвоить значение, иначе пустую строку
            EdiniciIzmereniaLabel.Content = MaterialComboBox.SelectedIndex != -1 ? ediniciIzmerenia[MaterialComboBox.SelectedIndex] : "";

            bool flag;

            // проверяем на ввод цифр
            try
            {
                Convert.ToInt32(DebitMaterialTextBox.Text);
                Convert.ToInt32(PriceTextBox.Text);
                flag = true;
            }
            catch (FormatException)
            {
                flag = false;
            }

            // Если не пустые строчи и конвертируется в цифры
            if (DebitMaterialTextBox.Text != "" & (PriceTextBox.Text != "") & flag)
            {
                if (Convert.ToInt32(DebitMaterialTextBox.Text) != 0 &
                    Convert.ToInt32(PriceTextBox.Text) != 0 &
                    MaterialComboBox.SelectedIndex != -1)
                {
                    AddButton.IsEnabled = true;
                }
                else // нельзя записать количество 0 или по цене 0 или не выбрав материал
                {
                    AddButton.IsEnabled = false;
                }

                LocalSumm.Content = Convert.ToString(Convert.ToInt32(DebitMaterialTextBox.Text) *
                                     Convert.ToInt32(PriceTextBox.Text)) + "руб";
            }
            else  // нельзя записывать если не конвертируется количество и цена в числа
            {
                AddButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// Нажатие на кнопку Новая счет-фактура 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewBillButton_OnClick(object sender, RoutedEventArgs e)
        {
            Window newBillWindow = new NewBillWindow();
            newBillWindow.Show();
            this.Close();
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
