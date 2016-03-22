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
    public partial class DebitMaterialWindow : Window
    {
        //TODO что за переменная
        //private int lastrow = 0; //???
        // массив для хранения индексов материалов
        string[] ediniciIzmerenia = new string[200];  //

        // для получения адреса файла
        ExcelClass excel = new ExcelClass();

        /// <summary>
        /// Переменная для подсчета количества записей в массиве = количество нажатий на кнопку Добавить
        /// </summary>
        private int NumberClickButton;

        public DebitMaterialWindow()
        {
            InitializeComponent();

            // Получаем из файла данные о метериалах. Заполняем свойства экзмепляра класса excel
            excel.GetMaterials();
           
            // проход по всем строкам листа Mat
            for (int i = 3; i <= excel.NamberMaterials; i++)
            {
                // заполняем comboBox значениями
                MaterialComboBox.Items.Add(excel.MaterialsNames[i-3]);
                // запоминаем единици измерения
                ediniciIzmerenia[i - 3] = excel.EdiniciIzmerenia[i-3];
            }
             
           
            // заполняем текущее время
            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");

            // обнуляем счетчик количества нажатй на кнопку дабавть
            // TODO нужно ли указывать значение при загрузке страници? Всегда ли при объявлении переменной будет присваиваться 0 по умолчанию?
            NumberClickButton = 0;

        }

        private void DataGrid_SelectedCellsChanged(object sender, System.Windows.Controls.SelectedCellsChangedEventArgs e)
        {
             
            // подсчитываем и выводи новую сумму локальную и общую
            double newSumm=0;
            for (int i = 0; i < Collection.Count; i++)
            {
                Collection[i].LocalSumm = Collection[i].Price * Collection[i].Debit;
                newSumm += Collection[i].LocalSumm;
            }
            SummTextBox.Content = newSumm;

            DataGrid.Items.Refresh();
            ColorSumm();

        }

        private void ColorSumm()
        {
            // подкращиваем введенную сумму
            Brush newBrush = Brushes.Yellow;

            if (SummTextBox.Content.ToString() == BillSummLabel.Content.ToString())
                newBrush = Brushes.LawnGreen;
            if (Convert.ToInt32(SummTextBox.Content) > Convert.ToInt32(BillSummLabel.Content))
                newBrush = Brushes.Red;

            SummTextBox.Background = newBrush;
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
            newDebitMaterial.DocumentNumber = DocNamberTexBox.Content.ToString();
            newDebitMaterial.Data = DatePicker.Text;
            newDebitMaterial.Summ = Convert.ToInt32(BillSummLabel.Content);
            newDebitMaterial.Material = MaterialComboBox.Text;
            newDebitMaterial.MaterialIndex = MaterialComboBox.SelectedIndex;
            newDebitMaterial.Debit = Convert.ToDouble(DebitMaterialTextBox.Text);
            newDebitMaterial.Price = Convert.ToDouble(PriceTextBox.Text);
            newDebitMaterial.LocalSumm = newDebitMaterial.Debit * newDebitMaterial.Price;
            newDebitMaterial.Edinici = ediniciIzmerenia[newDebitMaterial.MaterialIndex];
            newDebitMaterial.Row = Collection.Count + 1;

            ArrayDebitMaterial[NumberClickButton] = newDebitMaterial;

            // считаем очередное нажатие кнопки
            NumberClickButton += 1;

            // добавляем в коллекцию и отображаем в DataGrid
            Collection.Add(newDebitMaterial);
            DataGrid.ItemsSource = Collection;
            DataGrid.Items.Refresh();


            // обнуляем строки в ComboBox  которые уже использовали
            for (int i = 0; i < Collection.Count; i++)
            {
                for (int j = 0; j < MaterialComboBox.Items.Count; j++)
                {
                    if (MaterialComboBox.Items[j].ToString() == Collection[i].Material)
                    {
                        MaterialComboBox.Items[j] = "";
                    }
                }
            }


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

            // подкрашиваем сумму
            ColorSumm();

            // разблокируем кнопку сохранить
            SaveButton.IsEnabled = true;

        }

        /// <summary>
        /// сохранение Прихода материала в Excel
        /// </summary>
        private void SaveData()
        {
            // РАБОТА С ФАЙЛОМ
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet materialDebitWorksheet = (Worksheet)workbook.Sheets["Приход материалов"];
            Range debitRange = materialDebitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
            Worksheet priceDebitWorksheet = (Worksheet)workbook.Sheets["Цена прихода"];
            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;
            Worksheet costsWorksheet = workbook.Sheets["Стоимость прихода"];


            // TODO где используется?
            int column = 0;

            //Получаем номер последней заполненной строки TODO переместить
            int lastrow = materialDebitWorksheet.UsedRange.Rows.Count;

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
                materialDebitWorksheet.Cells[lastrow, column] = "";
                priceDebitWorksheet.Cells[lastrow, column] = "";
                costsWorksheet.Cells[lastrow, column] = "";

                // формируем новую формулу
                string formula = "=СУММ(" + letter + "3:" + letter + lastrow + ")";

                // записываем новую формулу суммы по столбцам в соответствующие ячейки 
                // ...для Прихода материалов
                materialDebitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;
                // ... для Цены прихода
                priceDebitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;
                // ... для Стоимости 
                costsWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                // формируем формулу для Аналитики
                string analitica = "='Приход материалов'!" + letter + (lastrow + 1) + "-'Расход материалов'!" +
                                   letter + creditRange.Rows.Count;
                // записываем новую формулу в аналитику
                Range analiticaRange = (Range)analiticaWorksheet.Cells[5, column];
                analiticaRange.FormulaLocal = analitica;
                
                #endregion
            }

            // заполняем ячейки файла данными из формы о ПРИХОДЕ материала, ЦЕНЕ и СТОИМОСТИ
            // порядковый номер
            materialDebitWorksheet.Cells[lastrow, 1] = lastrow - 2;
            priceDebitWorksheet.Cells[lastrow, 1] = lastrow - 2;
            costsWorksheet.Cells[lastrow, 1] = lastrow - 2;
            // дата 
            materialDebitWorksheet.Cells[lastrow, 2] = ArrayDebitMaterial[0].Data;
            priceDebitWorksheet.Cells[lastrow, 2] = ArrayDebitMaterial[0].Data;
            costsWorksheet.Cells[lastrow, 2] = ArrayDebitMaterial[0].Data;
            // номер документа
            materialDebitWorksheet.Cells[lastrow, 3] = ArrayDebitMaterial[0].DocumentNumber;
            priceDebitWorksheet.Cells[lastrow, 3] = ArrayDebitMaterial[0].DocumentNumber;
            costsWorksheet.Cells[lastrow, 3] = ArrayDebitMaterial[0].DocumentNumber;
            // количество материала цена и стоимость
            for (int i = 0; i < DataGrid.Items.Count; i++)
            {
                materialDebitWorksheet.Cells[lastrow, ArrayDebitMaterial[i].MaterialIndex + 4] = ArrayDebitMaterial[i].Debit;
                priceDebitWorksheet.Cells[lastrow, ArrayDebitMaterial[i].MaterialIndex + 4] = ArrayDebitMaterial[i].Price;
                costsWorksheet.Cells[lastrow, ArrayDebitMaterial[i].MaterialIndex + 4].FormulaLocal = ArrayDebitMaterial[i].Debit*ArrayDebitMaterial[i].Price;
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
            ProverkaDannih();
        }

        /// <summary>
        /// Изменение в поле ввода количества материала
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CreditMaterialTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            ProverkaDannih();
        }

        private void ProverkaDannih()
        {
            bool flag;

            // проверяем на ввод цифр
            try
            {
                Convert.ToDouble(DebitMaterialTextBox.Text);
                Convert.ToDouble(PriceTextBox.Text);
                flag = true;
            }
            catch (FormatException)
            {
                flag = false;
            }

            // Если не пустые строчи и конвертируется в цифры
            if (DebitMaterialTextBox.Text != "" & (PriceTextBox.Text != "") & flag & MaterialComboBox.SelectedIndex != -1)
            {
                if (Convert.ToInt32(DebitMaterialTextBox.Text) != 0 &
                    Convert.ToInt32(PriceTextBox.Text) != 0 &
                    MaterialComboBox.SelectedItem.ToString() != "")
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
            ProverkaDannih();
            // если не равно -1 то присвоить значение, иначе пустую строку
            EdiniciIzmereniaLabel.Content = MaterialComboBox.SelectedIndex != -1 ? ediniciIzmerenia[MaterialComboBox.SelectedIndex] : "";
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
