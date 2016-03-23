using System;
using System.Collections.ObjectModel;
using System.Windows;

using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для DebitMaterialWindow.xaml
    /// </summary>
    public partial class DebitMaterialWindow : Window
    {
        ExcelClass excel = new ExcelClass();

        ObservableCollection<DebitMaterial> debitMaterialsCollection = new ObservableCollection<DebitMaterial>();

        public DebitMaterialWindow()
        {
            InitializeComponent();

            // Получаем из файла данные о метериалах. Заполняем свойства экзмепляра класса excel
            excel.GetMaterials();

            // проход по всем строкам листа Mat и заполняем ComboBox значениями
            for (int i = 3; i <= excel.NamberMaterials; i++)
            {
                MaterialComboBox.Items.Add(excel.MaterialsNames[i-3]);
            }
                        
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
            // создаем экземпляр и заполняем его значениями
            DebitMaterial newDebitMaterial = new DebitMaterial()
            {
                DocumentNumber = DocNamberTexBox.Content.ToString(),
                Data = DatePicker.Text,
                Summ = Convert.ToInt32(BillSummLabel.Content),
                Material = MaterialComboBox.Text,
                MaterialIndex = MaterialComboBox.SelectedIndex,
                Debit = Convert.ToDouble(DebitMaterialTextBox.Text),
                Price = Convert.ToDouble(PriceTextBox.Text),
                Row = debitMaterialsCollection.Count + 1
            };
            newDebitMaterial.LocalSumm = newDebitMaterial.Debit*newDebitMaterial.Price;
            newDebitMaterial.Edinici = excel.EdiniciIzmerenia[newDebitMaterial.MaterialIndex];

            // добавляем в коллекцию и отображаем в MyDataGrid
            debitMaterialsCollection.Add(newDebitMaterial);
            MyDataGrid.ItemsSource = debitMaterialsCollection;
            MyDataGrid.Items.Refresh();

            // значение, которое выбрали в MaterialComboBox делаем пустой строкой чтобы исключить дольнейшее использование
            MaterialComboBox.Items[newDebitMaterial.MaterialIndex] = String.Empty;

            // подсчитываем введенную за все шаги сумму и показываем
            // TODO можно подсчитывать сумму через DataGrid
            SummTextBox.Content =
                Convert.ToString(Convert.ToInt32(SummTextBox.Content) +
                                 Convert.ToInt32(Functions.CutStringRub(LocalSumm)));

            // обнуляем значения элементов формы
            // индекс материала
            MaterialComboBox.SelectedIndex = -1;
            // локальную сумму поля для ввода количества и цены
            LocalSumm.Content = DebitMaterialTextBox.Text = PriceTextBox.Text = String.Empty;

            // подкрашиваем сумму
            Functions.ColorSumm(SummTextBox, BillSummLabel);

            // разблокируем кнопку сохранить
            // TODO можно оставить так пока запрещено удаление строк непосредственно через DataGrid
            SaveButton.IsEnabled = true;
        }


        /// <summary>
        /// Нажатие на кнопку Сохранить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // просим подтверждения если введенная сумма и сумма по документу не сопали
            if (Convert.ToInt32(SummTextBox.Content) != Convert.ToInt32(BillSummLabel.Content))
            {
                if (MessageBox.Show("Суммы не совпадают!\n Все равно продолжить?", "Все равно продолжить?",
                    MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    // сохраняем данные в файл //TODO Убедиться в однозначности связывания Collections и DataGrid
                    Functions.SaveDataDebit(excel.Filename, debitMaterialsCollection);

                    // выводим сообщение и закрываем текущее окно
                    MessageBox.Show("Данные успешно внесены в базу!");
                    this.Close();
                }
            }
            else
            {
                Functions.SaveDataDebit(excel.Filename, debitMaterialsCollection);
                // выводим сообщение и закрываем текущее окно
                MessageBox.Show("Данные успешно внесены в базу!");
                this.Close();
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


#region Обработчики изменений контролов
        /// <summary>
        /// Обработка изменеий в DataGrid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGrid_SelectedCellsChanged(object sender, System.Windows.Controls.SelectedCellsChangedEventArgs e)
        {
            //TODO работает корректно если есть хотя бы две строки в DataGread
            // подсчитываем и выводи новую сумму локальную и общую
            double newSumm = 0;
            foreach (DebitMaterial templDebitMaterial in debitMaterialsCollection)
            {
                templDebitMaterial.LocalSumm = templDebitMaterial.Price * templDebitMaterial.Debit;
                newSumm += templDebitMaterial.LocalSumm;
            }
            SummTextBox.Content = newSumm;

            // обновляем DataGrid
            MyDataGrid.Items.Refresh();

            // подкрашиваем сумму
            Functions.ColorSumm(SummTextBox, BillSummLabel);
        }
        
        /// <summary>
        /// Изменения в поле ввода цены материала
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PriceTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // проверяем на ввод цифр и добавляем "руб" Иначе пишем "Ошибка"
            try
            {
                Convert.ToDouble(DebitMaterialTextBox.Text);
                Convert.ToDouble(PriceTextBox.Text);
                LocalSumm.Content = Convert.ToString(Convert.ToInt32(DebitMaterialTextBox.Text) *
                                                     Convert.ToInt32(PriceTextBox.Text)) + " руб";
            }
            catch (Exception)
            {
                if (DebitMaterialTextBox.Text != "" & PriceTextBox.Text != "")
                    LocalSumm.Content = "Ошибка";
            }

            AddButton.IsEnabled = Functions.ProverkaDannih(DebitMaterialTextBox, PriceTextBox, MaterialComboBox);
        }

        /// <summary>
        /// Изменение в поле ввода количества материала
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CreditMaterialTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // проверяем на ввод цифр и добавляем "руб" Иначе пишем "Ошибка"
            try
            {
                Convert.ToDouble(DebitMaterialTextBox.Text);
                Convert.ToDouble(PriceTextBox.Text);
                LocalSumm.Content = Convert.ToString(Convert.ToInt32(DebitMaterialTextBox.Text) *
                                                     Convert.ToInt32(PriceTextBox.Text)) + " руб";
            }
            catch (Exception)
            {
                if (DebitMaterialTextBox.Text != "" & PriceTextBox.Text != "")
                    LocalSumm.Content = "Ошибка";
            }

            AddButton.IsEnabled = Functions.ProverkaDannih(DebitMaterialTextBox, PriceTextBox, MaterialComboBox);
        }

        /// <summary>
        /// Изменения в выборе материала из списка
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MaterialComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(DebitMaterialTextBox, PriceTextBox, MaterialComboBox);

            // показываем единицы измерения соответствующие выбранному материалу
            // если не равно -1 то присвоить значение, иначе пустую строку          
            EdiniciIzmereniaLabel.Content = MaterialComboBox.SelectedIndex != -1 ? excel.EdiniciIzmerenia[MaterialComboBox.SelectedIndex] : "";
        }
#endregion
  
    }
}
