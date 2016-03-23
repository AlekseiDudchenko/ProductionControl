using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class CreditMaterialWindow : Window
    {
        private ExcelClass excel = new ExcelClass();
   
        public CreditMaterialWindow()
        {
            InitializeComponent();

            // получаем список материалов из файла
            excel.GetMaterials();

            // заполняем MaterialComboBox значениями
            for (int i = 3; i <= excel.NamberMaterials; i++)
            {
                // заполняем comboBox значениями
                MaterialComboBox.Items.Add(excel.MaterialsNames[i - 3]);
            }

            // получаем текущую дату
            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");

            // Связываем DataGrid и коллекцию
            // TODO перенести в XAML
            DataGrid.ItemsSource = creditMaterialCollection;
        }

        ObservableCollection<CreditMaterial> creditMaterialCollection = new ObservableCollection<CreditMaterial>(); 

        private void Button_Click_Add(object sender, RoutedEventArgs e)
        {
            // создаем экземпляр записи и заполняем его поля
            var newCreditMaterial = new CreditMaterial
            {
                DocumentNumber = DocNamberTexBox.Text,
                Data = DatePicker.Text,
                MaterialName = MaterialComboBox.Text,
                MaterialIndex = MaterialComboBox.SelectedIndex,
                Credit = Convert.ToInt32(CreditMaterialTextBox.Text),
                Edinici = excel.EdiniciIzmerenia[MaterialComboBox.SelectedIndex]
            };

            // добавляем в коллекцию и в DataGrid
            creditMaterialCollection.Add(newCreditMaterial);
            // Обновляем значения DataGrid
            DataGrid.Items.Refresh();

            // значение, которое выбрали в MaterialComboBox делаем пустой строкой чтобы исключить дольнейшее использование
            MaterialComboBox.Items[newCreditMaterial.MaterialIndex] = String.Empty;

            // сбрасываем значения контролов
            MaterialComboBox.SelectedIndex = -1;
            CreditMaterialTextBox.Text = String.Empty;
            // TODO управлять блокировкой кнопки через проверку DataGrid
            SaveButton.IsEnabled = true;
        }


        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Functions.SaveDataCredit(excel.Filename, creditMaterialCollection);

            // выводим сообщение и закрываем текущее окно
            MessageBox.Show("Данные успешно внесены в базу!");
            this.Close();
        }

#region Обработчики изменений в контролах и Кнопка закрыть

        private void DocNamberTexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(CreditMaterialTextBox, DocNamberTexBox, MaterialComboBox);
        }

        private void MaterialComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {            
            if (MaterialComboBox.SelectedIndex != -1)
                EdiniciIzmereniaLabel.Content = excel.EdiniciIzmerenia[MaterialComboBox.SelectedIndex];

            AddButton.IsEnabled = Functions.ProverkaDannih(CreditMaterialTextBox, DocNamberTexBox, MaterialComboBox);
        }

        private void CreditMaterialTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(CreditMaterialTextBox, DocNamberTexBox, MaterialComboBox);
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            Close();
        }

#endregion

    }
}
