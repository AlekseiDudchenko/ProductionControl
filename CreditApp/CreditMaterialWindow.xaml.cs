
using System;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class CreditMaterialWindow : Window
    {
        private ExcelClass _excel = new ExcelClass();
        // массив для хранения индексов материалов
        private string[] _ediniciIzmerenia = new string[200];  
        

        public CreditMaterialWindow()
        {
            InitializeComponent();

            _excel.GetMaterials();
            // проход по всем строкам листа Mat
            for (int i = 3; i <= _excel.NamberMaterials; i++)
            {
                // заполняем comboBox значениями
                MaterialComboBox.Items.Add(_excel.MaterialsNames[i - 3]);
                // запоминаем единици измерения
                _ediniciIzmerenia[i - 3] = _excel.EdiniciIzmerenia[i - 3];
            }

            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");


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
                Edinici = _ediniciIzmerenia[MaterialComboBox.SelectedIndex]
            };

            // добавляем в коллекцию и в DataGrid
            creditMaterialCollection.Add(newCreditMaterial);
            DataGrid.ItemsSource = creditMaterialCollection;
            DataGrid.Items.Refresh();

            // обнуляем строки в ComboBox  которые уже использовали
            for (int i = 0; i < creditMaterialCollection.Count; i++)
            {
                for (int j = 0; j < MaterialComboBox.Items.Count; j++)
                {
                    if (MaterialComboBox.Items[j].ToString() == creditMaterialCollection[i].MaterialName)
                    {
                        MaterialComboBox.Items[j] = "";
                    }
                }
            }

            // сбрасываем значения контролов
            MaterialComboBox.SelectedIndex = -1;
            CreditMaterialTextBox.Text = "";
            SaveButton.IsEnabled = true;
        }


        private void SaveData()
        {
            // открываем excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(_excel.Filename);
            Worksheet creditMaterialWorksheet = (Worksheet)workbook.Sheets["Расход материалов"];
            Range creditMaterialRange = creditMaterialWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];

            // номер последней строки
            int lastRow = creditMaterialRange.Rows.Count;

            for (int column = 4; column <= creditMaterialRange.Columns.Count; column++)
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
                creditMaterialWorksheet.Cells[lastRow, column] = "";

                // формируем новую формулу
                string formula = "=СУММ(" + letter + "3:" + letter + lastRow + ")";

                // записываем новую формулу суммы по столбцам в соответствующие ячейки расхода материалов
                creditMaterialWorksheet.Cells[lastRow + 1, column].FormulaLocal = formula;


                // формируем формулу для Аналитики
                string analitica = "='Приход материалов'!" + letter + (lastRow + 1) + "-'Расход материалов'!" +
                                   letter + creditMaterialRange.Rows.Count;
                // записываем новую формулу в аналитику
                Range analiticaRange = (Range)analiticaWorksheet.Cells[5, column];
                analiticaRange.FormulaLocal = analitica;

                #endregion

                // записываем Порядковый номер, Дату, Номер документа
                creditMaterialWorksheet.Cells[lastRow, 1] = lastRow;
                creditMaterialWorksheet.Cells[lastRow, 2] = creditMaterialCollection[0].Data;
                creditMaterialWorksheet.Cells[lastRow, 3] = creditMaterialCollection[0].DocumentNumber;

                // Записываем данные о расходе по материалам
                for (int i = 0; i < DataGrid.Items.Count; i++)
                {
                    creditMaterialWorksheet.Cells[lastRow, creditMaterialCollection[i].MaterialIndex + 4] =
                        creditMaterialCollection[i].Credit;
                }
            }

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveData();
            // выводим сообщение и закрываем текущее окно
            MessageBox.Show("Данные успешно внесены в базу!");
            this.Close();
        }

        private void ProverkaDannih()
        {
            bool convertSuccess = false;

            try
            {
                Convert.ToDouble(CreditMaterialTextBox.Text);
                convertSuccess = true;
            }
            catch (Exception)
            {
                AddButton.IsEnabled = false;

            }
            if (MaterialComboBox.SelectedIndex != -1 & convertSuccess & DocNamberTexBox.Text != "")
            {
                if (Convert.ToDouble(CreditMaterialTextBox.Text) != 0 & MaterialComboBox.SelectedItem.ToString() != "")
                    AddButton.IsEnabled = true;
                else
                {
                    AddButton.IsEnabled = false;
                }
            }
            else
            {
                AddButton.IsEnabled = false;
            }
        }

        private void DocNamberTexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ProverkaDannih();
        }

        private void MaterialComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {            
            if (MaterialComboBox.SelectedIndex != -1)
                EdiniciIzmereniaLabel.Content = _ediniciIzmerenia[MaterialComboBox.SelectedIndex];

            ProverkaDannih();            
        }

        private void CreditMaterialTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ProverkaDannih();
        }



        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            Close();
        }



    }
}
