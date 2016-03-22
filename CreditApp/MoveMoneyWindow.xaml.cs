using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для DebitMoneyWindow.xaml
    /// </summary>
    public partial class MoveMoneyWindow : Window
    {
        // для получения адреса файла
        ExcelClass excel = new ExcelClass();
 

        /// <summary>
        /// Переменная для подсчета количества записей в массиве = количество нажатий на кнопку Добавить
        /// </summary>
        private int _numberClickButton;

        public MoveMoneyWindow()
        {
            InitializeComponent();
            
            // блокируем кнопку Добавить при загрузке формы
            AddButton.IsEnabled = false;
   
            // заполняем текущее время
            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");

            // обнуляем счетчик количества нажатй на кнопку дабавть
            //_numberClickButton = 0;
        }

        ObservableCollection<DebitMoney> MoveMoneyCollection = new ObservableCollection<DebitMoney>();
       
        /// <summary>
        ///  Нажатие кнопки ДОБАВИТЬ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            DebitMoney newMoveMoney = new DebitMoney();

            // заполняем экземпляр ПРИХОД ДЕНЕГ
            newMoveMoney.Data = DatePicker.Text;
            newMoveMoney.DocumentNumber = DocNamberTexBox.Text;
            newMoveMoney.Statia = MoveMoneyComboBox.Text;
            newMoveMoney.StatialIndex = MoveMoneyComboBox.SelectedIndex;
            newMoveMoney.Debit = Convert.ToDouble(DebitMoneyTextBox.Text);
            newMoveMoney.TypeMove = MoveComboBox.Text;
            newMoveMoney.TypeMoveIndex = MoveComboBox.SelectedIndex;
            newMoveMoney.Osnovanie = OsnovanieTextBox.Text;

            // вносим экземпляр в коллекцию
            MoveMoneyCollection.Add(newMoveMoney);

            DataGrid.ItemsSource = MoveMoneyCollection;
            DataGrid.Items.Refresh();
            
            // считаем нажатие на кнопку
            //_numberClickButton += 1;
            
            // сбрасываем значения контролов
            MoveComboBox.SelectedIndex = -1;
            DocNamberTexBox.Text = "";
            DebitMoneyTextBox.Text = "";
            OsnovanieTextBox.Text = "";

            SaveButton.IsEnabled = true;

        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet debitMoneyWorksheet = (Worksheet) workbook.Sheets["Приход ДС"];
            Range debitMoneyRange = debitMoneyWorksheet.UsedRange;

            int lastRow = debitMoneyRange.Rows.Count;

            debitMoneyWorksheet.Cells[lastRow, 6] = "";
            debitMoneyWorksheet.Cells[lastRow, 7] = "";

            // записываем в таблицу
            for (int i = 0; i < MoveMoneyCollection.Count; i++)
            {
                debitMoneyWorksheet.Cells[lastRow + i , 1] = lastRow + i;
                debitMoneyWorksheet.Cells[lastRow + i , 2] = MoveMoneyCollection[i].Data;
                debitMoneyWorksheet.Cells[lastRow + i , 3] = MoveMoneyCollection[i].TypeMove;
                debitMoneyWorksheet.Cells[lastRow + i , 4] = MoveMoneyCollection[i].DocumentNumber;
                debitMoneyWorksheet.Cells[lastRow + i , 5] = MoveMoneyCollection[i].Statia;
                if (MoveMoneyCollection[i].TypeMoveIndex == 0)
                    debitMoneyWorksheet.Cells[lastRow + i , 6] = MoveMoneyCollection[i].Debit;
                if (MoveMoneyCollection[i].TypeMoveIndex == 1)
                    debitMoneyWorksheet.Cells[lastRow + i , 7] = MoveMoneyCollection[i].Debit;
                debitMoneyWorksheet.Cells[lastRow + i , 8] = MoveMoneyCollection[i].Osnovanie;                
            }

            debitMoneyWorksheet.Cells[lastRow + DataGrid.Items.Count, 6].FormulaLocal = ("=СУММ(F" + ((int) lastRow +
                                                                                                      (int)
                                                                                                          DataGrid.Items
                                                                                                              .Count - 1) +
                                                                                         ":F2)");
            debitMoneyWorksheet.Cells[lastRow + DataGrid.Items.Count, 7].FormulaLocal = ("=СУММ(G" +
                                                                                         ((int) lastRow +
                                                                                          (int) DataGrid.Items.Count - 1) +
                                                                                         ":G2)");

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            MessageBox.Show("Данные успешно внесены в базу!");
            this.Close();           
        }


        private List<string> DocumentNames = new List<string>();
        public string DocumentName;

        /// <summary>
        /// Выбор в ComboBox Тип движения
        /// Обеспечивает подгрузку коллекции в ComboBox Статья 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MoveComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {           
            // Очищаем текущую коллекцию
            MoveMoneyComboBox.Items.Clear();
            // добавляем значения в зависимости от выбранного значения
            switch (MoveComboBox.SelectedIndex)
            {
                // ПРИХОД
                case 0:
                {
                    //DocumentNameLabel.Content = DocumentName;
                    MoveMoneyComboBox.Items.Add("Аванс");
                    MoveMoneyComboBox.Items.Add("Полная оплата заказа");
                    MoveMoneyComboBox.Items.Add("Окончание оплаты заказа");
                    MoveMoneyComboBox.Items.Add("Другое");
                }break;
                //РАСХОД
                case 1: 
                {
                    MoveMoneyComboBox.Items.Add("Аванс");
                    MoveMoneyComboBox.Items.Add("Заработная плата");
                    MoveMoneyComboBox.Items.Add("Оплата услуг");
                    MoveMoneyComboBox.Items.Add("Другое");
                }break;
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

        private void ProverkaDannih()
        {
            // проверка на корректность данных
            if (DebitMoneyTextBox.Text != "" & MoveMoneyComboBox.SelectedIndex != -1 & DocNamberTexBox.Text != "")
            {
                bool convertgood = false;
                try
                {
                    Convert.ToInt32(DebitMoneyTextBox.Text);
                    convertgood = true;
                }
                catch (Exception)
                {

                    //throw;
                }
                if (convertgood)
                {
                    AddButton.IsEnabled = true;
                }
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

        private void DebitMoneyTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
           ProverkaDannih();
        }

        private void DocNamberTexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ProverkaDannih();
        }

        private void MoveMoneyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ProverkaDannih();
            if (MoveMoneyComboBox.SelectedValue != null)
            {
                DocumentName = "Документ №";
                if (MoveMoneyComboBox.SelectedValue.ToString() == "Аванс")
                    DocumentName = "Наряд №";
                if (MoveMoneyComboBox.SelectedValue.ToString() == "Оплата услуг")
                    DocumentName = "Чек №";
                if (MoveMoneyComboBox.SelectedValue.ToString() == "Полная оплата заказа")
                    DocumentName = "Чек №";
            }

            DocumentNameLabel.Content = DocumentName;
        }
    }
}
