using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Логика взаимодействия для DebitMoneyWindow.xaml
    /// </summary>
    public partial class DebitMoneyWindow : Window
    {
        // для получения адреса файла
        ExcelClass excel = new ExcelClass();
 

        /// <summary>
        /// Переменная для подсчета количества записей в массиве = количество нажатий на кнопку Добавить
        /// </summary>
        private int NumberClickButton;

        public DebitMoneyWindow()
        {
            InitializeComponent();
            
            // блокируем кнопку Добавить при загрузке формы
            AddButton.IsEnabled = false;
   
            // заполняем текущее время
            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");

            // обнуляем счетчик количества нажатй на кнопку дабавть
            NumberClickButton = 0;
        }

        ObservableCollection<DebitMoney> DebitMoneyCollection = new ObservableCollection<DebitMoney>();
       
        /// <summary>
        ///  Нажатие кнопки ДОБАВИТЬ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            DebitMoney newDebitMoney = new DebitMoney();

            // заполняем экземпляр ПРИХОД ДЕНЕГ
            newDebitMoney.Data = DatePicker.Text;
            newDebitMoney.DocumentNumber = DocNamberTexBox.Text;
            newDebitMoney.Statia = DebitMoneyComboBox.Text;
            newDebitMoney.StatialIndex = DebitMoneyComboBox.SelectedIndex;
            newDebitMoney.Debit = Convert.ToDouble(DebitMoneyTextBox.Text);
            newDebitMoney.TypeMove = MoveComboBox.Text;
            newDebitMoney.TypeMoveIndex = MoveComboBox.SelectedIndex;
            newDebitMoney.Osnovanie = OsnovanieTextBox.Text;

            // вносим экземпляр в коллекцию
            DebitMoneyCollection.Add(newDebitMoney);

            DataGrid.ItemsSource = DebitMoneyCollection;
            DataGrid.Items.Refresh();
            
            // считаем нажатие на кнопку
            NumberClickButton += 1;
            
            // сбрасываем значения контролов
            MoveComboBox.SelectedIndex = -1;
            DocNamberTexBox.Text = "";
            DebitMoneyTextBox.Text = "";
            OsnovanieTextBox.Text = "";


        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excel.Filename);
            Worksheet debitMoneyWorksheet = (Worksheet) workbook.Sheets["Приход ДС"];
            Range debitMoneyRange = debitMoneyWorksheet.UsedRange;

            int lastRow = debitMoneyRange.Rows.Count;

            // записываем в таблицу
            for (int i = 0; i < NumberClickButton; i++)
            {
                debitMoneyWorksheet.Cells[lastRow + i + 1, 1] = lastRow + i;
                debitMoneyWorksheet.Cells[lastRow + i + 1, 2] = DebitMoneyCollection[i].Data;
                debitMoneyWorksheet.Cells[lastRow + 1 + i, 3] = DebitMoneyCollection[i].TypeMove;
                debitMoneyWorksheet.Cells[lastRow + i + 1, 4] = DebitMoneyCollection[i].DocumentNumber;
                debitMoneyWorksheet.Cells[lastRow + i + 1, 5] = DebitMoneyCollection[i].Statia;
                if (DebitMoneyCollection[i].TypeMoveIndex == 0)
                    debitMoneyWorksheet.Cells[lastRow + i + 1, 6] = DebitMoneyCollection[i].Debit;
                if (DebitMoneyCollection[i].TypeMoveIndex == 1)
                    debitMoneyWorksheet.Cells[lastRow + i + 1, 7] = DebitMoneyCollection[i].Debit;
                debitMoneyWorksheet.Cells[lastRow + i + 1, 8] = DebitMoneyCollection[i].Osnovanie;                
            }

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            MessageBox.Show("Данные успешно внесены в базу!");
            this.Close();           
        }


        /// <summary>
        /// Выбор в ComboBox Тип движения
        /// Обеспечивает подгрузку коллекции в ComboBox Статья 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MoveComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {           
            // Очищаем текущую коллекцию
            DebitMoneyComboBox.Items.Clear();
            // добавляем значения в зависимости от выбранного значения
            switch (MoveComboBox.SelectedIndex)
            {
                // ПРИХОД
                case 0: 
                {
                    DebitMoneyComboBox.Items.Add("Аванс");
                    DebitMoneyComboBox.Items.Add("Полная оплата заказа");
                    DebitMoneyComboBox.Items.Add("Окончание оплаты заказа");
                    DebitMoneyComboBox.Items.Add("Другое");
                }break;
                //РАСХОД
                case 1: 
                {
                    DebitMoneyComboBox.Items.Add("Аванс");
                    DebitMoneyComboBox.Items.Add("Заработная плата");
                    DebitMoneyComboBox.Items.Add("Оплата услуг");
                    DebitMoneyComboBox.Items.Add("Другое");
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

        private void DebitMoneyTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // проверка на корректность данных
            if (DebitMoneyTextBox.Text != "" & DebitMoneyComboBox.SelectedIndex != -1 & DocNamberTexBox.Text != "")
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

        private void DocNamberTexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            DebitMoneyTextBox_TextChanged(sender, e);
        }

        private void DebitMoneyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // проверка на корректность данных
            if (DebitMoneyTextBox.Text != "" & DebitMoneyComboBox.SelectedIndex != -1 & DocNamberTexBox.Text != "")
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
    }
}
