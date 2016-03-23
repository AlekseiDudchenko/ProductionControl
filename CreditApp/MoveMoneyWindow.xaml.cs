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
 
        public MoveMoneyWindow()
        {
            InitializeComponent();

            // заполняем текущее время
            DatePicker.Text = DateTime.Now.ToString("dd.MM.yyyy");
        }

        ObservableCollection<DebitMoney> moveMoneyCollection = new ObservableCollection<DebitMoney>();
       
        /// <summary>
        ///  Нажатие кнопки ДОБАВИТЬ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            // создаем экземпляр и заполняем поля
            DebitMoney newMoveMoney = new DebitMoney
            {
                Data = DatePicker.Text,
                DocumentNumber = DocNamberTexBox.Text,
                Statia = MoveMoneyComboBox.Text,
                StatialIndex = MoveMoneyComboBox.SelectedIndex,
                Debit = Convert.ToDouble(DebitMoneyTextBox.Text),
                TypeMove = MoveComboBox.Text,
                TypeMoveIndex = MoveComboBox.SelectedIndex,
                Osnovanie = OsnovanieTextBox.Text
            };

            // вносим экземпляр в коллекцию
            moveMoneyCollection.Add(newMoveMoney);
            DataGrid.ItemsSource = moveMoneyCollection;
            DataGrid.Items.Refresh();
       
            // сбрасываем значения контролов
            MoveComboBox.SelectedIndex = -1;
            DocNamberTexBox.Text = DebitMoneyTextBox.Text = OsnovanieTextBox.Text = String.Empty;

            // если DataGrid не пустой - разблокировать кнопку Сохранить
            SaveButton.IsEnabled = (DataGrid.Items != null);
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
            for (int i = 0; i < moveMoneyCollection.Count; i++)
            {
                debitMoneyWorksheet.Cells[lastRow + i , 1] = lastRow + i;
                debitMoneyWorksheet.Cells[lastRow + i , 2] = moveMoneyCollection[i].Data;
                debitMoneyWorksheet.Cells[lastRow + i , 3] = moveMoneyCollection[i].TypeMove;
                debitMoneyWorksheet.Cells[lastRow + i , 4] = moveMoneyCollection[i].DocumentNumber;
                debitMoneyWorksheet.Cells[lastRow + i , 5] = moveMoneyCollection[i].Statia;
                if (moveMoneyCollection[i].TypeMoveIndex == 0)
                    debitMoneyWorksheet.Cells[lastRow + i , 6] = moveMoneyCollection[i].Debit;
                if (moveMoneyCollection[i].TypeMoveIndex == 1)
                    debitMoneyWorksheet.Cells[lastRow + i , 7] = moveMoneyCollection[i].Debit;
                debitMoneyWorksheet.Cells[lastRow + i , 8] = moveMoneyCollection[i].Osnovanie;                
            }

            // записываем формулы суммы
            debitMoneyWorksheet.Cells[lastRow + DataGrid.Items.Count, 6].FormulaLocal = ("=СУММ(F" + (lastRow + DataGrid.Items.Count - 1) + ":F2)");
            debitMoneyWorksheet.Cells[lastRow + DataGrid.Items.Count, 7].FormulaLocal = ("=СУММ(G" + lastRow + (DataGrid.Items.Count - 1) + ":G2)");

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // TODO проверять успешность сохранения
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
            // очищаем текущую коллекцию
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
        /// Возможно будем использовать если окажется что номер документа может быть не только числовой
        /// </summary>
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
            AddButton.IsEnabled = Functions.ProverkaDannih(DocNamberTexBox, DebitMoneyTextBox, MoveMoneyComboBox);
        }

        private void DocNamberTexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            AddButton.IsEnabled = Functions.ProverkaDannih(DocNamberTexBox, DebitMoneyTextBox, MoveMoneyComboBox);
        }

        private void MoveMoneyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // меняем названия документоd в зависимости от выбранной статьи расхода
            string documentName = "Документ №";
            if (MoveMoneyComboBox.SelectedValue != null)
            {
                if (MoveMoneyComboBox.SelectedValue.ToString() == "Аванс")
                    documentName = "Наряд №";
                if (MoveMoneyComboBox.SelectedValue.ToString() == "Оплата услуг")
                    documentName = "Чек №";
                if (MoveMoneyComboBox.SelectedValue.ToString() == "Полная оплата заказа")
                    documentName = "Чек №";
            }
            DocumentNameLabel.Content = documentName;

            // упраялем доступностью кнопки добавить
            AddButton.IsEnabled = Functions.ProverkaDannih(DocNamberTexBox, DebitMoneyTextBox, MoveMoneyComboBox);
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

        private void DataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            SaveButton.IsEnabled = (DataGrid.Items != null);
        }
    }
}
