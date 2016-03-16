using System;
using System.Windows;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductionControl
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }
     
        /// <summary>
        /// Обрабатывает нажатие клавиши Сохранить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (DataTextBox.Text == "" || SupplyTexBox.Text == "" || ReasonSupplyTexBox.Text == "" ||
                SypplyNumberTextBox.Text == "" || ExpendTexBox.Text == "" || ExpendTexBox.Text == "" ||
                ReasonExpendTexBox.Text == "" || ExpendNumberTexBox.Text == "")
            {
                MessageBox.Show("Введены не все данные!\nВведите все данные и попробуйте снова.");
            }             
            else
            {
                DataUpdate();
            }           
        }


        private int _row;

        /// <summary>
        /// Осуществляет запись данных из формы в файл
        /// </summary>
        public void DataUpdate()
        {
            // создаеем экземпляр ExcelData для получения имени файла из которого читали данные и в который будем записыавть.
            // TODO: Обеспецить пользователю возможность выбора этого файла через интерфейс программы и, возможно, листа. 
            ExcelData nED = new ExcelData();

            // указываем документ с которым будем работать
            Excel.Application excelApp = new Excel.Application();             
            Excel.Workbook workbook = excelApp.Workbooks.Open(Environment.CurrentDirectory + nED.FileName); //адрес и имя файла
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets["Test Sheet"]; // Имя листа
            Excel.Range range = worksheet.UsedRange;

            int column = 0;
            
            //Получаем номер последней заполненной строки
            //TODO: найти другой способ
            for ( int row = 2; row <= range.Rows.Count; row++)
            {
                for (column = 1; column <= range.Columns.Count; column++)
                {
                }
                _row = row;
            }

            // номер первой пустой строки
            _row ++;
            // заполняем заполняем ячейки файла данными
            worksheet.Cells[_row, 1] = _row - 1;
            worksheet.Cells[_row, 2] = DataTextBox.Text;
            worksheet.Cells[_row, 3] = SupplyTexBox.Text;
            worksheet.Cells[_row, 4] = ReasonSupplyTexBox.Text;
            worksheet.Cells[_row, 5] = SypplyNumberTextBox.Text;
            worksheet.Cells[_row, 6] = ExpendTexBox.Text;
            worksheet.Cells[_row, 7] = ReasonExpendTexBox.Text;
            worksheet.Cells[_row, 8] = ExpendNumberTexBox.Text;
            // в ячейку девятого столбца вводим формулу для подсчета суммы. Таким образом храним в файле не сумму, а заставляем Excel её считать самостоятельно
            worksheet.Cells[_row, 9] = ("=C" + _row + "-F" + _row);
        
            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // подсчитываем сумму за день из формы
            int _summ = Convert.ToInt32(SupplyTexBox.Text) - Convert.ToInt32(ExpendTexBox.Text);

            // выводим сообщение на экран об успешности операции
            MessageBox.Show("Изменения сохранены.\nДата: " + DataTextBox.Text + "\nПриход: " + SupplyTexBox.Text +
                            " по документу № " + SypplyNumberTextBox.Text + "\nРасход:  " + ExpendTexBox.Text +
                            " по документу № " + ExpendNumberTexBox.Text + "\n\nБаланс за день: " + _summ);
        }


        /// <summary>
        /// Обрабатывает клик по кнопке Материалы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            // открываем форму Material
             var Matertial =  new MaterialView();
             Matertial.ShowDialog();
        }
    }
}









