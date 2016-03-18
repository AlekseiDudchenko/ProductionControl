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
                DebitNumberTextBox.Text == "" || CreditTexBox.Text == "" || CreditTexBox.Text == "" ||
                ReasonCreditTexBox.Text == "" || CreditNumberTexBox.Text == "")
            {
                MessageBox.Show("Введены не все данные!\nВведите все данные и попробуйте снова.");
            }             
            else
            {
                DataUpdate();
            }           
        }


        private int lastrow;

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
                lastrow = row;
            }

            // номер первой пустой строки
            lastrow ++;
            // заполняем заполняем ячейки файла данными
            worksheet.Cells[lastrow, 1] = lastrow - 1;
            worksheet.Cells[lastrow, 2] = DataTextBox.Text;
            worksheet.Cells[lastrow, 3] = SupplyTexBox.Text;
            worksheet.Cells[lastrow, 4] = ReasonSupplyTexBox.Text;
            worksheet.Cells[lastrow, 5] = DebitNumberTextBox.Text;
            worksheet.Cells[lastrow, 6] = CreditTexBox.Text;
            worksheet.Cells[lastrow, 7] = ReasonCreditTexBox.Text;
            worksheet.Cells[lastrow, 8] = CreditNumberTexBox.Text;
            // в ячейку девятого столбца вводим формулу для подсчета суммы. Таким образом храним в файле не сумму, а заставляем Excel её считать самостоятельно
            worksheet.Cells[lastrow, 9] = ("=C" + lastrow + "-F" + lastrow);
        
            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // подсчитываем сумму за день из формы
            int _summ = Convert.ToInt32(SupplyTexBox.Text) - Convert.ToInt32(CreditTexBox.Text);

            // выводим сообщение на экран об успешности операции
            MessageBox.Show("Изменения сохранены.\nДата: " + DataTextBox.Text + "\nПриход: " + SupplyTexBox.Text +
                            " по документу № " + DebitNumberTextBox.Text + "\nРасход:  " + CreditTexBox.Text +
                            " по документу № " + CreditNumberTexBox.Text + "\n\nБаланс за день: " + _summ);
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


        private void Button_Click_Credit(object sender, RoutedEventArgs e)
        {
            // открываем форму Material
            var creditMaterial = new CreditMaterialView();
            creditMaterial.ShowDialog();
        }

        private void Button_Click_Debit(object sender, RoutedEventArgs e)
        {
            // открываем форму Material
            var debitMaterial = new DebitMaterialView();
           debitMaterial.ShowDialog();
        }


    }
}









