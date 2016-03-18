using System;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;


namespace ProductionControl
{
    /// <summary>
    /// Логика взаимодействия для UserControl1.xaml
    /// </summary>
    public partial class CreditMaterialView : Window
    {
        private int lastrow = 0;

        public CreditMaterialView()
        {
            InitializeComponent();

            // открываем документ и лист для считывания данных для comboBox
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();           
            Microsoft.Office.Interop.Excel.Workbook workbook =
                excelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Materials.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet worksheet =
                (Microsoft.Office.Interop.Excel.Worksheet) workbook.Sheets["Mat"];
            


            // заполняем comboBox значениями
            for (int i = 3; i < 76; i++)
            {
                MaterialComboBox.Items.Add(worksheet.Cells[i, 2].Value);
            }

            // получем номер последнего документа
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Расход материалов"];
            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

            DocNamberTexBox.Text = worksheet.Cells[2, 2].Value;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();


            // заполняем текущее время
            DateTexBox.Text = DateTime.Now.ToString("dd.MM.yyyy");

        }

        private void Button_Click_Add(object sender, RoutedEventArgs e)
        {           

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();        
            Microsoft.Office.Interop.Excel.Workbook workbook =
                excelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Materials.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet worksheet =
                (Microsoft.Office.Interop.Excel.Worksheet) workbook.Sheets["Расход материалов"];
            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

            int column = 0;


            //Получаем номер последней заполненной строки
            //TODO: найти другой способ
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                for (column = 1; column <= range.Columns.Count; column++)
                {
                }
                lastrow = row;
            }

            //Стираем все с последней строки. Потом запишем это заново в конец
            for (column = 1; column <= range.Columns.Count; column++)
            {
                worksheet.Cells[lastrow, column] = "";
            }


            // заполняем заполняем ячейки файла данными
            worksheet.Cells[lastrow, 1] = lastrow - 2;
            worksheet.Cells[lastrow, 2] = DateTexBox.Text;
            worksheet.Cells[lastrow, 3] = DocNamberTexBox.Text;
            worksheet.Cells[lastrow, MaterialComboBox.SelectedIndex+4] = CreditMaterialTextBox.Text;
            // в ячейку девятого столбца вводим формулу для подсчета суммы. Таким образом храним в файле не сумму, а заставляем Excel её считать самостоятельно
            worksheet.Cells[lastrow+1, 3] = ("Сумма");  // ("=C" + lastrow + "-F" + lastrow);

            // записываем в файл формулу подсчета суммы, которую стерли ранее
            for (column = 1; column <= range.Columns.Count; column++)
            {
                worksheet.Cells[lastrow+1, column] = "Сумма";
            }



            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            // выводим информационно сообщение
            MessageBox.Show("Добавлен расход материала " + MaterialComboBox.Text + " в размере " + CreditMaterialTextBox.Text);

            MaterialComboBox.Text = "выберите материал"; // так не работает ))
            CreditMaterialTextBox.Text = String.Empty;

        }



    }
}
