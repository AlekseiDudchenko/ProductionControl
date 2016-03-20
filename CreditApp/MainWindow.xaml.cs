using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelClass excel = new ExcelClass();

        public MainWindow()
        {
            // проверяем наличие файла
            if (!File.Exists(excel.Filename))
            {
                MessageBox.Show("Файл не найден!");
                this.Close();
            }
            else
            {
                // загрузка компонентов
                InitializeComponent();
                // заполняем текущее время  
                Datelable.Content = Convert.ToString("Текущая дата: " + DateTime.Now.ToString("dd.MM.yyyy"));
            }                  
        }

        private void Button_Click_Debit_Material(object sender, RoutedEventArgs e)
        {         
                Window newBillWindow = new NewBillWindow();
                newBillWindow.ShowDialog();          
        }

        private void Button_Click_Credit_Material(object sender, RoutedEventArgs e)
        {       
                Window creditWindow = new CreditWindow();
                creditWindow.Show();           
        }

        private void Button_Click_Credit_Maney(object sender, RoutedEventArgs e)
        {
            Window creditMoneyWindow = new CreditMoneyWindow();
            creditMoneyWindow.Show();                       
        }

        private void Button_Click_Debit_Maney(object sender, RoutedEventArgs e)
        {
            Window debitmoneyWindow = new DebitMoneyWindow();
            debitmoneyWindow.Show();
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
