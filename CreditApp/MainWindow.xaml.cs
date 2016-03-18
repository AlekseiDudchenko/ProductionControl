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
            InitializeComponent();



            // заполняем текущее время         
            Datelable.Content = Convert.ToString("Текущая дата: " + DateTime.Now.ToString("dd.MM.yyyy"));
        }

        private void Button_Click_Debit_Material(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(excel.Filename))
            {
                MessageBox.Show("Файл не найден!");
                //this.Close();
            }
            else
            {
                Window debitWindow = new DebitWindow();
                debitWindow.Show();
            }

        }

        private void Button_Click_Credit_Material(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(excel.Filename))
            {
                MessageBox.Show("Файл не найден!");
                //this.Close();
            }
            else
            {
                Window creditWindow = new CreditWindow();
                creditWindow.Show();
            }
        }

        private void Button_Click_Credit_Maney(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(excel.Filename))
            {
                MessageBox.Show("Файл не найден!");
                //this.Close();
            }
            else
            {
                Window credirManeyWindow = new CreditWindow();
                credirManeyWindow.Show();
            }
        }

        private void Button_Click_Debit_Maney(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(excel.Filename))
            {
                MessageBox.Show("Файл не найден!");
                //this.Close();
            }
            else
            {
                Window debitmaneyWindow = new CreditWindow();
                debitmaneyWindow.Show();
            }
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
