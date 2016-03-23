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
            Window creditWindow = new CreditMaterialWindow();
            creditWindow.Show();
        }



        private void Button_Click_Debit_Maney(object sender, RoutedEventArgs e)
        {
            Window debitmoneyWindow = new MoveMoneyWindow();
            debitmoneyWindow.Show();
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void AddNewMaterial_Click(object sender, RoutedEventArgs e)
        {
            Window addNewMaterialWindow = new AddNewMaterialWindow();
            addNewMaterialWindow.ShowDialog();
        }

        private void AddNewProvider_Click(object sender, RoutedEventArgs e)
        {
            Window addNewProviderWindow = new AddNewProviderWindow();
            addNewProviderWindow.ShowDialog();
        }



    }
}
