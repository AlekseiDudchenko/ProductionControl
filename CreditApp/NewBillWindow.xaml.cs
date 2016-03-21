using System;
using System.Collections.Generic;
using System.Linq;
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

namespace CreditApp
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class NewBillWindow: Window
    {
        public NewBillWindow()
        {
            InitializeComponent();

            // блокируем кнопку
            CreateNewBillButton.IsEnabled = false;
        }

        private void CreateNewBill(object sender, RoutedEventArgs e)
        {

            DebitWindow debitWindow = new DebitWindow();
            debitWindow.Show();

            // передаем значения в новое окно
            debitWindow.DocNamberTexBox.Text = this.NomberBillTextBox.Text;
            debitWindow.BillSummLabel.Content = this.BillPriceTextBox.Text;

            // закрываем текущее окно
            this.Close();
        }


        private void BillPriceTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (NomberBillTextBox.Text != "" & BillPriceTextBox.Text != "")
            {
                CreateNewBillButton.IsEnabled = true;
            }
            else
            {
                CreateNewBillButton.IsEnabled = false;
            }
        }

        private void NomberBillTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (NomberBillTextBox.Text != "" & BillPriceTextBox.Text != "")
            {
                CreateNewBillButton.IsEnabled = true;
            }
            else
            {
                CreateNewBillButton.IsEnabled = false;
            }
        }
    }
}
