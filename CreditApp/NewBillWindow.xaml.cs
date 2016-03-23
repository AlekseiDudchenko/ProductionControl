using System;
using System.Windows;
using System.Windows.Controls;

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
            DebitMaterialWindow debitWindow = new DebitMaterialWindow();
            debitWindow.Show();

            // передаем значения в новое окно
            debitWindow.DocNamberTexBox.Content = NomberBillTextBox.Text;
            debitWindow.BillSummLabel.Content = BillPriceTextBox.Text;
            debitWindow.ProviderNameLabel.Content = ProviderNameComboBox.SelectedItem.ToString();

            // закрываем текущее окно
            this.Close();
        }


        private void BillPriceTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            bool convertToDouble = false;
            try
            {
                Convert.ToDouble(BillPriceTextBox.Text);
                convertToDouble = true;
            }
            catch (Exception)
            {
                CreateNewBillButton.IsEnabled = false;
                //throw;
            }
            if (NomberBillTextBox.Text != "" & BillPriceTextBox.Text != "" & convertToDouble)
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
            BillPriceTextBox_TextChanged(sender, e);
        }
    }
}
