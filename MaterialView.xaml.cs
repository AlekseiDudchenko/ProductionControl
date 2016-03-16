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

namespace ProductionControl
{
    /// <summary>
    /// Логика взаимодействия для Material.xaml
    /// </summary>
    public partial class MaterialView : Window
    {
        public MaterialView()
        {
            InitializeComponent();
        }

        public string[] MaterialList = new string[100];

        public string[] MaterialList1
        {
            get { return MaterialList; }
            set { MaterialList = value; }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}
