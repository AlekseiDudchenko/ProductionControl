using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace ProductionControl
{
    class MainViewModel
    {
        ObservableCollection<MaterialViewModel> MaterialList { get; set; }
        
        /*
        public MainViewModel(List<Material> materials)
        {
            MaterialList = new ObservableCollection<MaterialViewModel>(materials.Select(b => new MaterialViewModel(b)));

        }// это если описывать содержимое в App
         */

        public MainViewModel()
        {


        }

    }
}
