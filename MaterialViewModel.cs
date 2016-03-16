using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionControl
{
    class MaterialViewModel : ViewModelBase
    {
        public Material Material;

        public MaterialViewModel(Material material)
        {
            this.Material = material;
        }

        public string Name
        {
            get { return Material.Name; }
            set
            {
                Material.Name = value;
                OnPropertyChanged("Name");
            }
        }

        public string Cod
        {
            get { return Material.Cod; }
            set
            {
                Material.Cod = value;
                OnPropertyChanged("Cod");
            }
        }

        public int Count
        {
            get { return Material.Count; }
            set
            {
                Material.Count = value;
                OnPropertyChanged("Count");
            }
        }
    }
}
