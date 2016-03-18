using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionControl
{
    class Material
    {
        // имя материала
        public string Name;
        // код материала в локальном классификаторе
        public string Cod;
        // количество материала которое остальсь на складе
        public int Count; //Сколько осталось


        public Material(string name, string cod, int count)
        {
            this.Name = name;
            this.Cod = cod;
            this.Count = count;
        }



        

    }
}
