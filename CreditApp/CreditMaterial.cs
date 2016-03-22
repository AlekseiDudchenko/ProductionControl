using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreditApp
{
    /// <summary>
    /// Запись Расхода материала
    /// </summary>
    class CreditMaterial
    {
        private string _date;
        private string _materialName;
        private int _materialIndex;
        private double _creditMaterial;

        /// <summary>
        /// Дата
        /// </summary>
        public string Data { get; set; }

        /// <summary>
        /// Номер документа
        /// </summary>
        public string DocumentNumber { get; set; }

        /// <summary>
        /// Наименование материала
        /// </summary>
        public string MaterialName { get; set; }

        /// <summary>
        /// Порядковый номер материала. Соответствует индексу в MaterialComboBox
        /// </summary>
        public  int MaterialIndex { get; set; }

        /// <summary>
        /// Количество списываемого материала
        /// </summary>
        public double Credit { get { return _creditMaterial; } set { _creditMaterial = value; } }

        public string Edinici { get; set; }

    }
}
