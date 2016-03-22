using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CreditApp
{
    class DebitMaterial
    {

        private string documentNumber;
        private string data;
        private int summ;
        private string material;
        private int materialIndex;
        private double debit;
        private double price;
        private int row;
        private double localsumm;
        private string provider;


        public string Provider
        {
            get { return provider; }
            set { provider = value; }
        }

        /// <summary>
        /// Номер документа
        /// </summary>
        public string DocumentNumber
        {
            get { return documentNumber; }
            set { documentNumber = value; }
        }

        /// <summary>
        /// Дата в документе
        /// </summary>
        public string Data
        {
            get { return data; }
            set { data = value; }
        }

        /// <summary>
        /// Сумма документа (Счёта-фактуры)
        /// </summary>
        public int Summ
        {
            get { return summ; }
            set { summ = value; }
        }

        /// <summary>
        /// Наименование материала
        /// </summary>
        public string Material
        {
            get { return material; }
            set { material = value; }
        }

        /// <summary>
        /// Индекс материала в ComboBox
        /// </summary>
        public int MaterialIndex
        {
            get { return materialIndex; }
            set { materialIndex = value; }
        }

        /// <summary>
        /// Количество материала
        /// </summary>
        public double Debit
        {
            get { return debit; }
            set { debit = value; }
        }

        /// <summary>
        /// Цена материала
        /// </summary>
        public double Price
        {
            get { return price; }
            set { price = value; }
        }

        /// <summary>
        /// Строка в которую должен быть записан 
        /// </summary>
        public int Row
        {
            get { return row; }
            set { row = value; }
        }

        /// <summary>
        /// Сумма прзиции
        /// </summary>
        public double LocalSumm
        {
            get { return localsumm; }
            set { localsumm = value; }
        }

        /// <summary>
        /// Единици измерения материала
        /// </summary>
        public string Edinici { get; set; }



    }

}
