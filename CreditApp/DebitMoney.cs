using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreditApp
{
    class DebitMoney
    {
        private string documentNumber;
        private string data;
        private string statia;
        private int statiaIndex;
        private double debit;
        private string osnovanie;
        private string typeMove;
        private int typeMoveIndex;

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
        /// Наименование статьи
        /// </summary>
        public string Statia
        {
            get { return statia; }
            set { statia = value; }
        }

        /// <summary>
        /// Индекс статьи в ComboBox
        /// </summary>
        public int StatialIndex
        {
            get { return statiaIndex; }
            set { statiaIndex = value; }
        }

        /// <summary>
        /// Сумма прихода
        /// </summary>
        public double Debit
        {
            get { return debit; }
            set { debit = value; }
        }

        /// <summary>
        /// Содержит описание основания
        /// </summary>
        public string Osnovanie
        {
            get { return osnovanie; }
            set { osnovanie = value; }
        }

        /// <summary>
        /// Тип движения. Приход или Расход
        /// </summary>
        public string TypeMove
        {
            get { return typeMove; }
            set { typeMove = value; }
        }

        /// <summary>
        /// Индекс в ComboBox указывающий Приход или Расход
        /// </summary>
        public int TypeMoveIndex
        {
            get { return typeMoveIndex; }
            set { typeMoveIndex = value; }
        }

    }
}
