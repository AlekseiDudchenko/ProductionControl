using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreditApp
{
    class ExcelClass
    {
        // устанавливаем адрес файла
        private string filename = Environment.CurrentDirectory + "\\6.xlsx";
        //private string filename = "C:\\12.xlsx";

        public string Filename
        {
            get { return filename; }
        }

    }
}
