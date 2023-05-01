using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcleBirthday
{
    internal class OpenExcelFail
    {
        public static string fileName = "";
        public static void OpenFail()   //открывает диалоговое окно где можно выбрать нужный Excel файл
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.FileName = "data";
            ofd.DefaultExt = ".xlsx";
            ofd.Filter = "Excel file (.xlsx)|*.xlsx| Excel file(.xlsm)|*.xlsm";
            if (ofd.ShowDialog() == true)
            {
                fileName = ofd.FileName;
            }
        }
    }
}
