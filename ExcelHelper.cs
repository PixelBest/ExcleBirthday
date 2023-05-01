using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcleBirthday
{
    internal class ExcelHelper : IDisposable
    {
        private Excel.Application excel;
        private Excel.Workbook workbook;
        private string _failPath;

        public ExcelHelper() { excel = new Excel.Application(); }

        internal void Open(string failPath)    //открывает выбранный файл
        {
            if(File.Exists(failPath))
                workbook = excel.Workbooks.Open(failPath);
            else
            {
                workbook = excel.Workbooks.Add();
                _failPath = failPath;
            }
        }

        internal void Set(string colum, int row, string data)   //изменяет данные в Excel файле
        {
            try
            {
                ((Excel.Worksheet)excel.ActiveSheet).Cells[row, colum] = data;
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
        }

        internal object Get(string colum, int row)    //получает данные из Excel файла
        {
            try
            {
                if (((Excel.Worksheet)excel.ActiveSheet).Cells[row, colum].Value != null)
                    return ((Excel.Worksheet)excel.ActiveSheet).Cells[row, colum].Value;
                else
                    return "";
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
            return null;
        }

        internal void Save()    //сохраняет Excel файл
        {
            if(!string.IsNullOrEmpty(_failPath))
            {
                workbook.SaveAs(_failPath);
                _failPath = null;
            }
            else
            {
                workbook.Save();
            }
        }

        public void Dispose()    //закрывает процесс Excel файла
        {
            try
            {
                excel.Quit();
                workbook.Close();
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
