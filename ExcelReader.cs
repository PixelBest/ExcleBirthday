using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml.ExtendedProperties;
using LinqToExcel;
using eh = ExcelHelper;

namespace ExcleBirthday
{
    internal class ExcelReader
    {
        public static ObservableCollection<ExcelData> data = new ObservableCollection<ExcelData>();
        public static void ExcelRead()  //читает данные из Excel файла
        {
            bool check = true;
            int a = 2;
            eh.ExcelHelper eh = new eh.ExcelHelper();
            eh.Open(OpenExcelFail.fileName);
            while (check)
            {
                if (eh.Get(colum: "A", row: a).ToString() != "")
                {
                    ExcelData excel = new ExcelData();
                    excel.Name = eh.Get(colum: "A", row: a).ToString();
                    excel.TabNum = eh.Get(colum: "B", row: a).ToString();
                    excel.Position = eh.Get(colum: "C", row: a).ToString();
                    excel.DatePriema = eh.Get(colum: "D", row: a).ToString();
                    excel.DateYvolneniya = eh.Get(colum: "E", row: a).ToString();
                    excel.DataOfBirth = eh.Get(colum: "F", row: a).ToString();
                    excel.WillYear = Convert.ToInt32(eh.Get(colum: "J", row: a).ToString());
                    excel.SmallYbiley = eh.Get(colum: "K", row: a).ToString();
                    excel.Ybiley = eh.Get(colum: "L", row: a).ToString();
                    data.Add(excel);
                    a++;
                }
                else
                    break;
            }
            eh.Dispose();
        }
    }
}
