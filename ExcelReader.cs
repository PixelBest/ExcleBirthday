using System;
using System.Collections.ObjectModel;
using System.Windows;

namespace ExcleBirthday
{
    internal class ExcelReader
    {
        public static ObservableCollection<ExcelData> data = new ObservableCollection<ExcelData>();
        public static ObservableCollection<ExcelData> Birthday = new ObservableCollection<ExcelData>();
        public static string day;
        public static string month;
        public static string year;
        public static void ExcelRead()  //читает данные из Excel файла
        {
            string[] fullDate = DateTime.Now.ToString().Split(' ');
            string[] date = fullDate[0].Split('.');
            string dayNow = date[0];
            string monthNow = date[1];

            bool check = true;
            int a = 2;
            ExcelHelper eh = new ExcelHelper();
            eh.Open(OpenExcelFail.fileName);
            while (check)
            {
                if(eh.Get(colum: "A", row: a).ToString() != "")
                {
                    ExcelData excel = new ExcelData();
                    excel.Name = eh.Get(colum: "A", row: a).ToString();
                    excel.TabNum = eh.Get(colum: "B", row: a).ToString();
                    excel.Position = eh.Get(colum: "C", row: a).ToString();
                    excel.DatePriema = eh.Get(colum: "D", row: a).ToString();
                    excel.DateYvolneniya = eh.Get(colum: "E", row: a).ToString();
                    excel.DataOfBirth = eh.Get(colum: "F", row: a).ToString();
                    excel.WillYear = Convert.ToInt32(eh.Get(colum: "J", row: a));
                    if(excel.WillYear % 5 == 0)
                    {
                        if(excel.WillYear % 10 == 0)
                            excel.Ybiley = "Юбилей";
                        else
                            excel.SmallYbiley = "Малый юбилей";
                    }

                    data.Add(excel);
                    a++;

                    string[] checkD = excel.DataOfBirth.Split(' ');
                    string[] checkDate = checkD[0].Split('.');
                    string day = checkDate[0];
                    string month = checkDate[1];

                    if (day == dayNow && month == monthNow)
                    {
                        if (excel.WillYear % 5 == 0)
                        {
                            if (excel.WillYear % 10 == 0)
                                excel.Ybiley = "Сегодня юбилей";
                            else
                                excel.SmallYbiley = "Сегодня малый юбилей";
                        }
                        Birthday.Add(excel);
                        excel.WillYear += 1;
                    }
                }
                else
                    break;
            }
            eh.Dispose();
        }
    }
}
