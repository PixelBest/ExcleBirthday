using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcleBirthday
{
    internal class CheckBirthday
    {
        public static ObservableCollection<ExcelData> Birthday = new ObservableCollection<ExcelData>();
        public static void CheckBirth()
        {
            string[] fullDate = DateTime.Now.ToString().Split(' ');
            string[] date = fullDate[0].Split('.');
            string dayNow = date[0];
            string monthNow = date[1];
            string day = "";
            string month = "";

            foreach (ExcelData ed in ExcelReader.data)
            {
                string[] checkD = ed.DataOfBirth.Split(' ');
                string[] checkDate = checkD[0].Split('.');
                day = checkDate[0];
                month = checkDate[1];
                if(day == dayNow && month == monthNow)
                {
                    Birthday.Add(ed);
                }
            }
        }
    }
}
