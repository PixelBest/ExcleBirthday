using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcleBirthday
{
    public class ExcelData
    {
        public string Name { get; set; }
        public string TabNum { get; set; }
        public string Position { get; set; }
        public string DatePriema { get; set; }
        public string DateYvolneniya { get; set; }
        public string DataOfBirth { get; set; }
        public int WillYear { get; set; }
        public string SmallYbiley { get; set; }
        public string Ybiley { get; set; }

        public ExcelData()
        {
            Name = "";
            TabNum = "";
            Position = "";
            DatePriema = "";
            DateYvolneniya = "";
            DataOfBirth = "";
            WillYear = 0;
            SmallYbiley = "";
            Ybiley = "";
        }

        public ExcelData(string name, string tabNum, string position, string datePriema, string dateYvolneniya, string dataOfBirth, int willYear, string smallYbiley, string ybiley, string todayDate)
        {
            Name = name;
            TabNum = tabNum;
            Position = position;
            DatePriema = datePriema;
            DateYvolneniya = dateYvolneniya;
            DataOfBirth = dataOfBirth;
            WillYear = willYear;
            SmallYbiley = smallYbiley;
            Ybiley = ybiley;
        }
    }
}