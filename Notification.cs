using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcleBirthday
{
    internal class Notification
    {
        public static void ViewNotif()
        {
            string[] fullDate = DateTime.Now.ToString().Split(' ');
            string[] date = fullDate[0].Split('.');
            string dayNow = date[0];
            string monthNow = date[1];

            string birthday = "\n";
            foreach (ExcelData ed in CheckBirthday.Birthday)
            {
                birthday += $"{ed.Name}\n";
            }
            var notify = new ToastContentBuilder();
            notify.AddText($"Сегодня день рождения у {birthday}Дни рождения {dayNow}.{monthNow}");
            notify.Show();
        }
    }
}
