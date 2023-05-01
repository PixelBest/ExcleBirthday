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
        async public static Task ViewNotif()
        {
            string birthday = "\n";
            foreach(ExcelData ed in ExcelReader.Birthday)
            {
                birthday += $"{ ed.Name}\n";
            }
            Task.Run(() =>
            {
                MessageBox.Show($"Сегодня день рождение у {birthday}", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }
    }
}
