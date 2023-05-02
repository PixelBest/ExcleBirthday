using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections;
using System.Data.SqlTypes;
using ClosedXML.Excel;
using System.Collections.ObjectModel;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Microsoft.Toolkit.Uwp.Notifications;

namespace ExcleBirthday
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string path = Directory.GetCurrentDirectory();
        static bool check = true;
        public MainWindow() 
        {
            InitializeComponent();
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)    //выводит всех работников из Excel файла у которых сегодня день рождение
        {
            CheckBirthday.Birthday.Clear();
            CheckBirthday.CheckBirth();
            dataGrid1.ItemsSource = CheckBirthday.Birthday;
        }

        private void Button_Click1(object sender, RoutedEventArgs e)    //загружает всех работников из выбранного Excel файла
        {
            if(OpenExcelFail.fileName == "")
            {
                OpenExcelFail.OpenFail();
                MessageBox.Show("Это займёт некоторое время", "Уведомление", MessageBoxButton.OK);
                ExcelReader.ExcelRead();
                dataGrid1.ItemsSource = ExcelReader.data;
            }
            else
            {
                MessageBox.Show("Файл уже выбрал");
            }
        }

        private void Button_Click2(object sender, RoutedEventArgs e)    //выводит всех работников из Excel файла
        {
            dataGrid1.ItemsSource = ExcelReader.data;
        }

        private void Button_Click3(object sender, RoutedEventArgs e)    //переводит в фоновый процесс и уведомляет каждый день у кого сегодня день рождение
        {
            this.Hide();
            CheckBirthday.Birthday.Clear();
            CheckBirthday.CheckBirth();
            Notification.ViewNotif();
            CheckBirthday.Birthday.Clear();
            while (check)
            {
                Thread.Sleep(10000);
                CheckBirthday.CheckBirth();
                Notification.ViewNotif();
                CheckBirthday.Birthday.Clear();
            }
        }

        private void Button_Click4(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Лучшее приложения для просмотра ДР 1.0","Версия", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
