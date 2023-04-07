using HelpMe.Model;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace HelpMe
{
    public partial class MainPage : Page
    {
        public MainPage()
        {
            InitializeComponent();
            MyLogger.Instance.Info("Приложение запущено");
            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Готовые счета");
            
        }

        private void Start(object sender, RoutedEventArgs e)
        {
            AllDataBox.Text = "";
            if (Datas.DatasInitFromExcel(Properties.Settings.Default.PathToDatas))
            {
                AllDataBox.Text += "Выписки успешно загружены\n";
                MyLogger.Instance.Info("Выписки успешно загружены");
            }
            else
            {
                AllDataBox.Text += "Выписки не загружены\n";
                MyLogger.Instance.Info("Выписки не загружены");
            }

            if (Currencies.CurrenciesInitFromCSV(Properties.Settings.Default.PathToCurr))
            {
                AllDataBox.Text += "Валюты и их коды успешно загружены\n";
                MyLogger.Instance.Info("Валюты и их коды успешно загружены");
            }
            else
            {
                AllDataBox.Text += "Валюты и их коды не загружены\n";
                MyLogger.Instance.Info("Валюты и их коды не загружены");
            }

            if (Addresses.AddressesInitFromTXT(Properties.Settings.Default.PathToAd))
            {
                AllDataBox.Text += "Базы и адреса успешно загружены\n";
                MyLogger.Instance.Info("Базы и адреса успешно загружены");
            }
            else
            {
                AllDataBox.Text += "Базы и адреса не загружены\n";
                MyLogger.Instance.Info("Базы и адреса не загружены");
            }
            string pathToLastFile = Scores.GetLastScoresFilePath();
            if (Scores.ScoresInitFromTXT(pathToLastFile))
            {
                AllDataBox.Text += "Счета загружены из файла " + pathToLastFile;
                MyLogger.Instance.Info("Счета загружены из файла " + pathToLastFile);
            } else {
                AllDataBox.Text += "Счета не загружены!!! Файл: " + pathToLastFile;
                MyLogger.Instance.Info("Счета не загружены!!! Файл: " + pathToLastFile);
            }

            for (int i = 0; i < Scores.All.Count; i++)
            {
                TemplateBuilder.BuildScore(Scores.All[i]);
                AllDataBox.Text += "\nСоздан файл для счета: " + Scores.All[i];
            } 
        }

        private void SettingWindow(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Setting());
        }


    }
}
