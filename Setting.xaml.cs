using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace HelpMe
{
    /// <summary>
    /// Логика взаимодействия для Setting.xaml
    /// </summary>
    public partial class Setting : Page
    {
        public Setting()
        {
            InitializeComponent();
            PathToCurrTextBox.Text = Properties.Settings.Default.PathToCurr;
            PathToDataTextBox.Text = Properties.Settings.Default.PathToDatas;
            PathToAddressesTextBox.Text = Properties.Settings.Default.PathToAd;
        }

        private void PathToCurrBtn(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".csv";
            dlg.Filter = "СИЭСВЭ |*.csv";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                PathToCurrTextBox.Text = dlg.FileName;
            }
        }

        private void PathToDataBtn(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Эксельчик |*.xlsx";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                PathToDataTextBox.Text = dlg.FileName;
            }
        }

        private void PathToAddressesBtn(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".txt";
            dlg.Filter = "Текстовичек |*.txt";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                PathToAddressesTextBox.Text = dlg.FileName;
            }
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.PathToCurr = PathToCurrTextBox.Text;
            Properties.Settings.Default.PathToDatas = PathToDataTextBox.Text;
            Properties.Settings.Default.PathToAd = PathToAddressesTextBox.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("Сохранено");
        }

        private void GoBack(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}
