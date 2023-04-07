using NLog;
using System;
using System.Windows;

namespace HelpMe
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Content = new MainPage();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            MyLogger.Instance.Info("Приложение закрыто.");
        }
    }
}
