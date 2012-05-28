using System;
using System.Windows;

namespace xlsdiff
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _strFile1 = "";
        private string _strFile2;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void TitlebarMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void BtnMinimizeClick(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void BtnCloseClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnFile1Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
                          {FileName = "", DefaultExt = ".xls", Filter = "Excel files|*.xls;*.xlsx;*.csv"};
            var result = dlg.ShowDialog();
            if (result == true)
            {
                this._strFile1 = dlg.FileName;
                this.BtnFile2.IsEnabled = true;
                this.GetFileDetails();
            }
        }

        private void BtnFile2Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { FileName = "", DefaultExt = ".xls", Filter = "Excel files|*.xls;*.xlsx;*.csv" };
            var result = dlg.ShowDialog();
            if (result == true)
            {
                this._strFile2 = dlg.FileName;
                this.GetFileDetails();
            }
        }

        private void GetFileDetails()
        {
        }

        private void BtnShowClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Show");
        }
    }
}
