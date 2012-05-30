using System;
using System.Windows;
using System.Threading;

namespace xlsdiff
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _strFile1 = "";
        private string _strFile2 = "";

        public MainWindow()
        {
            //Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("de-DE");
            //Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("de-DE");

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
                this.LblFile1.Text = System.IO.Path.GetFileName(dlg.FileName);
                this.BtnFile2.IsEnabled = true;
                this.GetFileDetails();
            }
        }

        private void BtnFile2Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog 
                          { FileName = "", DefaultExt = ".xls", Filter = "Excel files|*.xls;*.xlsx;*.csv" };
            var result = dlg.ShowDialog();
            if (result == true)
            {
                this._strFile2 = dlg.FileName;
                this.LblFile2.Text = System.IO.Path.GetFileName(dlg.FileName);
                this.GetFileDetails();
            }
        }

        private void GetFileDetails()
        {
            if (this._strFile1 != "" && this._strFile2 != "")
            {
                this.BtnShow.IsEnabled = true;
            }
        }

        private void BtnShowClick(object sender, RoutedEventArgs e)
        {
            this.BtnFile1.IsEnabled = this.BtnFile2.IsEnabled = this.BtnShow.IsEnabled = false;
            this.PrgProgress.Visibility = this.LblProgress.Visibility = Visibility.Visible;
        }
    }
}
