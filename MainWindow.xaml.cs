using System.IO;
using System.Windows;
using System.Resources;
using System.Windows.Controls;
using System.Windows.Input;

namespace xlsdiff
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _strFile1 = "";
        private string _strFile2 = "";
        private FileType _typeFile1;
        private FileType _typeFile2;

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Window titlebar, base buttons
        /// <summary>
        /// Allow dragging the window by moving its titlebar
        /// </summary>
        private void TitlebarMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
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
        #endregion

        #region File pickers
        private void BtnFile1Click(object sender, RoutedEventArgs e)
        {
            string strFile = this.AskFile(sender);
            if (strFile != null)
            {
                this._strFile1 = strFile;
                this.LblFile1.Text = Path.GetFileNameWithoutExtension(strFile);
                this.LblFile1Ext.Text = Path.GetExtension(strFile);
                if (this._strFile2 != "")
                {
                    this.BtnShow.IsEnabled = true;
                }
            }
            else
            {
                this._strFile1 = this.LblFile1.Text = this.LblFile1Ext.Text = "";
                this.BtnShow.IsEnabled = false;
            }
        }

        private void BtnFile2Click(object sender, RoutedEventArgs e)
        {
            string strFile = this.AskFile(sender);
            if (strFile != null)
            {
                this._strFile2 = strFile;
                this.LblFile2.Text = Path.GetFileNameWithoutExtension(strFile);
                this.LblFile2Ext.Text = Path.GetExtension(strFile);
                if (this._strFile1 != "")
                {
                    this.BtnShow.IsEnabled = true;   
                }
            }
            else
            {
                this._strFile2 = this.LblFile2.Text = this.LblFile2Ext.Text = "";
                this.BtnShow.IsEnabled = false;
            }
        }

        /// <summary>
        /// Asks for an Excel file name and validates the file type
        /// </summary>
        private string AskFile(object sender)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog {
                FileName = "",
                DefaultExt = ".xls",
                Filter = Resource.Resource.DlgExcelFiles + "|*.xls;*.xlsx;*.csv"
            };
            var result = dlg.ShowDialog();

            if (result == false)
            {
                return null;
            }

            // detect file type
            FileType typeFile;
            try
            {
                typeFile = new FileTypeDetector(dlg.FileName).Detect();
            }
            catch (FileFormatException)
            {
                MessageBox.Show(
                    string.Format(Resource.Resource.MsgFileFormat, Path.GetFileName(dlg.FileName)),
                    "xlsdiff", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return null;
            }
            switch (((Button) sender).Name)
            {
                case "BtnFile1":
                    this._typeFile1 = typeFile;
                    break;
                case "BtnFile2":
                    this._typeFile2 = typeFile;
                    break;
            }
            return dlg.FileName;
        }
        #endregion

        private void BtnShowClick(object sender, RoutedEventArgs e)
        {
            // disable the controls as we're working now
            this.BtnFile1.IsEnabled = this.BtnFile2.IsEnabled = this.BtnShow.IsEnabled = false;

            // show some progress
            this.PrgProgress.Visibility = this.LblProgress.Visibility = Visibility.Visible;
            string strReadingFileX = Resource.Resource.LblReadingFileX;
            this.LblProgress.Text = string.Format(strReadingFileX, 1);
            this.PrgProgress.Value = 50;
        }
    }
}