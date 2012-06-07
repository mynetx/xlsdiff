using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace xlsdiff
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _strCurrentFile = "";
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
            DragMove();
        }

        private void BtnMinimizeClick(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void BtnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion

        #region File pickers

        private void BtnFile1Click(object sender, RoutedEventArgs e)
        {
            string strFile = AskFile(sender, _strFile2);
            if (strFile != null)
            {
                _strFile1 = strFile;
                LblFile1.Text = Path.GetFileNameWithoutExtension(strFile);
                LblFile1Ext.Text = Path.GetExtension(strFile);
                if (_strFile2 != "")
                {
                    BtnShow.IsEnabled = true;
                }
            }
            else
            {
                _strFile1 = LblFile1.Text = LblFile1Ext.Text = "";
                BtnShow.IsEnabled = false;
            }
        }

        private void BtnFile2Click(object sender, RoutedEventArgs e)
        {
            string strFile = AskFile(sender, _strFile1);
            if (strFile != null)
            {
                _strFile2 = strFile;
                LblFile2.Text = Path.GetFileNameWithoutExtension(strFile);
                LblFile2Ext.Text = Path.GetExtension(strFile);
                if (_strFile1 != "")
                {
                    BtnShow.IsEnabled = true;
                }
            }
            else
            {
                _strFile2 = LblFile2.Text = LblFile2Ext.Text = "";
                BtnShow.IsEnabled = false;
            }
        }

        /// <summary>
        /// Asks for an Excel file name and validates the file type
        /// </summary>
        private string AskFile(object sender, string strOtherFile)
        {
            var dlg = new OpenFileDialog
                          {
                              FileName = "",
                              DefaultExt = ".xls",
                              Filter = Resource.Resource.DlgExcelFiles + "|*.xls;*.xlsx;*.csv"
                          };
            bool? result = dlg.ShowDialog();

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
                MessageBox.Show(string.Format(Resource.Resource.MsgFileFormat, Path.GetFileName(dlg.FileName)),
                                "xlsdiff", MessageBoxButton.OK, MessageBoxImage.Information);
                return null;
            }
            // check if same file as other file
            if (dlg.FileName == strOtherFile)
            {
                MessageBox.Show(Resource.Resource.MsgFileTwice,
                                "xlsdiff", MessageBoxButton.OK, MessageBoxImage.Information);
                return null;
            }

            switch (((Button) sender).Name)
            {
                case "BtnFile1":
                    _typeFile1 = typeFile;
                    break;
                case "BtnFile2":
                    _typeFile2 = typeFile;
                    break;
            }
            return dlg.FileName;
        }

        #endregion

        private void BtnShowClick(object sender, RoutedEventArgs e)
        {
            // disable the controls as we're working now
            BtnFile1.IsEnabled = BtnFile2.IsEnabled = BtnShow.IsEnabled = false;

            // show some progress
            PrgProgress.Value = 5;
            LblProgress.Text = Resource.Resource.LblPreparing;
            PrgProgress.Visibility = LblProgress.Visibility = Visibility.Visible;

            // file 1
            _strCurrentFile = "1";
            FileConverter.ConvertFile(_strFile1, _typeFile1, ConversionProgressHandler);

            _strCurrentFile = "2";
            FileConverter.ConvertFile(_strFile2, _typeFile2, ConversionProgressHandler);

            MessageBox.Show("done");
        }

        private void ConversionProgressHandler(int intPercentage)
        {
            int intStartPercentage = _strCurrentFile == "1" ? 5 : 30;
            PrgProgress.Value = intStartPercentage + 25*(intPercentage/100);
            LblProgress.Text = string.Format(Resource.Resource.LblReadingFileXPercent, _strCurrentFile, intPercentage);
            MessageBox.Show(_strCurrentFile + intPercentage.ToString(CultureInfo.InvariantCulture));
        }
    }
}
