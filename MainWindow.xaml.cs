using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Win32;

namespace xlsdiff
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private const string StrProductCode = @"{58532E1E-8DB2-4E99-A883-BA84E4A0FC60}";
        internal static bool BoolCancel;
        private static readonly string StrDiffFolder = AppDomain.CurrentDomain.BaseDirectory + @"diff\";
        private static readonly string StrPhpInterpreter = StrDiffFolder + @"php.exe";
        private static readonly string StrDiffScript = StrDiffFolder + @"diff.php";
        private static readonly string[] ArrComponentFiles = {StrPhpInterpreter, StrDiffScript};
        private double _dblPercentageFile1;
        private string _strCurrentFile = "";
        private string _strFile1 = "";
        private string _strFile2 = "";
        private FileType _typeFile1;
        private FileType _typeFile2;

        public MainWindow()
        {
            // check if diff parser exists
            RepairIfNecessary();
            InitializeComponent();
        }

        private void RepairIfNecessary()
        {
            foreach (string strFile in ArrComponentFiles.Where(strFile => !File.Exists(strFile)))
            {
                MessageBox.Show(Resource.Resource.MsgAppComponentMissing, "xlsdiff", MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                Process.Start(@"msiexec.exe", @"/passive /norestart /fa " + StrProductCode);
                Hide();
            }
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            BtnFile1.Focus();
        }

        private void BtnShowClick(object sender, RoutedEventArgs e)
        {
            // disable the controls as we're working now
            BtnFile1.IsEnabled = BtnFile2.IsEnabled = BtnShow.IsEnabled = false;
            DoEvents();

            // show some progress
            PrgProgress.Value = .5;
            LblProgress.Text = Resource.Resource.LblPreparing;
            PrgProgress.Visibility = LblProgress.Visibility = BtnCancel.Visibility = Visibility.Visible;
            DoEvents();

            // calculate whether the first or second file is larger
            var objInfoFile1 = new FileInfo(_strFile1);
            var objInfoFile2 = new FileInfo(_strFile2);
            _dblPercentageFile1 = objInfoFile1.Length/Convert.ToDouble(objInfoFile1.Length + objInfoFile2.Length);

            // file 1
            _strCurrentFile = "1";
            string strTarget1 = FileConverter.ConvertFile(_strFile1, _typeFile1, ConversionProgressHandler);

            // file 2
            string strTarget2 = "";
            if (!BoolCancel)
            {
                _strCurrentFile = "2";
                strTarget2 = FileConverter.ConvertFile(_strFile2, _typeFile2, ConversionProgressHandler);
            }

            if (!BoolCancel)
            {
                PrgProgress.Value += 5;
                LblProgress.Text = Resource.Resource.LblComparingFiles;
                DoEvents();

                // get temporary file name
                string strDiffFile = Path.GetTempPath() + Guid.NewGuid() + ".html";

                // diff the two files
                var objDiffProcess = new Process
                                         {
                                             StartInfo =
                                                 {
                                                     FileName = StrPhpInterpreter,
                                                     WorkingDirectory = Path.GetDirectoryName(StrPhpInterpreter),
                                                     Arguments =
                                                         string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\" \"{5}\"",
                                                                       "diff.php",
                                                                       strTarget1,
                                                                       strTarget2,
                                                                       strDiffFile,
                                                                       Resource.Resource.LblOld,
                                                                       Resource.Resource.LblNew),
                                                     CreateNoWindow = true,
                                                     WindowStyle = ProcessWindowStyle.Hidden
                                                 }
                                         };
                objDiffProcess.Start();

                while (! objDiffProcess.HasExited)
                {
                    DoEvents();
                    if (BoolCancel)
                    {
                        objDiffProcess.Kill();
                    }
                }
                // check process exit code
                // launch diff result in default browser
                if (objDiffProcess.ExitCode == 0 && File.Exists(strDiffFile))
                {
                    PrgProgress.Value = 95;
                    LblProgress.Text = Resource.Resource.LblOpeningResult;
                    DoEvents();

                    Process.Start(strDiffFile);

                    PrgProgress.Value = 100;
                    LblProgress.Text = "";
                    DoEvents();
                }
            }

            // clean up the converted temp files
            try
            {
                File.Delete(strTarget1);
                File.Delete(strTarget2);
            }
            catch
            {
            }

            BoolCancel = false;
            BtnFile1.IsEnabled = BtnFile2.IsEnabled = BtnShow.IsEnabled = true;
            PrgProgress.Visibility = LblProgress.Visibility = BtnCancel.Visibility = Visibility.Hidden;
        }

        private void BtnCancelClick(object sender, RoutedEventArgs e)
        {
            BoolCancel = true;
            BtnFile1.IsEnabled = BtnFile2.IsEnabled = BtnShow.IsEnabled = true;
            PrgProgress.Visibility = LblProgress.Visibility = BtnCancel.Visibility = Visibility.Hidden;
        }

        private void ConversionProgressHandler(double dblPercentage)
        {
            double dblStartPercentage = 1.0;
            double dblProgressTotal = 50*_dblPercentageFile1;
            if (_strCurrentFile == "2")
            {
                dblStartPercentage += dblProgressTotal;
                dblProgressTotal = 50 - dblProgressTotal;
            }
            PrgProgress.Value = dblStartPercentage + dblProgressTotal*(dblPercentage/100.0);
            LblProgress.Text = string.Format(Resource.Resource.LblReadingFileXPercent, _strCurrentFile,
                                             (int) dblPercentage);
            DoEvents();
        }

        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        private static void DoEvents()
        {
            var frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
                                                     new DispatcherOperationCallback(ExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        private static object ExitFrame(object f)
        {
            ((DispatcherFrame) f).Continue = false;
            return null;
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
            if (! BtnFile1.IsEnabled)
            {
                if (
                    MessageBox.Show(Resource.Resource.MsgWorkingConfirmExit, "xlsdiff", MessageBoxButton.YesNo,
                                    MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                {
                    return;
                }
                Environment.Exit(1);
            }
            Close();
        }

        #endregion

        #region File pickers

        private void BtnFile1Click(object sender, RoutedEventArgs e)
        {
            string strFile = AskFile(sender, _strFile2);
            if (strFile == null) return;
            _strFile1 = strFile;
            LblFile1.Text = Path.GetFileNameWithoutExtension(strFile);
            LblFile1Ext.Text = Path.GetExtension(strFile);
            if (_strFile2 != "")
            {
                BtnShow.IsEnabled = true;
            }
        }

        private void BtnFile2Click(object sender, RoutedEventArgs e)
        {
            string strFile = AskFile(sender, _strFile1);
            if (strFile == null) return;
            _strFile2 = strFile;
            LblFile2.Text = Path.GetFileNameWithoutExtension(strFile);
            LblFile2Ext.Text = Path.GetExtension(strFile);
            if (_strFile1 != "")
            {
                BtnShow.IsEnabled = true;
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
                              Filter = Resource.Resource.DlgExcelFiles + "|*.xls;*.xlsx"
                          };
            do
            {
                bool? result = dlg.ShowDialog();

                if (result == false)
                {
                    return null;
                }
                // check if same file as other file
                if (dlg.FileName == strOtherFile)
                {
                    MessageBox.Show(Resource.Resource.MsgFileTwice,
                                    "xlsdiff", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            } while (dlg.FileName == strOtherFile);

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
    }
}