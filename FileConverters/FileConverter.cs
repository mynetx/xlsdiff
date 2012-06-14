using System;
using System.IO;

namespace xlsdiff
{
    internal class FileConverter
    {
        #region Delegates

        public delegate void ConversionProgressUpdatedEventHandler(double dblPercentage);

        #endregion

        private string _strSource;

        private string _strTarget;

        public string Source
        {
            get { return _strSource; }
            set
            {
                if (! File.Exists(value))
                {
                    throw new FileNotFoundException("The file to convert was not found.", value);
                }
                _strSource = value;
            }
        }

        public string Target
        {
            get
            {
                if (_strTarget == null)
                {
                    _strTarget = Path.GetTempFileName();
                    new FileInfo(_strTarget) {Attributes = FileAttributes.Temporary};
                }
                return _strTarget;
            }
            set { _strTarget = value; }
        }

        public FileType TargetType { get; set; }

        public virtual bool Convert(bool overwriteExistingTarget = false)
        {
            throw new NotImplementedException();
        }

        public static string ConvertFile(string strFile, FileType typeFile,
                                         ConversionProgressUpdatedEventHandler fnProgressUpdated)
        {
            switch (typeFile)
            {
                case FileType.Xls:
                    {
                        var objConvert = new XlsFileConverter {Source = strFile, TargetType = FileType.Csv};
                        objConvert.ConversionProgressUpdated += fnProgressUpdated;
                        try
                        {
                            if (objConvert.Convert())
                            {
                                return objConvert.Target;
                            }
                        }
                        catch (Exception)
                        {
                            return null;
                        }
                        break;
                    }
                case FileType.Xlsx:
                    {
                        var objConvert = new XlsxFileConverter {Source = strFile, TargetType = FileType.Csv};
                        objConvert.ConversionProgressUpdated += fnProgressUpdated;
                        try
                        {
                            if (objConvert.Convert())
                            {
                                return objConvert.Target;
                            }
                        }
                        catch (Exception)
                        {
                            return null;
                        }
                        break;
                    }
                case FileType.Csv:
                    {
                        var objConvert = new CsvFileConverter {Source = strFile, TargetType = FileType.Csv};
                        objConvert.ConversionProgressUpdated += fnProgressUpdated;
                        try
                        {
                            if (objConvert.Convert())
                            {
                                return objConvert.Target;
                            }
                        }
                        catch (Exception)
                        {
                            return null;
                        }
                        break;
                    }
            }

            return null;
        }
    }
}