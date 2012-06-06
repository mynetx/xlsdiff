using System;
using System.IO;

namespace xlsdiff
{
    public class FileConverter
    {
        public delegate void ConversionProgressUpdatedEventHandler(int intPercentage);

        private string _strSource;

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

        private string _strTarget;
        private bool _boolTargetIsTemp;

        public string Target
        {
            get
            {
                if (_strTarget == null)
                {
                    _strTarget = Path.GetTempFileName();
                    new FileInfo(_strTarget) {Attributes = FileAttributes.Temporary};
                    _boolTargetIsTemp = true;
                }
                return _strTarget;
            }
            set
            {
                _strTarget = value;
                _boolTargetIsTemp = false;
            }
        }

        public FileType TargetType { get; set; }

        public virtual bool Convert(bool overwriteExistingTarget = false)
        {
            throw new NotImplementedException();
        }

        ~FileConverter()
        {
            if (_boolTargetIsTemp)
            {
                try
                {
                    File.Delete(Target);
                }
                catch
                {
                    ; // well, if the tempfile cannot be deleted, then leave it
                }
            }
        }
    }
}