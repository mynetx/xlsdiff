using System;
using System.IO;
using System.Linq;

namespace xlsdiff
{
    class FileTypeDetector
    {
        private readonly string _strFile;

        public FileTypeDetector(string strFile)
        {
            this._strFile = strFile;
        }

        public FileType Detect()
        {
            if (this.DetectXls())
            {
                return FileType.Xls;
            }
            if (this.DetectXlsx())
            {
                return FileType.Xlsx;
            }
            if (this.DetectCsv())
            {
                return FileType.Csv;
            }
            throw new FileFormatException(new Uri(this._strFile));
        }

        private bool DetectXls()
        {
            // check file extension first
            if (! this._strFile.EndsWith(".xls", ignoreCase: true, culture: null))
            {
                return false;
            }

            // the XLS magic bytes
            byte[] bytMagic = new byte[] { 
                0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x3B };

            // open file in question
            var resFile = File.Open(this._strFile, FileMode.Open);
            BinaryReader resFileReader = new BinaryReader(resFile);

            // try to read the magic bytes
            byte[] bytInFile = resFileReader.ReadBytes(bytMagic.Length);

            resFileReader.Close();
            resFile.Close();

            return bytInFile.SequenceEqual(bytMagic);
        }

        private bool DetectXlsx()
        {
            // check file extension first
            if (!this._strFile.EndsWith(".xlsx", ignoreCase: true, culture: null))
            {
                return false;
            }

            // the ZIP magic bytes
            byte[] bytMagic = new byte[] { 0x50, 0x4B, 0x03, 0x04 };

            // open file in question
            var resFile = File.Open(this._strFile, FileMode.Open);
            BinaryReader resFileReader = new BinaryReader(resFile);

            // try to read the magic bytes
            byte[] bytInFile = resFileReader.ReadBytes(bytMagic.Length);

            resFileReader.Close();
            resFile.Close();

            return bytInFile.SequenceEqual(bytMagic);
        }

        private bool DetectCsv()
        {
            // check file extension first
            if (!this._strFile.EndsWith(".csv", ignoreCase: true, culture: null))
            {
                return false;
            }

            // try to read the first line
            string strLine = new StreamReader(this._strFile).ReadLine();
            if (strLine != null && strLine.IndexOfAny(new[] {',', ';'}) > -1)
            {
                return true;
            }
            return false;
        }
    }
}
