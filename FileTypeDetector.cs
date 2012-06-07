using System;
using System.IO;
using System.Linq;

namespace xlsdiff
{
    internal class FileTypeDetector
    {
        private readonly string _strFile;

        public FileTypeDetector(string strFile)
        {
            _strFile = strFile;
        }

        public FileType Detect()
        {
            if (DetectXls())
            {
                return FileType.Xls;
            }
            if (DetectXlsx())
            {
                return FileType.Xlsx;
            }
            if (DetectCsv())
            {
                return FileType.Csv;
            }
            throw new FileFormatException(new Uri(_strFile));
        }

        private bool DetectXls()
        {
            // check file extension first
            if (! _strFile.EndsWith(".xls", ignoreCase: true, culture: null))
            {
                return false;
            }

            // the XLS magic bytes
            var bytMagic = new byte[]
                               {
                                   0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
                                   0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                   0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x3B
                               };

            // open file in question
            FileStream resFile = File.Open(_strFile, FileMode.Open);
            var resFileReader = new BinaryReader(resFile);

            // try to read the magic bytes
            byte[] bytInFile = resFileReader.ReadBytes(bytMagic.Length);

            resFileReader.Close();
            resFile.Close();

            return bytInFile.SequenceEqual(bytMagic);
        }

        private bool DetectXlsx()
        {
            // check file extension first
            if (!_strFile.EndsWith(".xlsx", ignoreCase: true, culture: null))
            {
                return false;
            }

            // the ZIP magic bytes
            var bytMagic = new byte[] {0x50, 0x4B, 0x03, 0x04};

            // open file in question
            FileStream resFile = File.Open(_strFile, FileMode.Open);
            var resFileReader = new BinaryReader(resFile);

            // try to read the magic bytes
            byte[] bytInFile = resFileReader.ReadBytes(bytMagic.Length);

            resFileReader.Close();
            resFile.Close();

            return bytInFile.SequenceEqual(bytMagic);
        }

        private bool DetectCsv()
        {
            // check file extension first
            if (!_strFile.EndsWith(".csv", ignoreCase: true, culture: null))
            {
                return false;
            }

            // try to read the first line
            string strLine = new StreamReader(_strFile).ReadLine();
            if (strLine != null && strLine.IndexOfAny(new[] {',', ';'}) > -1)
            {
                return true;
            }
            return false;
        }
    }
}