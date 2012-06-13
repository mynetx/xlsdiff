using System;
using System.Windows;

namespace xlsdiff
{
    internal class CsvFileConverter : FileConverter
    {
        public event ConversionProgressUpdatedEventHandler ConversionProgressUpdated;

        public new bool Convert(bool overwriteExistingTarget = false)
        {
            if (Source == null)
            {
                throw new MissingFieldException("The source file to convert has not been set.");
            }

            try
            {
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
            }
            finally
            {
            }

            return false;
        }
    }
}