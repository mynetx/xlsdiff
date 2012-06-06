using System;

namespace xlsdiff
{
    public class XlsFileConverter : FileConverter
    {
        public event ConversionProgressUpdatedEventHandler ConversionProgressUpdated;
        
        public new bool Convert(bool overwriteExistingTarget = false)
        {
            if (this.Source == null)
            {
                throw new MissingFieldException("The source file to convert has not been set.");
            }

            // try to read the file
            ConversionProgressUpdated(10);

            return false;
        }
    }
}