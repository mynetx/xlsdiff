namespace xlsdiff
{
    public class XlsxFileConverter : XlsFileConverter
    {
        public event ConversionProgressUpdatedEventHandler ConversionProgressUpdated;

        protected string GetConnectionString()
        {
            return string.Format("Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\";", Source);
        }
    }
}
