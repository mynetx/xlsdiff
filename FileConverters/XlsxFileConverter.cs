namespace xlsdiff
{
    internal class XlsxFileConverter : XlsFileConverter
    {
        protected override string GetConnectionString()
        {
            return
                string.Format(
                    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\";",
                    Source);
        }
    }
}