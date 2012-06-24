using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows;

namespace xlsdiff
{
    internal class XlsFileConverter : FileConverter
    {
        public event ConversionProgressUpdatedEventHandler ConversionProgressUpdated;

        protected virtual string GetConnectionString()
        {
            return
                string.Format(
                    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";",
                    Source);
        }

        public new bool Convert(bool overwriteExistingTarget = false)
        {
            bool boolSuccess = false;

            if (Source == null)
            {
                throw new MissingFieldException("The source file to convert has not been set.");
            }

            // create ADO.NET connection to the XLS file
            string strConn = GetConnectionString();
            OleDbConnection resConn = null;
            StreamWriter resTarget = null;
            OleDbCommand objCommand = null;
            OleDbDataAdapter objAdapter = null;
            try
            {
                resConn = new OleDbConnection(strConn);
                resConn.Open();

                // load name of first worksheet
                string strSheet = resConn.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
                if (strSheet == "_xlnm#Print_Titles")
                {
                    strSheet = resConn.GetSchema("Tables").Rows[1]["TABLE_NAME"].ToString();
                }

                objCommand = new OleDbCommand("SELECT * FROM [" + strSheet + "]", resConn)
                                 {CommandType = CommandType.Text};
                resTarget = new StreamWriter(Target);

                objAdapter = new OleDbDataAdapter(objCommand);
                var objTable = new DataTable();
                objAdapter.Fill(objTable);

                int intRows = objTable.Rows.Count;

                for (int intRow = 0; intRow < intRows; intRow++)
                {
                    if (intRow%10 == 0)
                    {
                        ConversionProgressUpdated(intRow/(double) intRows*100);
                        // check if we are to cancel the operation
                        if (MainWindow.BoolCancel)
                        {
                            throw new Exception();
                        }
                    }

                    string rowString = "";
                    for (int y = 0; y < objTable.Columns.Count; y++)
                    {
                        rowString += "\"" + objTable.Rows[intRow][y] + "\";";
                    }
                    // remove last semicolon in line
                    if (rowString.Length > 0)
                    {
                        rowString = rowString.Substring(0, rowString.Length - 1);
                    }
                    resTarget.WriteLine(rowString);
                }
                ConversionProgressUpdated(100);
                boolSuccess = true;
            }
            catch (Exception exc)
            {
                if (!MainWindow.BoolCancel)
                {
                    MessageBox.Show(exc.ToString());
                    throw;
                }
            }
            finally
            {
                if (resConn.State == ConnectionState.Open)
                {
                    resConn.Close();
                }
                resConn.Dispose();
                objCommand.Dispose();
                objAdapter.Dispose();
                resTarget.Close();
                resTarget.Dispose();
            }
            return boolSuccess;
        }
    }
}