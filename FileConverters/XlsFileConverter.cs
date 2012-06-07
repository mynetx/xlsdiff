using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows;

namespace xlsdiff
{
    public class XlsFileConverter : FileConverter
    {
        public event ConversionProgressUpdatedEventHandler ConversionProgressUpdated;

        public new bool Convert(bool overwriteExistingTarget = false)
        {
            if (Source == null)
            {
                throw new MissingFieldException("The source file to convert has not been set.");
            }

            // create ADO.NET connection to the XLS file
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Source +
                             ";Extended Properties=\" Excel.0;HDR=Yes;IMEX=1\"";
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

                objCommand = new OleDbCommand("SELECT * FROM [" + strSheet + "$]", resConn)
                                 {CommandType = CommandType.Text};
                resTarget = new StreamWriter(Target);

                objAdapter = new OleDbDataAdapter(objCommand);
                var objTable = new DataTable();
                objAdapter.Fill(objTable);

                int intRows = objTable.Rows.Count;

                for (int intRow = 0; intRow < intRows; intRow++)
                {
                    if (intRow % 10 == 0)
                    {
                        ConversionProgressUpdated((int) (intRow / (float) intRows * 100));
                    }

                    string rowString = "";
                    for (int y = 0; y < objTable.Columns.Count; y++)
                    {
                        rowString += "\"" + objTable.Rows[intRow][y] + "\",";
                    }
                    resTarget.WriteLine(rowString);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
            }
            finally
            {
                if (resConn.State == ConnectionState.Open)
                    resConn.Close();
                resConn.Dispose();
                objCommand.Dispose();
                objAdapter.Dispose();
                resTarget.Close();
                resTarget.Dispose();
            }

            return false;
        }
    }
}