using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wwc.MsExcel
{
    public class MsExcelTable
    {
        public static DataSet ImportExcelXLS(string filename)
        {
            DataSet output = new DataSet();
            try
            {
                using (OleDbConnection conn = new OleDbConnection(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    filename +
                    ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';"))
                {
                    conn.Open();
                    DataTable dt = conn.GetOleDbSchemaTable(
                        OleDbSchemaGuid.Tables,
                        new object[] { null, null, null, "TABLE" });

                    foreach (DataRow row in dt.Rows)
                    {
                        string sheet = row["TABLE_NAME"].ToString();

                        OleDbCommand cmd = new OleDbCommand(
                            "SELECT * FROM [" + sheet + "]",
                            conn);

                        cmd.CommandType = CommandType.Text;
                        DataTable outputTable = new DataTable(sheet);
                        output.Tables.Add(outputTable);
                        new OleDbDataAdapter(cmd).Fill(outputTable);

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            return output;
        }
    }
}
