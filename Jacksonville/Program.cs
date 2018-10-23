using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Text;
//using  Microsoft.Office.Interop.Excel;
//using Excel;



namespace Jacksonville
{
    static class Program
    {
        static void Main()
        {

            Dictionary<string, string> props = new Dictionary<string, string>();
            
            // XLSX - Excel 2007, 2010, 2012, 2013
            //props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            //props["Extended Properties"] = "Excel 12.0 XML";
            //props["Data Source"] = "C:\\MyExcel.xlsx";

            // XLS - Excel 2003 and Older
            props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            props["Extended Properties"] = "Excel 8.0";
            props["Data Source"] = "C:\\Users\\Ahish N\\Desktop\\WORK\\AFS\\Jacksonville.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            Console.WriteLine("End"+sb.ToString());

            DataSet ds = new DataSet();

            string connectionString = sb.ToString();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {   
            conn.Open();
            Console.WriteLine("Connection Open");
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;

            // Get all Sheets in Excel File
            DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    Console.WriteLine(sheetName);

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    // cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    // DataTable dt = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //OleDbCommand cmd = new OleDbCommand("UPDATE [Sheet1$] SET done='yes' where id=1", oledbConn);
                    // cmd.ExecuteNonQuery();
                }
                // cmd.CommandText = "SELECT * FROM [JacksonvilleS1$] WHERE F6='43402076'";
                // int result= cmd.ExecuteNonQuery();
                 cmd.CommandText = "UPDATE [JacksonvilleS1$] SET F12='12,52,631',F13='1236', F14='5/19',F15='580075', F16='300,000', F17='900,000', F18='6/19', F19='6/20'  WHERE F6='43402076'";

                cmd.CommandText = "UPDATE [JacksonvilleS1$] SET F12='11,52,531',F13='22361', F14='4/19',F15='580075', F16='300,000', F17='900,000', F18='5/17', F19='5/18'  WHERE F6='43539496'";
                //Console.WriteLine(result);
                DataTable dt = new DataTable();
                dt.TableName = "JacksonvilleS1";
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dt.WriteXml(@"D:\AFS.xml");
                System.IO.StringWriter writer = new System.IO.StringWriter();
                dt.WriteXml(writer, XmlWriteMode.IgnoreSchema, false);
                string result = writer.ToString();
                Console.WriteLine(result);
                ds.Tables.Add(dt);
                //ds.con
                cmd = null;
                conn.Close();
            }
        }
    }
}
