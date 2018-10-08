using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Text;



namespace Jacksonville
{
    static class Program
    {
        static void Main()
        {
            /* string path;
             string constr;
             oledbconnection con;
             string selectstatement;
             datatable table;

             path = @"C:\users\ahish n\desktop\work\afs\jacksonville.xlsx";


             constr = "provider=microsoft.ace.oledb.12.0;data source=" + path + ";extended properties=\"excel 12.0 xml;hdr=yes;imex=1\";";

             con = new oledbconnection(constr);
             oledbdataadapter adaptor = new oledbdataadapter(selectstatement, con);
             table = new datatable();
             oledbconnection conn = new oledbconnection(constr);
             oledbcommand cmd = new oledbcommand();
             cmd.connection = conn;
             oledbdataadapter da = new oledbdataadapter(cmd);
             datatable dt = new datatable();
             conn.open();
             datatable dtsheet = conn.getoledbschematable(oledbschemaguid.tables, null);
             string sheetname = dtsheet.rows[0]["table_name"].tostring();
             cmd.commandtext = "select * from [" + sheetname + "]";
             da.selectcommand = cmd;
             da.fill(dt);
             conn.close();*/

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
            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

            DataTable dt = new DataTable();
            dt.TableName = sheetName;

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);

            ds.Tables.Add(dt);
             
        }

        cmd = null;
        conn.Close();
    }

   

        }
    }
}
