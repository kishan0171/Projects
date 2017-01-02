using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using Newtonsoft.Json;
using Microsoft.Office.Core;
using System.Data;
using System.Data.OleDb;
using System.Linq;


namespace DATA_PROJECT._Webservice
{
    /// <summary>
    /// Summary description for WebService1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    [System.Web.Script.Services.ScriptService]
    public class WebService1 : System.Web.Services.WebService
    {

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod(EnableSession = true)]
        public String GetData()
        {
            string result = String.Empty;

            try
            {
                //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("Test_DataSet.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                //Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                //int rowCount = xlRange.Rows.Count;
                //int colCount = xlRange.Columns.Count;

                //for (int i = 1; i <= rowCount; i++)
                //    {
                //        for (int j = 1; j <= colCount; j++)
                //        {
                //              //MessageBox.Show(xlWorksheet.Cells[i,j].ToString());
                //        }
                //    }

                //string filePath = "../DATA_PROJECT/Test_DataSet.xlsx";
                string filePath = Server.MapPath("Test_DataSet.xlsx");
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath +
                             ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\";";

                var output = new DataSet();
                List<DataName> oblist = new List<DataName>();
                DataName dataObj;

                using (var conn = new OleDbConnection(strConn))
                {
                    conn.Open();

                    var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    foreach (DataRow row in dt.Rows)
                    {
                        string sheet = row["TABLE_NAME"].ToString();
                        var cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                        cmd.CommandType = CommandType.Text;
                        OleDbDataAdapter xlAdapter = new OleDbDataAdapter(cmd);
                        xlAdapter.Fill(output, "School");
                    }

                    //result = JsonConvert.SerializeObject(output);

                    //foreach (DataTable table in output.Tables)
                    //{

                    //    foreach (DataRow dr in table.Rows)
                    //    {
                    //        dataObj = new DataName();
                    //        dataObj
                    //    }
                    //}
                    result = JsonConvert.SerializeObject(output.Tables[0]);

                }
            }
            catch(Exception Ex)
            {
                result =  "ERROR";
            }
           
            //result = JsonConvert.SerializeObject(orderBookDetails);
            return result;
        }


        [WebMethod(EnableSession = true)]
        public String GetDataDup()
        {
            string result = String.Empty;

            try
            {
                //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("Test_DataSet.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                //Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                //int rowCount = xlRange.Rows.Count;
                //int colCount = xlRange.Columns.Count;

                //for (int i = 1; i <= rowCount; i++)
                //    {
                //        for (int j = 1; j <= colCount; j++)
                //        {
                //              //MessageBox.Show(xlWorksheet.Cells[i,j].ToString());
                //        }
                //    }

                //string filePath = "../DATA_PROJECT/Test_DataSet.xlsx";
                string filePath = Server.MapPath("Test_DataSet.xlsx");
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath +
                             ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\";";

                var output = new DataSet();
                List<DataName> oblist = new List<DataName>();
                DataName dataObj;

                using (var conn = new OleDbConnection(strConn))
                {
                    conn.Open();

                    var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    foreach (DataRow row in dt.Rows)
                    {
                        string sheet = row["TABLE_NAME"].ToString();
                        var cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                        cmd.CommandType = CommandType.Text;
                        OleDbDataAdapter xlAdapter = new OleDbDataAdapter(cmd);
                        xlAdapter.Fill(output, "School");
                    }

                    //result = JsonConvert.SerializeObject(output);

                    foreach (DataTable table in output.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            dataObj = new DataName();
                            dataObj.F1 = Convert.ToString(dr["F1"]);
                            dataObj.F2 = Convert.ToString(dr["F2"]);
                            dataObj.F3 = Convert.ToString(dr["F3"]);
                            dataObj.F4 = Convert.ToString(dr["F4"]);
                            dataObj.F5 = Convert.ToString(dr["F5"]);
                            dataObj.F6 = Convert.ToString(dr["F6"]);
                            dataObj.F7 = Convert.ToString(dr["F7"]);
                            dataObj.F8 = Convert.ToString(dr["F8"]);
                            oblist.Add(dataObj);
                        }
                    }
                    oblist.RemoveAt(0);
                    oblist = oblist.GroupBy(d => new { d.F2, d.F3, d.F4 , d.F5})
                   .Select(grp => grp.First())
                   .ToList();

                    result = JsonConvert.SerializeObject(oblist);

                }
            }
            catch (Exception Ex)
            {
                result = "ERROR";
            }

            //result = JsonConvert.SerializeObject(orderBookDetails);
            return result;
        }


        [WebMethod(EnableSession = true)]
        public String Update()
        {
            string result = String.Empty;

            try
            {
                //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("Test_DataSet.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                //Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                //int rowCount = xlRange.Rows.Count;
                //int colCount = xlRange.Columns.Count;

                //for (int i = 1; i <= rowCount; i++)
                //    {
                //        for (int j = 1; j <= colCount; j++)
                //        {
                //              //MessageBox.Show(xlWorksheet.Cells[i,j].ToString());
                //        }
                //    }

                //string filePath = "../DATA_PROJECT/Test_DataSet.xlsx";
                string filePath = Server.MapPath("Test_DataSet.xlsx");
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath +
                             ";Extended Properties=\"Excel 12.0;HDR=No;\"";

                var output = new DataSet();
                List<DataName> oblist = new List<DataName>();
                DataName dataObj;
                //con.Open();

                using (var conn = new OleDbConnection(strConn))
                {
                    conn.Open();
                    System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                    myCommand.Connection = conn;

                    var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, "Sheet1$", null });
                    foreach (DataRow row in dt.Rows)
                    {
                        //conn.Open();
                        //string sheet = row["TABLE_NAME"].ToString();
                        //var columnNameColumn = row["COLUMN_NAME"].ToString(); ;
                        string sql = null;
                        //sql = "Update [" + sheet + "] set F2 = 'New Name' where F1='1'";
                        sql = "Update [Sheet1$] set [F3] = 'New Name' where [F1]='1'";
                        myCommand.CommandText = sql;
                        myCommand.ExecuteNonQuery();
                    }

                    
                    conn.Close();
                }             
            }
            catch (Exception Ex)
            {
                result = "ERROR";
            }

            //result = JsonConvert.SerializeObject(orderBookDetails);
            return result;
        }

    }





    public class DataName
    {
        public string F1 { get; set; }
        public string F2 { get; set; }
        public string F3 { get; set; }
        public string F4 { get; set; }
        public string F5 { get; set; }
        public string F6 { get; set; }
        public string F7 { get; set; }
        public string F8 { get; set; }
    }



}
