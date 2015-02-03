using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.ServiceProcess;

namespace DFSystem
{
    class ImportExcel
    {
        private string excelFile;
        private string tableName;

        public ImportExcel(string excelFile, string tableName)
        {
            this.excelFile = excelFile;
            this.tableName = tableName;
        }

        public int StartWork(bool needDelete)
        {
            int count = 0;
            count = TransferData(needDelete);
            return count;
        }

        private System.Data.DataTable importFromExcel()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;

            object oMissing = System.Reflection.Missing.Value;//相当null
            workbook = excel.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            worksheet = (Worksheet)workbook.Worksheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range1;
            System.Data.DataTable dt = new System.Data.DataTable();
            // 获得表头，即第一行数据
            for (int i = 0; i < colCount; i++)
            {
                //去掉”主要技术指标“和”主要功能“和”主要附件“这3列
                range1 = worksheet.Range[worksheet.Cells[1, i + 1], worksheet.Cells[1, i + 1]];
                string t = range1.Value2.ToString();
                if (t.Equals("主要技术指标") || t.Equals("主要功能") || t.Equals("主要附件"))
                {
                    continue;
                }
                dt.Columns.Add(range1.Value2.ToString());
            }
   
            // 获得表数据，从第二行开始
            for (int j = 2; j < rowCount; j++)
            {
                DataRow dr = dt.NewRow();
                int k = 0;
                for (int i = 0; i < colCount; i++)
                {
                    range1 = worksheet.Range[worksheet.Cells[j, i + 1], worksheet.Cells[j, i + 1]];
                    object text = range1.Value2;
                    if (text == null)
                    {
                        text = "";
                    }
                    range1 = worksheet.Range[worksheet.Cells[1, i + 1], worksheet.Cells[1, i + 1]];
                    string t = range1.Value2.ToString();
                    if (t.Equals("主要技术指标") || t.Equals("主要功能") || t.Equals("主要附件"))
                    {
                        continue;
                    }
                    dr[k++] = text.ToString();
                }
                dt.Rows.Add(dr);
                Console.WriteLine("read excel:" + rowCount + "-" + j);
            }
            //dataGridView1.DataSource = dt;
            excel.Quit();
            return dt;
        }

        // 向sql server中批量导入数据
        public int TransferData(bool needDelete)     
        {
            string sheetName = "";
            DataSet ds = new DataSet();
            System.Data.DataTable dt = new System.Data.DataTable();
            try    
            {
                //string strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + Server.MapPath("ExcelFiles/MyExcelFile.xls") + ";Extended Properties='Excel 8.0; HDR=Yes; IMEX=1'"; //此连接只能操作Excel2007之前(.xls)文件
                //获取全部数据   
                string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "Data Source=" + excelFile + ";" + ";Extended Properties='Excel 12.0'";     
                OleDbConnection conn = new OleDbConnection(strConn);  
                conn.Open();  
                string strExcel = "";  
                OleDbDataAdapter myCommand = null;
                System.Data.DataTable sheetNames = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                //获取所有的sheet name
                string[] vSheets = new string[sheetNames.Rows.Count];
                string vName = string.Empty;

                //获取所有的sheet names
                for (int i = 0; i < sheetNames.Rows.Count; i++)
                {
                    vSheets[i] = sheetNames.Rows[i][2].ToString().Trim().Replace("$", "");
                    if(vSheets[i].Equals("Sheet1"))
                    {
                        sheetName = vSheets[i];
                    }
                }
                if (sheetName.Equals(""))
                {
                    //获取第一个sheet名称
                    sheetName = vSheets[0];
                }
                strExcel = string.Format("select * from [{0}$] ", sheetName);  
                myCommand = new OleDbDataAdapter(strExcel, strConn);  
                myCommand.Fill(ds, sheetName);

                String strSql = "";
                //如果目标表不存在则创建,excel文件的第一行为列标题,从第二行开始全部都是数据记录   
                if (needDelete)
                {
                    strSql = string.Format("if exists(select * from sysobjects where name='{0}')drop table {0}; create table {0}(", tableName);
                }

                dt = ds.Tables[0];
                
                //去掉”主要技术指标“和”主要功能“和”主要附件“这3列
                if(dt.Columns.Contains("主要技术指标"))
                {
                    dt.Columns.Remove("主要技术指标");
                }
                if (dt.Columns.Contains("主要功能"))
                {
                    dt.Columns.Remove("主要功能");
                }
                if (dt.Columns.Contains("主要附件"))
                {
                    dt.Columns.Remove("主要附件");
                }
                if (dt.Columns.Contains("问题、需求及政策建议"))
                {
                    dt.Columns.Remove("问题、需求及政策建议");
                }
                //删除多余的列
                for (int i = 0; i < dt.Columns.Count; i++ )
                {
                    string v = dt.Columns[i].ToString();
                    int index = v.IndexOf('F');
                    if (index == 0 || index == 1)
                    {
                        dt.Columns.Remove(v.ToString());
                        i--;
                    }
                }

                if (needDelete)
                {
                    foreach (System.Data.DataColumn c in dt.Columns)
                    {
                        strSql += string.Format("[{0}] varchar(2048),", c.ColumnName);
                    }
                    strSql = strSql.Trim(',') + ")";

                    using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(DBHelper.connString))
                    {
                        try
                        {
                            sqlconn.Open();
                            System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                            command.CommandText = strSql;
                            command.ExecuteNonQuery();
                            sqlconn.Close();
                        }
                        catch (Exception err)
                        {
                            Log.RecordLog("ImportExcel.TransferData:" + err.Message);
                            System.Windows.Forms.MessageBox.Show(err.Message);
                            return 0;
                        }
                    }
                }

                //用bcp导入数据      
                //excel文件中列的顺序必须和数据表的列顺序一致，因为数据导入时，是从excel文件的第二行数据开始，不管数据表的结构是什么样的，
                //反正就是第一列的数据会插入到数据表的第一列字段中，第二列的数据插入到数据表的第二列字段中，以此类推，它本身不会去判断要插入的数据是对应数据表中哪一个字段的   
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(DBHelper.connString))     
                {     
                    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(RunningForm.bcp_SqlRowsCopied);     
                    bcp.BatchSize = 3000;//每次传输的行数      
                    bcp.NotifyAfter = 100;//进度提示的行数      
                    bcp.DestinationTableName = tableName;//目标表      
                    bcp.WriteToServer(dt);  
                }     
            }     
            catch (Exception ex)
            {
                Log.RecordLog("ImportExcel.TransferData: " + tableName + ": " + ex.Message); 
                return 0;
            }
            Log.RecordLog("\"" + tableName + "\" 导入成功！共导入" + dt.Rows.Count.ToString() + "条数据！");
            return dt.Rows.Count;
        }

        //进度显示      
        public static void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {
            
        }  
    }
}
