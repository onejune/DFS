using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Data.SqlClient;
using System.Diagnostics;

namespace DFSystem
{
    public partial class RunningForm : Form
    {
        //计算程序执行时间
        System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();

        private string excelFile;  //excel文件名
        private string groupName;  //分类名称
        private List<string> displayedColumns;
        private Form lastForm;  //上一个窗体
        private int fromValue;
        private int toValue;
        float[] columnsWidth;

        public RunningForm(Form lastForm, string excelFile, List<string> displayedColumns, string groupName, int fromValue, int toValue, float[] columnsWidth)
        {
            InitializeComponent();
            this.lastForm = lastForm;
            this.excelFile = excelFile;
            this.displayedColumns = displayedColumns;
            this.groupName = groupName;
            SharedData.isCompleted = false;
            this.fromValue = fromValue;
            this.toValue = toValue;
            this.columnsWidth = columnsWidth;
            SharedData.excelFile = excelFile;
        }

        public void run()
        {
            btnExcape.Enabled = false;
            Thread th = new Thread(new ThreadStart(StartRunning));
            th.Start();
            //th.Join();
        }

        private void StartRunning()
        {
            oTime.Start();

            //过滤数据，存放到temp表中
            GetFilteredDataIntoNewTable();

            GetSumRecord();
            SharedData.isReady = true;
            //将数据从sql server按类别写入到word
            OutputWord ow = new OutputWord(displayedColumns, groupName, this, fromValue, toValue, columnsWidth);
            try
            {
                ow.CreateWordFile();
            }
            catch (Exception err)
            {
                Log.RecordLog(err.ToString());
                Environment.Exit(0);
            }

            oTime.Stop();
            long time2 = oTime.ElapsedMilliseconds / 1000;
            decimal dd = Math.Round((decimal)time2 / 60, 2);
            Log.RecordLog("按分类生成word成功，耗时 " + dd.ToString() + " 分钟！ ");

            decimal d = Math.Round((decimal)time2 / 60, 2);

            Log.RecordLog("全部工作结束！共耗时：" + d.ToString() + " 分钟！");
            SharedData.sequenceNumber = SharedData.sumRows;
            if (DialogResult.OK == MessageBox.Show(@"最终生成的word文件为：" + SharedData.fileName, "导出word成功！"))
            {
                Environment.Exit(0);
            }

            //MessageBox.Show("请查看错误日志！", "哦活，出错了！");
        }

        //将用户筛选之后的数据写入新表temp
        private void GetFilteredDataIntoNewTable()
        {
            string sql = "";
            string sharedPattern = "";
            sharedPattern = GetSharedPattern();
            string sheetName = "temp";
            sql = string.Format("if exists(select * from sysobjects where name = '{0}')drop table {0}; ", sheetName);   //以sheetName为表名   

            if (fromValue == 0 && toValue == 0)
            {
                sql += @"select * into temp from Sheet1 where (" + sharedPattern + ") and 上级行政主管部门 != '公安部'";
            }
            else if (fromValue == 0 && toValue != 0)
            {
                sql += @"select * into temp from Sheet1 where (" + sharedPattern + ") and 上级行政主管部门 != '公安部' and (cast(原值 AS float)<=" + toValue + ")";
            }
            else if (fromValue != 0 && toValue == 0)
            {
                sql += @"select * into temp from Sheet1 where (" + sharedPattern + ") and 上级行政主管部门 != '公安部' and (cast(原值 AS float)>=" + fromValue + ")";
            }
            else if (fromValue != 0 && toValue != 0)
            {
                sql += @"select * into temp from Sheet1 where (" + sharedPattern + ") and 上级行政主管部门 != '公安部' and (cast(原值 AS float) between " + fromValue + " and " + toValue + ")";
            }

            if (DBHelper.only211)
            {
                sql += " and 单位名称 in (select 学校 from [DFS].[dbo].[211名单])";
            }
            else if (DBHelper.only985)
            {
                sql += " and 单位名称 in (select 学校 from [DFS].[dbo].[985名单])";
            }

            if (groupName.Equals("基地类型"))
            {
                sql += " and 基地级别='国家级' ";
            }

            if (DBHelper.unselectedDeviceTypeList.Count != 0)
            {
                string deviceType = "";
                for (int i = 0; i < DBHelper.unselectedDeviceTypeList.Count; i++)
                {
                    if (DBHelper.unselectedDeviceTypeList[i] != null)
                    {
                        deviceType += "'";
                        deviceType += DBHelper.unselectedDeviceTypeList[i];
                        deviceType += "',";
                    }
                }
                deviceType = deviceType.Substring(0, deviceType.Length - 1);
                sql += " and (仪器分类大类 NOT IN(" + deviceType + "))";
            }
            try
            {
                SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
                sqlConn.Open();
                SqlCommand command = new SqlCommand(sql, sqlConn);
                command.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
                Log.RecordLog(err.ToString());
            }
        }

        private void GetSumRecord()
        {
            string sql = "";
            sql = @"select count(*) as 'sum' from temp";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            SqlCommand command = new SqlCommand(sql, sqlConn);
            SqlDataReader reader = command.ExecuteReader();

            if (reader.Read())
            {
                SharedData.sumRows = int.Parse(reader["sum"].ToString());
            }
            reader.Close();
            sqlConn.Close();

            Log.RecordLog("符合筛选条件的数据共 " + SharedData.sumRows.ToString() + " 条！");
        }

        private string GetSharedPattern()
        {
            string sql = "共享模式 !=''";
            if (DBHelper.sharedPatternList.Count != 0)
            {
                sql += " and (";
                for (int i = 0; i < DBHelper.sharedPatternList.Count; i++)
                {
                    if (DBHelper.sharedPatternList[i] != null)
                    {
                        sql += "(共享模式='" + DBHelper.sharedPatternList[i] + "'";
                        if (DBHelper.serviceTimeList[i])
                        {
                            sql += " and cast(年对外服务机时 AS float) > 0 )";
                        }
                        else
                        {
                            sql += ")";
                        }
                    }
                    if (i < DBHelper.sharedPatternList.Count - 1)
                    {
                        sql += "or ";
                    }
                }
                sql += ")";
            }
            return sql;
        }

        

        private void btnExcape_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void btnReturn_Click(object sender, EventArgs e)
        {
            
        }
        //进度显示      
        public static void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (SharedData.isReady)
            {
                progressBar1.Minimum = 0;
                progressBar1.Maximum = SharedData.sumRows;
                //label1.Text = SharedData.sequenceNumber.ToString() + SharedData.sumRows.ToString();
                label1.Text = ((int)SharedData.sequenceNumber * 100 / (int)SharedData.sumRows).ToString() + "%";
                if ((int)SharedData.sequenceNumber < SharedData.sumRows)
                {
                    progressBar1.Value = (int)SharedData.sequenceNumber;
                }
            }
            if (SharedData.isCompleted)
            {
                progressBar1.Value = progressBar1.Maximum;
                label1.Text = "100%";
                btnExcape.Enabled = true;
            }
        }

        private void RunningForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //KillProcess("WINWORD.EXE");
            DBHelper.Execute("taskkill /f /im winword.exe");
            Environment.Exit(0);
        }

        private void KillProcess(string processName)
        {
            Process[] myproc = Process.GetProcesses();
            foreach (Process item in myproc)
            {
                if (item.ProcessName == processName)
                {
                    item.Kill();
                }
            }
        }
    }
}
