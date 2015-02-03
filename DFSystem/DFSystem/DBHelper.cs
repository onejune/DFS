using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SQLite;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;

namespace DFSystem
{
    /* 此类维护数据库连接字符串和 Connection 对象 */
    class DBHelper
    {
         /* 连接sql server数据库 */
        //public static string connString = @"server=XP-201305101104\MSSQLSERVER;database=DFS;uid=sa; pwd=passw0rd ";
        public static string connString = @"server=(local);database=DFS;uid=sa; pwd=passw0rd ";

        //不显示的仪器类型
        public static List<string> unselectedDeviceTypeList = new List<string>();
        //共享模式
        public static List<string> sharedPatternList = new List<string>();
        //年服务机时
        public static List<bool> serviceTimeList = new List<bool>();
        //是否对分析仪器按仪器分类中类再次分组
        public static bool groupToAnalyseDevice = false;
        //是否只显示211
        public static bool only211 = false;
        //是否只显示985
        public static bool only985 = false;
        //是否按单位名称再次分组
        public static bool groupToOrgName = false;

        public static List<string> GetAllColumnNamesFromServer(string tableName)
        {
            List<string> names = new List<string>();
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            string sql = "select top 1 * from " + tableName;

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand(sql, sqlConn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);

            System.Data.DataTable dt = ds.Tables[0];
            foreach (System.Data.DataColumn c in dt.Columns)
            {
                string text = dt.Rows[0][c].ToString();
                if (text != "")
                {
                    names.Add(c.ColumnName);
                }
            }
            sqlConn.Close();
            return names;
        }

        //dosCommand Dos命令语句  
        public static string Execute(string dosCommand)
        {
            return Execute(dosCommand, 0);
        }
        /// <summary>  
        /// 执行DOS命令，返回DOS命令的输出  
        /// </summary>  
        /// <param name="dosCommand">dos命令</param>  
        /// <param name="milliseconds">等待命令执行的时间（单位：毫秒），  
        /// 如果设定为0，则无限等待</param>  
        /// <returns>返回DOS命令的输出</returns>  
        public static string Execute(string command, int seconds)
        {
            string output = ""; //输出字符串  
            if (command != null && !command.Equals(""))
            {
                Process process = new Process();//创建进程对象  
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "cmd.exe";//设定需要执行的命令  
                startInfo.Arguments = "/C " + command;//“/C”表示执行完命令后马上退出  
                startInfo.UseShellExecute = false;//不使用系统外壳程序启动  
                startInfo.RedirectStandardInput = false;//不重定向输入  
                startInfo.RedirectStandardOutput = true; //重定向输出  
                startInfo.CreateNoWindow = true;//不创建窗口  
                process.StartInfo = startInfo;
                try
                {
                    if (process.Start())//开始进程  
                    {
                        if (seconds == 0)
                        {
                            process.WaitForExit();//这里无限等待进程结束  
                        }
                        else
                        {
                            process.WaitForExit(seconds); //等待进程结束，等待时间为指定的毫秒  
                        }
                        output = process.StandardOutput.ReadToEnd();//读取进程的输出  
                    }
                }
                catch
                {
                }
                finally
                {
                    if (process != null)
                    {
                        process.Close();
                    }
                }
            }
            return output;
        }
    }
}
