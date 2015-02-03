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
using System.ServiceProcess;
using System.Threading;
using System.IO;

namespace DFSystem
{
    public partial class MainForm : Form
    {
        //计算程序执行时间
        System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();
        SelectColumnsForm selectColumnsForm;

        public MainForm()
        {
            InitializeComponent();
            /*初始化日志目录*/
            if (!Directory.Exists("c:\\DFS"))
            {
                Directory.CreateDirectory("c:\\DFS");
            }
            /*启动sql server */
            Thread th = new Thread(new ThreadStart(StartServer));
            th.Start();
        }

        //单击共享目录按钮
        private void btnImport_Click(object sender, EventArgs e)
        {
            Log.RecordLog("---------------------------------------------------------------------------------------------");
            String excelFile = null;
            excelFile = openFile();
            if (excelFile == null)
            {
                return;
            }
            selectColumnsForm = new SelectColumnsForm(this, excelFile);
            this.Hide();
            selectColumnsForm.Show();
        }

        //启动sql server
        private void StartServer()
        {
            ServiceController sc = new ServiceController("MSSQLSERVER");
            if (sc.Status == ServiceControllerStatus.Stopped)
            {
                sc.Start();
                Log.RecordLog("SQL数据库服务SQLEXPRESS启动成功！");
            }
        }

        private string openFile()
        {
            string filePath = "";
            openFileDialog1.Title = "选择Excel数据源文件！";
            openFileDialog1.Filter = "Excel文件(*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
            }
            else
            {
                return null;
            }
        
            return filePath;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        //数据汇编功能
        private void btnDataCompile_Click(object sender, EventArgs e)
        {
            SelectCompileSourceForm selectDataSourceForm = new SelectCompileSourceForm();
            selectDataSourceForm.Show();
            this.Hide();
        }
    }
}
