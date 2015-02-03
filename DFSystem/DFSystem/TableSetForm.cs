using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace DFSystem
{
    public partial class TableSetForm : Form
    {
        string[] defaultColumns = new string[] { "中文名称", "型号规格", "仪器类型", "产地", "单位名称", "联系人", "联系电话" };
        private SelectColumnsForm selectColumnsForm;
        private string excelFile;
        private List<string> displayedColumns;
        private string groupName;
        private int fromValue;
        private int toValue;

        public TableSetForm()
        {
            InitializeComponent();
        }

        public TableSetForm(SelectColumnsForm selectColumnsForm, string excelFile, List<string> displayedColumns, string groupName, int fromValue, int toValue)
        {
            InitializeComponent();
            this.selectColumnsForm = selectColumnsForm;
            this.excelFile = excelFile;
            this.displayedColumns = displayedColumns;
            this.groupName = groupName;
            this.fromValue = fromValue;
            this.toValue = toValue;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            string displayedColumnString = getDisplayedColumnString();
            string sql = "select TOP 1000 " + displayedColumnString + " from Sheet1 ";
            SqlCommand command = new SqlCommand(sql, sqlConn);
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            da.Fill(ds, "temp");

            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "temp";
            sqlConn.Close();
        }

        //将用户选择的column数组组合成sql字符串
        private string getDisplayedColumnString()
        {
            int i = 0;
            string str = "";
            for (i = 0; i < displayedColumns.Count; i++)
            {
                //如果出现原excel表中没有的“仪器类型”这一列，将“仪器分类大类”和“仪器分类中类”组合成该列
                if (!displayedColumns[i].Equals("仪器类型"))
                {
                    str += displayedColumns[i] + ",";
                }
                else
                {
                    str += "仪器分类大类+ '/' + 仪器分类中类 as 仪器类型,";
                }
            }
            return str.Substring(0, str.Length - 1);
        }

        private void button1_Click(object sender, EventArgs e)
        {           
            DataGridViewColumnCollection columnCollection = dataGridView1.Columns;
            string[] selectedColumns = new string[10];
            float[] columnsWidth = new float[10];
            float sumWidth = 0;

            //计算所有列列宽之和
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                sumWidth += dataGridView1.Columns[i].Width;
            }
            foreach (DataGridViewColumn col in columnCollection)
            {
                selectedColumns[col.DisplayIndex] = col.HeaderText;
                columnsWidth[col.DisplayIndex] = col.Width / sumWidth;
            }
            displayedColumns.Clear();
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (selectedColumns[i].Length != 0)
                {
                    displayedColumns.Add(selectedColumns[i]);
                }
            }
            //开始生成数据
            RunningForm runForm = new RunningForm(this, excelFile, displayedColumns, groupName, fromValue, toValue, columnsWidth);
            this.Hide();
            runForm.Show();
            runForm.run();
        }

        private void btnLast_Click(object sender, EventArgs e)
        {
            selectColumnsForm.Show();
            this.Dispose();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }
    }
}
