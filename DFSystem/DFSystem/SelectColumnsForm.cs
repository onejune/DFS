using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace DFSystem
{
    public partial class SelectColumnsForm : Form
    {
        private string excelFile;
        private Form lastForm;
        //选定的列
        private List<string> displayedColumns = new List<string>();
        //默认按区域分组
        private string groupName = "区域";
        //默认选择的列
        private List<string> defaultColumnsForArea = new List<string> { "中文名称", "型号规格", "仪器分类大类", "仪器分类中类", "产地", "单位名称", "联系人", "联系电话" };
        private List<string> defaultColumnsForDevice = new List<string> { "中文名称", "型号规格", "所在省", "产地", "单位名称", "联系人", "联系电话" };
        private List<string> defaultColumnsForMember = new List<string> { "中文名称", "型号规格", "仪器分类大类", "仪器分类中类", "所在省", "单位名称", "联系人", "联系电话" };

        public SelectColumnsForm(Form lastForm, string excelFile)
        {
            InitializeComponent();
            this.excelFile = excelFile;
            this.lastForm = lastForm;
        }

        //从excel表中获取当前所有列名
        private List<string> GetColumns()
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
            List<string> columnList = new List<string>();

            // 获得表头，即第一行数据
            for (int i = 0; i < colCount; i++)
            {
                range1 = worksheet.Range[worksheet.Cells[1, i + 1], worksheet.Cells[1, i + 1]];
                columnList.Add(range1.Value2.ToString());
            }
            excel.Quit();
            return columnList;
        }

        private List<string> GetAllColumnNamesFromServer()
        {
            List<string> names = new List<string>();
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            string sql = "select top 1 * from Sheet1";

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

        private void btnReturn_Click(object sender, EventArgs e)
        {
            lastForm.Show();
            this.Dispose();
        }

        //开始执行
        private void btnStart_Click(object sender, EventArgs e)
        {
            SharedData.isReady = false;
            getColumnsAndGroupName();
            if (displayedColumns.Count > 7)
            {
                MessageBox.Show("选定的列过多，不能超过7列！", "Error!");
                return;
            }
            if (displayedColumns.Count == 0)
            {
                MessageBox.Show("请选择需要显示的列名！", "Error!");
                return;
            }
            int fromValue, toValue;
            if (textBox1.Text.Length == 0)
            {
                fromValue = 0;
            }
            else
            {
                fromValue = int.Parse(textBox1.Text);
            }
            if (textBox2.Text.Length == 0)
            {
                toValue = 0;
            }
            else
            {
                toValue = int.Parse(textBox2.Text);
            }

            GetUserChoice();

            //设置列的顺序和宽度
            TableSetForm tableSetForm = new TableSetForm(this, excelFile, displayedColumns, groupName, fromValue, toValue);
            tableSetForm.Show();
            this.Hide();
        }

        private void GetUserChoice()
        {
            //获取不需要显示的仪器类型列表
            for (int i = 0; i < chkListDeviceType.Items.Count; i++)
            {
                if (chkListDeviceType.GetItemCheckState(i) == CheckState.Unchecked)
                {
                    DBHelper.unselectedDeviceTypeList.Add(chkListDeviceType.Items[i].ToString());
                }
            }
            //获取共享模式
            if (cmb1.Text != "")
            {
                DBHelper.sharedPatternList.Add(cmb1.Text);
                DBHelper.serviceTimeList.Add(cb1.Checked);
            }
            if (cmb2.Text != "")
            {
                DBHelper.sharedPatternList.Add(cmb2.Text);
                DBHelper.serviceTimeList.Add(cb2.Checked);
            }
            if (cmb3.Text != "")
            {
                DBHelper.sharedPatternList.Add(cmb3.Text);
                DBHelper.serviceTimeList.Add(cb3.Checked);
            }
        }

        private void btnEscape_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        //获取分类名称以及选择的列名
        private void getColumnsAndGroupName()
        {
            if (groupColumn.SelectedItem != null)
            {
                groupName = groupColumn.CheckedItems[0].ToString(); //按什么分类？默认按所在省分类
            }
            int i = 0, flag = 0;
            displayedColumns.Clear();

            for (i = 0; i < columnListBox.Items.Count; i++)
            {
                if (columnListBox.GetItemChecked(i))
                {
                    string text = columnListBox.Items[i].ToString();
                    if (!text.ToString().Equals("仪器分类大类") && !text.ToString().Equals("仪器分类中类"))
                    {
                        displayedColumns.Add(columnListBox.Items[i].ToString());
                    }
                    else
                    {
                        flag = 1;
                    }
                }
            }
            if (flag == 1)
            {
                displayedColumns.Add("仪器类型");
            }
        }

        private void groupColumn_ItemCheck(object sender, ItemCheckEventArgs e)
        {

            for (int i = 0; i < groupColumn.CheckedIndices.Count; i++)
            {
                if (groupColumn.CheckedIndices[i] != e.Index)
                {
                    groupColumn.SetItemChecked(groupColumn.CheckedIndices[i], false);
                }
            }
            //if (groupColumn.CheckedIndices.Count == 0)
            //{
            //    return;
            //}
            List<string> defaultColumns;
            if (e.NewValue == CheckState.Unchecked)
            {
                defaultColumns = null;
            }
            else
            {
                groupName = groupColumn.SelectedItem.ToString();

                if (groupName.Equals("仪器类型"))
                {
                    defaultColumns = defaultColumnsForDevice;
                }
                else if (groupName.Equals("区域"))
                {
                    defaultColumns = defaultColumnsForArea;
                }
                else
                {
                    defaultColumns = defaultColumnsForMember;
                }
                //如果当前选中了“基地类型”，判断是否有“基地级别”和“基地大类-汇编用”列
                if (groupName.Equals("基地类型"))
                {
                    if (!columnListBox.Items.Contains("基地级别") || !columnListBox.Items.Contains("基地大类-汇编用"))
                    {
                        MessageBox.Show("当前数据库中不存在“基地级别”和“基地大类-汇编用”，无法按基地类型汇编！", "错误！");
                        return;
                    }
                }
            }
            for (int i = 0; i < columnListBox.Items.Count; i++)
            {
                string str = "";
                str = columnListBox.Items[i].ToString();
                if (defaultColumns != null && defaultColumns.Contains(columnListBox.Items[i].ToString()))
                {
                    columnListBox.SetItemChecked(i, true);
                }
                else
                {
                    columnListBox.SetItemChecked(i, false);
                }
            }

        }

        private void SelectColumnsForm_Load(object sender, EventArgs e)
        {
            //计算程序执行时间
            System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();

            oTime.Start();
            //从excel文件导入数据到sql server
            ImportExcel import = new ImportExcel(excelFile, "Sheet1");
            int count = import.StartWork(true);
            oTime.Stop();
            long time1 = oTime.ElapsedMilliseconds / 1000;
            Log.RecordLog("从" + excelFile + "导入数据完成，耗时 " + time1.ToString() + " 秒！ 共导入 " + count + "条数据！");

            List<string> columnList = GetAllColumnNamesFromServer();

            //待显示列名
            columnListBox.Items.Clear();

            foreach (string col in columnList)
            {
                columnListBox.Items.Add(col);
                //目前仅支持按区域（所在省）、按仪器类型（仪器分类大类）、按隶属关系（？）分类     
            }

            DBHelper.unselectedDeviceTypeList.Clear();
            DBHelper.sharedPatternList.Clear();

            for (int i = 0; i < chkListDeviceType.Items.Count; i++)
            {
                chkListDeviceType.SetItemChecked(i, true);
            }
        }

        private void SelectColumnsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string str = textBox1.Text;
            for (int i = 0; i < str.Length; i++)
            {
                if (!Char.IsNumber(str[i]))
                {
                    textBox1.Text = string.Empty;
                    if (MessageBox.Show("请输入数字", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    {
                        textBox1.Focus();
                        textBox1.SelectAll();
                        break;
                    }
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string str = textBox2.Text;
            for (int i = 0; i < str.Length; i++)
            {
                if (!Char.IsNumber(str[i]))
                {
                    textBox1.Text = string.Empty;
                    if (MessageBox.Show("请输入数字", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    {
                        textBox2.Focus();
                        textBox2.SelectAll();
                        break;
                    }
                }
            }
        }

        private void chkListDeviceType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void chkListDeviceType_ItemCheck(object sender, ItemCheckEventArgs e)
        {

        }

        private void cbGroupAgain_CheckedChanged(object sender, EventArgs e)
        {
            DBHelper.groupToAnalyseDevice = (cbGroupAgain.CheckState == CheckState.Checked ? true : false);
        }

        private void groupColumn_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cb211_CheckedChanged(object sender, EventArgs e)
        {
            DBHelper.only211 = (cb211.CheckState == CheckState.Checked ? true : false);
            if (DBHelper.only211)
            {
                DBHelper.only985 = false;
                cb985.CheckState = CheckState.Unchecked;
            }
        }

        private void cb985_CheckedChanged(object sender, EventArgs e)
        {
            DBHelper.only985 = (cb985.CheckState == CheckState.Checked ? true : false);
            if (DBHelper.only985)
            {
                DBHelper.only211 = false;
                cb211.CheckState = CheckState.Unchecked;
            }
        }

        private void cbOrgName_CheckedChanged(object sender, EventArgs e)
        {
            DBHelper.groupToOrgName = (cbOrgName.CheckState == CheckState.Checked ? true : false);
        }
    }
}
