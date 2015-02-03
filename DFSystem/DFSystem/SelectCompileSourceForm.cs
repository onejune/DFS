using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ServiceProcess;
using System.IO;
using Microsoft.Office.Interop.Word;
using DFSystem.Tables;
using System.Diagnostics;

namespace DFSystem
{
    public partial class SelectCompileSourceForm : Form
    {
        TableGenerateForDataCompile generate = null;
        //excel文件存放目录
        string dirName = "";
        //excel文件和table对应关系
        List<ExcelMap> excelMapList = new List<ExcelMap>();

        public SelectCompileSourceForm()
        {
            InitializeComponent();
        }

        private void SelectDataSourceForm_Load(object sender, EventArgs e)
        {
            Log.RecordLog("-------------------------------------------------------------------------------------");
            Log.RecordLog("-----------------------------资源调查数据汇编-----------------------------------------");

            //选择excel目录
            Log.RecordLog("选择数据汇编数据目录");
            dirName = SelectDataSourceDir();
            if (dirName == "")
            {
                return;
            }

            int i = 0, j = 0;
            excelMapList.Clear();
            //获取指定目录下的非隐藏word文件
            foreach (string filePath in Directory.GetFileSystemEntries(dirName, "*.xls", SearchOption.AllDirectories))
            {
                DirectoryInfo d = new DirectoryInfo(filePath);
                if ((d.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
                {
                    //string fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1, (filePath.LastIndexOf(".") - filePath.LastIndexOf("\\") - 1));
                    //不能添加重复的excel文件
                    i = 0;
                    j = 0;
                    for (i = 0; i < excelMapList.Count(); i++)
                    {
                        for (j = 0; j < excelMapList[i].FilePathList.Count(); j++)
                        {
                            string fileName2 = excelMapList[i].FilePathList[j];
                            if (fileName2.Equals(filePath))
                            {
                                break;
                            }
                        }
                        
                    }
                    //没有重复的excel
                    if (i == excelMapList.Count())
                    {
                        //看看有没有连续的excel文件
                        string fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1, (filePath.LastIndexOf(".") - filePath.LastIndexOf("\\") - 1));
                        fileName = fileName.Substring(0, fileName.Length - 1);
                        int k = 0;
                        for (k = 0; k < excelMapList.Count(); k++)
                        {
                            string path = excelMapList[k].FilePathList[0];
                            string name = path.Substring(path.LastIndexOf("\\") + 1, (path.LastIndexOf(".") - path.LastIndexOf("\\") - 1));
                            name = name.Substring(0, name.Length - 1);
                            if (fileName.Equals(name))
                            {
                                break;
                            }
                        }
                        if (k < excelMapList.Count())
                        {
                            excelMapList[k].FilePathList.Add(filePath);
                        }
                        else
                        {
                            ExcelMap em = new ExcelMap();
                            em.FilePathList.Add(filePath);
                            excelMapList.Add(em);
                        } 
                    }
                }
            }
            if (excelMapList.Count() == 0)
            {
                return;
            }
            i = 0;
            //左侧添加excel文件名
            foreach (ExcelMap em in excelMapList)
            {
                //根据路径获取文件名
                string fileName = em.FileName();
                excelFileListBox.Items.Add(fileName);
                em.ListBoxIndex = i++;
            }
            //保证两边的文件数目相同
            int n = SharedData.tableListInDB.Count() - excelMapList.Count();
            String nullStr = "null";
            while (n > 0)
            {
                ExcelMap em = new ExcelMap();
                excelMapList.Add(em);
                excelFileListBox.Items.Add(nullStr);
                n--;
            }
            //右侧添加table名
            for (i = 0; i < SharedData.tableListInDB.Count(); i++)
            {
                TableAndView tbl = SharedData.tableListInDB[i];
                fixedFileNameList.Items.Add(tbl.TableName);
                excelMapList[i].TableListIndex = i;
            }
        }

        //选择数据汇编数据源目录
        private string SelectDataSourceDir()
        {
            string dir = "";
            folderBrowserDialog1.Description = "选择数据汇编数据目录！";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                dir = folderBrowserDialog1.SelectedPath;
            }
            else
            {
                return null;
            }
            //string fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1, (filePath.LastIndexOf(".") - filePath.LastIndexOf("\\") - 1));
            return dir;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (excelMapList.Count == 0)
            {
                return;
            }
            this.Hide();
            generate = new TableGenerateForDataCompile();
            generate.StartWordGenerate(excelMapList);
            //this.Show();
            MessageBox.Show("全部汇编结束！耗时 " + SharedData.oTime.ElapsedMilliseconds / 60000 + " 分钟！");
            //清理word进程
            RunCmd();
            Environment.Exit(0);
        }

        private void SelectCompileSourceForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //清理word进程
            RunCmd();
        }

        /// <summary>   
        /// 运行DOS命令   
        /// DOS关闭进程命令(ntsd -c q -p PID )PID为进程的ID   
        /// </summary>   
        /// <param name="command"></param>   
        /// <returns></returns>   
        public string RunCmd()
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.Arguments = "/c " + "taskkill /IM winword.exe /f";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;

            p.Start();

            return p.StandardOutput.ReadToEnd();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private void excelFileListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (excelFileListBox.CheckedItems.Count > 0)
            {
                for (int i = 0; i < excelFileListBox.Items.Count; i++)
                {
                    if (i != e.Index)
                    {
                        this.excelFileListBox.SetItemCheckState(i, System.Windows.Forms.CheckState.Unchecked);
                    }
                }
            }  
        }

        private void fixedFileNameList_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (fixedFileNameList.CheckedItems.Count > 0)
            {
                for (int i = 0; i < fixedFileNameList.Items.Count; i++)
                {
                    if (i != e.Index)
                    {
                        this.fixedFileNameList.SetItemCheckState(i, System.Windows.Forms.CheckState.Unchecked);
                    }
                }
            }  
        }

        //手动匹配
        private void btnMove_Click(object sender, EventArgs e)
        {
            if (fixedFileNameList.CheckedItems.Count > 0 && excelFileListBox.CheckedItems.Count > 0)
            {
                int right = fixedFileNameList.CheckedIndices[0];
                int left = excelFileListBox.CheckedIndices[0];
                if((right < excelFileListBox.Items.Count) && (right != left))
                {
                    string str = excelFileListBox.Items[left].ToString();
                    excelFileListBox.Items[left] = excelFileListBox.Items[right];
                    excelFileListBox.Items[right] = str;

                    int k = excelMapList[left].TableListIndex;
                    excelMapList[left].TableListIndex = right;
                    excelMapList[right].TableListIndex = k;

                    ExcelMap em = excelMapList[left];
                    excelMapList[left] = excelMapList[right];
                    excelMapList[right] = em;

                    this.excelFileListBox.SetItemCheckState(right, System.Windows.Forms.CheckState.Checked);
                    this.excelFileListBox.SetSelected(right, true);
                }
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }  
    }
}
