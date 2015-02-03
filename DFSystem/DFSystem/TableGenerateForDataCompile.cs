using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Data;
using DFSystem.Tables;
using System.ServiceProcess;
using System.IO;

namespace DFSystem
{
    class TableGenerateForDataCompile
    {
        //每个table对应一个job
        private List<AbstractTable> dataCompileJobList = new List<AbstractTable>();

        private float maxTableHeight = 646F;
        private int sumRows;
        private int sumColumns;

        private List<string> finalListA = new List<string>();
        private List<string> finalListB = new List<string>();

        private List<double> sumList = new List<double>();
        private List<double> assistSumList = new List<double>();

        private DataCompile dataCompile;
        private string sqlStrForCompile;

        private int sumMaxLine = 0;
        private int sumMinLine = 0;

        //当无法显示全部列时，只显示部分列+“其他”，如2-5
        private string otherColumn = "";

        private string sumOrCount;

        private List<DataCompile> researchInstitutionDataCompileList;

        public string Sql
        {
            get { return sqlStrForCompile; }
            set { sqlStrForCompile = value; }
        }

        internal DataCompile DataCompile
        {
            get { return dataCompile; }
            set { dataCompile = value; }
        }

        public TableGenerateForDataCompile()
        {


        }

        public TableGenerateForDataCompile(DataCompile dc)
        {
            this.dataCompile = dc;
        }

        //初始化job list
        private void InitDataCompileJobList()
        {
            dataCompileJobList.Clear();
            //向job list中添加job
            dataCompileJobList.Add(new ResearchInstitution1_1());
            dataCompileJobList.Add(new ResearchInstitution1_2());
            dataCompileJobList.Add(new ResearchInstitution1_3());
            dataCompileJobList.Add(new ResearchInstitution1_4());
            dataCompileJobList.Add(new ResearchInstitution1_5());
            dataCompileJobList.Add(new ResearchInstitution1_6());
            dataCompileJobList.Add(new ResearchInstitution1_7());
            dataCompileJobList.Add(new ResearchInstitution1_8());
            dataCompileJobList.Add(new ResearchInstitution1_9());
            dataCompileJobList.Add(new ResearchInstitution1_10_1());
            dataCompileJobList.Add(new ResearchInstitution1_10_2());
            dataCompileJobList.Add(new ResearchInstitution1_10_3());
            dataCompileJobList.Add(new ResearchInstitution1_11());
            dataCompileJobList.Add(new ResearchInstitution1_12());
            dataCompileJobList.Add(new ResearchInstitution2_1_1());
            dataCompileJobList.Add(new ResearchInstitution2_1_2());
            dataCompileJobList.Add(new ResearchInstitution2_2_1());
            dataCompileJobList.Add(new ResearchInstitution2_2_2());
            dataCompileJobList.Add(new ResearchInstitution2_3_1());
            dataCompileJobList.Add(new ResearchInstitution2_3_2());
            dataCompileJobList.Add(new ResearchInstitution2_4_1());
            dataCompileJobList.Add(new ResearchInstitution2_4_2());
            dataCompileJobList.Add(new ResearchInstitution2_5_1());
            dataCompileJobList.Add(new ResearchInstitution2_5_2());
            dataCompileJobList.Add(new ResearchInstitution2_6_1());
            dataCompileJobList.Add(new ResearchInstitution2_6_2());
            dataCompileJobList.Add(new ResearchInstitution2_7_1());
            dataCompileJobList.Add(new ResearchInstitution2_7_2());
            dataCompileJobList.Add(new ResearchInstitution2_8_1());
            dataCompileJobList.Add(new ResearchInstitution2_8_2());
            dataCompileJobList.Add(new ResearchInstitution2_9_1());
            dataCompileJobList.Add(new ResearchInstitution2_9_2());
            dataCompileJobList.Add(new ResearchInstitution2_10_1());
            dataCompileJobList.Add(new ResearchInstitution2_10_2());
            dataCompileJobList.Add(new ResearchInstitution2_11_1());
            dataCompileJobList.Add(new ResearchInstitution2_11_2());
            dataCompileJobList.Add(new ResearchInstitution2_12());
            dataCompileJobList.Add(new ResearchInstitution2_13());
            dataCompileJobList.Add(new ResearchInstitution2_14());
            dataCompileJobList.Add(new ResearchInstitution2_15());
            dataCompileJobList.Add(new ResearchInstitution3_1_1());
            dataCompileJobList.Add(new ResearchInstitution3_1_2());
            dataCompileJobList.Add(new ResearchInstitution3_1_3());
            dataCompileJobList.Add(new ResearchInstitution3_1_4());
            dataCompileJobList.Add(new ResearchInstitution3_1_5());
            dataCompileJobList.Add(new ResearchInstitution3_1_6());
            dataCompileJobList.Add(new ResearchInstitution3_2_1());
            dataCompileJobList.Add(new ResearchInstitution3_2_2());
            dataCompileJobList.Add(new ResearchInstitution3_3_1());
            dataCompileJobList.Add(new ResearchInstitution3_3_2());
            dataCompileJobList.Add(new ResearchInstitution3_4_1());
            dataCompileJobList.Add(new ResearchInstitution3_4_2());
            dataCompileJobList.Add(new ResearchInstitution3_5_1());
            dataCompileJobList.Add(new ResearchInstitution3_5_2());
            dataCompileJobList.Add(new ResearchInstitution3_6_1());
            dataCompileJobList.Add(new ResearchInstitution3_6_2());
            dataCompileJobList.Add(new ResearchInstitution3_7_1());
            dataCompileJobList.Add(new ResearchInstitution3_7_2());
            dataCompileJobList.Add(new ResearchInstitution3_8_1());
            dataCompileJobList.Add(new ResearchInstitution3_8_2());
            dataCompileJobList.Add(new ResearchInstitution3_9_1());
            dataCompileJobList.Add(new ResearchInstitution3_9_2());
            dataCompileJobList.Add(new ResearchInstitution4_1());
            dataCompileJobList.Add(new ResearchInstitution4_2());
            dataCompileJobList.Add(new ResearchInstitution4_3());
            dataCompileJobList.Add(new ResearchInstitution4_4());
            dataCompileJobList.Add(new ResearchInstitution4_5());
            dataCompileJobList.Add(new ResearchInstitution4_6());
            dataCompileJobList.Add(new ResearchInstitution4_7());
            dataCompileJobList.Add(new ResearchInstitution4_9());
            dataCompileJobList.Add(new ResearchInstitution4_10());
            dataCompileJobList.Add(new ResearchInstitution4_11_1());
            dataCompileJobList.Add(new ResearchInstitution4_11_2());
            dataCompileJobList.Add(new ResearchInstitution4_11_3());
            dataCompileJobList.Add(new ResearchInstitution4_12());
            dataCompileJobList.Add(new ResearchInstitution4_13());
            dataCompileJobList.Add(new ResearchInstitution4_14());
            dataCompileJobList.Add(new ResearchInstitution4_15());
            dataCompileJobList.Add(new ResearchInstitution4_16());
            dataCompileJobList.Add(new ResearchInstitution4_17());
            dataCompileJobList.Add(new ResearchInstitution4_18());
            dataCompileJobList.Add(new ResearchInstitution4_19());
            dataCompileJobList.Add(new ResearchInstitution4_20());
            dataCompileJobList.Add(new ResearchInstitution4_21());
            dataCompileJobList.Add(new ResearchInstitution4_22());
            dataCompileJobList.Add(new ResearchInstitution4_22());
            dataCompileJobList.Add(new ResearchInstitution4_23());
            dataCompileJobList.Add(new ResearchInstitution4_24());
            dataCompileJobList.Add(new ResearchInstitution4_25());
            dataCompileJobList.Add(new ResearchInstitution4_26());
        }

        //启动sql server
        private void StartServer()
        {
            ServiceController sc = new ServiceController("MSSQLSERVER");
            if (sc.Status == ServiceControllerStatus.Stopped)
            {
                sc.Start();
            }
            Log.RecordLog("SQL数据库服务SQLEXPRESS启动成功！");
        }

        //执行数据汇编生成table的入口函数
        public void StartWordGenerate(List<ExcelMap> excelMapList)
        {
            SharedData.oTime.Start();

            //初始化目录
            InitDirectory();

            //启动sql server
            StartServer();

            //从excel中导入数据到数据库
            ImportAllExcelFiles(excelMapList);

            //全国数据汇编
            Log.RecordLog("----------------------------------------------------全国数据汇编----------------------------------------------------------");
            Generate(excelMapList, 1);
            //中央数据汇编
            Log.RecordLog("----------------------------------------------------中央数据汇编----------------------------------------------------------");
            Generate(excelMapList, 2);
            //地方数据汇编
            Log.RecordLog("----------------------------------------------------地方数据汇编----------------------------------------------------------");
            Generate(excelMapList, 3);

            SharedData.oTime.Stop();
            long time = SharedData.oTime.ElapsedMilliseconds;
            Log.RecordLog("**************************************************************************************************");
            Log.RecordLog(@"数据汇编结束，耗时：" + (time / 60000).ToString() + "分钟！");
        }

        private void Generate(List<ExcelMap> excelMapList, int type)
        {
            SharedData.AREATYPE = type;

            //过滤数据，建立视图
            foreach (ExcelMap f in excelMapList)
            {
                if ((f.TableListIndex == -1) || (f.FileName() == ""))
                {
                    continue;
                }
                int k = 0;
                k = f.TableListIndex;
                string tableName = f.GetTaleName();
                CreateView(tableName, k);
                SharedData.tableListInDB[k].IsExisted = true;
            }
            //Environment.Exit(0);

            //把所在省和设备类型列表写入数据库中
            SharedData.WriteProvincesToDB();
            SharedData.WriteDeviceTypesToDB();
            int res = SharedData.WriteBaseTypesToDB();
            if (res == -1)
            {
                return;
            }

            //初始化job list
            InitDataCompileJobList();

            for (int i = 0; i < dataCompileJobList.Count; i++)
            {
                AbstractTable tbl = dataCompileJobList[i];
                try
                {
                    tbl.RunJob();
                }
                catch (System.Exception ex)
                {
                    Log.RecordLog("RunJob():\n" + ex.ToString());
                }
            }
            //生成插页
            CreateBlankPageFile();

            //合并word文件
            FileMerge();
        }

        //将指定的文件夹中的所有excel文件导入到sql server
        private void ImportAllExcelFiles(List<ExcelMap> excelMapList)
        {
            //string[] files = Directory.GetFiles(path, "*.xls", SearchOption.AllDirectories);
            foreach (ExcelMap f in excelMapList)
            {
                if ((f.TableListIndex == -1))
                {
                    continue;
                }
                if(f.FileName() == "")
                {
                    //解除excel文件和table之间的映射
                    SharedData.tableListInDB[f.TableListIndex].IsExisted = false;
                    SharedData.tableListInDB[f.TableListIndex].TableName = "";
                    SharedData.tableListInDB[f.TableListIndex].ViewName = "";
                    SharedData.tableListInDB[f.TableListIndex].ColumnNames = null;
                    f.TableListIndex = -1;
                    continue;
                }
                string tableName = f.GetTaleName();

                int i = 0;
                bool needDeleteOldTable = true;
                for (i = 0; i < f.FilePathList.Count(); i++)
                {
                    string filePath = f.FilePathList[i];
                    //导入用户选择的excel文件
                    ImportExcel importExcel = new ImportExcel(filePath, tableName);
                    int res = importExcel.StartWork(needDeleteOldTable);
                    if (res == 0)
                    {
                        continue;
                    }
                    else
                    {
                        Log.RecordLog("ImportAllExcelFiles: 成功导入" + res + "条数据到“" + tableName + "”中！");
                    }
                    needDeleteOldTable = false;
                }
                //保存列名
                i = f.TableListIndex;
                SharedData.tableListInDB[i].ColumnNames.Clear();
                List<string> columns = GetAllColumnNamesFromServer(tableName);
                foreach (var v in columns)
                {
                    SharedData.tableListInDB[i].ColumnNames.Add(v);
                }
            }
        }

        //过滤数据并生成视图
        private void CreateView(string tableName, int index)
        {
            string sql = "";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            string viewName = "view_" + tableName;

            if (SharedData.tableListInDB[index].ColumnNames.Contains("隶属") == false)
            {
                Log.RecordLog("CreateView：error: 表\"" + tableName + "\"中不存在\"隶属\"字段！");
                return;
            }

            sql = "if exists(select * from INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME='" + viewName + "') drop view " + viewName;
            sqlConn.Open();
            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.ExecuteNonQuery();
            sql = "";
            try
            {
                if (SharedData.AREATYPE == 3)//地方
                {
                    //临时word保存目录
                    SharedData.currentWordPath = @"C:\DFS\word\地方\";
                    SharedData.currentFilePath = @"C:/DFS/word/地方/";
                    sql = "create view " + viewName + " as select * from " + tableName + " where 隶属 != '国家级'";
                }
                else if (SharedData.AREATYPE == 2)//中央
                {
                    //临时word保存目录
                    SharedData.currentWordPath = @"C:\DFS\word\中央\";
                    SharedData.currentFilePath = @"C:/DFS/word/中央/";
                    sql = "create view " + viewName + " as select * from " + tableName + " where 隶属 = '国家级'";
                }
                else//全国
                {
                    //临时word保存目录
                    SharedData.currentWordPath = @"C:\DFS\word\全国\";
                    SharedData.currentFilePath = @"C:/DFS/word/全国/";
                    sql = "create view " + viewName + " as select * from " + tableName;
                }
                command.CommandText = sql;
                command.ExecuteNonQuery();

                Log.RecordLog("\"" + viewName + "\" 生成成功！");
            }
            catch (System.Exception ex)
            {
                Log.RecordLog(ex.ToString());
            }
            finally
            {
                sqlConn.Close();
            }
        }

        private string GetConvertStringByUnitName(string str)
        {
            if (dataCompile.UnitName.Equals("亿元"))
            {
                return " convert(float, sum(cast(\"" + str + "\" as float)))/10000 ";
            }
            else
            {
                return " convert(float, sum(cast(\"" + str + "\" as float))) ";
            }
        }

        private void CreateBlankPageFile()
        {
            CreateBlankPageFile("1-0 一、科研院所和高校概况");
            CreateBlankPageFile("2-0 二、大型科学仪器设备");
            CreateBlankPageFile("3-0 三、研究实验基地");
            CreateBlankPageFile("4-0 四、生物种质资源");
        }

        //生成只有一页的word文件
        private void CreateBlankPageFile(string caption)
        {
            //创建Word文档
            Object Nothing;
            Microsoft.Office.Interop.Word.Application wordApp;
            Microsoft.Office.Interop.Word.Document wordDoc;
            //wordDoc.Paragraphs.Last.Range.Select();

            //创建Word文档
            Nothing = System.Reflection.Missing.Value;
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //设置布局方向
            wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;
            //wordDoc.Paragraphs.First.Range.Select();

            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落

            wordApp.Selection.Font.Size = 13;
            wordApp.Selection.Font.Name = "楷体";
            wordApp.Selection.Font.Bold = 2;
            wordApp.Selection.TypeText("全国科研院所和高校重点科技基础条件资源调查数据汇编");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落

            wordApp.Selection.TypeText(caption.Substring(3));

            // 设置标题样式，以方便插入目录
            object oStyleName = "标题 1";
            wordApp.ActiveDocument.Styles[oStyleName].Font.Size = 26;
            wordApp.ActiveDocument.Styles[oStyleName].Font.Name = "微软雅黑";
            wordApp.ActiveDocument.Styles[oStyleName].Font.Bold = 2;
            wordApp.Selection.set_Style(ref oStyleName);
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落


            wordDoc.Paragraphs.First.Range.Select();
            // 插入文本框
            Shape newTextbox;
            newTextbox = wordApp.ActiveDocument.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 440, 180, 50, 40, Nothing);

            newTextbox.TextFrame.TextRange.Text = SharedData.currYear;
            newTextbox.TextFrame.TextRange.Font.Size = 14;
            newTextbox.TextFrame.TextRange.Font.Bold = 3;
            newTextbox.TextFrame.TextRange.Font.Name = "楷体_GB2312";

            //插入分页符
            wordDoc.Paragraphs.Last.Range.Select();
            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            wordApp.Application.Selection.InsertBreak(oPageBreak);
            wordApp.Selection.TypeText(" ");
            //插入分页符
            wordDoc.Paragraphs.Last.Range.Select();
            wordApp.Application.Selection.InsertBreak(oPageBreak);

            //保存
            string name = caption + ".doc";
            string filename = SharedData.currentFilePath + name;   //文件保存路径

            try
            {
                wordDoc.SaveAs(filename);
            }
            catch (Exception err)
            {
                Log.RecordLog(err.Message);
            }
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        private void InitDirectory()
        {
            //清空目录
            string path = @"C:\DFS\word\全国";
            CreatePath(path);
            path = @"C:\DFS\word\中央";
            CreatePath(path);
            path = @"C:\DFS\word\地方";
            CreatePath(path);
        }

        private void CreatePath(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
                Directory.CreateDirectory(path);
            }
            catch (Exception err)
            {
                Log.RecordLog("删除目录失败！" + err.ToString());
            }
        }

        //合并生成的word文件
        private void FileMerge()
        {
            System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();
            oTime.Start();

            //建立一个空word用以存放合并后的文件
            string filename = "";
            if (SharedData.AREATYPE == 3)//地方
            {
                filename = @"C:\DFS\国家重点科技基础条件资源调查数据汇编（地方）.doc";
            }
            else if (SharedData.AREATYPE == 2)//中央
            {
                filename = @"C:\DFS\国家重点科技基础条件资源调查数据汇编（中央）.doc";
            }
            else//全国
            {
                filename = @"C:\DFS\国家重点科技基础条件资源调查数据汇编（全国）.doc";
            }

            object Nothing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            Log.RecordLog("新建空白" + filename + "成功！等待合并word文档！");
            SharedData.fileName = filename;

            wordDoc.SaveAs2(filename);
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

            //合并各个word文件
            wordDocumentMerger merger = new wordDocumentMerger();
            List<string> fileList = new List<string>();
            fileList = GetFileList();
            merger.InsertMerge(fileList);

            oTime.Stop();
            long time = oTime.ElapsedMilliseconds;
            Log.RecordLog(@"word合并完成，最终生成文件为 " + SharedData.fileName + " 耗时：" + (time / 1000).ToString() + "秒！");
        }

        private List<string> GetFileList()
        {
            List<string> fileList = new List<string>();

            if (!Directory.Exists(SharedData.currentWordPath))
            {
                return null;
            }
            //获取指定目录下的非隐藏word文件
            foreach (string file in Directory.GetFileSystemEntries(SharedData.currentWordPath))
            {
                DirectoryInfo d = new DirectoryInfo(file);
                if ((d.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
                {
                    fileList.Add(d.ToString());
                }
            }
            fileList.Sort(delegate(string A, string B)
            {
                A = DataCompile.GetFileName(A);
                B = DataCompile.GetFileName(B);
                int A1 = -1, B1 = -1;
                int A2 = -1, B2 = -1;
                A1 = A.Trim()[0] - '0';
                if (A1 >= 0 && A1 <= 9)
                {
                    int i = A.Trim().IndexOf('-');
                    if (i != -1)
                    {
                        int j = A.Trim().IndexOf(' ');
                        if (j != -1)
                        {
                            A2 = int.Parse(A.Trim().Substring(i + 1, j - i - 1));
                            if (A2 <= 0 || A2 >= 100)
                            {
                                A2 = -1;
                            }
                        }
                    }
                    B1 = B.Trim()[0] - '0';
                    if (B1 >= 0 && B1 <= 9)
                    {
                        i = B.Trim().IndexOf('-');
                        if (i != -1)
                        {
                            int j = B.Trim().IndexOf(' ');
                            if (j != -1)
                            {
                                B2 = int.Parse(B.Trim().Substring(i + 1, j - i - 1));
                                if (B2 <= 0 || B2 >= 100)
                                {
                                    B2 = -1;
                                }
                            }
                        }
                    }
                    else
                    {
                        B1 = -1;
                    }
                }
                else
                {
                    A1 = -1;
                }

                if (A1 != -1 && B1 != -1 && A2 != -1 && B2 != -1)
                {
                    if (A1 < B1)
                    {
                        return -1;
                    }
                    else if (A1 == B1)
                    {
                        if (A2 < B2)
                        {
                            return -1;
                        }
                        else if (A2 == B2)
                        {
                            return A.CompareTo(B);
                        }
                        else
                        {
                            return 1;
                        }
                    }
                    else
                    {
                        return 1;
                    }
                }
                else
                {
                    return A.CompareTo(B);
                }
            });
            return fileList;
        }



        private void GetSumRowsAndColumns()
        {
            Log.RecordLog(dataCompile.Caption + ": GetSumRowsAndColumns");
            try
            {
                SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
                sqlConn.Open();
                DataSet ds = new DataSet();
                SqlCommand cmd = new SqlCommand(sqlStrForCompile, sqlConn);
                cmd.CommandTimeout = 180;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(ds);

                System.Data.DataTable dt = ds.Tables[0];
                if (dataCompile.SelectedRowName.Contains("所在"))
                {
                    sumRows = SharedData.areaListSorted.Count;
                }
                else
                {
                    sumRows = SharedData.deviceTypeSorted.Count;
                }
                if (dataCompile.SelectedRowName.Contains("仪器分类大类"))
                {
                    sumColumns = dt.Columns.Count - 1;
                }
                else
                {
                    sumColumns = dt.Columns.Count;
                }
                sqlConn.Close();
            }
            catch (System.Exception ex)
            {
                Log.RecordLog(ex.ToString());
            }
        }

        private List<string> GetAllColumnNamesFromServer(string tableName)
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
                names.Add(c.ColumnName);
            }
            sqlConn.Close();
            return names;
        }

        // 有多少续表
        private int GetSumTable()
        {
            int n = 0;
            for (int i = 0; i < researchInstitutionDataCompileList.Count(); i++)
            {
                if (researchInstitutionDataCompileList[i].Caption.Substring(0, 5).Equals(this.dataCompile.Caption.Substring(0, 5)))
                {
                    n++;
                }
            }
            return n;
        }

        // 续表号
        private int GetSeqTable()
        {
            int n = 0;
            for (int i = 0; i < researchInstitutionDataCompileList.Count() && i < index; i++)
            {
                if (researchInstitutionDataCompileList[i].Caption.Substring(0, 5).Equals(this.dataCompile.Caption.Substring(0, 5)))
                {
                    n++;
                }
            }
            return n;
        }



        public int index { get; set; }
    }
}
