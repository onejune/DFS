using System;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using DFSystem;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace DFSystem
{
    class WordGeneration
    {
        private string typeName;
        private int sumDevices;
        private object pages;
        private string sqlOrgNames;

        //分析仪器中仪器分类中类名称及数量
        Dictionary<string, int> secondTypeNames = new Dictionary<string, int>();
        //单位名称及数量
        Dictionary<string, int> organizationNames = new Dictionary<string, int>();

        //计算程序执行时间
        System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();

        //创建Word文档
        private Object Nothing;
        private Microsoft.Office.Interop.Word.Application wordApp;
        private Microsoft.Office.Interop.Word.Document wordDoc;
        private object count = 1;

        //用户选项参数
        private List<string> displayedColumns = new List<string>(); //需要显示的列
        private string groupName = "所在省"; //按什么分类？默认按所在省分类
        private int rowsPerPage = 0;
        float[] columnsWidth;
        private string statisticsString;
        private int fromValue;
        private int toValue;
        private string strSql = "";

        //单个分类开始生成word
        public WordGeneration(string typeName, int n, List<string> columns, string groupName, string statisticsString, float[] columnsWidth, int fromValue, int toValue)
        {
            //创建Word文档
            Nothing = System.Reflection.Missing.Value;
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordDoc.SpellingChecked = false;
            wordDoc.ShowSpellingErrors = false;

            this.displayedColumns = columns;
            this.groupName = groupName;
            this.typeName = typeName;
            this.columnsWidth = columnsWidth;
            this.fromValue = fromValue;
            this.toValue = toValue;
            sumDevices = n;

            this.statisticsString = statisticsString;
        }

        public void GeneratePages()
        {
            string message;
            Word.Table newTable = null;
            int i = 0;
            int time = 0;
            object pageCount = 0;
            string sql = "";
            string displayedColumnString = getDisplayedColumnString();
            Log.RecordLog("开始生成" + typeName + ".doc");

            try
            {
                setPage();
                insertPage();
                pageHeaderText = "";

                //是否对分析仪器按仪器分类中类再次分组
                if (DBHelper.groupToAnalyseDevice && typeName.Equals("分析仪器"))
                {
                    //分析仪器中仪器分类中类类型名称及数量
                    getSecondTypeNames();

                    foreach (KeyValuePair<string, int> p in secondTypeNames)
                    {
                        //如果需要按单位名称再分组的话先获取单位名称及仪器数
                        if (DBHelper.groupToOrgName)
                        {
                            getOrganizationNames(p.Key);
                        }
                        int rows = p.Value + p.Value / rowsPerPage + 1;
                        pageHeaderText = p.Key + "/" + p.Value + "台（套）";
                        newTable = insertTable(rows);

                        sql = "select " + displayedColumnString + ",a.单位名称 from temp a where a.仪器分类大类='分析仪器' and a.仪器分类中类='" + p.Key + "' order by a.仪器分类大类, a.仪器分类中类";
                        
                        //如果需要对单位名称分组，先按单位名称排序
                        if (DBHelper.groupToOrgName)
                        {
                            sql = "select " + displayedColumnString + ",a.单位名称 from temp a where a.仪器分类大类='分析仪器' and a.仪器分类中类='" + p.Key + "' order by a.单位名称, a.仪器分类大类, a.仪器分类中类";
                        }

                        fillTable(newTable, sql);
                        wordDoc.Paragraphs.Last.Range.Select();
                    }
                }
                else
                {
                    //如果需要按单位名称再分组的话先获取单位名称及仪器数
                    if (DBHelper.groupToOrgName)
                    {
                        getOrganizationNames(null);
                    }
                    int rows = sumDevices + sumDevices / rowsPerPage + 1;
                    
                    pageHeaderText = typeName + "/" + sumDevices + "台（套）";
                    newTable = insertTable(rows);

                    if (DBHelper.only985 || DBHelper.only211)
                    {
                        sql = "select " + displayedColumnString + ",a.单位名称, cast(CONVERT(DECIMAL(4,1), CONVERT(float,a.年有效工作机时)/1600*100) as varchar(10))+'%' as 利用率," +
                        "共享率=case when a.年有效工作机时!='0' then cast(CONVERT(DECIMAL(4,1), CONVERT(float,(CONVERT(float,a.年对外服务机时)/CONVERT(float,a.年有效工作机时))*100)) as varchar(10))+'%' else '0%' end" +
                        " from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                            "where a.仪器分类大类=b.仪器分类大类 and a.[" + groupName + "]='" + typeName + "' order by b.sum desc";
                        //如果需要对单位名称分组，先按单位名称排序
                        if (DBHelper.groupToOrgName)
                        {
                            sql = "select " + displayedColumnString + ",a.单位名称, cast(CONVERT(DECIMAL(4,1), CONVERT(float,a.年有效工作机时)/1600*100) as varchar(10))+'%' as 利用率," +
                       "共享率=case when a.年有效工作机时!='0' then cast(CONVERT(DECIMAL(4,1), CONVERT(float,(CONVERT(float,a.年对外服务机时)/CONVERT(float,a.年有效工作机时))*100)) as varchar(10))+'%' else '0%' end" +
                       " from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                           "where a.仪器分类大类=b.仪器分类大类 and a.[" + groupName + "]='" + typeName + "' order by a.单位名称, b.sum desc";
                        }
                    }
                    else
                    {
                        sql = "select " + displayedColumnString + ",a.单位名称 from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                            "where a.仪器分类大类=b.仪器分类大类 and a.[" + groupName + "]='" + typeName + "' order by b.sum desc";
                        //如果需要对单位名称分组，先按单位名称排序
                        if (DBHelper.groupToOrgName)
                        {
                            sql = "select " + displayedColumnString + ",a.单位名称 from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                            "where a.仪器分类大类=b.仪器分类大类 and a.[" + groupName + "]='" + typeName + "' order by a.单位名称, b.sum desc";
                        }
                    }
                    if(groupName.Equals("隶属"))
                    {
                        sql = strSql;
                    }
                    fillTable(newTable, sql);
                }
                InsertPageFooterNumber();  
                message = saveFile();
                //停止计时
                oTime.Stop();
                //总共用时
                time = oTime.Elapsed.Milliseconds;
                message = message + " /n共耗时：" + time / 60000 + "分钟！";
            }
            catch (Exception ee)
            {
                message = ee.ToString();
                //Console.WriteLine(typeName + ": GeneratePages :" + message);
                Log.RecordLog(typeName + ".doc 导出失败！");
                Log.RecordLog(typeName + ": GeneratePages :" + message);
                Log.RecordLog("查询语句为：" + sql);
                lock (SharedData.failedWorkList)
                {
                    SharedData.failedWorkList.Add(typeName);
                }
                return;
            }
            Log.RecordLog(typeName + ".doc 导出成功！");
        }

        public void SetExtraSQL(string sql)
        {
            strSql = sql;
        }

        //获取每个分类名称的单位名称及数量
        private void getOrganizationNames(string deviceType)
        {
            //读取所有的类型名称
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            string sql = "";

            if (DBHelper.groupToAnalyseDevice && typeName.Equals("分析仪器"))
            {
                sql = @"select 单位名称, count(*) as '总数' from temp where " + groupName + "='分析仪器' and 仪器分类中类='" + deviceType + "'group by 单位名称 order by 单位名称";
            }
            else if (groupName.Equals("隶属"))
            {
                sql = sqlOrgNames;
            }
            else
            {
                sql = @"select 单位名称, count(*) as '总数' from temp where [" + groupName + "]='" + typeName + "' group by 单位名称 order by 单位名称";
            }
            
            SqlCommand command = new SqlCommand(sql, sqlConn);
            SqlDataReader reader = command.ExecuteReader();
            organizationNames.Clear();
            while (reader.Read())
            {
                string name = reader["单位名称"].ToString();
                if (name.Trim().Length != 0)
                {
                    organizationNames.Add(name, int.Parse(reader["总数"].ToString()));
                }
            }
            reader.Close();
            sqlConn.Close();
        }

        //获取分析仪器中仪器分类中类类型名称
        private void getSecondTypeNames()
        {
            int i = 0;
            //读取所有的类型名称
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            string sql = "";

            sql = @"select 仪器分类中类, count(*) as '总数' from temp where 仪器分类大类='分析仪器' group by 仪器分类中类 order by 总数 desc";
          
            SqlCommand command = new SqlCommand(sql, sqlConn);
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                string name = reader["仪器分类中类"].ToString();
                if (name.Trim().Length != 0)
                {
                    secondTypeNames.Add(name, int.Parse(reader["总数"].ToString()));
                }
            }
            reader.Close();
            sqlConn.Close();
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

        // 对每个省对应的table进行数据填充
        private void fillTable(Microsoft.Office.Interop.Word.Table newTable, string sql)
        {
            int i = 1, k = 1, j = 0;
            //SQLiteDataReader reader = DBHelper.SqlInquiry("select count(*) as '总数', 仪器分类大类 from sheet1 where 所在省='" + typeName + "' group by 仪器分类大类");
            string displayedColumnString = getDisplayedColumnString();
            //SQLiteDataReader reader = DBHelper.SqlInquiry("select " + displayedColumnString + " from Sheet1 where 所在省='" + typeName + "'");     

            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            SqlCommand command = new SqlCommand(sql, sqlConn);
            SqlDataReader reader = command.ExecuteReader();
            string preName = "";
            int orgNameSeq = 1;

            while (reader.Read())
            {
                lock (SharedData.sequenceNumber)
                {
                    int n = (int)SharedData.sequenceNumber;
                    SharedData.sequenceNumber = ++n;
                }              

                //如果需要对单位名称分组，在不同的单位之间空3行，插入表头，并重新编号
                string currentOrgName = reader["单位名称"].ToString();
                if (DBHelper.groupToOrgName)
                {
                    if (!currentOrgName.Equals(preName))
                    {
                        preName = currentOrgName;
                        for (int e = 1; e <= displayedColumns.Count; e++)
                        {
                            newTable.Cell(k, 1).Merge(newTable.Cell(k, 2));
                            newTable.Cell(k + 1, 1).Merge(newTable.Cell(k + 1, 2));
                        }
                        newTable.Cell(k, 1).Range.Text = "";
                       
                        newTable.Cell(k, 1).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                        newTable.Cell(k, 1).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                        newTable.Cell(k + 1, 1).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                        newTable.Cell(k + 1, 1).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                        newTable.Cell(k + 1, 1).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;

                        //插入单位名称
                        newTable.Cell(k + 1, 1).Range.Text = orgNameSeq + ". " + currentOrgName;
                        object oStyleName = "标题 2";
                        wordApp.ActiveDocument.Styles[oStyleName].Font.Size = 9;
                        wordApp.ActiveDocument.Styles[oStyleName].Font.Name = "微软雅黑";
                        wordApp.ActiveDocument.Styles[oStyleName].Font.Bold = 1;
                        wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceBefore = 0;
                        wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceAfter = 0;
                        wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceAfterAuto = 0;
                        wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceBeforeAuto = 0;
                        wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                        newTable.Cell(k + 1, 1).Range.set_Style(ref oStyleName);

                        newTable.Cell(k + 1, 1).Range.Font.Bold = 1;
                        orgNameSeq++;
                        fillTableHeader(newTable.Rows[k + 2].Range);
                        if (k % rowsPerPage == 1)
                        {
                            newTable.Cell(k, 1).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                        }

                        k += 3;
                        i = 1;
                    }
                }
                //填充表头，每一页都有表头,rows从1开始编号
                if (k % rowsPerPage == 1)//表头
                {
                    if (k < newTable.Rows.Count)
                    {
                        fillTableHeader(newTable.Rows[k].Range);
                        int page = k / rowsPerPage;
                        k++;
                    }
                    else
                    {
                        break;
                    }
                }
                wordApp.Selection.Font.Size = 8f;
                newTable.Cell(k, 1).Range.Text = i.ToString();
                i++;

                for (j = 0; j < displayedColumns.Count; j++)
                {
                    newTable.Cell(k, j + 1).Range.Select();
                    if (!displayedColumns[j].Equals("仪器类型"))
                    {
                        string text = reader[displayedColumns[j]].ToString();
                        if(displayedColumns[j].Equals("所在省"))
                        {
                            text = text.Replace("省", "");
                            text = text.Replace("市", "");
                            text = text.Replace("自治区", "");
                            text = text.Replace("维吾尔", "");
                            text = text.Replace("回族", "");
                            text = text.Replace("壮族", "");
                        }
                        newTable.Cell(k, j + 2).Range.Text = text;
                    }
                    else
                    {
                        newTable.Cell(k, j + 2).Range.Text = reader["仪器分类大类"].ToString() + "/" + reader["仪器分类中类"].ToString();
                    }
                }
                if (DBHelper.only985 || DBHelper.only211)
                {
                    newTable.Cell(k, j + 2).Range.Text = reader["共享率"].ToString();
                    newTable.Cell(k, j + 3).Range.Text = reader["利用率"].ToString();
                }
                if (k % rowsPerPage != 1)
                {
                    k++;
                }
                //Console.WriteLine(typeName + "——fillTable:" + i);
            }
            reader.Close();
            sqlConn.Close();
        }

        //设置表头的内容和格式
        private void fillTableHeader(Range range)
        {
            //range.Rows.Height = 15f;
            range.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
            range.Font.Bold = 1;
            //range.Font.Size = 9f;
            int i = 0;
            range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            range.Cells[1].Range.Text = "序号";
            for (i = 0; i < displayedColumns.Count; i++)
            {
                if (displayedColumns[i].Equals("所在省"))
                {
                    range.Cells[i + 2].Range.Text = "所在地";
                }
                else
                {
                    range.Cells[i + 2].Range.Text = displayedColumns[i];
                }
            }

            if (DBHelper.only211 || DBHelper.only985)
            {
                range.Cells[i + 2].Range.Text = "共享率";
                range.Cells[i + 3].Range.Text = "利用率";
            }
            //插入页眉
            //wordDoc.ActiveWindow.View.Type = WdViewType.wdPrintView;
            wordDoc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.LinkToPrevious = false;

            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.Range.Text = pageHeaderText;
            wordDoc.ActiveWindow.ActivePane.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//跳出页眉设置
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
                    str += "a." + displayedColumns[i] + ",";
                }
                else
                {
                    str += "a.仪器分类大类,a.仪器分类中类,";
                }
            }
            return str.Substring(0, str.Length - 1);
        }

        // 插入插页
        private void insertPage()
        {
            pageHeaderText = "";

            //插入页眉
            wordDoc.ActiveWindow.View.Type = WdViewType.wdPrintView;
            wordDoc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.LinkToPrevious = false;

            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.Range.Text = pageHeaderText;
            //wordDoc.ActiveWindow.ActivePane.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//跳出页眉设置

            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;

            //设置“区域”字符串的对齐方式并输出区域
            wordApp.Selection.TypeText(typeName);
            // 设置标题样式，以方便插入目录
            object oStyleName = "标题 1";
            wordApp.ActiveDocument.Styles[oStyleName].Font.Size = 27;
            wordApp.ActiveDocument.Styles[oStyleName].Font.Name = "微软雅黑";
            wordApp.ActiveDocument.Styles[oStyleName].Font.Bold = 1;
            wordApp.Selection.set_Style(ref oStyleName);
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落
            wordApp.Selection.TypeParagraph();//插入段落

            count = 7;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//按行移动
            wordApp.Selection.MoveUp(ref WdLine, ref count, ref Nothing);//移动焦点

            // 插入文本框
            Shape newTextbox;
            newTextbox = wordDoc.Application.ActiveDocument.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 170, 170, 520, 300, ref Nothing);

            newTextbox.TextFrame.TextRange.Text = statisticsString;
            newTextbox.TextFrame.TextRange.Font.Size = 18f;
            newTextbox.TextFrame.TextRange.Font.Name = "楷体_GB2312";

            //插入分页符
            wordDoc.Paragraphs.Last.Range.Select();
            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            wordApp.Application.Selection.InsertBreak(oPageBreak);
            wordDoc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.LinkToPrevious = false;

            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.Range.Text = pageHeaderText;
            //wordDoc.ActiveWindow.ActivePane.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//跳出页眉设置
            wordApp.Application.Selection.InsertBreak(oPageBreak);
            wordDoc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.LinkToPrevious = false;

            wordDoc.ActiveWindow.ActivePane.Selection.HeaderFooter.Range.Text = pageHeaderText;
            //wordDoc.ActiveWindow.ActivePane.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//跳出页眉设置
            wordApp.Selection.Font.Size = 8f;
            wordApp.Selection.Font.Bold = 0;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        //文件保存
        private string saveFile()
        {
            string name = typeName + ".doc";
            string filename = @"C:/DFS/word/" + name;   //文件保存路径
            while (true)
            {
                int m_iErrCnt = 0;
                try
                {
                    // 选中当前文档的末尾,保证插页在奇数页
                    wordDoc.Paragraphs.Last.Range.Select();

                    int pageNumber = (int)wordApp.Selection.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
                    if (pageNumber % 2 == 0)
                    {
                        object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
                        wordApp.Application.Selection.InsertBreak(oPageBreak);
                    }   
                    wordDoc.SaveAs2(filename);
                    break;
                }
                catch (Exception err)
                {
                    m_iErrCnt++;
                    if (m_iErrCnt < 10)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
                        throw err;
                    }
                    Log.RecordLog(err.Message);
                }

            }
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            string message = "文档生成成功，已保存到C:\\DFS\\word下";
            return message;
        }

        // 页面设置
        private void setPage()
        {
            Directory.CreateDirectory(@"C:/DFS/word");   //创建文件所在目录
            wordApp.Selection.TypeParagraph();
            //设置布局方向
            wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;
            //设置纸张类型
            wordDoc.PageSetup.PageWidth = wordApp.CentimetersToPoints(29.7F);
            wordDoc.PageSetup.PageHeight = wordApp.CentimetersToPoints(21F);
            //设置页边距
            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.0F);
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(0.5F);
            wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1.1F);
            wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.1F);
            //页眉距离
            wordDoc.PageSetup.HeaderDistance = wordApp.CentimetersToPoints(1.3F);

            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.HeaderFooter.LinkToPrevious = false;

            //wordApp.ActiveWindow.ActivePane.Selection.InsertAfter(typeName + "/" + sumDevices + "台（套）");
            //wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//跳出页眉设置

            //页脚距离
            wordDoc.PageSetup.FooterDistance = wordApp.CentimetersToPoints(0.9F);

            wordApp.Selection.ParagraphFormat.LineSpacing = 10f;//设置文档的行间距
            //wordApp.Selection.Font.Name = "黑体";   //设置字体   
            wordApp.Selection.Font.Size = 8f;
            //wordApp.Selection.Font.Bold = 1;

            //每页行数
            rowsPerPage = 40;
        }

        // 根据用户选择的列数插入空白表格
        private Table insertTable(int rows)
        {
            //如果需要对单位名称分组，需要添加额外行
            if (DBHelper.groupToOrgName)
            {
                rows += organizationNames.Count * 3;
            }
            //文档中创建表格
            int columns = displayedColumns.Count + 1;
            if (DBHelper.only211 || DBHelper.only985)
            {
                columns += 2;
            }
            Microsoft.Office.Interop.Word.Table newTable = wordDoc.Tables.Add(wordApp.Selection.Range, rows,
                columns, ref Nothing, ref Nothing);

            //设置表格样式
            newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            newTable.Range.Font.Size = 9f;

            //设置表格列宽
            float sumWidth = 0;
            newTable.Columns[1].Width = 35f;
            if (DBHelper.only985 || DBHelper.only211)
            {
                newTable.Columns[newTable.Columns.Count].Width = 40f;
                newTable.Columns[newTable.Columns.Count - 1].Width = 40f;
            }
            for (int i = 2; i <= displayedColumns.Count; i++)
            {
                sumWidth += newTable.Columns[i].Width;
            }
            for (int i = 2; i <= displayedColumns.Count; i++)
            {
                newTable.Columns[i].Width = 630 * columnsWidth[i - 2];
            }

            //设置行高
            newTable.Rows.Height = 12.5f;
            newTable.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly; //固定行高

            wordApp.ActiveDocument.Paragraphs.Last.Range.Select();

            //插入分页符
            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            wordApp.Application.Selection.InsertBreak(oPageBreak);   
            return newTable;
        }

        //插入页脚页边距
        private void InsertPageFooterNumber()
        {
            try
            {
                wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader;
                wordApp.Selection.WholeStory();
                wordApp.Selection.ParagraphFormat.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone;

                wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;
                wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter;
                //wordApp.Selection.TypeText("第");
                object page = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
                wordApp.Selection.Fields.Add(wordApp.Selection.Range, ref page, ref Nothing, ref Nothing);
                wordApp.Selection.TypeText("/");
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                pages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages;
                wordApp.Selection.Fields.Add(wordApp.Selection.Range, ref pages, ref Nothing, ref Nothing);

                wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;

            }
            catch (Exception ex)
            {
                Log.RecordLog(ex.Message);
            }
        }

        private void InsertPageFooter()
        {
            try
            {
                object page = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
                wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter;
                int p = (int)page;
                wordApp.Selection.InsertAfter(p.ToString());
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }
            catch (Exception ex)
            {
                Log.RecordLog(ex.Message);
            }
        }

        public string pageHeaderText = "";

        public void SetOrgNamesSQL(string sqlForOrgNames)
        {
            sqlOrgNames = sqlForOrgNames;
        }
    }
}
