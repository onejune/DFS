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
using System.Threading;
using System.Data.SqlClient;
using System.Diagnostics;

/// <summary>
/// Class1 的摘要说明
/// </summary>
public class OutputWord
{
    //计算程序执行时间
    System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();

    //创建Word文档
    private Object Nothing;
    private Word.Application wordApp;
    private Word.Document wordDoc;
    private object count = 1;

    //处理进度条的委托
    delegate void RunProgressBarDelegate();

    //所有出现的省份以及其仪器套数
    Dictionary<string, int> allTypeNames = new Dictionary<string, int>();
    int allDevicesCount = 0;

    //中央各部门名称
    Dictionary<string, int> centralDepartmentNames = new Dictionary<string, int>();
    //地方各省市名称
    Dictionary<string, int> localProvinceNames = new Dictionary<string, int>();
    //特殊城市
    Dictionary<string, string> specialCity = new Dictionary<string, string>() { 
        { "大连市", "辽宁省" },
        { "宁波市", "浙江省" },
        { "武汉市", "湖北省" },
        { "青岛市", "山东省" },
        { "厦门市", "福建省" },
        { "成都市", "四川省" },
        { "济南市", "山东省" },
        { "西安市", "陕西省" },
        { "深圳市", "广东省" },
        { "新疆建设兵团", "新疆维吾尔自治区" }
    };

    //中央部门顺序
    static List<string> centralDepartmentSorted = new List<string>{
            "最高人民检察院",
            "国家档案局",
            "国家民族事务委员会",
            "民政部",
            "司法部",
            "人力资源和社会保障部",
            "农业部",
            "水利部",
            "住房和城乡建设部",
            "国土资源部",
            "工业和信息化部",
            "铁道部",
            "交通运输部",
            "文化部",
            "教育部（国家语言文字工作委员会）",
            "卫生部",
            "国家人口和计划生育委员会",
            "中国气象局",
            "中国民用航空局",
            "国家海洋局",
            "中国地震局",
            "国家新闻出版总署（国家版权局）",
            "国家质量监督检验检疫总局",
            "国家广播电影电视总局",
            "国家林业局",
            "国务院侨务办公室",
            "中华全国供销合作总社",
            "国家邮政局",
            "国家粮食局",
            "国家安全生产监督管理总局",
            "国家体育总局",
            "国家文物局",
            "国家烟草专卖局",
            "国家测绘局",
            "环境保护部（国家核安全局）",
            "国家中医药管理局",
            "中国科学院",
            "中国社会科学院",
            "社科院",
            "中国残疾人联合会"
        };

    //用户选项参数
    private List<string> displayedColumns = new List<string>(); //需要显示的列
    private string groupName = "所在省"; //按什么分类？默认按区域(所在省)分类
    private string[] defaultColumns = new string[] { "中文名称", "型号规格", "仪器类型", "产地", "单位名称", "联系人", "联系电话" };
    private int fromValue;
    private int toValue;
    float[] columnsWidth;
    private string strSql = "";
    //按单位名称分组用于统计单位名称及数量
    private string sqlForOrgNames = "";

    [DllImport("shell32.dll ")]
    public static extern int ShellExecute(IntPtr hwnd, String lpszOp, String lpszFile, String lpszParams, String lpszDir, int FsShowCmd);

    public OutputWord(List<string> columns, string groupName, Form form, int fromValue, int toValue, float[] columnsWidth)
    {
        //创建Word文档
        Nothing = System.Reflection.Missing.Value;
        wordApp = new Word.Application();
        wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

        //默认的列名为
        if (columns.Count == 0 || columns.Count > 7)
        {
            for (int i = 0; i < defaultColumns.Length; i++)
            {
                displayedColumns.Add(defaultColumns[i]);
            }
        }
        this.displayedColumns = columns;

        //分类名称映射
        if (groupName.Equals("仪器类型"))
        {
            groupName = "仪器分类大类";
        }
        else if (groupName.Equals("区域"))
        {
            groupName = "所在省";
        }
        else if (groupName.Equals("隶属关系"))
        {
            groupName = "隶属";
        }
        else
        {
            groupName = "基地大类-汇编用";
        }
        this.groupName = groupName;
        this.fromValue = fromValue;
        this.toValue = toValue;
        this.columnsWidth = columnsWidth;
    }

    public void CreateWordFile()
    {
        string message = "";
        int i = 0;
        long time = 0;

        try
        {
            //开始计时
            oTime.Reset();
            oTime.Start();

            //清空目录
            string path = @"C:\DFS\word";
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


            //对于按照隶属关系分类，需要另作处理
            if (groupName.Equals("隶属"))
            {
                //获取数据来源名称和中央各部门名称
                GetDataSourceTypeAndDepartmentName();
                
                //求总设备台数
                foreach (KeyValuePair<string, int> p in localProvinceNames)
                {
                    allDevicesCount += p.Value;
                }
                foreach (KeyValuePair<string, int> p in centralDepartmentNames)
                {
                    allDevicesCount += p.Value;
                }

                //地方各省市
                RunThreadToCreatePages(localProvinceNames);
                //中央各部门
                RunThreadToCreatePages(centralDepartmentNames);

                CreateBlankPageFile("中央各部门");
                CreateBlankPageFile("地方各省市");

                //插入汇总表页
                CreateSummaryTablePageFile();
            }
            else
            {
                //根据分类名获取所有的类别名称，比如按区域分类则获取所有的省份
                getAllTypeNames();
                //计算所有设备台数
                foreach (KeyValuePair<string, int> p in allTypeNames)
                {
                    allDevicesCount += p.Value;
                }
                RunThreadToCreatePages(allTypeNames);
            }

            //停止计时
            oTime.Stop();
            //总共用时
            time = oTime.ElapsedMilliseconds / 1000;
        }
        catch (Exception ee)
        {
            message = ee.ToString();
            Log.RecordLog(message);
        }
        Log.RecordLog("word生成完成，耗时" + time.ToString() + "秒！");

        //合并word文件
        oTime.Reset();
        oTime.Start();

        //合并生成的所有word文件
        Merge();

        SharedData.isCompleted = true;
        oTime.Stop();
        long time3 = oTime.ElapsedMilliseconds / 1000;
        Log.RecordLog("word合并成功，耗时 " + time3.ToString() + " 秒！ ");

        //KillProcess("WINWORD.EXE");
        //关闭word进程
        DBHelper.Execute("taskkill /f /im winword.exe");
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

    //如果是按隶属分类，插入汇总表页
    private void CreateSummaryTablePageFile()
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
        wordDoc.SpellingChecked = false;
        wordDoc.ShowSpellingErrors = false;

        InsertSummaryTable(wordDoc, wordApp);

        //保存
        string name = "汇总表.doc";
        string filename = @"C:/DFS/word/" + name;   //文件保存路径

        try
        {
            wordDoc.SaveAs2(filename);
        }
        catch (Exception err)
        {
            Log.RecordLog(err.Message);
        }
        wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
        wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
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
        wordDoc.SpellingChecked = false;
        wordDoc.ShowSpellingErrors = false;

        //设置布局方向
        wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;
        //wordDoc.Paragraphs.First.Range.Select();

        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落

        wordApp.Selection.TypeText(caption);

        // 设置标题样式，以方便插入目录
        object oStyleName = "标题 1";
        wordApp.ActiveDocument.Styles[oStyleName].Font.Size = 29;
        wordApp.ActiveDocument.Styles[oStyleName].Font.Name = "微软雅黑";
        wordApp.ActiveDocument.Styles[oStyleName].Font.Bold = 2;
        wordApp.Selection.set_Style(ref oStyleName);
        wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落
        wordApp.Selection.TypeParagraph();//插入段落

        //插入分页符
        wordDoc.Paragraphs.Last.Range.Select();
        object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
        wordApp.Application.Selection.InsertBreak(oPageBreak);

        //插入分页符
        wordDoc.Paragraphs.Last.Range.Select();
        wordApp.Application.Selection.InsertBreak(oPageBreak);

        //保存
        string name = caption + ".doc";
        string filename = @"C:/DFS/word/" + name;   //文件保存路径

        try
        {
            wordDoc.SaveAs2(filename);
        }
        catch (Exception err)
        {
            Log.RecordLog(err.Message);
        }
        wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
        wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
    }

    private void Merge()
    {
        System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();
        oTime.Start();

        //建立一个空word用以存放合并后的文件
        string filename = @"C:\DFS\" + SharedData.excelFile.Substring(SharedData.excelFile.LastIndexOf("\\") + 1,
            (SharedData.excelFile.LastIndexOf(".") - SharedData.excelFile.LastIndexOf("\\") - 1)) + ".doc";
        object Nothing = System.Reflection.Missing.Value;
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        wordDoc.SpellingChecked = false;
        wordDoc.ShowSpellingErrors = false;


        Log.RecordLog("新建空白" + filename + "成功！等待合并word文档！");
        SharedData.fileName = filename;

        wordDoc.SaveAs2(filename);
        wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
        wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

        //合并各个word文件
        wordDocumentMerger merger = new wordDocumentMerger();
        merger.InsertMerge(groupName);

        oTime.Stop();
        long time = oTime.ElapsedMilliseconds;
        Log.RecordLog(@"word合并完成，最终生成文件为 " + SharedData.fileName + " 耗时：" + (time / 1000).ToString() + "秒！");
        //MessageBox.Show("word合并完成，耗时：" + (time / 1000).ToString() + "秒！");
    }

    //按隶属分类时插入汇总表
    private void InsertSummaryTable(Microsoft.Office.Interop.Word.Document wordDoc, Microsoft.Office.Interop.Word.Application wordApp)
    {
        //设置布局方向
        wordApp.Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;
        wordApp.Selection.PageSetup.TopMargin = wordApp.CentimetersToPoints(2.54F);
        wordApp.Selection.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2.54F);
        wordApp.Selection.PageSetup.LeftMargin = wordApp.CentimetersToPoints(3.18F);
        wordApp.Selection.PageSetup.RightMargin = wordApp.CentimetersToPoints(3.18F);

        //中央各部门
        wordDoc.Paragraphs.First.Range.Select();
        //wordApp.Selection.TypeParagraph();
        wordApp.Selection.Font.Size = 15F;
        wordApp.Selection.Font.Bold = 2;
        wordApp.Selection.TypeText("中央级各部门大型科学仪器数量汇总表");
        wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
  
        wordApp.Selection.TypeParagraph();
        wordApp.Selection.TypeParagraph();

        Microsoft.Office.Interop.Word.Table newTable = wordDoc.Tables.Add(wordApp.Selection.Range, centralDepartmentNames.Count + 2,
                3, ref Nothing, ref Nothing);

        //设置表格样式
        newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
        newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
        newTable.Range.Font.Size = 12F;
        newTable.Range.Font.Bold = 0;
        newTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


        //设置表格列宽
        float sumWidth = newTable.Columns[1].Width * 3;
        newTable.Columns[1].Width = 60f;
        newTable.Columns[3].Width = 60f;
        newTable.Columns[2].Width = sumWidth - 120;

        //设置行高
        newTable.Rows.Height = 17f;
        newTable.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly; //固定行高

        Range range = newTable.Rows[1].Range;

        range.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
        range.Font.Bold = 1;
        range.Font.Size = 11f;
        range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        range.Cells[1].Range.Text = "序号";
        range.Cells[2].Range.Text = "中央级部门";
        range.Cells[3].Range.Text = "数量";

        int i = 1;
        int sum = 0;
        foreach (string name in centralDepartmentSorted)
        {
            if (centralDepartmentNames.ContainsKey(name))
            {
                range = newTable.Rows[i + 1].Range;
                range.Font.Size = 11f;
                range.Cells[1].Range.Text = i.ToString();
                range.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                range.Cells[2].Range.Text = name;
                range.Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                range.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                range.Cells[3].Range.Text = centralDepartmentNames[name].ToString();
                range.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                sum += centralDepartmentNames[name];
                i++;
            }
        }
        i++;
        newTable.Rows[i].Range.Cells[1].Range.Text = "总计";
        newTable.Rows[i].Range.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        newTable.Cell(i, 3).Range.Text = sum.ToString();
        newTable.Cell(i, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        newTable.Cell(i, 3).Range.Font.Size = 12;
        newTable.Cell(i, 1).Range.Font.Bold = 3;
        newTable.Cell(i, 1).Range.Font.Size = 12;
        newTable.Cell(i, 3).Range.Font.Bold = 3;

        object count = newTable.Rows.Count;
        object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//按行移动
        wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点

        //插入分页符
        object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
        wordApp.Application.Selection.InsertBreak(oPageBreak);

        wordApp.ActiveDocument.Paragraphs.Last.Range.Select();
        wordApp.Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;

        //地方各省市
        wordApp.Selection.TypeText("地方各省市大型科学仪器数量汇总表");
        wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        wordApp.Selection.Font.Size = 15F;
        wordApp.Selection.Font.Bold = 2;
        wordApp.Selection.TypeParagraph();
        wordApp.Selection.TypeParagraph();

        newTable = wordDoc.Tables.Add(wordApp.Selection.Range, localProvinceNames.Count + 2,
                3, ref Nothing, ref Nothing);

        //设置表格样式
        newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
        newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
        newTable.Range.Font.Size = 11f;
        newTable.Range.Font.Bold = 0;

        //设置表格列宽
        newTable.Columns[1].Width = 60f;
        newTable.Columns[3].Width = 60f;
        newTable.Columns[2].Width = sumWidth - 120;

        //设置行高
        newTable.Rows.Height = 17f;
        newTable.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly; //固定行高

        range = newTable.Rows[1].Range;

        range.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
        range.Font.Bold = 1;
        range.Font.Size = 11f;
        range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        range.Cells[1].Range.Text = "序号";
        range.Cells[2].Range.Text = "地方各省份";
        range.Cells[3].Range.Text = "数量";

        i = 1;
        sum = 0;
        foreach (string name in SharedData.areaListSorted)
        {
            if (localProvinceNames.ContainsKey(name))
            {
                range = newTable.Rows[i + 1].Range;
                range.Font.Size = 11f;
                range.Cells[1].Range.Text = i.ToString();
                range.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                range.Cells[2].Range.Text = name;
                range.Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                range.Cells[3].Range.Text = localProvinceNames[name].ToString();
                range.Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                sum += localProvinceNames[name];
                i++;
            }
        }
        i++;
        newTable.Rows[i].Range.Cells[1].Range.Text = "总计";
        newTable.Rows[i].Range.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        newTable.Cell(i, 3).Range.Text = sum.ToString();
        newTable.Cell(i, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        newTable.Cell(i, 3).Range.Font.Size = 12;
        newTable.Cell(i, 1).Range.Font.Bold = 3;
        newTable.Cell(i, 1).Range.Font.Size = 12;
        newTable.Cell(i, 3).Range.Font.Bold = 3;

        wordApp.Selection.TypeParagraph();
        wordApp.ActiveDocument.Paragraphs.Last.Range.Select();

        //插入分页符
        oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
        wordApp.Application.Selection.InsertBreak(oPageBreak);

        Log.RecordLog(@"汇总表生成成功！ ");
    }

    private void RunThreadToCreatePages(Dictionary<string, int> typeNames)
    {
        string displayedColumnString = getDisplayedColumnString();
        List<Thread> unStartedThreadList = new List<Thread>();
        List<Thread> startedThreadList = new List<Thread>();
        int i = 0;

        //每个分类建立一个线程
        foreach (KeyValuePair<string, int> p in typeNames)
        {
            //debug
            if (!p.Key.Equals("上海市"))
            {
                continue;
            }

            i++;
            string statisticsString = GetStatisticString(p.Key);
            WordGeneration work = new WordGeneration(p.Key, p.Value, displayedColumns, groupName, statisticsString, columnsWidth, fromValue, toValue);
            if (groupName.Equals("隶属"))
            {
                work.SetExtraSQL(strSql);
                work.SetOrgNamesSQL(sqlForOrgNames);
            }
            Thread th = new Thread(new ThreadStart(work.GeneratePages));
            th.Name = "thread-" + p.Key;

            th.Start();
            startedThreadList.Add(th);
        }
        foreach (var v in startedThreadList)
        {
            v.Join();
        }

        // 重新运行失败的任务
        if (SharedData.failedWorkList.Count > 0)
        {
            //重新运行失败的任务
            ReRunFailedWork();
        }
    }

    private void GetDataSourceTypeAndDepartmentName()
    {
        int i = 0;
        //读取所有的类型名称
        SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
        sqlConn.Open();
        string sql = "";

        sql = @"select 数据来源, count(*) as '总数' from temp where 数据来源!='国家' group by 数据来源 order by 总数 desc";

        SqlCommand command = new SqlCommand(sql, sqlConn);
        SqlDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {
            string name = reader["数据来源"].ToString();
            if (name.Trim().Length != 0)
            {
                if (specialCity.ContainsKey(name))
                {
                    if (localProvinceNames.ContainsKey(specialCity[name]))
                    {
                        localProvinceNames[specialCity[name]] += int.Parse(reader["总数"].ToString());
                    }
                    else
                    {
                        localProvinceNames.Add(specialCity[name], int.Parse(reader["总数"].ToString()));
                    }
                }
                else
                {
                    if (!localProvinceNames.ContainsKey(name))
                    {
                        localProvinceNames.Add(name, int.Parse(reader["总数"].ToString()));
                    }
                    else
                    {
                        localProvinceNames[name] += int.Parse(reader["总数"].ToString());
                    }
                }
            }
        }
        reader.Close();
        //从高到低排序
        localProvinceNames = DictionarySort(localProvinceNames);

        sql = @"select 上级行政主管部门, count(*) as '总数' from temp where 数据来源='国家' group by 上级行政主管部门 order by 总数 desc";
        command = new SqlCommand(sql, sqlConn);
        reader = command.ExecuteReader();
        while (reader.Read())
        {
            string name = reader["上级行政主管部门"].ToString();
            if (name.Trim().Length != 0 && !localProvinceNames.ContainsKey(name))
            {
                centralDepartmentNames.Add(name, int.Parse(reader["总数"].ToString()));
            }
        }
        reader.Close();
        centralDepartmentNames = DictionarySort(centralDepartmentNames);
        sqlConn.Close();
    }

    private Dictionary<string, int> DictionarySort(Dictionary<string, int> localProvinceNames)
    {
        List<KeyValuePair<string, int>> myList = new List<KeyValuePair<string, int>>(localProvinceNames);
        myList.Sort(delegate(KeyValuePair<string, int> s1, KeyValuePair<string, int> s2)
        {
            return s2.Value.CompareTo(s1.Value);
        });
        localProvinceNames.Clear();
        foreach (KeyValuePair<string, int> pair in myList)
        {
            localProvinceNames.Add(pair.Key, pair.Value);
        }
        return localProvinceNames;
    }

    // 重新运行失败的任务
    private void ReRunFailedWork()
    {
        Log.RecordLog("重新运行失败的任务.......................");
        int i = 0;
        List<Thread> startedThreadList = new List<Thread>();
        string typeName = "";
        int sumDevice = 0;
        if(groupName.Equals("隶属"))
        {
            if(localProvinceNames.ContainsKey(typeName))
            {
                sumDevice = localProvinceNames[typeName];
            }
            else
            {
                sumDevice = centralDepartmentNames[typeName];
            }
        }
        else
        {
            sumDevice = allTypeNames[typeName];
        }

        for (i = 0; i < SharedData.failedWorkList.Count; i++)
        {
            typeName = SharedData.failedWorkList[i];
            string statisticsString = GetStatisticString(typeName);
            WordGeneration work = new WordGeneration(typeName, sumDevice, displayedColumns, groupName, statisticsString, columnsWidth, fromValue, toValue);
            Thread th = new Thread(new ThreadStart(work.GeneratePages));
            if (groupName.Equals("隶属"))
            {
                work.SetExtraSQL(strSql);
                work.SetOrgNamesSQL(sqlForOrgNames);
            }
            th.Name = "thread-" + typeName;
            th.Start();
            startedThreadList.Add(th);
        }
        foreach (var v in startedThreadList)
        {
            v.Join();
        }
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
                str += "仪器分类大类,仪器分类中类,";
            }
        }
        return str.Substring(0, str.Length - 1);
    }

    //根据groupName获取所有的类型名称(按仪器类型和按区域分类）
    private void getAllTypeNames()
    {
        int i = 0;
        //读取所有的类型名称
        SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
        sqlConn.Open();
        string sql = "";
        string deviceType = "";

        if (groupName.Equals("仪器分类大类") && DBHelper.unselectedDeviceTypeList.Count != 0)
        {
            for (i = 0; i < DBHelper.unselectedDeviceTypeList.Count; i++)
            {
                if (DBHelper.unselectedDeviceTypeList[i] != null)
                {
                    deviceType += "'";
                    deviceType += DBHelper.unselectedDeviceTypeList[i];
                    deviceType += "',";
                }
            }
            deviceType = deviceType.Substring(0, deviceType.Length - 1);
            sql = @"select [" + groupName + "], count(*) as '总数' from temp where 仪器分类大类 NOT IN(" + deviceType + ") group by [" + groupName + "] order by 总数";
        }
        else
        {
            sql = @"select [" + groupName + "], count(*) as '总数' from temp group by [" + groupName + "] order by 总数";
        }
        SqlCommand command = new SqlCommand(sql, sqlConn);
        SqlDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {
            string name = reader[groupName].ToString();
            if (name.Trim().Length != 0)
            {
                allTypeNames.Add(reader[groupName].ToString(), int.Parse(reader["总数"].ToString()));
            }
        }
        reader.Close();
        sqlConn.Close();
        SharedData.groupCount = allTypeNames.Count;
    }

    // 获取某个类型的数据统计字符串
    private string GetStatisticString(string typeName)
    {
        int currentDevicesCount = 0;
        decimal d ;
        string percent;
        string statistics = "";
        string sql = "";
        string prefixString = "";

        if (DBHelper.only985)
        {
            prefixString += "“985工程”高校";
        }
        else if (DBHelper.only211)
        {
            prefixString += "“211工程”高校";
        }
        if (fromValue == 0)
        {
            prefixString += "原值50万以上的";
        }
        else if (fromValue != 0)
        {
            prefixString += "原值" + fromValue / 10000 + "万以上的";
        }

        if (groupName.Equals("所在省"))
        {
            currentDevicesCount = allTypeNames[typeName];
            d = Math.Round((decimal)(currentDevicesCount * 100) / allDevicesCount, 1);
            percent = d.ToString() + "%";
            statistics = "    " + prefixString + "对外提供共享服务的大型科学仪器设备总量为" + allDevicesCount + "台（套），其中" + typeName + "的仪器的数量为" + currentDevicesCount + "台（套），占总量的"
                 + percent + "。" + typeName + "大型科学仪器中，";
            sql = @"Select 仪器分类大类,count(*) as '总数' from temp where 所在省='" + typeName + "' group by 仪器分类大类 order by 总数 desc";
        }
        else if (groupName.Equals("仪器分类大类"))
        {
            currentDevicesCount = allTypeNames[typeName];
            d = Math.Round((decimal)(currentDevicesCount * 100) / allDevicesCount, 1);
            percent = d.ToString() + "%";
            statistics = "    " + prefixString + "对外提供共享服务的大型科学仪器设备总量为" + allDevicesCount + "台（套），其中" + typeName + "的数量为" + currentDevicesCount + "台（套），占总量的"
                 + percent + "。" + typeName + "中，";
            sql = @"Select 仪器分类中类,count(*) as '总数' from temp where 仪器分类大类='" + typeName + "' group by 仪器分类中类 order by 总数 desc";
        }
        else if (groupName.Equals("隶属"))
        {
            if (centralDepartmentNames.ContainsKey(typeName))
            {
                currentDevicesCount = centralDepartmentNames[typeName];
            }
            else
            {
                currentDevicesCount = localProvinceNames[typeName];
            }
            d = Math.Round((decimal)(currentDevicesCount * 100) / allDevicesCount, 1);
            percent = d.ToString() + "%";
            statistics = "    " + prefixString + "对外提供共享服务的大型科学仪器设备总量为" + allDevicesCount + "台（套），其中" + typeName + "的仪器的数量为" + currentDevicesCount + "台（套），占总量的"
                 + percent + "。" + typeName + "大型科学仪器中，";
            //地市级
            if (localProvinceNames.ContainsKey(typeName))
            {
                if (!specialCity.ContainsValue(typeName))
                {
                    sql = @"Select 仪器分类大类,count(*) as '总数' from temp where 数据来源='" + typeName + "' group by 仪器分类大类 order by 总数 desc";
                    strSql = "select " + getDisplayedColumnStringWithAlias() + " from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                        "where a.仪器分类大类=b.仪器分类大类 and a.数据来源='" + typeName + "' order by b.sum desc";
                    if (DBHelper.groupToOrgName)
                    {                        
                        strSql = "select " + getDisplayedColumnStringWithAlias() + ",a.单位名称 from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                            "where a.仪器分类大类=b.仪器分类大类 and a.数据来源='" + typeName + "' order by a.单位名称, b.sum desc";
                        sqlForOrgNames = "select a.单位名称, count(*) as '总数' from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                            "where a.仪器分类大类=b.仪器分类大类 and a.数据来源='" + typeName + "' group by a.单位名称 order by a.单位名称";
                    }
                }
                else
                {
                    sql = @"Select 仪器分类大类,count(*) as '总数' from temp where 数据来源='" + typeName + "'";
                    strSql = "select " + getDisplayedColumnStringWithAlias() + ", a.单位名称 from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                        "where a.仪器分类大类=b.仪器分类大类 and (a.数据来源='" + typeName + "'";
                    sqlForOrgNames = "select a.单位名称, count(*) as '总数' from temp a where a.数据来源='" + typeName + "'";

                    //考虑计划单列市
                    foreach (KeyValuePair<string, string> p in specialCity)
                    {
                        if (p.Value.Equals(typeName))
                        {
                            sql += " or 数据来源='" + p.Key + "'";
                            strSql += " or 数据来源='" + p.Key + "'";
                            sqlForOrgNames += " or 数据来源='" + p.Key + "'";
                        }
                    }
                    sql += " group by 仪器分类大类 order by 总数 desc";
                    
                    if (DBHelper.groupToOrgName)
                    {
                        strSql += ") order by a.单位名称, b.sum desc";
                        sqlForOrgNames += " group by a.单位名称 order by a.单位名称";
                    }
                    else
                    {
                        strSql += ") order by b.sum desc";
                    }
                }
            }
            //中央各部门
            else
            {
                sql = @"Select 仪器分类大类,count(*) as '总数' from temp where 上级行政主管部门='" + typeName + "' group by 仪器分类大类 order by 总数 desc";
                strSql = "select " + getDisplayedColumnStringWithAlias() + ", a.单位名称 from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                       "where a.仪器分类大类=b.仪器分类大类 and a.上级行政主管部门='" + typeName + "' order by b.sum desc";
                if (DBHelper.groupToOrgName)
                {
                    strSql = "select " + getDisplayedColumnStringWithAlias() + ", a.单位名称 from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                       "where a.仪器分类大类=b.仪器分类大类 and a.上级行政主管部门='" + typeName + "' order by a.单位名称, b.sum desc";
                    sqlForOrgNames = "select a.单位名称, count(*) as '总数'  from temp a, (select COUNT(*) as sum,仪器分类大类 from temp group by 仪器分类大类) b " +
                       "where a.仪器分类大类=b.仪器分类大类 and a.上级行政主管部门='" + typeName + "' group by a.单位名称 order by a.单位名称";
                }
            }
        }
        else if (groupName.Equals("基地大类-汇编用"))
        {
            currentDevicesCount = allTypeNames[typeName];
            d = Math.Round((decimal)(currentDevicesCount * 100) / allDevicesCount, 1);
            percent = d.ToString() + "%";
            statistics = "    " + prefixString + "对外提供共享服务的大型科学仪器设备总量为" + allDevicesCount + "台（套），其中" + typeName + "的仪器的数量为" + currentDevicesCount + "台（套），占总量的"
                 + percent + "。" + typeName + "大型科学仪器中，";
            sql = @"Select 仪器分类大类,count(*) as '总数' from temp where [" + groupName + "]='" + typeName + "' group by 仪器分类大类 order by 总数 desc";
        }

        SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
        sqlConn.Open();
        SqlCommand command = new SqlCommand(sql, sqlConn);
        SqlDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {
            if (groupName.Equals("所在省"))
            {
                statistics += reader["仪器分类大类"].ToString();
            }
            else if (groupName.Equals("仪器分类大类"))
            {
                statistics += reader["仪器分类中类"].ToString();
            }
            else
            {
                statistics += reader["仪器分类大类"].ToString();
            }
            statistics += reader["总数"].ToString();
            statistics += "台（套），";
        }
        statistics = statistics.Substring(0, statistics.Length - 1);
        statistics += '。';

        reader.Close();
        sqlConn.Close();
        return statistics;
    }

    //将用户选择的column数组组合成sql字符串
    private string getDisplayedColumnStringWithAlias()
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
}
