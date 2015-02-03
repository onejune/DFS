using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace DFSystem
{
    abstract class AbstractTable
    {
        //表标题
        protected string tableCaption;
        //保存的文件名
        protected string fileName;
        //单位
        protected string unitName;
        //类别
        protected string tableCategory;
        //listA中的各列
        protected List<string> finalListA = new List<string>();
        protected List<string> finalListB = new List<string>();
        //第一个汇总列
        protected string summaryColumnA;
        protected string summaryColumnB;
        //当无法显示全部列时，只显示部分列+“其他”，如2-5
        private string otherColumn = "";

        //是否需要转成“万元”
        private bool isNeedTransTenThousand = true;

        public bool IsNeedTransTenThousand
        {
            get { return isNeedTransTenThousand; }
            set { isNeedTransTenThousand = value; }
        }

        //数据库中的列，用于检查列名
        protected List<string> columnCheckList = new List<string>();
        //表格高度
        protected float maxTableHeight = 646F;
        //数据库表名
        protected string dataTableName;
        //sql
        protected string sqlStrForCompile;

        protected DataCompile dataCompile = null;
        protected int sumMaxLine = 0;
        protected int sumMinLine = 0;
        protected List<double> sumList = new List<double>();
        protected List<double> assistSumList = new List<double>();
        protected bool tableIsExisted = true;
        protected int sumRows;
        protected int sumColumns;
        protected int seqNo ;
        protected int sumSameTable ;

        //创建Word文档
        protected Object Nothing;
        protected Microsoft.Office.Interop.Word.Application wordApp;
        protected Microsoft.Office.Interop.Word.Document wordDoc;

        //错误信息
        private bool hasError = false;

        public bool HasError
        {
            get { return hasError; }
            set { hasError = value; }
        }

        public void RunJob()
        {
            bool res = true;

            if (hasError == true)
            {
                return;
            }

            //检查数据表是否存在
            if (!CheckTableIsExisted())
            {
                return;
            }

            InitWordPage();
            res = GenerateSQL();

            if (res == false)
            {
                return;
            }

            if (dataCompile != null)
            {
                //获取列数和行数
                GetSumRowsAndColumns();

                //初始化每一列总计列表
                for (int i = 0; i < sumColumns; i++)
                {
                    sumList.Add(0);
                    assistSumList.Add(0);
                }
                if (dataCompile != null && dataCompile.SelectedRowName.Equals("仪器分类大类"))
                {
                    sumMaxLine = 33;
                    sumMinLine = 6;
                }
                else
                {
                    sumMaxLine = 39;
                    sumMinLine = 4;
                }
            }

            //插入空白表
            List<Table> blankTableList = InsertTable();

            //填充数据
            FillTable(blankTableList);

            InsertPageFooterNumber();
            string message = SaveFile();
        }

        public bool CheckTableIsExisted()
        {
            int i = 0;
            for (i = 0; i < SharedData.tableListInDB.Count(); i++)
            {
                if ((SharedData.tableListInDB[i].TableName.Equals(dataTableName)) && (SharedData.tableListInDB[i].IsExisted == true))
                {
                    return true;
                }
            }
            Log.RecordLog("表 " + dataTableName + " 不存在！");
            return false;
        }

        //检查列名是否存在
        public bool CheckColumnNames(List<string> columnList)
        {
            //test
            return true;

            bool res = true, isExisted = true;
            int tableIndex = 0;
            for (tableIndex = 0; tableIndex < SharedData.tableListInDB.Count; tableIndex++)
            {
                if (SharedData.tableListInDB[tableIndex].ViewName.Equals(dataTableName))
                {
                    break;
                }
            }
            List<string> columnsInDB = SharedData.tableListInDB[tableIndex].ColumnNames;

            for (int i = 0; i < columnList.Count; i++)
            {
                res = columnsInDB.Contains(columnList[i]);
                if (res == false)
                {
                    Log.RecordLog(tableCaption + ": 列名检查失败！'" + columnList[i] +"'不存在！");
                    isExisted &= res;
                }
            }
            return isExisted;
        }

        private void InitWordPage()
        {
            //创建Word文档
            Nothing = System.Reflection.Missing.Value;
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //设置页边距
            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1F);
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1F);
            wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2.5F);
            wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(2.5F);

            fileName = tableCaption;
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
                wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;
            }
            catch (Exception ex)
            {
                Log.RecordLog(ex.Message);
            }
        }

        //文件保存
        private string SaveFile()
        {
            string name = fileName + ".doc";
            if (dataCompile != null)
            {
                name = dataCompile.Caption + ".doc";
            }
            string filename = SharedData.currentFilePath + name;

            while (true)
            {
                int m_iErrCnt = 0;
                try
                {
                    wordDoc.SaveAs(filename);
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
                    Log.RecordLog(tableCaption + " SaveFile: " + err.Message);
                }

            }
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            string message = "文档生成成功，已保存到C:\\DFS\\word下";
            Log.RecordLog(tableCaption + ": " + message);
            return message;
        }

        protected List<string> GetFinalColumnNames(string name)
        {
            List<string> columnList = new List<string>();
            string sqlStr = @"select " + name + " from [" + dataCompile.TableName + "] group by " + name;
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();

            try
            {
                SqlCommand command = new SqlCommand(sqlStr, sqlConn);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (reader[name] != null && reader[name].ToString().Length > 0)
                    {
                        columnList.Add(reader[name].ToString());
                    }
                }
                reader.Close();
                sqlConn.Close();
            }
            catch (System.Exception ex)
            {
                Log.RecordLog(ex.ToString());
            }
            return columnList;
        }

        //生成sql语句
        public virtual bool GenerateSQL()
        {
            Log.RecordLog(tableCaption + ": GenerateSQL");

            //判断列是否存在
            if (!CheckColumnNames(columnCheckList))
            {
                return false;
            }

            List<string> list = dataCompile.SelectedColumnListA;
            List<string> listA = GetFinallyColumnNames(list, 'A');
            List<string> listB = null;
            string sumOrCount;

            string where = "";
            string and = "";

            int i = 0;
            List<string> columnNamesOfTable = GetAllColumnNamesFromServer();

            if (dataCompile.SelectedColumnListA[0].Equals("仪器分类中类"))
            {
                where = " where 仪器分类大类='分析仪器' ";
            }
            if (dataCompile.SelectedColumnListA[0].Equals("全国汇编用基地小类"))
            {
                where = " where 基地大类='" + dataCompile.SummaryColumnNameA + "'";
            }
            if (dataCompile.SummaryColumnNameA.Equals("植物种质资源") || dataCompile.SummaryColumnNameA.Equals("动物种质资源")
                || dataCompile.SummaryColumnNameA.Equals("微生物菌种资源"))
            {
                where = " where 资源领域='" + dataCompile.SummaryColumnNameA + "' ";
                and = " and 资源领域='" + dataCompile.SummaryColumnNameA + "' ";
            }
            if (dataCompile.SumColumnName.Equals(""))
            {
                sumOrCount = " count(*) ";
            }
            else
            {
                sumOrCount = " sum(cast(\"" + dataCompile.SumColumnName + "\" as float)) ";
                if (dataCompile.UnitName.Contains("亿元"))
                {
                    sumOrCount = " convert(bigint," + sumOrCount + "/ 10000) ";
                }
            }

            list = dataCompile.SelectedColumnListB;
            if (list != null && list.Count != 0)
            {
                listB = GetFinallyColumnNames(list, 'B');
            }
            finalListA = listA;
            finalListB = listB;

            string notInSet = "";

            // 第一种类型：行名是所在省
            if (dataCompile.SelectedRowName.Contains("所在"))
            {
                // 如果selectedColumnListA只指定了一个字段而且第一个字母为#，表示该字符串为最终列名，否则是以该列的内容进行汇总
                if (dataCompile.SelectedColumnListA.Count == 1 && dataCompile.SelectedColumnListA[0][0] != '#')
                {
                    string sqlMiddle = "";
                    string aliasStr = "";
                    char alias = 'a';

                    //处理listA的部分
                    for (i = 0; i < listA.Count; i++)
                    {
                        string column = listA[i];
                        if (!otherColumn.Equals("其他") || !column.Equals("其他"))
                        {
                            notInSet += "'" + column + "',";
                            sqlMiddle += "(select x.provinceName as \"" + dataCompile.SelectedRowName + "\", y.[" + column + "] from (select provinceName from [DFS].[dbo].[所在省]) x left join " +
                                "(select \"" + dataCompile.SelectedRowName + "\"," + sumOrCount + " as '" + column + "' from [" + dataCompile.TableName + "] where [" + dataCompile.SelectedColumnListA[0] + "] like '%" + column + "%'" + and +
                                " group by \"" + dataCompile.SelectedRowName + "\") y on x.provinceName like '%' + y.\"" + dataCompile.SelectedRowName + "\" + '%') " + alias + ",";
                        }
                        else
                        {
                            notInSet = notInSet.Substring(0, notInSet.Length - 1);
                            sqlMiddle += "(select x.provinceName as \"" + dataCompile.SelectedRowName + "\", y.[" + column + "] from (select provinceName from [DFS].[dbo].[所在省]) x left join " +
                                "(select \"" + dataCompile.SelectedRowName + "\"," + sumOrCount + " as '" + column + "' from [" + dataCompile.TableName + "] where [" + dataCompile.SelectedColumnListA[0] + "] not in(" + notInSet + ")" + and +
                                " group by \"" + dataCompile.SelectedRowName + "\") y on x.provinceName like '%' + y.\"" + dataCompile.SelectedRowName + "\" + '%') " + alias + ",";
                        }
                        aliasStr += alias;
                        alias = (char)((int)alias + 1);
                    }
                    //处理汇总部分
                    sqlMiddle += "(select x.provinceName as \"" + dataCompile.SelectedRowName + "\", y.[" + dataCompile.SummaryColumnNameA + "] from (select provinceName from [DFS].[dbo].[所在省]) x left join "
                        + "(select \"" + dataCompile.SelectedRowName + "\"," + sumOrCount + " as '" + dataCompile.SummaryColumnNameA + "' from [" + dataCompile.TableName + "] " + where +
                        " group by \"" + dataCompile.SelectedRowName + "\") y on x.provinceName like '%' +y.\"" + dataCompile.SelectedRowName + "\" + '%') " + alias;
                    aliasStr += alias;
                    alias = (char)((int)alias + 1);

                    //处理listB的部分
                    if (listB != null)
                    {
                        sqlMiddle += ",";
                        for (i = 0; i < listB.Count; i++)
                        {
                            string column = listB[i];
                            if (column.Contains("对外服务机时"))
                            {
                                sumOrCount = " CONVERT(float," + " sum(cast(\"" + column + "\" as float))/10000) ";

                                sqlMiddle += "(select x.provinceName as [" + dataCompile.SelectedRowName + "], y.[" + column + "] from (select provinceName from [DFS].[dbo].[所在省]) x left join " + @"(
                        select [" + dataCompile.SelectedRowName + "]," + sumOrCount + " as '" + column + "' from [" + dataCompile.TableName +
                                "] group by [" + dataCompile.SelectedRowName + "]) y on x.provinceName like '%' + y.[" + dataCompile.SelectedRowName + "] + '%') " + alias + ",";
                            }
                            else
                            {
                                sqlMiddle += "(select x.provinceName as [" + dataCompile.SelectedRowName + "], y.[" + column + "] from (select provinceName from [DFS].[dbo].[所在省]) x left join " + @"(
                        select [" + dataCompile.SelectedRowName + "]," + sumOrCount + " as '" + column + "' from [" + dataCompile.TableName + "] where [" + dataCompile.SelectedColumnListB[0] + "]='" + column + "'" +
                                " group by [" + dataCompile.SelectedRowName + "]) y on x.provinceName like '%' + y.[" + dataCompile.SelectedRowName + "] + '%') " + alias + ",";
                            }
                            aliasStr += alias;
                            alias = (char)((int)alias + 1);
                        }
                        if (dataCompile.SummaryColumnNameB.Contains("有效工作机时"))
                        {
                            sumOrCount = " CONVERT(float," + " sum(cast(\"" + dataCompile.SummaryColumnNameB + "\" as float))/10000) ";
                        }
                        sqlMiddle += "(select x.provinceName as [" + dataCompile.SelectedRowName + "], y.[" + dataCompile.SummaryColumnNameB + "] from (select provinceName from [DFS].[dbo].[所在省]) x left join " + @"(
                        select [" + dataCompile.SelectedRowName + "]," + sumOrCount + " as '" + dataCompile.SummaryColumnNameB + "' from [" + dataCompile.TableName + "] group by [" + dataCompile.SelectedRowName + "]) y on x.provinceName like '%' + y.[" + dataCompile.SelectedRowName + "] + '%') " + alias;
                        aliasStr += alias;
                    }

                    string sqlHead = "select ";
                    sqlHead += aliasStr[0] + ".\"" + dataCompile.SelectedRowName + "\",";
                    for (i = 0; i < listA.Count; i++)
                    {
                        sqlHead += aliasStr[i] + ".\"" + listA[i] + "\",";
                    }
                    sqlHead += aliasStr[i] + ".\"" + dataCompile.SummaryColumnNameA + "\"";
                    int x = i + 1;
                    if (listB != null)
                    {
                        sqlHead += ",";
                        for (i = 0; i < listB.Count; i++)
                        {
                            sqlHead += aliasStr[x++] + ".\"" + listB[i] + "\",";
                        }
                        sqlHead += aliasStr[x] + ".\"" + dataCompile.SummaryColumnNameB + "\"";
                    }

                    sqlStrForCompile = sqlHead + " from " + sqlMiddle + " where ";
                    for (i = 0; i < aliasStr.Length - 1; i++)
                    {
                        sqlStrForCompile += aliasStr[i] + ".[" + dataCompile.SelectedRowName + "]=" + aliasStr[i + 1] + ".[" + dataCompile.SelectedRowName + "]";
                        if (i < aliasStr.Length - 2)
                        {
                            sqlStrForCompile += " and ";
                        }
                    }
                }
                else
                {
                    dataCompile.SelectedColumnListA[0] = dataCompile.SelectedColumnListA[0].Replace("#", "");
                    if (dataCompile.SelectedColumnListB != null && dataCompile.SelectedColumnListB.Count() > 0)
                    {
                        dataCompile.SelectedColumnListB[0] = dataCompile.SelectedColumnListB[0].Replace("#", "");
                    }
                    sqlStrForCompile = "select [" + dataCompile.SelectedRowName + "] ";
                    for (i = 0; i < listA.Count; i++)
                    {
                        sqlStrForCompile += "," + GetConvertStringByUnitName(listA[i]) + " as \"" + listA[i] + "\"";
                    }
                    sqlStrForCompile += ", (";
                    if (dataCompile.SummaryColumnFromNameA != null && dataCompile.SummaryColumnFromNameA.Length != 0)
                    {
                        sqlStrForCompile += GetConvertStringByUnitName(dataCompile.SummaryColumnFromNameA);
                    }
                    else
                    {
                        if (!columnNamesOfTable.Contains(dataCompile.SummaryColumnNameA))
                        {
                            for (i = 0; i < listA.Count; i++)
                            {
                                sqlStrForCompile += GetConvertStringByUnitName(listA[i]) + " +";
                            }
                        }
                        else
                        {
                            sqlStrForCompile += GetConvertStringByUnitName(dataCompile.SummaryColumnNameA);
                        }
                    }
                    sqlStrForCompile = sqlStrForCompile.Substring(0, sqlStrForCompile.Length - 1);
                    sqlStrForCompile += ") as '" + dataCompile.SummaryColumnNameA + "'";
                    if (listB != null)
                    {
                        for (i = 0; i < listB.Count; i++)
                        {
                            sqlStrForCompile += "," + GetConvertStringByUnitName(listB[i]) + " as \"" + listB[i] + "\"";
                        }
                        sqlStrForCompile += ", (";
                        if (dataCompile.SummaryColumnFromNameB != null && dataCompile.SummaryColumnFromNameB.Length != 0)
                        {
                            sqlStrForCompile += GetConvertStringByUnitName(dataCompile.SummaryColumnFromNameB);
                        }
                        else
                        {
                            if (!columnNamesOfTable.Contains(dataCompile.SummaryColumnNameB))
                            {
                                for (i = 0; i < listB.Count; i++)
                                {
                                    sqlStrForCompile += GetConvertStringByUnitName(listB[i]) + " +";
                                }
                            }
                            else
                            {
                                sqlStrForCompile += GetConvertStringByUnitName(dataCompile.SummaryColumnNameB);
                            }
                        }
                        sqlStrForCompile = sqlStrForCompile.Substring(0, sqlStrForCompile.Length - 1);
                        sqlStrForCompile += ") as " + dataCompile.SummaryColumnNameB;
                    }
                    sqlStrForCompile += " from [" + dataCompile.TableName + "] group by [" + dataCompile.SelectedRowName + "] ";
                }
            }
            // 第二种类型：行名是仪器分类大类
            else if (dataCompile.SelectedRowName.Equals("仪器分类大类"))
            {
                //如果选定的列名只有一个
                if (dataCompile.SelectedColumnListA.Count == 1 && dataCompile.SelectedColumnListA[0][0] != '#')
                {
                    string sqlMiddle = "";
                    string aliasStr = "";
                    char alias = 'a';
                    sqlMiddle = "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameA + "\",y.ord from(select deviceType as rowName from 设备类型) x left join " + @"(
                        select '0' as ord, '按仪器类型分组' as rowName,'' as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        union
                        select '1' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        from  [" + dataCompile.TableName + @"]  
                        where 仪器分类大类='分析仪器'
                        group by 仪器分类大类 having 仪器分类大类='分析仪器'
                        union
                        select '2' as ord, 仪器分类中类, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        from  [" + dataCompile.TableName + @"]  
                        where 仪器分类大类='分析仪器'
                        group by 仪器分类中类
                        union
                        select '3' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        from  [" + dataCompile.TableName + @"]  
                        where 仪器分类大类!='分析仪器'
                        group by 仪器分类大类
                        union
                        select '4' as ord, '按所属单位类型分组',''
                        union
                        select '5' as ord, 单位属性, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        from  [" + dataCompile.TableName + @"]  
                        group by 单位属性
                        union
                        select '6' as ord, '属于研究实验基地', " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        from  [" + dataCompile.TableName + @"] 
                        group by 是否属于研究实验基地
                        having 是否属于研究实验基地='是'
                        union
                        select '7' as ord, '属于生物种质资源保存机构', " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameA + "'" + @"
                        from  [" + dataCompile.TableName + @"] 
                        group by 是否属于保存机构
                        having 是否属于保存机构='是'
                        ) y on x.rowName=y.rowName ) " + alias + ",";
                    aliasStr += alias;
                    //处理listA的部分
                    for (i = 0; i < finalListA.Count; i++)
                    {
                        string column = finalListA[i];
                        alias = (char)((int)alias + 1);
                        if (!otherColumn.Equals("其他") || !column.Equals("其他"))
                        {
                            notInSet += "'" + column + "',";
                            sqlMiddle += "(select x.rowName, y.\"" + column + "\",y.ord from(select deviceType as rowName from 设备类型) x left join " + @"(
                                select '0' as ord, '按仪器类型分组' as rowName,'' as " + "'" + column + "' " + @"
                                union
                                select '1' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' and " + dataCompile.SelectedColumnListA[0] + "='" + column + "'" + @"
                                group by 仪器分类大类 having 仪器分类大类='分析仪器'
                                union
                                select '2' as ord, 仪器分类中类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' and " + dataCompile.SelectedColumnListA[0] + "='" + column + "'" + @"
                                group by 仪器分类中类
                                union
                                select '3' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类!='分析仪器' and " + dataCompile.SelectedColumnListA[0] + "='" + column + "'" + @"
                                group by 仪器分类大类
                                union
                                select '4' as ord, '按所属单位类型分组',''
                                union
                                select '5' as ord, 单位属性, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where " + dataCompile.SelectedColumnListA[0] + "='" + column + "'" + @"
                                group by 单位属性
                                union
                                select '6' as ord, '属于研究实验基地', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                where " + dataCompile.SelectedColumnListA[0] + "='" + column + "'" + @"
                                group by 是否属于研究实验基地
                                having 是否属于研究实验基地='是'
                                union
                                select '7' as ord, '属于生物种质资源保存机构', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                where " + dataCompile.SelectedColumnListA[0] + "='" + column + "'" + @"
                                group by 是否属于保存机构
                                having 是否属于保存机构='是'
                                ) y on x.rowName=y.rowName ) " + alias + ",";
                            aliasStr += alias;
                        }
                        else
                        {
                            dataCompile.SelectedColumnListA[0] = dataCompile.SelectedColumnListA[0].Replace("#", "");
                            if (dataCompile.SelectedColumnListB != null && dataCompile.SelectedColumnListB.Count() != 0)
                            {
                                dataCompile.SelectedColumnListB[0] = dataCompile.SelectedColumnListB[0].Replace("#", "");
                            }
                            notInSet = notInSet.Substring(0, notInSet.Length - 1);
                            sqlMiddle += "(select x.rowName, y.\"" + column + "\",y.ord from(select deviceType as rowName from 设备类型) x left join " + @"(
                                select '0' as ord, '按仪器类型分组' as rowName,'' as " + "'" + column + "' " + @"
                                union
                                select '1' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' and " + dataCompile.SelectedColumnListA[0] + " not in(" + notInSet + ")" + @"
                                group by 仪器分类大类 having 仪器分类大类='分析仪器'
                                union
                                select '2' as ord, 仪器分类中类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' and " + dataCompile.SelectedColumnListA[0] + " not in(" + notInSet + ")" + @"
                                group by 仪器分类中类
                                union
                                select '3' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类!='分析仪器' and " + dataCompile.SelectedColumnListA[0] + " not in(" + notInSet + ")" + @"
                                group by 仪器分类大类
                                union
                                select '4' as ord, '按所属单位类型分组',''
                                union
                                select '5' as ord, 单位属性, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where " + dataCompile.SelectedColumnListA[0] + " not in(" + notInSet + ")" + @"
                                group by 单位属性
                                union
                                select '6' as ord, '属于研究实验基地', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                where " + dataCompile.SelectedColumnListA[0] + " not in(" + notInSet + ")" + @"
                                group by 是否属于研究实验基地
                                having 是否属于研究实验基地='是'
                                union
                                select '7' as ord, '属于生物种质资源保存机构', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                where " + dataCompile.SelectedColumnListA[0] + " not in(" + notInSet + ")" + @"
                                group by 是否属于保存机构
                                having 是否属于保存机构='是'
                                ) y on x.rowName=y.rowName ) " + alias + ",";
                            aliasStr += alias;
                        }
                    }
                    //处理listB的部分
                    if (listB != null)
                    {
                        if (dataCompile.SummaryColumnNameB.Equals("年有效工作机时"))
                        {
                            sumOrCount = " CONVERT(float," + " sum(cast(\"" + dataCompile.SummaryColumnNameB + "\" as float))/10000) ";
                        }
                        alias = (char)((int)alias + 1);
                        sqlMiddle += "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameB + "\",y.ord from(select deviceType as rowName from 设备类型) x left join " + @"(
                                select '0' as ord, '按仪器类型分组' as rowName,'' as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                union
                                select '1' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器'
                                group by 仪器分类大类 having 仪器分类大类='分析仪器'
                                union
                                select '2' as ord, 仪器分类中类, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器'
                                group by 仪器分类中类
                                union
                                select '3' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类!='分析仪器' 
                                group by 仪器分类大类
                                union
                                select '4' as ord, '按所属单位类型分组',''
                                union
                                select '5' as ord, 单位属性, " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                group by 单位属性
                                union
                                select '6' as ord, '属于研究实验基地', " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                group by 是否属于研究实验基地
                                having 是否属于研究实验基地='是'
                                union
                                select '7' as ord, '属于生物种质资源保存机构', " + sumOrCount + " as " + "'" + dataCompile.SummaryColumnNameB + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                group by 是否属于保存机构
                                having 是否属于保存机构='是'
                                ) y on x.rowName=y.rowName ) " + alias + ",";
                        aliasStr += alias;
                        for (i = 0; i < finalListB.Count; i++)
                        {
                            string column = finalListB[i];
                            if (column.Equals("年对外服务机时"))
                            {
                                sumOrCount = " CONVERT(float," + " sum(cast(\"" + column + "\" as float))/10000) ";
                                alias = (char)((int)alias + 1);
                                sqlMiddle += "(select x.rowName, y.\"" + column + "\",y.ord from(select deviceType as rowName from 设备类型) x left join " + @"(
                                select '0' as ord, '按仪器类型分组' as rowName,'' as " + "'" + column + "' " + @"
                                union
                                select '1' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' " + @"
                                group by 仪器分类大类 having 仪器分类大类='分析仪器'
                                union
                                select '2' as ord, 仪器分类中类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' " + @"
                                group by 仪器分类中类
                                union
                                select '3' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类!='分析仪器' " + @"
                                group by 仪器分类大类
                                union
                                select '4' as ord, '按所属单位类型分组',''
                                union
                                select '5' as ord, 单位属性, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  " + @"
                                group by 单位属性
                                union
                                select '6' as ord, '属于研究实验基地', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                group by 是否属于研究实验基地
                                having 是否属于研究实验基地='是'
                                union
                                select '7' as ord, '属于生物种质资源保存机构', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                group by 是否属于保存机构
                                having 是否属于保存机构='是'
                                ) y on x.rowName=y.rowName ) " + alias + ",";
                                aliasStr += alias;
                            }
                            else
                            {
                                alias = (char)((int)alias + 1);
                                sqlMiddle += "(select x.rowName, y.\"" + column + "\",y.ord from(select deviceType as rowName from 设备类型) x left join " + @"(
                                select '0' as ord, '按仪器类型分组' as rowName,'' as " + "'" + column + "' " + @"
                                union
                                select '1' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' and " + dataCompile.SelectedColumnListB[0] + "='" + column + "'" + @"
                                group by 仪器分类大类 having 仪器分类大类='分析仪器'
                                union
                                select '2' as ord, 仪器分类中类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类='分析仪器' and " + dataCompile.SelectedColumnListB[0] + "='" + column + "'" + @"
                                group by 仪器分类中类
                                union
                                select '3' as ord, 仪器分类大类, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where 仪器分类大类!='分析仪器' and " + dataCompile.SelectedColumnListB[0] + "='" + column + "'" + @"
                                group by 仪器分类大类
                                union
                                select '4' as ord, '按所属单位类型分组',''
                                union
                                select '5' as ord, 单位属性, " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"]  
                                where " + dataCompile.SelectedColumnListB[0] + "='" + column + "'" + @"
                                group by 单位属性
                                union
                                select '6' as ord, '属于研究实验基地', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                where " + dataCompile.SelectedColumnListB[0] + "='" + column + "'" + @"
                                group by 是否属于研究实验基地
                                having 是否属于研究实验基地='是'
                                union
                                select '7' as ord, '属于生物种质资源保存机构', " + sumOrCount + " as " + "'" + column + "'" + @"
                                from  [" + dataCompile.TableName + @"] 
                                where " + dataCompile.SelectedColumnListB[0] + "='" + column + "'" + @"
                                group by 是否属于保存机构
                                having 是否属于保存机构='是'
                                ) y on x.rowName=y.rowName ) " + alias + ",";
                                aliasStr += alias;
                            }
                        }
                    }
                    sqlMiddle = sqlMiddle.Substring(0, sqlMiddle.Length - 1);
                    string sqlHead = "select ";
                    sqlHead += "a.ord, a.rowName, a.[" + dataCompile.SummaryColumnNameA + "],";
                    for (i = 0; i < listA.Count; i++)
                    {
                        sqlHead += aliasStr[i + 1] + ".\"" + listA[i] + "\",";
                    }
                    int x = i + 1;
                    if (listB != null)
                    {
                        sqlHead += aliasStr[x++] + ".\"" + dataCompile.SummaryColumnNameB + "\",";
                        for (i = 0; i < listB.Count; i++)
                        {
                            sqlHead += aliasStr[x++] + ".\"" + listB[i] + "\",";
                        }
                    }
                    sqlHead = sqlHead.Substring(0, sqlHead.Length - 1);
                    sqlStrForCompile = sqlHead + " from " + sqlMiddle + " where ";
                    for (i = 0; i < aliasStr.Length - 1; i++)
                    {
                        sqlStrForCompile += aliasStr[i] + ".rowName=" + aliasStr[i + 1] + ".rowName ";
                        if (i < aliasStr.Length - 2)
                        {
                            sqlStrForCompile += " and ";
                        }
                    }
                    sqlStrForCompile = "select * from(" + sqlStrForCompile + ") as z order by z.ord";
                }
            }
            //第三种类型，行名是实验基地
            else if (dataCompile.SelectedRowName.Equals("实验基地"))
            {
                if (dataCompile.SelectedColumnListA.Count == 1 && dataCompile.SelectedColumnListA[0][0] != '#')
                {
                    string sqlMiddle = "";
                    string aliasStr = "";
                    char alias = 'a';
                    sqlMiddle = "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameA
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + "FROM [" + dataCompile.TableName + "]" + " group by 基地大类 union "
                        + "SELECT 全国汇编用基地小类, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + "FROM  [" + dataCompile.TableName + "]" + " group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName ) " + alias + ",";
                    aliasStr += alias;
                    //处理listA的部分
                    for (i = 0; i < finalListA.Count; i++)
                    {
                        string column = finalListA[i];
                        alias = (char)((int)alias + 1);
                        if (!otherColumn.Equals("其他") || !column.Equals("其他"))
                        {
                            notInSet += "'" + column + "',";
                            sqlMiddle += "(select x.rowName, y.\"" + column
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName," + sumOrCount + " as \"" + column + "\" "
                        + "FROM  [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListA[0] + "]='" + column + "' group by 基地大类 "
                        + " union "
                        + "SELECT 全国汇编用基地小类, " + sumOrCount + " as \"" + column + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListA[0] + "]='" + column + "' group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName )" + alias + ",";
                            aliasStr += alias;
                        }
                        else
                        {
                            notInSet = notInSet.Substring(0, notInSet.Length - 1);
                            sqlMiddle += "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameA
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListA[0] + "] not in(" + notInSet + ")  group by 基地大类 "
                        + " union "
                        + "SELECT 全国汇编用基地小类, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListA[0] + "] not in(" + notInSet + ") group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName ) " + alias + ",";
                            aliasStr += alias;
                        }
                    }
                    //处理listB的部分
                    if (listB != null)
                    {
                        alias = (char)((int)alias + 1);
                        sqlMiddle += "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameB
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameB + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListB[0] + "]='" + dataCompile.SummaryColumnNameB + "' group by 基地大类"
                        + " union "
                        + "SELECT 全国汇编用基地小类, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameB + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListB[0] + "]='" + dataCompile.SummaryColumnNameB + "' group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName ) " + alias + ",";
                        aliasStr += alias;

                        for (i = 0; i < finalListB.Count; i++)
                        {
                            string column = finalListB[i];
                            alias = (char)((int)alias + 1);
                            sqlMiddle += "(select x.rowName, y.\"" + column
                       + "\" from(select type as rowName from 实验基地) x left join "
                       + "(SELECT 基地大类 as rowName, " + sumOrCount + " as \"" + column + "\" "
                       + "FROM [" + dataCompile.TableName + "] "
                       + "where [" + dataCompile.SelectedColumnListB[0] + "]='" + column + "' group by 基地大类 "
                       + " union "
                       + "SELECT 全国汇编用基地小类, " + sumOrCount + " as \"" + column + "\" "
                       + "FROM [" + dataCompile.TableName + "] "
                       + "where [" + dataCompile.SelectedColumnListB[0] + "]='" + column + "' group by 全国汇编用基地小类"
                       + ") y on x.rowName=y.rowName ) " + alias + ",";
                            aliasStr += alias;
                        }
                    }
                    sqlMiddle = sqlMiddle.Substring(0, sqlMiddle.Length - 1);
                    string sqlHead = "select ";
                    sqlHead += "a.rowName, a.[" + dataCompile.SummaryColumnNameA + "],";
                    for (i = 0; i < listA.Count; i++)
                    {
                        sqlHead += aliasStr[i + 1] + ".\"" + listA[i] + "\",";
                    }
                    int x = i + 1;
                    if (listB != null)
                    {
                        sqlHead += aliasStr[x++] + ".\"" + dataCompile.SummaryColumnNameB + "\",";
                        for (i = 0; i < listB.Count; i++)
                        {
                            sqlHead += aliasStr[x++] + ".\"" + listB[i] + "\",";
                        }
                    }
                    sqlHead = sqlHead.Substring(0, sqlHead.Length - 1);
                    sqlStrForCompile = sqlHead + " from " + sqlMiddle + " where ";
                    for (i = 0; i < aliasStr.Length - 1; i++)
                    {
                        sqlStrForCompile += aliasStr[i] + ".rowName=" + aliasStr[i + 1] + ".rowName ";
                        if (i < aliasStr.Length - 2)
                        {
                            sqlStrForCompile += " and ";
                        }
                    }
                }
                else
                {
                    dataCompile.SelectedColumnListA[0] = dataCompile.SelectedColumnListA[0].Replace("#", "");
                    if (dataCompile.SelectedColumnListB != null && dataCompile.SelectedColumnListB.Count() > 0)
                    {
                        dataCompile.SelectedColumnListB[0] = dataCompile.SelectedColumnListB[0].Replace("#", "");
                    }

                    string sqlMiddle = "";
                    string aliasStr = "";
                    char alias = 'a';
                    string sumString = "";
                    if (dataCompile.UnitName.Contains("亿元") && (dataCompile.SummaryColumnNameA.Contains("原值") || dataCompile.SummaryColumnNameA.Contains("总额") || dataCompile.SummaryColumnNameA.Contains("费")))
                    {
                        sumString = " CONVERT(bigint," + " sum(cast(\"" + dataCompile.SummaryColumnNameA + "\" as float))/10000) ";
                    }
                    else
                    {
                        sumString = " CONVERT(bigint," + " sum(cast(\"" + dataCompile.SummaryColumnNameA + "\" as float))) ";
                    }

                    sqlMiddle = "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameA
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName, " + sumString + " as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + " FROM [" + dataCompile.TableName + "]" + " group by 基地大类 union "
                        + "SELECT 全国汇编用基地小类, " + sumString + " as \"" + dataCompile.SummaryColumnNameA + "\""
                        + " FROM  [" + dataCompile.TableName + "]" + " group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName ) " + alias + ",";
                    aliasStr += alias;
                    //处理listA的部分

                    for (i = 0; i < finalListA.Count; i++)
                    {
                        string column = finalListA[i];
                        if (dataCompile.UnitName.Contains("亿元") && (dataCompile.SummaryColumnNameA.Contains("原值") || dataCompile.SummaryColumnNameA.Contains("总额") || dataCompile.SummaryColumnNameA.Contains("费")))
                        {
                            sumString = " CONVERT(bigint," + " sum(cast(\"" + column + "\" as float))/10000) ";
                        }
                        else
                        {
                            sumString = " CONVERT(bigint," + " sum(cast(\"" + column + "\" as float))) ";
                        }

                        alias = (char)((int)alias + 1);
                        if (!otherColumn.Equals("其他") || !column.Equals("其他"))
                        {
                            notInSet += "'" + column + "',";
                            sqlMiddle += "(select x.rowName, y.\"" + column
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName," + sumString + "  as \"" + column + "\" "
                        + "FROM  [" + dataCompile.TableName + "] "
                        + " group by 基地大类 "
                        + " union "
                        + "SELECT 全国汇编用基地小类, " + sumString + " as \"" + column + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + " group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName )" + alias + ",";
                            aliasStr += alias;
                        }
                        else
                        {
                            notInSet = notInSet.Substring(0, notInSet.Length - 1);
                            sqlMiddle += "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameA
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName, " + sumOrCount + " as float))) as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListA[0] + "] not in(" + notInSet + ")  group by 基地大类 "
                        + " union "
                        + "SELECT 全国汇编用基地小类, " + sumOrCount + " as \"" + dataCompile.SummaryColumnNameA + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + "where [" + dataCompile.SelectedColumnListA[0] + "] not in(" + notInSet + ") group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName ) " + alias + ",";
                            aliasStr += alias;
                        }
                    }
                    //处理listB的部分
                    if (listB != null)
                    {

                        alias = (char)((int)alias + 1);
                        sqlMiddle += "(select x.rowName, y.\"" + dataCompile.SummaryColumnNameB
                        + "\" from(select type as rowName from 实验基地) x left join "
                        + "(SELECT 基地大类 as rowName, " + "CONVERT(bigint, sum(cast(\"" + dataCompile.SummaryColumnNameB + "\" as float))) as \"" + dataCompile.SummaryColumnNameB + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + " group by 基地大类"
                        + " union "
                        + "SELECT 全国汇编用基地小类, " + "CONVERT(bigint, sum(cast(\"" + dataCompile.SummaryColumnNameB + "\" as float))) as \"" + dataCompile.SummaryColumnNameB + "\" "
                        + "FROM [" + dataCompile.TableName + "] "
                        + " group by 全国汇编用基地小类"
                        + ") y on x.rowName=y.rowName ) " + alias + ",";
                        aliasStr += alias;

                        for (i = 0; i < finalListB.Count; i++)
                        {
                            string column = finalListB[i];

                            if (dataCompile.UnitName.Contains("亿元") && (dataCompile.SummaryColumnNameB.Contains("原值") || dataCompile.SummaryColumnNameA.Contains("总额") || dataCompile.SummaryColumnNameA.Contains("费")))
                            {
                                sumString = " CONVERT(bigint," + " sum(cast(\"" + column + "\" as float))/10000) ";
                            }
                            else
                            {
                                sumString = " CONVERT(bigint," + " sum(cast(\"" + column + "\" as float))) ";
                            }
                            alias = (char)((int)alias + 1);
                            sqlMiddle += "(select x.rowName, y.\"" + column
                       + "\" from(select type as rowName from 实验基地) x left join "
                       + "(SELECT 基地大类 as rowName, " + sumString + " as \"" + column + "\" "
                       + "FROM [" + dataCompile.TableName + "] "
                       + " group by 基地大类 "
                       + " union "
                       + "SELECT 全国汇编用基地小类, " + sumString + " as \"" + column + "\" "
                       + "FROM [" + dataCompile.TableName + "] "
                       + " group by 全国汇编用基地小类"
                       + ") y on x.rowName=y.rowName ) " + alias + ",";
                            aliasStr += alias;
                        }
                    }
                    sqlMiddle = sqlMiddle.Substring(0, sqlMiddle.Length - 1);
                    string sqlHead = "select ";
                    sqlHead += "a.rowName, a." + dataCompile.SummaryColumnNameA + ",";
                    for (i = 0; i < listA.Count; i++)
                    {
                        sqlHead += aliasStr[i + 1] + ".\"" + listA[i] + "\",";
                    }
                    int x = i + 1;
                    if (listB != null)
                    {
                        sqlHead += aliasStr[x++] + ".\"" + dataCompile.SummaryColumnNameB + "\",";
                        for (i = 0; i < listB.Count; i++)
                        {
                            sqlHead += aliasStr[x++] + ".\"" + listB[i] + "\",";
                        }
                    }
                    sqlHead = sqlHead.Substring(0, sqlHead.Length - 1);
                    sqlStrForCompile = sqlHead + " from " + sqlMiddle + " where ";
                    for (i = 0; i < aliasStr.Length - 1; i++)
                    {
                        sqlStrForCompile += aliasStr[i] + ".rowName=" + aliasStr[i + 1] + ".rowName ";
                        if (i < aliasStr.Length - 2)
                        {
                            sqlStrForCompile += " and ";
                        }
                    }
                }
            }
            return true;
        }

        protected List<string> GetAllColumnNamesFromServer()
        {
            List<string> names = new List<string>();
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            string sql = "select top 1 * from " + dataCompile.TableName;

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

        protected string GetConvertStringByUnitName(string str)
        {
            if (dataCompile.UnitName.Equals("亿元"))
            {
                return " convert(bigint, sum(cast(\"" + str + "\" as float)))/10000 ";
            }
            else
            {
                return " convert(bigint, sum(cast(\"" + str + "\" as float))) ";
            }
        }

        protected List<string> GetFinallyColumnNames(List<string> list, char type)
        {
            Log.RecordLog(dataCompile.Caption + ": GetFinallyColumnNames");
            List<string> listA = new List<string>();

            if ((list != null && list.Count > 1) || list[0][0] == '#' || list[0].Equals("年对外服务机时"))
            {
                //如果选择的列A有多项，表明该选择的列为最终需要显示的列，否则是该列的成员
                listA = list;
            }
            else
            {
                if (list[0].Equals("产地"))
                {
                    //如果列标签是产地的话只显示5个国家
                    listA.Add("美国");
                    listA.Add("中国");
                    listA.Add("德国");
                    listA.Add("日本");
                    listA.Add("英国");
                    listA.Add("其他");
                    otherColumn = "其他";
                }
                else if (list[0].Contains("仪器分类大类"))
                {
                    listA = wordDocumentMerger.deviceTypeList;
                }
                else if (list[0].Contains("仪器分类中类"))
                {
                    listA = SharedData.analyseDeviceTypeSorted;
                }
                else if (list[0].Contains("共享方式"))
                {
                    listA = SharedData.sharedTypeForGeneticResources;
                }
                else if (list[0].Contains("保存资源类型"))
                {
                    if (dataCompile.Caption.Contains("植物"))
                    {
                        listA = SharedData.savedTypeForPlantGeneticResources;
                    }
                    else if (dataCompile.Caption.Contains("动物"))
                    {
                        listA = SharedData.savedTypeForAnimalGeneticResources;
                    }
                }
                else if (list[0].Contains("保藏类型"))
                {
                    listA = SharedData.reservedTypeForGeneticResources;
                }
                else if (list[0].Contains("保存方法"))
                {
                    listA = SharedData.reservedMethodForGeneticResources;
                }
                else if (list[0].Contains("全国汇编用基地小类"))
                {
                    //表3-1(4) 其他类型基地只显示一个
                    if (type == 'B' && dataCompile.SummaryColumnNameB.Contains("其他"))
                    {
                        listA.Add("其他国家级基地");
                        return listA;
                    }
                    //需要根据该列查询最终的列名
                    string name = list[0];
                    string sqlStr = "";
                    if (type == 'A')
                    {
                        if (dataCompile.SummaryColumnFromNameA.Equals(""))
                        {
                            sqlStr = @"select " + name + " from " + dataCompile.TableName + " where 基地大类='" + dataCompile.SummaryColumnNameA + "' group by " + name;
                        }
                        else
                        {
                            sqlStr = @"select " + name + " from " + dataCompile.TableName + " where 基地大类='" + dataCompile.SummaryColumnFromNameA + "' group by " + name;
                        }
                    }
                    else
                    {
                        if (dataCompile.SummaryColumnFromNameB.Equals(""))
                        {
                            sqlStr = @"select " + name + " from " + dataCompile.TableName + " where 基地大类='" + dataCompile.SummaryColumnNameB + "' group by " + name;
                        }
                        else
                        {
                            sqlStr = @"select " + name + " from " + dataCompile.TableName + " where 基地大类='" + dataCompile.SummaryColumnFromNameB + "' group by " + name;
                        }
                    }
                    try
                    {
                        SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
                        sqlConn.Open();
                        SqlCommand command = new SqlCommand(sqlStr, sqlConn);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            if (reader[name] != null && reader[name].ToString().Length > 0)
                            {
                                listA.Add(reader[name].ToString());
                            }
                        }
                        reader.Close();
                        sqlConn.Close();
                    }
                    catch (System.Exception ex)
                    {
                        Log.RecordLog(ex.ToString());
                    }
                }
                else
                {
                    //需要根据该列查询最终的列名
                    string name = list[0];
                    name = name.Replace("#", "");
                    string sqlStr = @"select [" + name + "] from [" + dataCompile.TableName + "] group by [" + name + "]";
                    SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
                    sqlConn.Open();

                    try
                    {
                        SqlCommand command = new SqlCommand(sqlStr, sqlConn);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            if (reader[name] != null && reader[name].ToString().Length > 0)
                            {
                                listA.Add(reader[name].ToString());
                            }
                        }
                        reader.Close();
                        sqlConn.Close();
                    }
                    catch (System.Exception ex)
                    {
                        Log.RecordLog(ex.ToString());
                    }
                }
            }
            //对list排序
            ListSort(listA);
            return listA;
        }

        private void ListSort(List<string> list)
        {
            List<int> numberList = new List<int>();
            int otherIndex = -1;
            int i = 0;
            for (i = 0; i < list.Count(); i++)
            {
                string ss = list[i];
                int k = 0;
                while(ss[k] >= '0' && ss[k] <= '9') k++;
                if (k == 0)
                {
                    if (ss.Contains("其他") || ss.Contains("其它"))
                    {
                        otherIndex = i;
                    }
                }
                else
                {
                    numberList.Add(System.Int32.Parse(ss.Substring(0, k)));
                }
            }
            if (otherIndex != -1)
            {
                string tem = list[otherIndex];
                list.RemoveAt(otherIndex);
                list.Add(tem);
                return;
            }
            if (numberList.Count() != list.Count())
            {
                return;
            }
            int j = 0;
            for (i = 0; i < list.Count(); i++)
            {
                for (j = 0; j < list.Count() - i - 1; j++)
                {
                    if (numberList[j] > numberList[j + 1])
                    {
                        int t = numberList[j];
                        numberList[j] = numberList[j + 1];
                        numberList[j + 1] = t;
                        string s = list[j];
                        list[j]= list[j + 1];
                        list[j + 1] = s;
                    }
                }
            }
        }

        // 对每个省对应的table进行数据填充
        public virtual void FillTable(List<Table> newTableList)
        {
            Log.RecordLog(dataCompile.Caption + ": FillTable");
            int i = 1, k = 1, j = 0, tableSeq, listAColumnSeq = 0;
            List<int> pageNumberOfListAColumn = new List<int>();
            List<int> columnNumberOfListAColumn = new List<int>();

            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            SqlCommand command = new SqlCommand(sqlStrForCompile, sqlConn);
            command.CommandTimeout = 180;
            SqlDataReader reader = command.ExecuteReader();

            for (tableSeq = 0; tableSeq < newTableList.Count(); tableSeq++)
            {
                Table newTable = newTableList[tableSeq];
                int tableColumns = newTable.Columns.Count;
                newTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                newTable.Cell(1, 1).Merge(newTable.Cell(2, 1));
                if (tableSeq == 0)
                {
                    newTable.Cell(1, 2).Merge(newTable.Cell(2, 2));
                }

                //合并第一行的单元格
                int begin, mergeCount;
                if (tableSeq == 0)
                {
                    begin = 3;
                    mergeCount = tableColumns - 3;
                }
                else
                {
                    begin = 2;
                    mergeCount = tableColumns - 2;
                }
                if (finalListB != null && finalListB.Count() > 0)
                {
                    mergeCount = mergeCount - finalListB.Count - 1;
                }
                for (i = 0; i < mergeCount; i++)
                {
                    newTable.Cell(1, begin).Merge(newTable.Cell(1, begin + 1));
                }
                if (finalListB != null && finalListB.Count() > 0)
                {
                    for (i = 0; i < finalListB.Count() - 1; i++)
                    {
                        newTable.Cell(1, 5).Merge(newTable.Cell(1, 6));
                    }
                    newTable.Cell(1, 4).Merge(newTable.Cell(2, 3 + finalListA.Count));
                }
                if (tableSeq == 0)
                {
                    newTable.Cell(1, 3).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                }
                if (finalListB != null && finalListB.Count() > 0)
                {
                    newTable.Cell(1, 5).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                }
                if (tableSeq == 0)
                {
                    newTable.Cell(1, 2).Range.Text = dataCompile.SummaryColumnNameA;
                    newTable.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                if (dataCompile.SummaryColumnNameB != null && dataCompile.SummaryColumnNameB.Length != 0)
                {
                    newTable.Cell(1, 4).Range.Text = dataCompile.SummaryColumnNameB;
                    newTable.Cell(1, 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                //填充列名
                if (tableSeq == 0)
                {
                    for (i = 3; i <= tableColumns && i < finalListA.Count + 3; i++)
                    {
                        newTable.Cell(2, i).Range.Text = finalListA[listAColumnSeq++];
                        newTable.Cell(2, i).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        pageNumberOfListAColumn.Add(tableSeq);
                        columnNumberOfListAColumn.Add(i);
                    }
                }
                else
                {
                    for (i = 2; i <= tableColumns; i++)
                    {
                        newTable.Cell(2, i).Range.Text = finalListA[listAColumnSeq++];
                        newTable.Cell(2, i).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        pageNumberOfListAColumn.Add(tableSeq);
                        columnNumberOfListAColumn.Add(i);
                    }
                }

                if (finalListB != null && finalListB.Count() > 0)
                {
                    for (i = 0; i < finalListB.Count; i++)
                    {
                        newTable.Cell(2, i + finalListA.Count + 4).Range.Text = finalListB[i];
                        newTable.Cell(2, i + finalListA.Count + 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    }
                }

                newTable.Cell(3, 1).Range.Text = "总计";

                for (i = 0; i < tableColumns; i++)
                {
                    newTable.Cell(3, i + 1).Range.Font.Bold = 2;
                }
                //设置边线
                for (i = 3; i <= sumRows + 3; i++)
                {
                    for (j = 1; j <= tableColumns; j++)
                    {
                        if (i > 3)
                        {
                            newTable.Cell(i, j).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                        }
                        if (j != 1)
                        {
                            newTable.Cell(i, j).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                        }
                    }
                }
                //填充第一列,如果第一列是所在省，需要指定顺序
                if (dataCompile.SelectedRowName.Contains("所在"))
                {
                    for (i = 0; i < SharedData.areaListSorted.Count; i++)
                    {
                        string text = SharedData.areaListSorted[i];
                        text = text.Replace("省", "");
                        text = text.Replace("市", "");
                        text = text.Replace("自治区", "");
                        text = text.Replace("维吾尔", "");
                        text = text.Replace("回族", "");
                        text = text.Replace("壮族", "");
                        newTable.Cell(i + 4, 1).Range.Text = text;
                    }
                }
                else if (dataCompile.SelectedRowName.Equals("仪器分类大类"))
                {
                    for (i = 0; i < SharedData.deviceTypeSorted.Count; i++)
                    {
                        string text = SharedData.deviceTypeSorted[i];
                        if (i >= 1 && i <= 14)
                        {
                            text = "  " + text;
                        }
                        newTable.Cell(i + 4, 1).Range.Text = text;
                        newTable.Cell(i + 4, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    newTable.Cell(4, 1).Range.Font.Bold = 2;
                    newTable.Cell(33, 1).Range.Font.Bold = 2;
                    newTable.Cell(37, 1).Range.Font.Bold = 2;
                    newTable.Cell(38, 1).Range.Font.Bold = 2;

                }
                else if (dataCompile.SelectedRowName.Equals("实验基地"))
                {
                    for (i = 0; i < SharedData.experimentalBaseSorted.Count; i++)
                    {
                        string text = SharedData.experimentalBaseSorted[i];
                        if ((i >= 2 && i <= 8) || (i >= 10 && i <= 12) || (i >= 14 && i <= 18) || (i >= 20 && i <= 24) || (i >= 26 && i <= 29) ||
                            (i >= 31 && i <= 34))
                        {
                            text = "  " + text;
                        }
                        newTable.Cell(i + 4, 1).Range.Text = text;
                        newTable.Cell(i + 4, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                }
                newTable.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                newTable.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                newTable.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            }
            k = 4;
            j = 1;
            while (reader.Read())
            {
                //寻找行号，k是该条纪录对应的table的行号
                if (dataCompile.SelectedRowName.Contains("所在"))
                {
                    string text = reader[dataCompile.SelectedRowName].ToString();
                    for (k = 0; k < SharedData.areaListSorted.Count; k++)
                    {
                        if (SharedData.areaListSorted[k].Contains(text))
                        {
                            break;
                        }
                    }
                    if (k == SharedData.areaListSorted.Count)
                    {
                        continue;
                    }
                    k += 4;
                }
                else if (dataCompile.SelectedRowName.Contains("仪器分类大类"))
                {
                    string text = reader["rowName"].ToString();
                    for (k = 0; k < SharedData.deviceTypeSorted.Count; k++)
                    {
                        if (SharedData.deviceTypeSorted[k].Equals(text))
                        {
                            break;
                        }
                    }
                    if (k == SharedData.deviceTypeSorted.Count)
                    {
                        continue;
                    }
                    k += 4;
                }
                else if (dataCompile.SelectedRowName.Contains("实验基地"))
                {
                    string text = reader["rowName"].ToString();
                    for (k = 0; k < SharedData.experimentalBaseSorted.Count; k++)
                    {
                        if (SharedData.experimentalBaseSorted[k].Equals(text))
                        {
                            break;
                        }
                    }
                    if (k == SharedData.experimentalBaseSorted.Count)
                    {
                        continue;
                    }
                    k += 4;
                }

                //第二列，汇总列
                string value = reader[dataCompile.SummaryColumnNameA].ToString().Trim();
                // **************************结果要除以10000**************************
                if (isNeedTransTenThousand == true && dataCompile.UnitName.Equals("万元") && value != "")
                {
                    value = (System.Double.Parse(value) / 10000).ToString();
                }
                // **************************************************************
                if (value.IndexOf('.') != -1)
                {
                    value = value.Substring(0, value.IndexOf('.') + 2);
                }
                if (!value.Equals("0") && !value.Equals(""))
                {
                    newTableList[0].Cell(k, 2).Range.Text = value;
                    newTableList[0].Cell(k, 2).Range.Font.Name = "Calibri";
                    if (value.IndexOf('.') != -1)
                    {
                        value = value.Substring(0, value.IndexOf('.') + 2);
                    }
                    if (value != null && value.Length > 0 && k <= sumMaxLine && k >= sumMinLine)
                    {
                        sumList[1] += double.Parse(value);
                        if (k == 4)
                        {
                            assistSumList[1] = double.Parse(value);
                        }
                    }
                }
                //listA中的各列
                for (i = 0; i < finalListA.Count; i++)
                {
                    value = reader[finalListA[i]].ToString().Trim();
                    // **************************结果要除以10000**************************
                    if (isNeedTransTenThousand == true && dataCompile.UnitName.Equals("万元") && value != "")
                    {
                        value = (System.Double.Parse(value) / 10000).ToString();
                    }
                    // **************************************************************
                    if (value.IndexOf('.') != -1)
                    {
                        value = value.Substring(0, value.IndexOf('.') + 2);
                    }
                    if (!value.Equals("0") && !value.Equals(""))
                    {
                        newTableList[pageNumberOfListAColumn[i]].Cell(k, columnNumberOfListAColumn[i]).Range.Text = value;
                        newTableList[pageNumberOfListAColumn[i]].Cell(k, columnNumberOfListAColumn[i]).Range.Font.Name = "Calibri";
                        if (value != null && value.Length > 0 && k <= sumMaxLine && k >= sumMinLine)
                        {
                            sumList[i + 2] += double.Parse(value);
                            if (k == 4)
                            {
                                assistSumList[i + 2] = double.Parse(value);
                            }
                        }
                    }
                }
                i += 2;
                //listB中的列
                if ((dataCompile.SelectedRowName.Contains("所在") && i < sumColumns) || (dataCompile.SelectedRowName.Equals("仪器分类大类") && i < sumColumns - 1) || (dataCompile.SelectedRowName.Contains("实验基地") && i < sumColumns))
                {
                    value = reader[dataCompile.SummaryColumnNameB].ToString().Trim();
                    if (value.IndexOf('.') != -1)
                    {
                        value = value.Substring(0, value.IndexOf('.') + 2);
                    }
                    newTableList[0].Cell(k, i + 1).Range.Text = value;
                    newTableList[0].Cell(k, i + 1).Range.Font.Name = "Calibri";
                    if (value != null && value.Length > 0 && k <= sumMaxLine && k >= sumMinLine)
                    {
                        sumList[i] += double.Parse(value);
                        if (k == 4)
                        {
                            assistSumList[i] = double.Parse(value);
                        }
                    }
                    for (int m = 0; m < finalListB.Count; m++)
                    {
                        value = reader[finalListB[m]].ToString().Trim();
                        if (value.IndexOf('.') != -1)
                        {
                            value = value.Substring(0, value.IndexOf('.') + 2);
                        }
                        if (!value.Equals("0") && !value.Equals(""))
                        {
                            newTableList[0].Cell(k, i + 2).Range.Text = value;
                            newTableList[0].Cell(k, i + 2).Range.Font.Name = "Calibri";
                            if (value != null && value.Length > 0 && k <= sumMaxLine && k >= sumMinLine)
                            {
                                sumList[i + 1] += double.Parse(value);
                                if (k == 4)
                                {
                                    assistSumList[i + 1] = double.Parse(value);
                                }
                            }
                            i++;
                        }
                    }
                }
            }

            reader.Close();
            sqlConn.Close();

            //如果航标签是“实验基地”，sum是重复统计的
            if (dataCompile.SelectedRowName.Contains("实验基地"))
            {
                for (i = 1; i < sumList.Count; i++)
                {
                    sumList[i] -= assistSumList[i];
                    sumList[i] /= 2;
                    sumList[i] += assistSumList[i];
                }
            }

            //填充总计行
            string v = sumList[1].ToString();
            if (v.IndexOf('.') != -1)
            {
                v = v.Substring(0, v.IndexOf('.') + 2);
            }
            newTableList[0].Cell(3, 2).Range.Text = v;

            if (finalListB == null)
            {
                for (i = 2; i < sumList.Count; i++)
                {
                    v = sumList[i].ToString();
                    if (v.IndexOf('.') != -1)
                    {
                        v = v.Substring(0, v.IndexOf('.') + 2);
                    }
                    newTableList[pageNumberOfListAColumn[i - 2]].Cell(3, columnNumberOfListAColumn[i - 2]).Range.Text = v;
                    newTableList[pageNumberOfListAColumn[i - 2]].Cell(3, columnNumberOfListAColumn[i - 2]).Range.Font.Name = "Calibri";
                }
            }
            else
            {
                for (i = 1; i < sumList.Count; i++)
                {
                    v = sumList[i].ToString();
                    if (v.IndexOf('.') != -1)
                    {
                        v = v.Substring(0, v.IndexOf('.') + 2);
                    }
                    newTableList[0].Cell(3, i + 1).Range.Text = v;
                    newTableList[0].Cell(3, i + 1).Range.Font.Name = "Calibri";
                }
            }
        }

        public virtual List<Table> InsertTable()
        {
            Log.RecordLog(dataCompile.Caption + ": InsertTable");
            List<Microsoft.Office.Interop.Word.Table> tableList = new List<Table>();

            //文档中创建表格
            List<int> columnsPerPage = new List<int>();
            Microsoft.Office.Interop.Word.Table newTable;

            if (dataCompile.SelectedColumnListA[0].Equals("仪器分类大类"))
            {
                columnsPerPage.Add(6);
                columnsPerPage.Add(7);
                columnsPerPage.Add(6);
            }
            else if (dataCompile.SelectedColumnListA[0].Equals("仪器分类中类"))
            {
                columnsPerPage.Add(8);
                columnsPerPage.Add(8);
            }
            else if (dataCompile.SelectedColumnListA[0].Equals("共享方式"))
            {
                columnsPerPage.Add(6);
                columnsPerPage.Add(6);
            }
            else
            {
                columnsPerPage.Add(sumColumns);
            }
            wordDoc.Paragraphs.First.Range.Select();
            int n = 0;

            string cap = dataCompile.Caption;
            if (cap[cap.Length - 1] > '0' && cap[cap.Length - 1] <= '9')
            {
                cap = cap.Substring(0, dataCompile.Caption.Length - 1);
            }
            string fontName = "";

            while (n < columnsPerPage.Count())
            {
                if (sumSameTable > 1)
                {
                    wordApp.Selection.Font.Size = 14F;
                    wordApp.Selection.Font.Bold = 2;
                    fontName = wordApp.Selection.Font.Name;
                    wordApp.Selection.Font.Name = "微软雅黑";
                }
                wordApp.Selection.TypeText(cap);
                if (seqNo == 0)
                {
                    // 设置标题样式，以方便插入目录
                    object oStyleName = "标题 2";
                    wordApp.ActiveDocument.Styles[oStyleName].Font.Size = 14;
                    wordApp.ActiveDocument.Styles[oStyleName].Font.Name = "微软雅黑";
                    wordApp.ActiveDocument.Styles[oStyleName].Font.Bold = 2;
                    wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceBefore = 0;
                    wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceAfter = 0;
                    wordApp.Selection.set_Style(ref oStyleName);
                }

                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                wordApp.Selection.TypeParagraph();
                wordApp.Selection.Font.Size = 10F;
                wordApp.Selection.Font.Bold = 0;
                string headerText = "";
                if (seqNo > 0)
                {
                    wordApp.Selection.Font.Name = "宋体";
                    headerText += "续表" + seqNo;
                    headerText += "\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t 单位：" + dataCompile.UnitName;
                    wordApp.Selection.TypeText(headerText);
                    
                    wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else
                {
                    headerText += "单位：" + dataCompile.UnitName;
                    wordApp.Selection.TypeText(headerText);
                    wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                }
                seqNo++;
                wordApp.Selection.TypeParagraph();
                newTable = wordDoc.Tables.Add(wordApp.Selection.Range, sumRows + 3, columnsPerPage[n], ref Nothing, ref Nothing);

                //设置表格样式
                newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                newTable.Range.Font.Size = 9F;

                //设置表格列宽
                if ((dataCompile.SelectedRowName.Equals("仪器分类大类") || dataCompile.SelectedRowName.Equals("实验基地")) && n == 0)//第一页需要将第一列设置较宽
                {
                    if (newTable.Columns.Count < 7)
                    {
                        newTable.Columns[1].Width *= 2;
                    }
                    else
                    {
                        newTable.Columns[1].Width *= 3;
                    }
                    //按窗口内容自动调整表格宽度
                    newTable.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
                }
                //设置行高
                newTable.Rows.Height = maxTableHeight / (sumRows + 4);
                //newTable.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly; //固定行高
                newTable.Rows[2].Height = newTable.Rows.Height * 2;
                wordApp.ActiveDocument.Paragraphs.Last.Range.Select();

                tableList.Add(newTable);

                //插入分页符
                object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
                wordApp.Application.Selection.InsertBreak(oPageBreak);

                wordDoc.Paragraphs.Last.Range.Select();
                n++;
            }

            return tableList;
        }

        private void GetSumRowsAndColumns()
        {
            if (dataCompile == null)
            {
                return;
            }
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

        protected List<string> GetBaseBigType()
        {
            List<string> typeList = new List<string>();

            //view_研究实验基地概况
            string tableName = SharedData.tableListInDB[4].ViewName;
            string groupBy1 = "";
            int i = 0;
            for (i = 0; i < SharedData.tableListInDB[4].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[4].ColumnNames[i].IndexOf("基地大类") != -1)
                {
                    groupBy1 = SharedData.tableListInDB[4].ColumnNames[i];
                    break;
                }
            }
            if (groupBy1 == "")
            {
                return typeList;
            }
            string sqlStr = @"select [" + groupBy1 + "] from view_研究实验基地概况 group by [" + groupBy1 + "]";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();

            try
            {
                SqlCommand command = new SqlCommand(sqlStr, sqlConn);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (reader[groupBy1] != null && reader[groupBy1].ToString().Length > 0)
                    {
                        typeList.Add(reader[groupBy1].ToString());
                    }
                }
                reader.Close();
                sqlConn.Close();
            }
            catch (System.Exception ex)
            {
                Log.RecordLog(ex.ToString());
            }
            return typeList;
        }
    }
}
