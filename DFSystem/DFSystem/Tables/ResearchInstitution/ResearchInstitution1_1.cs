using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Word;

namespace DFSystem.Tables
{
    //科研院所和高校概况：1-1 科研院所和高校数
    class ResearchInstitution1_1 : AbstractTable
    {
        public ResearchInstitution1_1()
        {
            tableCaption = "1-1 科研院所和高校数";
            dataTableName = SharedData.tableListInDB[0].ViewName;
            unitName = "个";
        }

        public override bool GenerateSQL()
        {
            Log.RecordLog(tableCaption + ": GenerateSQL");

            //用于检查列名
            columnCheckList.Add("所在省");
            columnCheckList.Add("单位属性");

            //判断列是否存在
            if (!CheckColumnNames(columnCheckList))
            {
                return false;
            }

            sqlStrForCompile = "select a.\"所在省\",c.\"科研机构\",a.\"高等学校\",d.\"转制院所\",b.\"检测机构\",e.\"单位数\" from (select x.provinceName as \"所在省\", y.[高等学校] from (select provinceName from [DFS].[dbo].[所在省]) x left join (select \"所在省\", count(*)  as '高等学校' from [" + dataTableName + "] where 单位属性 like '%高等学校%' group by \"所在省\") y on x.provinceName like '%' + y.\"所在省\" + '%') a,(select x.provinceName as \"所在省\", y.[检测机构] from (select provinceName from [DFS].[dbo].[所在省]) x left join (select \"所在省\", count(*)  as '检测机构' from [" + dataTableName + "] where 单位属性 like '%检测机构%' group by \"所在省\") y on x.provinceName like '%' + y.\"所在省\" + '%') b,(select x.provinceName as \"所在省\", y.[科研机构] from (select provinceName from [DFS].[dbo].[所在省]) x left join (select \"所在省\", count(*)  as '科研机构' from [" + dataTableName + "] where 单位属性 like '%科研机构%' group by \"所在省\") y on x.provinceName like '%' + y.\"所在省\" + '%') c,(select x.provinceName as \"所在省\", y.[转制院所] from (select provinceName from [DFS].[dbo].[所在省]) x left join (select \"所在省\", count(*)  as '转制院所' from [" + dataTableName + "] where 单位属性 like '%转制院所%' group by \"所在省\") y on x.provinceName like '%' + y.\"所在省\" + '%') d,(select x.provinceName as \"所在省\", y.[单位数] from (select provinceName from [DFS].[dbo].[所在省]) x left join (select \"所在省\", count(*)  as '单位数' from [" + dataTableName + "]  group by \"所在省\") y on x.provinceName like '%' +y.\"所在省\" + '%') e where a.[所在省]=b.[所在省] and b.[所在省]=c.[所在省] and c.[所在省]=d.[所在省] and d.[所在省]=e.[所在省]";

            return true;
        }

        public override List<Table> InsertTable()
        {
            Log.RecordLog(tableCaption + ": InsertTable");
            List<Table> tableList = new List<Table>();

            Microsoft.Office.Interop.Word.Table newTable;

            wordDoc.Paragraphs.First.Range.Select();
            wordApp.Selection.TypeText(tableCaption);

            // 设置标题样式，以方便插入目录
            object oStyleName = "标题 2";
            wordApp.ActiveDocument.Styles[oStyleName].Font.Size = 14;
            wordApp.ActiveDocument.Styles[oStyleName].Font.Name = "微软雅黑";
            wordApp.ActiveDocument.Styles[oStyleName].Font.Bold = 2;
            wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceBefore = 0;
            wordApp.ActiveDocument.Styles[oStyleName].ParagraphFormat.SpaceAfter = 0;
            wordApp.Selection.set_Style(ref oStyleName);

            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            wordApp.Selection.TypeParagraph();
            wordApp.Selection.Font.Size = 10F;
            wordApp.Selection.Font.Bold = 0;
            string headerText = "";

            headerText += "单位：" + unitName;
            wordApp.Selection.TypeText(headerText);
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            wordApp.Selection.TypeParagraph();
            newTable = wordDoc.Tables.Add(wordApp.Selection.Range, 34, 5, ref Nothing, ref Nothing);

            //设置表格样式
            newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            newTable.Range.Font.Size = 9F;

            //设置行高
            newTable.Rows.Height = maxTableHeight / 34;
            //newTable.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly; //固定行高
            newTable.Rows[2].Height = newTable.Rows.Height * 2;
            wordApp.ActiveDocument.Paragraphs.Last.Range.Select();

            tableList.Add(newTable);

            //插入分页符
            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            wordApp.Application.Selection.InsertBreak(oPageBreak);

            wordDoc.Paragraphs.Last.Range.Select();
            return tableList;
        }

        public override void FillTable(List<Table> table)
        {
            List<double> sumList = new List<double>();
            
            Log.RecordLog(tableCaption + ": FillTable");
            int i = 1, k = 1, j = 0;

            //listA中的各列
            finalListA.Add("科研机构");
            finalListA.Add("高等学校");
            finalListA.Add("转制院所");

            //初始化每一列总计列表
            for (i = 0; i < finalListA.Count + 2; i++)
            {
                sumList.Add(0);
            }
            List<int> columnNumberOfListAColumn = new List<int>();
            

            Table newTable = table[0];

            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();
            SqlCommand command = new SqlCommand(sqlStrForCompile, sqlConn);
            command.CommandTimeout = 180;
            SqlDataReader reader = command.ExecuteReader();

            int tableColumns = newTable.Columns.Count;
            newTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            newTable.Cell(1, 1).Merge(newTable.Cell(2, 1));
            newTable.Cell(1, 2).Merge(newTable.Cell(2, 2));

            //合并第一行的单元格
            int begin, mergeCount;
            begin = 3;
            mergeCount = tableColumns - 3;

            for (i = 0; i < mergeCount; i++)
            {
                newTable.Cell(1, begin).Merge(newTable.Cell(1, begin + 1));
            }
            newTable.Cell(1, 3).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;

            newTable.Cell(1, 2).Range.Text = "单位数";
            newTable.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            //填充列名
            int listAColumnSeq = 0;
            for (i = 3; i <= 5 && i < finalListA.Count + 3; i++)
            {
                newTable.Cell(2, i).Range.Text = finalListA[listAColumnSeq++];
                newTable.Cell(2, i).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                columnNumberOfListAColumn.Add(i);
            }

            newTable.Cell(3, 1).Range.Text = "总计";

            for (i = 0; i < 5; i++)
            {
                newTable.Cell(3, i + 1).Range.Font.Bold = 2;
            }
            //设置边线
            for (i = 3; i <= 34; i++)
            {
                for (j = 1; j <= 5; j++)
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

            newTable.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            newTable.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
            newTable.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

            k = 4;
            j = 1;
            while (reader.Read())
            {
                //寻找行号，k是该条纪录对应的table的行号
                string text = reader["所在省"].ToString();
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

                //第二列，汇总列
                string value = reader["单位数"].ToString().Trim();
                if (value.IndexOf('.') != -1)
                {
                    value = value.Substring(0, value.IndexOf('.') + 2);
                }
                if (!value.Equals("0") && !value.Equals(""))
                {
                    newTable.Cell(k, 2).Range.Text = value;
                    if (value.IndexOf('.') != -1)
                    {
                        value = value.Substring(0, value.IndexOf('.') + 2);
                    }
                    if (value != null && value.Length > 0)
                    {
                        sumList[1] += double.Parse(value);
                    }
                }

                for (i = 0; i < finalListA.Count; i++)
                {
                    value = reader[finalListA[i]].ToString().Trim();
                    if (value.IndexOf('.') != -1)
                    {
                        value = value.Substring(0, value.IndexOf('.') + 2);
                    }
                    if (!value.Equals("0") && !value.Equals(""))
                    {
                        newTable.Cell(k, columnNumberOfListAColumn[i]).Range.Text = value;
                        if (value != null && value.Length > 0)
                        {
                            sumList[i + 2] += double.Parse(value);
                        }
                    }
                }
            }

            reader.Close();
            sqlConn.Close();

            //填充总计行
            string v = sumList[1].ToString();
            if (v.IndexOf('.') != -1)
            {
                v = v.Substring(0, v.IndexOf('.') + 2);
            }
            newTable.Cell(3, 2).Range.Text = v;
            for (i = 2; i < sumList.Count; i++)
            {
                v = sumList[i].ToString();
                if (v.IndexOf('.') != -1)
                {
                    v = v.Substring(0, v.IndexOf('.') + 2);
                }
                newTable.Cell(3, columnNumberOfListAColumn[i - 2]).Range.Text = v;
            }
        }
    }
}
