using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Data.SqlClient;

namespace DFSystem.Tables
{
    //科研院所和高校概况：1-12 科研院所和高校高层次人员概况
    class ResearchInstitution1_12 : AbstractTable
    {
        public ResearchInstitution1_12()
        {
            tableCaption = "1-12 科研院所和高校高层次人员概况";

            dataTableName = SharedData.tableListInDB[2].ViewName;

            unitName = "人";

            //用于检查列名
            columnCheckList.Add("本年新增正高级职称");
            columnCheckList.Add("本年新增副高级职称");
            columnCheckList.Add("本年新增博士");
            columnCheckList.Add("本年新增人员总数");
            columnCheckList.Add("本年减少正高级职称");
            columnCheckList.Add("本年减少副高级职称");
            columnCheckList.Add("本年减少博士");
            columnCheckList.Add("本年减少人员总数");

            //实际显示的列名，名称和sql语句中列的别名一致
            finalListA.Add("新增");
            finalListA.Add("本年新增正高级职称");
            finalListA.Add("本年新增副高级职称");
            finalListA.Add("本年新增博士");
            finalListA.Add("减少");
            finalListA.Add("本年减少正高级职称");
            finalListA.Add("本年减少副高级职称");
            finalListA.Add("本年减少博士");
        }

        public override bool GenerateSQL()
        {
            Log.RecordLog(tableCaption + ": GenerateSQL");
            //判断列是否存在
            if (!CheckColumnNames(columnCheckList))
            {
                return false;
            }

            sqlStrForCompile = "select 所在省, convert(bigint, sum(cast(\"本年新增正高级职称\" as float)))  as \"本年新增正高级职称\", convert(bigint, sum(cast(\"本年新增副高级职称\" as float)))  as \"本年新增副高级职称\", convert(bigint, sum(cast(\"本年新增博士\" as float)))  as \"本年新增博士\", ( convert(bigint, sum(cast(\"本年新增人员总数\" as float)))) as '新增', convert(bigint, sum(cast(\"本年减少正高级职称\" as float)))  as \"本年减少正高级职称\", convert(bigint, sum(cast(\"本年减少副高级职称\" as float)))  as \"本年减少副高级职称\", convert(bigint, sum(cast(\"本年减少博士\" as float)))  as \"本年减少博士\", ( convert(bigint, sum(cast(\"本年减少人员总数\" as float)))) as \"减少\" FROM ["
            + dataTableName + "] group by 所在省";
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
            newTable = wordDoc.Tables.Add(wordApp.Selection.Range, 34, finalListA.Count + 1, ref Nothing, ref Nothing);

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

            //填充列名
            int listAColumnSeq = 0;
            for (i = 2; i < finalListA.Count + 2; i++)
            {
                newTable.Cell(2, i).Range.Text = finalListA[listAColumnSeq++];
                newTable.Cell(2, i).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //记录每一个列名对应的列号
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
                for (j = 1; j < finalListA.Count + 2; j++)
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

                string value;

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

            //处理表头，合并单元格
            newTable.Cell(1, 1).Merge(newTable.Cell(2, 1));
            newTable.Cell(1, 2).Merge(newTable.Cell(2, 2));
            newTable.Cell(1, 6).Merge(newTable.Cell(2, 6));
            newTable.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            newTable.Cell(1, 6).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            for (i = 3; i < finalListA.Count + 2; i++)
            {
                if (i != 6)
                {
                    newTable.Cell(1, i).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                }
            }
            newTable.Cell(2, 3).Range.Text = newTable.Cell(2, 7).Range.Text = "#正高级";
            newTable.Cell(2, 4).Range.Text = newTable.Cell(2, 8).Range.Text = "#副高级";
            newTable.Cell(2, 5).Range.Text = newTable.Cell(2, 9).Range.Text = "#博士";
        }
    }
}
