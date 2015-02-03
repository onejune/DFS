using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Data.SqlClient;

namespace DFSystem.Tables
{
    //科研院所和高校概况：2-2 按单台原值分段的大型科学仪器设备原值1
    class ResearchInstitution2_2_1 : AbstractTable
    {
        public ResearchInstitution2_2_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-2 按单台原值分段的大型科学仪器设备原值1";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("原值区间");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SelectedRowName = "仪器分类大类";
            dataCompile.SumColumnName = "原值";

            //续表号
            seqNo = 0;
            sumSameTable = 2;
            summaryColumnA = "原值区间";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }

        // 对每个省对应的table进行数据填充
        public override void FillTable(List<Table> newTableList)
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
                if (finalListB != null)
                {
                    mergeCount = mergeCount - finalListB.Count - 1;
                }
                for (i = 0; i < mergeCount; i++)
                {
                    newTable.Cell(1, begin).Merge(newTable.Cell(1, begin + 1));
                }
                if (finalListB != null)
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
                if (finalListB != null)
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

                if (finalListB != null)
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

                // **************************结果要除以1000**************************
                if (value != "")
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

                    // **************************结果要除以1000**************************
                    if (value != "")
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
                }
            }
        }
    }
}
