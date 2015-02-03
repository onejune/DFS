using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-1 研究实验基地数4
    class ResearchInstitution3_1_4 : AbstractTable
    {
        public ResearchInstitution3_1_4()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-1  研究实验基地数4";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "个";
            dataCompile.SelectedColumnListA.Add("全国汇编用基地小类");
            //获取基地大类
            List<string> baseBigType = GetBaseBigType();
            dataCompile.SummaryColumnNameA = "";
            for (int i = 0; i < baseBigType.Count(); i++)
            {
                if (baseBigType[i].IndexOf("分析测试中心") != -1)
                {
                    dataCompile.SummaryColumnNameA = baseBigType[i];
                    break;
                }
            }
            if (dataCompile.SummaryColumnNameA == "")
            {
                HasError = true;
                Log.RecordLog(dataCompile.Caption + ": " + dataCompile.TableName + ": 找不到\"分析测试中心\"列！");
                return;
            }

            dataCompile.SummaryColumnFromNameA = "";
            dataCompile.SelectedRowName = "";
            for (int i = 0; i < SharedData.tableListInDB[4].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[4].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[4].ColumnNames[i];
                    break;
                }
            }
            if (dataCompile.SelectedRowName == "")
            {
                HasError = true;
                Log.RecordLog(dataCompile.Caption + ": " + dataCompile.TableName + ": 找不到\"所在省\"列！");
            }

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;

            //用于检查列名
            columnCheckList.Add("全国汇编用基地小类");
            columnCheckList.Add("基地大类");
        }
    }
}
