using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-12 生物种质资源保存机构科技活动经费收入
    class ResearchInstitution4_12: AbstractTable
    {
        public ResearchInstitution4_12()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-12 生物种质资源保存机构科技活动经费收入";
            dataCompile.TableName = SharedData.tableListInDB[8].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("来源：中央财政资助");
            dataCompile.SelectedColumnListA.Add("来源：地方财政资助");
            dataCompile.SelectedColumnListA.Add("来源：企事业资金");
            dataCompile.SelectedColumnListA.Add("来源：单位自有资金");
            dataCompile.SelectedColumnListA.Add("来源：其他资金");

            dataCompile.SummaryColumnNameA = "科技活动经费收入";
            dataCompile.SummaryColumnFromNameA = "";

            if (SharedData.DEBUGGING == true)
            {
                dataCompile.SelectedRowName = "机构所在省";
            }
            else
            {
                dataCompile.SelectedRowName = "";
            }
            bool hasFound = false;
            for (int i = 0; i < SharedData.tableListInDB[8].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[8].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[8].ColumnNames[i];
                    hasFound = true;
                    break;
                }
            }
            if (SharedData.DEBUGGING == false && hasFound != true)
            {
                HasError = true;
                Log.RecordLog(dataCompile.Caption + ": " + dataCompile.TableName + ": 找不到\"所在省\"列！");
            }

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 1;
            //除以10000
            IsNeedTransTenThousand = false;

            //用于检查列名
            for (int i = 0; i < dataCompile.SelectedColumnListA.Count(); i++)
            {
                columnCheckList.Add(dataCompile.SelectedColumnListA[i]);
            }
        }
    }
}
