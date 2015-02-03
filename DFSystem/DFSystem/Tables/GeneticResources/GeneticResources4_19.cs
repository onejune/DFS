using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-19 动物种质资源类型
    class ResearchInstitution4_19 : AbstractTable
    {
        public ResearchInstitution4_19()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-19 动物种质资源类型";
            dataCompile.TableName = SharedData.tableListInDB[10].ViewName;
            dataCompile.UnitName = "种";
            dataCompile.SelectedColumnListA.Add("资源类型");
            dataCompile.SummaryColumnNameA = "动物种质资源数";
            dataCompile.SummaryColumnFromNameA = "";

            if (SharedData.DEBUGGING == true)
            {
                dataCompile.SelectedRowName = "单位所在省";
            }
            else
            {
                dataCompile.SelectedRowName = "";
            }
            bool hasFound = false;
            for (int i = 0; i < SharedData.tableListInDB[10].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[10].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[10].ColumnNames[i];
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
            columnCheckList.Add("资源类型");
            columnCheckList.Add("动物种质资源数");
        }
    }
}
