using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-15 植物种质资源类型
    class ResearchInstitution4_15 : AbstractTable
    {
        public ResearchInstitution4_15()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-15 植物种质资源类型";
            dataCompile.TableName = SharedData.tableListInDB[11].ViewName;
            dataCompile.UnitName = "份";
            dataCompile.SelectedColumnListA.Add("资源类型");
            dataCompile.SummaryColumnNameA = "植物资源总数";
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
            for (int i = 0; i < SharedData.tableListInDB[11].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[11].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[11].ColumnNames[i];
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
            columnCheckList.Add("植物资源总数");
        }
    }
}
