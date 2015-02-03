using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-6 生物种质资源保存机构信息化情况
    class ResearchInstitution4_6 : AbstractTable
    {
        public ResearchInstitution4_6()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-6  生物种质资源保存机构信息化情况";
            dataCompile.TableName = SharedData.tableListInDB[5].ViewName;
            dataCompile.UnitName = "个";
            dataCompile.SelectedColumnListA.Add("资源信息化情况");
            dataCompile.SummaryColumnNameA = "保存机构数";
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
            for (int i = 0; i < SharedData.tableListInDB[5].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[5].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[5].ColumnNames[i];
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
            //不需要除以10000
            IsNeedTransTenThousand = false;

            //用于检查列名
            columnCheckList.Add("资源信息化情况");
            columnCheckList.Add("保存机构数");
        }
    }
}
