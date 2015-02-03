using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-14 植物、动物种质资源编目数
    class ResearchInstitution4_14 : AbstractTable
    {
        public ResearchInstitution4_14()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-14 植物、动物种质资源编目数";
            dataCompile.TableName = SharedData.tableListInDB[5].ViewName;
            dataCompile.UnitName = "份、种";
            dataCompile.SelectedColumnListA.Add("#植物：国家编目数");
            dataCompile.SummaryColumnNameA = "植物资源总数";
            dataCompile.SummaryColumnFromNameA = "";

            dataCompile.SelectedColumnListB.Add("#动物：国家编目数");
            dataCompile.SummaryColumnNameB = "动物资源总数";
            dataCompile.SummaryColumnFromNameB = "";

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
            //除以10000
            IsNeedTransTenThousand = false;

            //用于检查列名
            columnCheckList.Add("植物：国家编目数");
            columnCheckList.Add("动物：国家编目数");
            columnCheckList.Add("动物资源总数");
            columnCheckList.Add("植物资源总数");
        }
    }
}
