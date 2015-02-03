using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-11 生物种质资源保存机构人才培养情况3
    class ResearchInstitution4_11_3 : AbstractTable
    {
        public ResearchInstitution4_11_3()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-11  生物种质资源保存机构人才培养情况3";
            dataCompile.TableName = SharedData.tableListInDB[7].ViewName;
            dataCompile.UnitName = "人";
            dataCompile.SelectedColumnListA.Add("博士后入学（进站）人数");
            dataCompile.SelectedColumnListA.Add("博士后毕业（出站）人数");
            dataCompile.SelectedColumnListA.Add("博士后在读（在站）人数");

            dataCompile.SummaryColumnNameA = "博士后人数";
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
            for (int i = 0; i < SharedData.tableListInDB[7].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[7].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[7].ColumnNames[i];
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
            seqNo = 2;
            sumSameTable =3;
            //除以11000
            IsNeedTransTenThousand = false;

            //用于检查列名
            columnCheckList.Add("博士后在读（在站）人数");
            columnCheckList.Add("博士后入学（进站）人数");
            columnCheckList.Add("博士后毕业（出站）人数");
            columnCheckList.Add("博士后在读");
        }
    }
}
