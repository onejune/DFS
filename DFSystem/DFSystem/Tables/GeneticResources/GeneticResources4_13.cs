using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-13 生物种质资源保存机构科技活动经费支出
    class ResearchInstitution4_13 : AbstractTable
    {
        public ResearchInstitution4_13()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-13 生物种质资源保存机构科技活动经费支出";
            dataCompile.TableName = SharedData.tableListInDB[8].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("支出：设备运行费");
            dataCompile.SelectedColumnListA.Add("支出：房屋使用费");
            dataCompile.SelectedColumnListA.Add("支出：水电能源费");
            dataCompile.SelectedColumnListA.Add("支出：材料费");
            dataCompile.SelectedColumnListA.Add("支出：人员费");
            dataCompile.SelectedColumnListA.Add("支出：办公费");

            dataCompile.SummaryColumnNameA = "科技活动经费支出";
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
