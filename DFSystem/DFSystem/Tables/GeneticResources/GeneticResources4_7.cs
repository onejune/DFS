using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-7 生物种质资源保存机构固定资产
    class ResearchInstitution4_7 : AbstractTable
    {
        public ResearchInstitution4_7()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-7  生物种质资源保存机构固定资产";
            dataCompile.TableName = SharedData.tableListInDB[5].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("设施用房（万元）");
            dataCompile.SelectedColumnListA.Add("仪器设备（万元）");
            dataCompile.SelectedColumnListA.Add("图书资料（万元）");
            dataCompile.SelectedColumnListA.Add("其他（万元）");

            dataCompile.SummaryColumnNameA = "固定资产";
            dataCompile.SummaryColumnFromNameA = "固定资产总额（万元）";
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
            //除以10000
            IsNeedTransTenThousand = false;

            //用于检查列名
            columnCheckList.Add("固定资产总额（万元）");
            columnCheckList.Add("仪器设备（万元）");
            columnCheckList.Add("图书资料（万元）");
            columnCheckList.Add("设施用房（万元）");
            columnCheckList.Add("其他（万元）");
        }
    }
}
