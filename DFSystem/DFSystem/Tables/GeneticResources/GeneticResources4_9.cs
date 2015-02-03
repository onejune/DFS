﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //生物种质资源：4-9 生物种质资源保存机构人员职称情况
    class ResearchInstitution4_9 : AbstractTable
    {
        public ResearchInstitution4_9()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "4-9  生物种质资源保存机构人员职称情况";
            dataCompile.TableName = SharedData.tableListInDB[6].ViewName;
            dataCompile.UnitName = "人";
            dataCompile.SelectedColumnListA.Add("职称级别");

            dataCompile.SummaryColumnNameA = "科技活动人员数";
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
            for (int i = 0; i < SharedData.tableListInDB[6].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[6].ColumnNames[i].IndexOf("所在省") != -1)
                {
                    dataCompile.SelectedRowName = SharedData.tableListInDB[6].ColumnNames[i];
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
            columnCheckList.Add("职称级别");
            columnCheckList.Add("科技活动人员数");
        }
    }
}
