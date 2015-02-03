﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Data.SqlClient;

namespace DFSystem.Tables
{
    //科研院所和高校概况：2-3 按购置年代分段的大型科学仪器设备数量1
    class ResearchInstitution2_3_1 : AbstractTable
    {
        public ResearchInstitution2_3_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-3 按购置年代分段的大型科学仪器设备数量1";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台（套）";
            dataCompile.SelectedColumnListA.Add("建账区间");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SelectedRowName = "仪器分类大类";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;
            summaryColumnA = "建账区间";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }
    }
}
