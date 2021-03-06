﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Data.SqlClient;

namespace DFSystem.Tables
{
    //科研院所和高校概况：2-7 按获取方式分类的大型科学仪器设备数量2
    class ResearchInstitution2_7_2 : AbstractTable
    {
        public ResearchInstitution2_7_2()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-7 按获取方式分类的大型科学仪器设备数量2";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台（套）";
            dataCompile.SelectedColumnListA.Add("获取方式");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SelectedRowName = "所在省";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 1;
            sumSameTable = 2;
            summaryColumnA = "获取方式";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }
    }
}
