﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-3 研究实验基地依托单位类型2
    class ResearchInstitution3_3_2 : AbstractTable
    {
        public ResearchInstitution3_3_2()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-3  研究实验基地依托单位类型2";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "个";
            dataCompile.SelectedColumnListA.Add("单位属性-汇编用");

            dataCompile.SummaryColumnNameA = "基地数";
            dataCompile.SummaryColumnFromNameA = "";
            dataCompile.SelectedRowName = "所在省";

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 1;
            sumSameTable = 2;

            //用于检查列名
            columnCheckList.Add("单位属性-汇编用");
            columnCheckList.Add("全国汇编用基地小类");
            columnCheckList.Add("所在省");
        }
    }
}
