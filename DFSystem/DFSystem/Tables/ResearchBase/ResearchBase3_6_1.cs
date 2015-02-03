using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-6 研究实验基地科技活动人员基本情况1
    class ResearchInstitution3_6_1 : AbstractTable
    {
        public ResearchInstitution3_6_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-6  研究实验基地科技活动人员基本情况1";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "人";
            dataCompile.SelectedColumnListA.Add("正高级");
            dataCompile.SelectedColumnListA.Add("副高级");
            dataCompile.SelectedColumnListA.Add("博士");

            dataCompile.SummaryColumnNameA = "科技活动人员数";
            dataCompile.SelectedRowName = "实验基地";

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;

            //用于检查列名
            columnCheckList.Add("正高级");
            columnCheckList.Add("副高级");
            columnCheckList.Add("博士");
            columnCheckList.Add("科技活动人员数");
            columnCheckList.Add("全国汇编用基地小类");
        }
    }
}
