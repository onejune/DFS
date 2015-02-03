using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-8 研究实验基地科技活动经费收入1
    class ResearchInstitution3_8_1 : AbstractTable
    {
        public ResearchInstitution3_8_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-8  研究实验基地科技活动经费收入1";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("其中：政府资金");
            dataCompile.SelectedColumnListA.Add("其中：技术性收入");

            dataCompile.SummaryColumnNameA = "科技活动经费收入";
            dataCompile.SelectedRowName = "实验基地";

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;
            //不需要除以10000
            IsNeedTransTenThousand = false;

            //用于检查列名
            columnCheckList.Add("其中：政府资金");
            columnCheckList.Add("其中：技术性收入");
            columnCheckList.Add("科技活动经费收入");
            columnCheckList.Add("全国汇编用基地小类");
        }
    }
}
