using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-9 研究实验基地科技活动经费支出2
    class ResearchInstitution3_9_2 : AbstractTable
    {
        public ResearchInstitution3_9_2()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-9  研究实验基地科技活动经费支出2";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("其中：基地运行费");
            dataCompile.SelectedColumnListA.Add("其中：仪器设备购置费");

            dataCompile.SummaryColumnNameA = "科技活动经费支出";
            dataCompile.SelectedRowName = "所在省";

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 1;
            sumSameTable = 2;
            //不需要除以10000
            IsNeedTransTenThousand = false;

            //用于检查列名
            columnCheckList.Add("其中：基地运行费");
            columnCheckList.Add("其中：仪器设备购置费");
            columnCheckList.Add("科技活动经费支出");
            columnCheckList.Add("全国汇编用基地小类");
        }
    }
}
