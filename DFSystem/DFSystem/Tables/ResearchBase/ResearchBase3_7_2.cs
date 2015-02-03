using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-7 研究实验基地科技活动人员变动情况2
    class ResearchInstitution3_7_2 : AbstractTable
    {
        public ResearchInstitution3_7_2()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-7  研究实验基地科技活动人员变动情况2";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "人";
            dataCompile.SelectedColumnListA.Add("当年科技活动人员新增");
            dataCompile.SelectedColumnListA.Add("当年科技活动人员减少");

            dataCompile.SummaryColumnNameA = "科技活动人员变动总数";
            dataCompile.SummaryColumnFromNameA = "";
            dataCompile.SelectedRowName = "所在省";

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 1;
            sumSameTable = 2;

            //用于检查列名
            columnCheckList.Add("当年科技活动人员新增");
            columnCheckList.Add("当年科技活动人员减少");
            columnCheckList.Add("全国汇编用基地小类");
        }
    }
}
