using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //大型科学仪器设备：2-15 各类型大型分析仪器设备原值
    class ResearchInstitution2_15 : AbstractTable
    {
        public ResearchInstitution2_15()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-15 各类型大型分析仪器设备原值";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "万元";
            dataCompile.SelectedColumnListA.Add("仪器分类中类");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SummaryColumnFromNameA = "";

            dataCompile.SelectedRowName = "所在省";
            dataCompile.SumColumnName = "原值";

            //续表号
            seqNo = 0;
            sumSameTable = 2;

            //用于检查列名
        }
    }
}
