using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //大型科学仪器设备：2-9 大型科学仪器设备运行状态2
    class ResearchInstitution2_9_2 : AbstractTable
    {
        public ResearchInstitution2_9_2()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-9 大型科学仪器设备运行状态2";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台（套）";
            dataCompile.SelectedColumnListA.Add("运行状态");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SelectedRowName = "所在省";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 1;
            sumSameTable = 2;
            summaryColumnA = "运行状态";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }
    }
}
