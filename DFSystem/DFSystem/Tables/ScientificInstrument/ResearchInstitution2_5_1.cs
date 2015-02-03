using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //科研院所和高校概况：2-5 不同产地的大型科学仪器设备数量1
    class ResearchInstitution2_5_1 : AbstractTable
    {
        public ResearchInstitution2_5_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-5 不同产地的大型科学仪器设备数量1";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台（套）";
            dataCompile.SelectedColumnListA.Add("产地");
            dataCompile.SummaryColumnNameA = "数量";
            dataCompile.SelectedRowName = "仪器分类大类";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;
            summaryColumnA = "产地";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }
    }
}
