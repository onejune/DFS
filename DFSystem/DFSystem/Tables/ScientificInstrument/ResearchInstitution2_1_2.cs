using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Data.SqlClient;

namespace DFSystem.Tables
{
    //科研院所和高校概况：2-1 按单台原值分段的大型科学仪器设备数量2
    class ResearchInstitution2_1_2 : AbstractTable
    {
        public ResearchInstitution2_1_2()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-1 按单台原值分段的大型科学仪器设备数量2";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台（套）";
            dataCompile.SelectedColumnListA.Add("原值区间");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SelectedRowName = "所在省";

            //续表号
            seqNo = 1;
            sumSameTable = 2;
            summaryColumnA = "原值区间";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }
    }
}
