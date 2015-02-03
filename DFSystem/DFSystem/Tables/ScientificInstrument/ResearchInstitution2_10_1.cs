using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //大型科学仪器设备：2-10 大型科学仪器设备共享情况1
    class ResearchInstitution2_10_1 : AbstractTable
    {
        public ResearchInstitution2_10_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-10 大型科学仪器设备共享情况1";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台（套）";
            dataCompile.SelectedColumnListA.Add("共享模式");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SelectedRowName = "仪器分类大类";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;
            summaryColumnA = "共享模式";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);

        }
    }
}
