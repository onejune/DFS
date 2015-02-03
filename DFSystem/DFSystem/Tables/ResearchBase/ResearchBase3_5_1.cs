using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-5 研究实验基地仪器设备概况1
    class ResearchInstitution3_5_1 : AbstractTable
    {
        public ResearchInstitution3_5_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-5  研究实验基地仪器设备概况1";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "台(套)、万元";

            dataCompile.SelectedColumnListA.Add("#大型仪器表设备总数");
            dataCompile.SummaryColumnNameA = "科研仪器设备总数";
            dataCompile.SummaryColumnFromNameA = "";
            dataCompile.SelectedColumnListB.Add("#大型仪器表原值总额");
            dataCompile.SummaryColumnNameB = "科研仪器设备原值总额";
            dataCompile.SummaryColumnFromNameB = "";

            dataCompile.SelectedRowName = "实验基地";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;

            //用于检查列名
            columnCheckList.Add("科研仪器设备总数");
            columnCheckList.Add("科研仪器设备原值总额");
        }
    }
}
