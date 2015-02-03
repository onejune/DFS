using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //大型科学仪器设备：2-11 大型科学仪器设备使用状态及机时1
    class ResearchInstitution2_11_1 : AbstractTable
    {
        public ResearchInstitution2_11_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-11 大型科学仪器设备使用状态及机时1";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台(套)、万小时";
            dataCompile.SelectedColumnListA.Add("使用状态");
            dataCompile.SummaryColumnNameA = "仪器设备数";
            dataCompile.SummaryColumnFromNameA = "";
            dataCompile.SelectedColumnListB.Add("年对外服务机时");
            dataCompile.SummaryColumnNameB = "年有效工作机时";
            dataCompile.SummaryColumnFromNameB = "年有效工作机时";

            dataCompile.SelectedRowName = "仪器分类大类";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;
            summaryColumnA = "使用状态";
            summaryColumnB = "年对外服务机时";

            //用于检查列名
            columnCheckList.Add(summaryColumnA);
            columnCheckList.Add(summaryColumnB);
        }
    }
}
