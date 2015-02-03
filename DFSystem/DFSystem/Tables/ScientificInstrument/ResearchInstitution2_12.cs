using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //大型科学仪器设备：2-12 各类型大型科学仪器设备数量
    class ResearchInstitution2_12 : AbstractTable
    {
        public ResearchInstitution2_12()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "2-12 各类型大型科学仪器设备数量1";
            dataCompile.TableName = SharedData.tableListInDB[3].ViewName;
            dataCompile.UnitName = "台(套)";
            dataCompile.SelectedColumnListA.Add("仪器分类大类");
            dataCompile.SummaryColumnNameA = "合计";
            dataCompile.SummaryColumnFromNameA = "";

            dataCompile.SelectedRowName = "所在省";
            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;

            //用于检查列名
        }
    }
}
