using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem.Tables
{
    //研究实验基地：3-7 研究实验基地科技活动人员变动情况1
    class ResearchInstitution3_7_1 : AbstractTable
    {
        public ResearchInstitution3_7_1()
        {
            dataCompile = new DataCompile();
            dataCompile.Caption = "3-7  研究实验基地科技活动人员变动情况1";
            dataCompile.TableName = SharedData.tableListInDB[4].ViewName;
            dataCompile.UnitName = "人";
            dataCompile.SelectedColumnListA.Add("当年科技活动人员新增");
            dataCompile.SelectedColumnListA.Add("当年科技活动人员减少");

            dataCompile.SummaryColumnNameA = "科技活动人员变动总数";
            dataCompile.SummaryColumnFromNameA = "";
            dataCompile.SelectedRowName = "实验基地";

            dataCompile.SumColumnName = "";

            //续表号
            seqNo = 0;
            sumSameTable = 2;

            //用于检查列名
            columnCheckList.Add("当年科技活动人员新增");
            columnCheckList.Add("当年科技活动人员减少");
            columnCheckList.Add("全国汇编用基地小类");

            finalListA.Add("当年科技活动人员新增");
            finalListA.Add("当年科技活动人员减少");
        }

        public override bool GenerateSQL()
        {
            Log.RecordLog(tableCaption + ": GenerateSQL");
            //判断列是否存在
            if (!CheckColumnNames(columnCheckList))
            {
                return false;
            }

            sqlStrForCompile = "select b.rowName, b.\"当年科技活动人员新增\"+c.\"当年科技活动人员减少\" as \"科技活动人员变动总数\",b.\"当年科技活动人员新增\",c.\"当年科技活动人员减少\" from (select x.rowName, y.\"当年科技活动人员新增\" from(select type as rowName from 实验基地) x  left join (SELECT 基地大类 as rowName, CONVERT(bigint, sum(cast(\"当年科技活动人员新增\" as float)))  as \"当年科技活动人员新增\" FROM  [" + dataCompile.TableName + "]  group by 基地大类  union SELECT 全国汇编用基地小类,  CONVERT(bigint, sum(cast(\"当年科技活动人员新增\" as float)))  as \"当年科技活动人员新增\" FROM [" + dataCompile.TableName + "]  group by 全国汇编用基地小类) y on x.rowName=y.rowName )b, (select x.rowName, y.\"当年科技活动人员减少\" from(select type as rowName from 实验基地) x  left join (SELECT 基地大类 as rowName, CONVERT(bigint, sum(cast(\"当年科技活动人员减少\" as float)))  as \"当年科技活动人员减少\" FROM  [" + dataCompile.TableName + "]  group by 基地大类  union SELECT 全国汇编用基地小类,  CONVERT(bigint, sum(cast(\"当年科技活动人员减少\" as float)))  as \"当年科技活动人员减少\" FROM [" + dataCompile.TableName + "]  group by 全国汇编用基地小类) y on x.rowName=y.rowName )c where  b.rowName=c.rowName ";
            return true;
        }
    }
}
