using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFSystem
{
    class DataCompile
    {
        //单位
        private string unitName;

        //如果指定了该属性，则使用sum求和
        private string sumColumnName;

        public string SumColumnName
        {
            get { return sumColumnName; }
            set { sumColumnName = value; }
        }

        public string UnitName
        {
            get { return unitName; }
            set { unitName = value; }
        }
        //目录中的标题
        private string caption;
        public string Caption
        {
            get { return caption; }
            set { caption = value; }
        }
        //excel数据源文件路径
        private string filePath;
        public string FilePath
        {
            get { return filePath; }
            set { filePath = value; }
        }
        //分类汇总的列名(最多两列)
        private List<string> selectedColumnListA;
        public List<string> SelectedColumnListA
        {
            get { return selectedColumnListA; }
            set { selectedColumnListA = value; }
        }
        private List<string> selectedColumnListB;
        public List<string> SelectedColumnListB
        {
            get { return selectedColumnListB; }
            set { selectedColumnListB = value; }
        }
        //汇总列名称
        private string summaryColumnNameA;
        public string SummaryColumnNameA
        {
            get { return summaryColumnNameA; }
            set { summaryColumnNameA = value; }
        }
        private string summaryColumnNameB;
        public string SummaryColumnNameB
        {
            get { return summaryColumnNameB; }
            set { summaryColumnNameB = value; }
        }
        //汇总列来源列名
        private string summaryColumnFromNameA;

        public string SummaryColumnFromNameA
        {
            get { return summaryColumnFromNameA; }
            set { summaryColumnFromNameA = value; }
        }
        private string summaryColumnFromNameB;

        public string SummaryColumnFromNameB
        {
            get { return summaryColumnFromNameB; }
            set { summaryColumnFromNameB = value; }
        }

        //选择的行
        private string selectedRowName;
        public string SelectedRowName
        {
            get { return selectedRowName; }
            set { selectedRowName = value; }
        }

        //对应的sql数据库中的表名
        private string tableName;
        public string TableName
        {
            get { return tableName; }
            set { tableName = value; }
        }
        //所有的列名
        private List<string> columnList;
        public List<string> ColumnList
        {
            get { return columnList; }
            set { columnList = value; }
        }
        //所属分类
        private string bigType;

        public string BigType
        {
            get { return bigType; }
            set { bigType = value; }
        }


        public DataCompile(string bigType, string caption, string filePath, List<string> selectedColumnListA,  string summaryNameA, string summaryFromNameA,
            List<string> selectedColumnListB, string summaryNameB, string summaryFromNameB,
            string selectedRow, string unitName, string sumColumnName)
        {
            this.bigType = bigType;
            this.caption = caption;
            this.filePath = filePath;
            this.selectedColumnListA = selectedColumnListA;
            this.tableName = CheckFileName(GetFileName(filePath));
            this.selectedRowName = selectedRow;
            this.summaryColumnNameA = summaryNameA;
            this.summaryColumnNameB = summaryNameB;
            this.selectedColumnListB = selectedColumnListB;
            this.unitName = unitName;
            this.summaryColumnFromNameA = summaryFromNameA;
            this.summaryColumnFromNameB = summaryFromNameB;
            this.sumColumnName = sumColumnName; //对哪一列求和
        }

        public DataCompile()
        {
            this.bigType = "";
            this.caption = "";
            this.filePath = "";
            this.selectedColumnListA = new List<string>();
            this.tableName = "";
            this.selectedRowName = "";
            this.summaryColumnNameA = "";
            this.summaryColumnNameB = "";
            this.selectedColumnListB = new List<string>();
            this.unitName = "";
            this.summaryColumnFromNameA = "";
            this.summaryColumnFromNameB = "";
            this.sumColumnName = ""; //对哪一列求和
        }
        public static string GetFileName(string filePath)
        {
            string fileName = "";
            if (filePath != null && filePath.Length != 0)
            {
                fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1, (filePath.LastIndexOf(".") - filePath.LastIndexOf("\\") - 1));
            }
            return fileName;
        }

        //对表名进行处理，去掉非法字符
        public static string CheckFileName(string tableName)
        {
            //对表名进行处理，首字母不能为数字，表名中只能包含字母、数字
            StringBuilder sb = new StringBuilder(tableName);
            for (int i = 0; i < sb.Length; i++)
            {
                if (!((sb[i] >= 0 && sb[i] <= '9') || (sb[i] >= 'A' && sb[i] <= 'Z') || (sb[i] >= 'a' && sb[i] <= 'z') || (sb[i] >= 0x4e00 && sb[i] <= 0x9fa5)))
                {
                    sb[i] = ' ';
                }
            }

            for (int i = 0; i < sb.Length; i++)
            {
                if (sb[i] >= '0' && sb[i] <= '9')
                {
                    sb[i] = ' ';
                }
                else
                {
                    break;
                }
            }
            sb.Replace("-", "");
            sb.Replace(" ", "");
            return sb.ToString();
        }
    }
}
