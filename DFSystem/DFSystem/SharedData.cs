using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace DFSystem
{
    //表、视图对应关系
    class TableAndView
    {
        private string tableName;

        public string TableName
        {
            get { return tableName; }
            set { tableName = value; }
        }
        private string viewName;

        public string ViewName
        {
            get { return viewName; }
            set { viewName = value; }
        }
        private bool isExisted;

        public bool IsExisted
        {
            get { return isExisted; }
            set { isExisted = value; }
        }

        private List<string> columnNames;

        public List<string> ColumnNames
        {
            get { return columnNames; }
            set { columnNames = value; }
        }

        public TableAndView(string tableName)
        {
            this.tableName = tableName;
            this.viewName = "view_" + tableName;
            this.isExisted = false;
            this.columnNames = new List<string>();
        }
    }

    //基地大类和基地小类对应关系
    class BaseType
    {
        private string bigType;

        public string BigType
        {
            get { return bigType; }
            set { bigType = value; }
        }
        private List<string> smallTypeList;

        public List<string> SmallTypeList
        {
            get { return smallTypeList; }
            set { smallTypeList = value; }
        }

        public BaseType()
        {
            bigType = "";
            smallTypeList = new List<string>();
        }
    }


    //线程共享数据
    class SharedData
    {
        //debug开关
        public static bool DEBUGGING = false;

        //1：全国 2：中央 3：地方
        public static int AREATYPE = 1;

        //word临时目录
        public static string currentWordPath = "";
        //文件路径
        public static string currentFilePath = "";

        //计时
        public static System.Diagnostics.Stopwatch oTime = new System.Diagnostics.Stopwatch();

        //设置年份
        public static string currYear = "2014";

        // 当前写入word的记录的序号
        public static object sequenceNumber = 0;

        // 总记录数
        public static int sumRows = 0;

        // 是否开始显示
        public static bool isReady = false;

        // 是否出现异常
        public static object isOK = true;

        // 最终生成的word文件名
        public static string fileName = "";

        // 是否完成
        public static bool isCompleted = false;

        // 记录失败的任务
        public static List<string> failedWorkList = new List<string>();

        // 分组个数
        public static int groupCount = 0;

        // excel文件路径
        public static string excelFile;

        //科研院所和高校概况
        public static List<string> researchInstitutionList;
        //大型科学仪器设备
        public static List<string> largeScientificInstrumentList;
        //研究实验基地
        public static List<string> experimentalResearchBaseList;
        //生物种质资源
        public static List<string> geneticResourcesList;

        //数据库中存储的对应的表名,实际上是SQL中视图名
        public static List<TableAndView> tableListInDB = new List<TableAndView>
        {
            new TableAndView("法人单位概况"),//0
            new TableAndView("法人单位仪器设备概况表"),//1
            new TableAndView("法人单位科技活动人员概况"),//2
            new TableAndView("法人单位大型科学仪器设备基本信息"),//3
            new TableAndView("研究实验基地概况"),//4
            new TableAndView("保存机构概况"),//5
            new TableAndView("保存机构科技活动人员基本信息"),//6
            new TableAndView("保存机构人才培养情况"),//7
            new TableAndView("保存机构运行经费情况"),//8
            new TableAndView("保存机构微生物种质资源"),//9
            new TableAndView("保存机构动物种质资源"),//10
            new TableAndView("保存机构植物种质资源")//11
        };

        public static Dictionary<string, List<string>> dataCompileBase = new Dictionary<string, List<string>>
        {
            {"科研院所和高校概况", researchInstitutionList},
            {"大型科学仪器设备", largeScientificInstrumentList},
            {"研究实验基地", experimentalResearchBaseList},
            {"生物种质资源", geneticResourcesList}
        };

        //已经排好序的仪器类型以及所属单位类型、研究实验基地
        public static List<string> deviceTypeSorted = new List<string>
        {
            "按仪器类型分组",
            "分析仪器",
            "电子光学仪器",
            "质谱仪器",
            "X射线仪器",
            "光谱仪器",
            "色谱仪器",
            "波谱仪器",
            "电化学仪器",
            "显微镜及图像分析仪器",
            "热分析仪器",
            "生化分离分析仪器",
            "环境与农业分析仪器",
            "样品前处理及制备仪器",
            "其他",
            "物理性能测试仪器",
            "计量仪器",
            "电子测量仪器",
            "海洋仪器",
            "地球探测仪器",
            "大气探测仪器",
            "天文仪器",
            "医学诊断仪器",
            "核仪器",
            "特种检测仪器",
            "工艺实验设备",
            "计算机及其配套设备",
            "激光器",
            "其他仪器",
            "按所属单位类型分组",
            "科研机构",
            "高等学校",
            "转制院所",
            "属于研究实验基地",
            "属于生物种质资源保存机构"
        };
        //分析仪器列表
        public static List<string> analyseDeviceTypeSorted = new List<string>
        {
            "电子光学仪器",
            "质谱仪器",
            "X射线仪器",
            "光谱仪器",
            "色谱仪器",
            "波谱仪器",
            "电化学仪器",
            "显微镜及图像分析仪器",
            "热分析仪器",
            "生化分离分析仪器",
            "环境与农业分析仪器",
            "样品前处理及制备仪器",
            "其他"
        };
        //实验基地列表
        public static List<string> experimentalBaseSorted = new List<string>
        {
            "国家重大科学工程",
            "各类重点实验室",
            "国家重点实验室",
            "国家工程实验室",
            "部属重点（开放）实验室",
            "省部共建重点实验室",
            "省属重点(开放)实验室",
            "生物安全实验室",
            "其他实验室",
            "野外台站",
            "国家级野外站",
            "部属野外站",
            "省属野外站",
            "工程(技术)研究中心",
            "国家工程技术研究中心",
            "国家工程研究中心",
            "部属工程(技术)中心",
            "省属工程(技术)中心",
            "其他工程(技术)中心",
            "分析测试中心",
            "国家级分析测试中心",
            "国家大型仪器中心",
            "部属分析测试中心",
            "省属分析测试中心",
            "其他分析测试中心",
            "研发（技术）中心",
            "部属研发(技术)中心",
            "省属研发(技术)中心",
            "其他研发(技术)中心",
            "国家级企业技术中心",
            "其他",
            "其他国家级基地",
            "其他部属基地",
            "其他省属基地",
            "其他地属基地"
        };
        public static List<string> sharedTypeForGeneticResources = new List<string>
        {
            "公益性共享",
            "公益性借用共享",
            "合作研究共享",
            "知识产权交易性共享",
            "资源纯交易性共享",
            "资源交换性共享",
            "收藏地共享",
            "行政许可性共享",
            "不共享"
        };
        //植物种质资源保存类型
        public static List<string> savedTypeForAnimalGeneticResources = new List<string>
        {
            "活体",
            "精子",
            "卵子",
            "胚胎",
            "细胞株",
            "组织器官",
            "DNA材料",
            "活体标本",
            "其他"
        };
        //动物种质资源保存类型
        public static List<string> savedTypeForPlantGeneticResources = new List<string>
        {
            "植株",
            "种子",
            "种茎",
            "块根",
            "花粉",
            "培养物",
            "DNA",
            "其他"
        };
        //保藏类型
        public static List<string> reservedTypeForGeneticResources = new List<string>
        {
            "培养物",
            "二元培养物",
            "基因",
            "其他"
        };
        public static List<string> reservedMethodForGeneticResources = new List<string>
        {
            "液氮超低温冻结",
            "-80℃冰箱冻结",	
            "真空冷冻干燥",	
            "矿物油",	
            "定期移植",	
            "其他"	
        };
        //区域顺序:用于汇编之前写入到数据库中
        public static List<string> areaListSorted = new List<string> { "北京市", "天津市", "河北省", "山西省", "内蒙古自治区", "辽宁省", "吉林省", "黑龙江省", "上海市", "江苏省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "河南省", "湖北省", "湖南省", "广东省", "广西壮族自治区", "海南省", "重庆市", "四川省", "贵州省", "云南省", "西藏自治区", "陕西省", "甘肃省", "青海省", "宁夏回族自治区", "新疆维吾尔自治区" };

        //将所在省写入数据库
        public static void WriteProvincesToDB()
        {
            string sql = "";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();

            sql = "if exists(select * from sysobjects where name='所在省') drop table 所在省;create table 所在省([provinceName] [varchar](1024) NULL)ON [PRIMARY];";
            foreach (var v in areaListSorted)
            {
                sql += "insert into 所在省 values('" + v + "');";
            }

            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandTimeout = 180;
            command.ExecuteNonQuery();
            sqlConn.Close();
        }
        //将设备类型写入数据库
        public static void WriteDeviceTypesToDB()
        {
            string sql = "";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();

            sql = "if exists(select * from sysobjects where name='设备类型') drop table 设备类型;create table 设备类型([deviceType] [varchar](1024) NULL)ON [PRIMARY];";
            foreach (var v in deviceTypeSorted)
            {
                if (!v.ToString().Contains("分组"))
                {
                    sql += "insert into 设备类型 values('" + v + "');";
                }
            }

            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandTimeout = 180;
            command.ExecuteNonQuery();
            sqlConn.Close();
        }

        //将实验基地表写入数据库，成功返回1，失败返回-1
        public static int WriteBaseTypesToDB()
        {
            int res = GetExperimentalBaseList();
            if (res == -1)
            {
                return -1;
            }
            string sql = "";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            sqlConn.Open();

            sql = "if exists(select * from sysobjects where name='实验基地') drop table 实验基地;create table 实验基地([type] [varchar](1024) NULL)ON [PRIMARY];";
            foreach (var v in experimentalBaseSorted)
            {
                sql += "insert into 实验基地 values('" + v + "');";
            }

            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandTimeout = 180;
            command.ExecuteNonQuery();
            sqlConn.Close();
            return 1;
        }

        //获取指定顺序的实验基地列表，默认的列表某些字段可能与数据表中不同，成功返回1，失败返回-1
        private static int GetExperimentalBaseList()
        {
            experimentalBaseSorted.Clear();
            List<BaseType> baseTypeList = new List<BaseType>();

            string sql = "";
            SqlConnection sqlConn = new SqlConnection(DBHelper.connString);
            //view_研究实验基地概况
            string tableName = tableListInDB[4].ViewName;
            string groupBy1 = "", groupBy2 = "";
            int i = 0;
            for (i = 0; i < SharedData.tableListInDB[4].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[4].ColumnNames[i].IndexOf("基地大类") != -1)
                {
                    groupBy1 = SharedData.tableListInDB[4].ColumnNames[i];
                    break;
                }
            }
            for (i = 0; i < SharedData.tableListInDB[4].ColumnNames.Count(); i++)
            {
                if (SharedData.tableListInDB[4].ColumnNames[i].IndexOf("基地小类") != -1)
                {
                    groupBy2 = SharedData.tableListInDB[4].ColumnNames[i];
                    break;
                }
            }
            if (groupBy1 == "" || groupBy2 == "")
            {
                Log.RecordLog("GetExperimentalBaseList: 从“" + tableName + "”中获取“基地大类”和“汇编用基地小类”失败！");
                return -1;
            }

            sql = "select [" + groupBy1 + "],[" + groupBy2 + "] from [" + tableName + "] group by [" + groupBy1 + "],[" + groupBy2 + "]";

            sqlConn.Open();
            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandTimeout = 180;
            SqlDataReader reader = command.ExecuteReader();

            string bigType = "";
            try
            {
                while (reader.Read())
                {
                    string preBigType = bigType;
                    bigType = reader[groupBy1].ToString();
                    if (!preBigType.Equals(bigType))
                    {
                        BaseType b = new BaseType();
                        b.BigType = bigType;
                        b.SmallTypeList.Add(reader[groupBy2].ToString());
                        baseTypeList.Add(b);
                    }
                    else
                    {
                        baseTypeList[baseTypeList.Count() - 1].SmallTypeList.Add(reader[groupBy2].ToString());
                    }
                }
                reader.Close();
                sqlConn.Close();
            }
            catch (System.Exception ex)
            {
                Log.RecordLog("GetExperimentalBaseList: " + "获取基地列表失败！");
            }
            if (baseTypeList.Count() < 6)
            {
                Log.RecordLog("GetExperimentalBaseList: " + "获取基地列表失败！");
                return -1;
            }
            experimentalBaseSorted.Add("国家重大科学工程");
            GetBigType(baseTypeList, "实验室");
            GetBigType(baseTypeList, "野外");
            GetBigType(baseTypeList, "工程");
            GetBigType(baseTypeList, "测试");
            GetBigType(baseTypeList, "研发");
            GetBigType(baseTypeList, "其他");

            return 1;
        }
        private static void GetBigType(List<BaseType> baseTypeList, string bigType)
        {
            int i = 0;
            for (i = 0; i < baseTypeList.Count(); i++)
            {
                BaseType type = baseTypeList[i];
                if (type.BigType.Contains(bigType))
                {
                    GetSmallList(type);
                    break;
                }
            }
            baseTypeList.RemoveAt(i);
        }
        private static void GetSmallList(BaseType type)
        {
            int k = 0;
            experimentalBaseSorted.Add(type.BigType);
            for (k = 0; k < type.SmallTypeList.Count(); k++)
            {
                string str = type.SmallTypeList[k];
                if (str.Contains("国家"))
                {
                    experimentalBaseSorted.Add(str);
                    type.SmallTypeList[k] = "";
                }
            }
            for (k = 0; k < type.SmallTypeList.Count(); k++)
            {
                string str = type.SmallTypeList[k];
                if (str.Contains("部属"))
                {
                    experimentalBaseSorted.Add(str);
                    type.SmallTypeList[k] = "";
                }
            }
            for (k = 0; k < type.SmallTypeList.Count(); k++)
            {
                string str = type.SmallTypeList[k];
                if (str.Contains("省属"))
                {
                    experimentalBaseSorted.Add(str);
                    type.SmallTypeList[k] = "";
                }
            }
            for (k = 0; k < type.SmallTypeList.Count(); k++)
            {
                string str = type.SmallTypeList[k];
                if (!str.Contains("其他") && str != "")
                {
                    experimentalBaseSorted.Add(str);
                    type.SmallTypeList[k] = "";
                }
            }
            for (k = 0; k < type.SmallTypeList.Count(); k++)
            {
                string str = type.SmallTypeList[k];
                if (str.Contains("其他"))
                {
                    experimentalBaseSorted.Add(str);
                    type.SmallTypeList[k] = "";
                }
            }
        }
    }

    //excel文件路径和数据库中固定的table表之间的映射
    public class ExcelMap
    {
        private List<string> filePathList;//excel文件路径

        public List<string> FilePathList
        {
            get 
            {
                return filePathList; 
            }
            set 
            {
                filePathList = value;
            }
        }

        private int tableListIndex;//对应tableListInDB的索引

        public int TableListIndex
        {
            get { return tableListIndex; }
            set { tableListIndex = value; }
        }

        private int listBoxIndex;//在excelFileListBox中的索引

        public int ListBoxIndex
        {
            get { return listBoxIndex; }
            set { listBoxIndex = value; }
        }

        public ExcelMap()
        {
            filePathList = new List<string>();
            tableListIndex = -1;
            listBoxIndex = -1;
        }

        public string GetTaleName()
        {
            return SharedData.tableListInDB[tableListIndex].TableName;
        }

        public string FileName()
        {
            string fileName = "";
            if (filePathList.Count() > 1)
            {
                fileName = filePathList[0].Substring(filePathList[0].LastIndexOf("\\") + 1,
                    (filePathList[0].LastIndexOf(".") - filePathList[0].LastIndexOf("\\") - 1));
                fileName = fileName.Substring(0, fileName.Length - 1);
            }
            else if (filePathList.Count() == 1)
            {
                fileName = filePathList[0].Substring(filePathList[0].LastIndexOf("\\") + 1,
                    (filePathList[0].LastIndexOf(".") - filePathList[0].LastIndexOf("\\") - 1));
            }
            return fileName;
        }
    }
}
