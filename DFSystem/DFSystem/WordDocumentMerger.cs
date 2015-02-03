using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
namespace DFSystem
{
    /// <summary>
    /// Word文档合并类
    /// </summary>
    public class wordDocumentMerger
    {
        private Application objApp = null;
        private Document objDocLast = null;
        private Document objDocBeforeLast = null;
        private int contentNum = 0;

        //仪器类型顺序
        public static List<string> deviceTypeList = new List<string>{"分析仪器","物理性能测试仪器","计量仪器","电子测量仪器","海洋仪器","地球探测仪器","大气探测仪器","天文仪器","医学诊断仪器",
            "核仪器","特种检测仪器","工艺实验设备","计算机及其配套设备","激光器","其他仪器"};

        //中央部门顺序
        static List<string> centralDepartment = new List<string>{
            "最高人民检察院",
            "国家档案局",
            "国家民族事务委员会",
            "民政部",
            "司法部",
            "人力资源和社会保障部",
            "农业部",
            "水利部",
            "住房和城乡建设部",
            "国土资源部",
            "工业和信息化部",
            "铁道部",
            "交通运输部",
            "文化部",
            "教育部（国家语言文字工作委员会）",
            "卫生部",
            "国家人口和计划生育委员会",
            "中国气象局",
            "中国民用航空局",
            "国家海洋局",
            "中国地震局",
            "国家新闻出版总署（国家版权局）",
            "国家质量监督检验检疫总局",
            "国家广播电影电视总局",
            "国家林业局",
            "国务院侨务办公室",
            "中华全国供销合作总社",
            "国家邮政局",
            "国家粮食局",
            "国家安全生产监督管理总局",
            "国家体育总局",
            "国家文物局",
            "国家烟草专卖局",
            "国家测绘局",
            "环境保护部（国家核安全局）",
            "国家中医药管理局",
            "中国科学院",
            "中国社会科学院",
            "社科院",
            "中国残疾人联合会"
        };
        //基地大类顺序
        static List<string> baseTypeList = new List<string>
        {
            "国家重大科学工程",
            "实验室",
            "工程(技术)研究中心",
            "研发(技术)中心",
            "分析测试中心",
            "野外站",
            "其他类型基地"
        };

        public wordDocumentMerger()
        {
            objApp = new Application();
            contentNum = 0;
        }
        #region 打开文件
        private void Open(string tempDoc)
        {
            object objTempDoc = tempDoc;
            object objMissing = System.Reflection.Missing.Value;

            objDocLast = objApp.Documents.Open(
                 ref objTempDoc,    //FileName
                 ref objMissing,   //ConfirmVersions
                 ref objMissing,   //ReadOnly
                 ref objMissing,   //AddToRecentFiles
                 ref objMissing,   //PasswordDocument
                 ref objMissing,   //PasswordTemplate
                 ref objMissing,   //Revert
                 ref objMissing,   //WritePasswordDocument
                 ref objMissing,   //WritePasswordTemplate
                 ref objMissing,   //Format
                 ref objMissing,   //Enconding
                 ref objMissing,   //Visible
                 ref objMissing,   //OpenAndRepair
                 ref objMissing,   //DocumentDirection
                 ref objMissing,   //NoEncodingDialog
                 ref objMissing    //XMLTransform
                 );

            objDocLast.Activate();
            objDocLast.SpellingChecked = false;
            objDocLast.ShowSpellingErrors = false;
        }
        #endregion

        #region 保存文件到输出模板
        private void SaveAs(string outDoc)
        {
            object objMissing = System.Reflection.Missing.Value;
            object objOutDoc = outDoc;
            objDocLast.SaveAs(
              ref objOutDoc,      //FileName
              ref objMissing,     //FileFormat
              ref objMissing,     //LockComments
              ref objMissing,     //PassWord     
              ref objMissing,     //AddToRecentFiles
              ref objMissing,     //WritePassword
              ref objMissing,     //ReadOnlyRecommended
              ref objMissing,     //EmbedTrueTypeFonts
              ref objMissing,     //SaveNativePictureFormat
              ref objMissing,     //SaveFormsData
              ref objMissing,     //SaveAsAOCELetter,
              ref objMissing,     //Encoding
              ref objMissing,     //InsertLineBreaks
              ref objMissing,     //AllowSubstitutions
              ref objMissing,     //LineEnding
              ref objMissing      //AddBiDiMarks
              );
        }
        #endregion

        #region 循环合并多个文件（复制合并重复的文件）
        /// <summary>
        /// 循环合并多个文件（复制合并重复的文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void CopyMerge(string tempDoc, string[] arrCopies, string outDoc)
        {
            object objMissing = Missing.Value;
            object objFalse = false;
            object objTarget = WdMergeTarget.wdMergeTargetSelected;
            object objUseFormatFrom = WdUseFormattingFrom.wdFormattingFromSelected;
            try
            {
                //打开模板文件
                Open(tempDoc);
                foreach (string strCopy in arrCopies)
                {
                    objDocLast.Merge(
                      strCopy,                //FileName    
                      ref objTarget,          //MergeTarget
                      ref objMissing,         //DetectFormatChanges
                      ref objUseFormatFrom,   //UseFormattingFrom
                      ref objMissing          //AddToRecentFiles
                      );
                    objDocBeforeLast = objDocLast;
                    objDocLast = objApp.ActiveDocument;
                    if (objDocBeforeLast != null)
                    {
                        objDocBeforeLast.Close(
                          ref objFalse,     //SaveChanges
                          ref objMissing,   //OriginalFormat
                          ref objMissing    //RouteDocument
                          );
                    }
                }
                //保存到输出文件
                SaveAs(outDoc);
                foreach (Document objDocument in objApp.Documents)
                {
                    objDocument.Close(
                      ref objFalse,     //SaveChanges
                      ref objMissing,   //OriginalFormat
                      ref objMissing    //RouteDocument
                      );
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.ToString());
                Log.RecordLog("InsertMerge:" + err.ToString());
            }
            finally
            {
                objApp.Quit(
                  ref objMissing,     //SaveChanges
                  ref objMissing,     //OriginalFormat
                  ref objMissing      //RoutDocument
                  );
                objApp = null;
            }
        }
        /// <summary>
        /// 循环合并多个文件（复制合并重复的文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void CopyMerge(string tempDoc, string strCopyFolder, string outDoc)
        {
            string[] arrFiles = Directory.GetFiles(strCopyFolder);
            CopyMerge(tempDoc, arrFiles, outDoc);
        }
        #endregion

        #region 循环合并多个文件（插入合并文件）
        /// <summary>
        /// 循环合并多个文件（插入合并文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void InsertMerge(List<string> arrCopies)
        {
            object objMissing = Missing.Value;
            object objRange = "";
            object objFalse = false;
            object confirmConversion = false;
            object link = true;
            object attachment = false;
            try
            {
                //打开模板文件
                Open(SharedData.fileName);
                objApp.Selection.Range.Paragraphs.Last.Range.Select();
                foreach (string strCopy in arrCopies)
                {
                    objApp.Selection.InsertFile(
                        strCopy,
                        ref objRange,
                        ref confirmConversion,
                        ref link,
                        ref attachment
                        );
                    Console.WriteLine("merge " + strCopy);
                }

                //插入目录
                InsertContent();

                //保存到输出文件
                objDocLast.Save();

                foreach (Document objDocument in objApp.Documents)
                {
                    objDocument.Close(
                      ref objFalse,     //SaveChanges
                      ref objMissing,   //OriginalFormat
                      ref objMissing    //RouteDocument
                      );
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.ToString());
                Log.RecordLog("InsertMerge:" + err.ToString());
            }
            finally
            {
                objApp.Quit(
                  ref objMissing,     //SaveChanges
                  ref objMissing,     //OriginalFormat
                  ref objMissing      //RoutDocument
                  );
                objApp = null;
            }
        }

        //插入目录
        private void InsertContent()
        {
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            object count = 1;
            object objMissing = Missing.Value;
            //Microsoft.Office.Interop.Word.Document wordDoc = objApp.Documents.Add(objMissing, objMissing, objMissing, objMissing);
            objDocLast.Paragraphs.First.Range.Select();

            //插入分节符
            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            objApp.Application.Selection.InsertBreak(oPageBreak);

            objDocLast.Paragraphs.First.Range.Select();

            //设置布局方向
            objApp.Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;
            objApp.Selection.PageSetup.TopMargin = objApp.CentimetersToPoints(2.54F);
            objApp.Selection.PageSetup.BottomMargin = objApp.CentimetersToPoints(2.54F);
            objApp.Selection.PageSetup.LeftMargin = objApp.CentimetersToPoints(3.18F);
            objApp.Selection.PageSetup.RightMargin = objApp.CentimetersToPoints(3.18F);

            objApp.Selection.Font.Size = 20;
            objApp.Selection.Font.Name = "微软雅黑";
            objApp.Selection.Font.Bold = 1;
            objApp.Selection.TypeText("目录");
            objApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            objApp.Selection.TypeParagraph();

            objApp.Selection.ParagraphFormat.LineSpacing = 15F;
            objApp.Selection.Font.Size = 20;
            objApp.ActiveDocument.TablesOfContents.Add(objApp.Selection.Range, objMissing, 1, 2, objMissing, objMissing, true, true, objMissing, objMissing, objMissing, objMissing);
            objApp.ActiveDocument.TablesOfContents[1].Range.Select();

            objApp.ActiveDocument.TablesOfContents[1].Update();

            count = 5;
            object extend = true;
            objApp.Selection.MoveUp(ref WdLine, ref count, ref extend);//移动焦点
            objApp.Selection.Font.Size = 12;
            objApp.Selection.Font.Name = "微软雅黑";

            objDocLast.Paragraphs.First.Range.Select();
            WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            count = SharedData.groupCount + 1;
            objApp.Selection.MoveDown(ref WdLine, ref count, ref objMissing);//移动焦点

            //获得当前的页码
            int pageNumber = (int)objApp.Selection.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
            if (pageNumber % 2 != 0)
            {
                objApp.Application.Selection.InsertBreak(oPageBreak);
            }
            objApp.ActiveDocument.TablesOfContents[1].Update();
        }

        /// <summary>
        /// 循环合并多个文件（插入合并文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void InsertMerge(string groupName)
        {
            List<string> fileList = new List<string>();
            if (!Directory.Exists(@"C:\DFS\word\"))
            {
                return;
            }
            //获取指定目录下的非隐藏word文件
            foreach (string file in Directory.GetFileSystemEntries(@"C:\DFS\word\"))
            {
                DirectoryInfo d = new DirectoryInfo(file);
                if ((d.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
                {
                    fileList.Add(d.ToString());
                    contentNum++;
                }
            }
            //指定顺序
            List<string> sortList = null;
            if (groupName.Equals("区域"))
            {
                sortList = SharedData.areaListSorted;
            }
            else if (groupName.Equals("仪器分类大类"))
            {
                sortList = deviceTypeList;
            }
            else if (groupName.Equals("隶属"))
            {
                List<string> fileListForMemberGroup = new List<string>();
                fileListForMemberGroup.Add("汇总表");
                fileListForMemberGroup.Add("中央各部门");
                foreach (string name in centralDepartment)
                {
                    fileListForMemberGroup.Add(name);
                }
                fileListForMemberGroup.Add("地方各省市");
                foreach (string name in SharedData.areaListSorted)
                {
                    fileListForMemberGroup.Add(name);
                }
                sortList = fileListForMemberGroup;
            }
            else if (groupName.Equals("基地大类-汇编用"))
            {
                sortList = baseTypeList;
            }

            List<string> tempFileList = new List<string>();
            foreach (string name in sortList)
            {
                foreach (string file in fileList)
                {
                    if (file.Contains(name))
                    {
                        tempFileList.Add(file);
                        break;
                    }
                }
            }
            InsertMerge(tempFileList);
        }
        #endregion
    }
}
