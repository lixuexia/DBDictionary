using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data;

namespace DBDictionary_IDWORD
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("处理开始");
            string FileName = @"C:\数据字典(" + DateTime.Now.ToString("yyyy-MM-dd") + ").docx";
            CreateWordFile(FileName);

            Console.WriteLine("处理完成，文件位置：" + FileName);
            Console.WriteLine("输入任意字符关闭...");
            Console.ReadLine();
        }

        public static void CreateWordFile(string FileName)
        {
            Missing mis = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.ApplicationClass WordApp = new ApplicationClass();
            Document WordDoc = WordApp.Documents.Add(mis, mis, mis, mis);

            DataTable TabDT = DBStructure.GetTables();
            if (TabDT != null && TabDT.Rows.Count > 0)
            {
                Console.WriteLine("共计： " + TabDT.Rows.Count.ToString() + " 个表");
                Console.WriteLine("序号\t\t表名\t\t表描述");
                int j = 1;
                foreach (DataRow dr in TabDT.Rows)
                {
                    string TabName = dr["name"].ToString().Trim();
                    string TabDescription = (dr["desctxt"] == DBNull.Value ? "" : dr["desctxt"].ToString().Trim());
                    DataTable ColsDT = DBStructure.GetTableInfo(TabName);
                    if (ColsDT != null && ColsDT.Rows.Count > 0)
                    {
                        //移动焦点并换行
                        object count = 14;
                        object WdLine = WdUnits.wdLine;//换一行;
                        WordApp.Selection.MoveDown(WdLine, count, mis);//移动焦点
                        WordApp.Selection.TypeParagraph();//插入段落
                        Table NewTab = WordDoc.Tables.Add(WordApp.Selection.Range, ColsDT.Rows.Count + 2, 7, mis, mis);

                        NewTab.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        WordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        WordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        WordApp.Selection.Cells.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        WordApp.Selection.Cells.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;

                        NewTab.Rows.AllowBreakAcrossPages = 0;
                        NewTab.Rows.First.HeadingFormat = -1;
                        NewTab.Rows.WrapAroundText = 0;

                        NewTab.Cell(1, 1).Merge(NewTab.Cell(1, 7));
                        NewTab.Cell(1, 1).Range.Text = TabName + "：" + TabDescription;
                        NewTab.Cell(1, 1).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(1, 1).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;

                        NewTab.Cell(2, 1).Range.Text = "列名";
                        NewTab.Cell(2, 1).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 1).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        NewTab.Cell(2, 2).Range.Text = "类型";
                        NewTab.Cell(2, 2).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 2).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        NewTab.Cell(2, 3).Range.Text = "长度";
                        NewTab.Cell(2, 3).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 3).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        NewTab.Cell(2, 4).Range.Text = "能否为空";
                        NewTab.Cell(2, 4).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 4).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        NewTab.Cell(2, 5).Range.Text = "是否自增";
                        NewTab.Cell(2, 5).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 5).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        NewTab.Cell(2, 6).Range.Text = "是否主键";
                        NewTab.Cell(2, 6).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 6).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                        NewTab.Cell(2, 7).Range.Text = "描述";
                        NewTab.Cell(2, 7).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        NewTab.Cell(2, 7).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;

                        int i = 3;
                        foreach (DataRow subdr in ColsDT.Rows)
                        {
                            NewTab.Cell(i, 1).Range.Text = subdr["列名"].ToString();
                            NewTab.Cell(i, 1).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 1).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            NewTab.Cell(i, 2).Range.Text = subdr["类型"].ToString();
                            NewTab.Cell(i, 2).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 2).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            NewTab.Cell(i, 3).Range.Text = subdr["长度"].ToString();
                            NewTab.Cell(i, 3).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 3).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            NewTab.Cell(i, 4).Range.Text = subdr["能否为空"].ToString();
                            NewTab.Cell(i, 4).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 4).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            NewTab.Cell(i, 5).Range.Text = subdr["是否自增"].ToString();
                            NewTab.Cell(i, 5).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 5).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            NewTab.Cell(i, 6).Range.Text = subdr["是否主键"].ToString();
                            NewTab.Cell(i, 6).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 6).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            NewTab.Cell(i, 7).Range.Text = subdr["描述"].ToString();
                            NewTab.Cell(i, 7).Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            NewTab.Cell(i, 7).Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                            i++;
                        }
                        Console.WriteLine(string.Format("{0}\t\t{1}\t\t{2}", j, TabName, TabDescription));
                        j++;
                        Object objUnit = WdUnits.wdStory;
                        WordApp.Selection.EndKey(ref objUnit); 
                    }
                }
            }
            WordDoc.SaveAs(FileName, mis, mis, mis, mis, mis, mis, mis, mis, mis, mis, mis, mis, mis, mis, mis);
            WordDoc.Close(mis, mis, mis);
            WordApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(WordApp);
            GC.Collect();
        }
    }
}