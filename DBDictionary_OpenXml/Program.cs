using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DBDictionary_OpenXml
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("处理开始");
            string FileName = @"C:\N数据字典(" + DateTime.Now.ToString("yyyy-MM-dd") + ").docx";
            CreateWordFile(FileName);

            Console.WriteLine("处理完成，文件位置：" + FileName);
            Console.WriteLine("输入任意字符关闭...");
            Console.ReadLine();
        }

        public static void CreateWordFile(string FileName)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(FileName, WordprocessingDocumentType.Document))
            {
                //文档主对象
                MainDocumentPart mdp = doc.AddMainDocumentPart();
                mdp.Document = new Document();
                Body body = mdp.Document.AppendChild(new Body());
                //获取表数据
                DataTable TabDT = DBStructure.GetTables();
                if (TabDT != null && TabDT.Rows.Count > 0)
                {
                    Console.WriteLine("共计： " + TabDT.Rows.Count.ToString() + " 个表");
                    Console.WriteLine("序号\t表名\t\t\t\t表描述");
                    int j = 0;
                    foreach (DataRow dr in TabDT.Rows)
                    {
                        j++;
                        //逐个创建表
                        string TabName = dr["name"].ToString().Trim();
                        string TabDescription = (dr["desctxt"] == DBNull.Value ? "" : dr["desctxt"].ToString().Trim());
                        DataTable ColsDT = DBStructure.GetTableInfo(TabName);

                        if (ColsDT != null && ColsDT.Rows.Count > 0)
                        {
                            #region 插入空段落
                            Paragraph p = mdp.Document.Body.AppendChild(new Paragraph() { RsidParagraphAddition = "007557D9", RsidRunAdditionDefault = "007557D9" });
                            p.AppendChild(new Run(new Text("")));
                            #endregion

                            #region 添加表和表头
                            Table table1 = new Table();

                            TableProperties tableProperties = new TableProperties();
                            TableWidth tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

                            TableBorders tableBorders = new TableBorders();
                            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

                            tableBorders.Append(topBorder);
                            tableBorders.Append(leftBorder);
                            tableBorders.Append(bottomBorder);
                            tableBorders.Append(rightBorder);

                            TableLayout tableLayout = new TableLayout() { Type = TableLayoutValues.Fixed };
                            TableLook tableLook = new TableLook() { Val = "0000" };

                            tableProperties.Append(tableWidth);
                            tableProperties.Append(tableBorders);
                            tableProperties.Append(tableLayout);
                            tableProperties.Append(tableLook);

                            TableGrid tableGrid = new TableGrid();
                            int[] ColWidths = new[] { 1800, 1000, 700, 700, 700, 700, 3100 };
                            for (int i = 0; i < 7; i++)
                            {
                                GridColumn gCol = new GridColumn() { Width = ColWidths[i].ToString() };
                                tableGrid.Append(gCol);
                            }

                            TableRow nameRow = new TableRow() { RsidTableRowAddition = "007557D9", RsidTableRowProperties = "00096DED" };

                            TablePropertyExceptions tablePropertyExceptions_name = new TablePropertyExceptions();
                            TableCellMarginDefault tableCellMarginDefault_name = new TableCellMarginDefault();
                            TopMargin topMargin_name = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
                            BottomMargin bottomMargin_name = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

                            tableCellMarginDefault_name.Append(topMargin_name);
                            tableCellMarginDefault_name.Append(bottomMargin_name);

                            tablePropertyExceptions_name.Append(tableCellMarginDefault_name);

                            TableRowProperties tableRowProperties_name = new TableRowProperties();
                            CantSplit cantSplit_name = new CantSplit();
                            TableHeader tableHeader_name = new TableHeader();

                            tableRowProperties_name.Append(cantSplit_name);
                            tableRowProperties_name.Append(tableHeader_name);

                            TableCell tableCell_name = new TableCell();

                            TableCellProperties tableCellProperties_name = new TableCellProperties();
                            TableCellWidth tableCellWidth_name = new TableCellWidth() { Width = "8500", Type = TableWidthUnitValues.Dxa };
                            GridSpan gridSpan_name = new GridSpan() { Val = 7 };

                            TableCellBorders tableCellBorders_name = new TableCellBorders();
                            TopBorder topBorder_name = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                            BottomBorder bottomBorder_name = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                            RightBorder rightBorder_name = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                            LeftBorder leftBorder_name = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

                            tableCellBorders_name.Append(topBorder_name);
                            tableCellBorders_name.Append(rightBorder_name);
                            tableCellBorders_name.Append(bottomBorder_name);
                            tableCellBorders_name.Append(leftBorder_name);
                            Shading shading_name = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                            TableCellVerticalAlignment tableCellVerticalAlignment_name = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                            tableCellProperties_name.Append(tableCellWidth_name);
                            tableCellProperties_name.Append(gridSpan_name);
                            tableCellProperties_name.Append(tableCellBorders_name);
                            tableCellProperties_name.Append(shading_name);
                            tableCellProperties_name.Append(tableCellVerticalAlignment_name);

                            Paragraph paragraph_name = new Paragraph() { RsidParagraphAddition = "007557D9", RsidRunAdditionDefault = "007557D9" };

                            Run run_name = new Run();

                            RunProperties runProperties_name = new RunProperties();
                            RunFonts runFonts_name = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

                            runProperties_name.Append(runFonts_name);
                            Text text_name = new Text();
                            text_name.Text = TabName + "：" + TabDescription;

                            run_name.Append(runProperties_name);
                            run_name.Append(text_name);

                            paragraph_name.Append(run_name);

                            tableCell_name.Append(tableCellProperties_name);
                            tableCell_name.Append(paragraph_name);

                            nameRow.Append(tablePropertyExceptions_name);
                            nameRow.Append(tableRowProperties_name);
                            nameRow.Append(tableCell_name);
                            table1.AppendChild(nameRow);
                            #endregion

                            //表头定义
                            TableRow headerRow = new TableRow() { RsidTableRowAddition = "007557D9", RsidTableRowProperties = "007557D9" };
                            TablePropertyExceptions tablePropertyExceptions_header = new TablePropertyExceptions();
                            TableCellMarginDefault tableCellMarginDefault_header = new TableCellMarginDefault();
                            TopMargin topMargin_header = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
                            BottomMargin bottomMargin_header = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
                            tableCellMarginDefault_header.Append(topMargin_header);
                            tableCellMarginDefault_header.Append(bottomMargin_header);
                            tablePropertyExceptions_header.Append(tableCellMarginDefault_header);
                            TableRowProperties tableRowProperties_header = new TableRowProperties();
                            CantSplit cantSplit_header = new CantSplit();
                            tableRowProperties_header.Append(cantSplit_header);

                            string[] HeaderArray = new string[] { "列名", "类型", "长度", "可空", "自增", "主键", "描述" };
                            int HeadIndex = 0;
                            foreach (string headstr in HeaderArray)
                            {
                                TableCell titleCell = new TableCell();

                                TableCellProperties tableCellProperties_header = new TableCellProperties();
                                TableCellWidth tableCellWidth_header = new TableCellWidth() { Width = ColWidths[HeadIndex].ToString(), Type = TableWidthUnitValues.Dxa };

                                TableCellBorders tableCellBorders_header = new TableCellBorders();
                                TopBorder topBorder_header = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                                BottomBorder bottomBorder_header = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                                RightBorder rightBorder_header = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                                LeftBorder leftBorder_header = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

                                tableCellBorders_header.Append(topBorder_header);
                                tableCellBorders_header.Append(bottomBorder_header);
                                tableCellBorders_header.Append(rightBorder_header);
                                tableCellBorders_header.Append(leftBorder_header);
                                Shading shading_header = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                                tableCellProperties_header.Append(tableCellWidth_header);
                                tableCellProperties_header.Append(tableCellBorders_header);
                                tableCellProperties_header.Append(shading_header);

                                Paragraph paragraph_header = new Paragraph() { RsidParagraphAddition = "007557D9", RsidRunAdditionDefault = "007557D9" };

                                Run run_header = new Run();

                                RunProperties runProperties_header = new RunProperties();
                                RunFonts runFonts_header = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

                                runProperties_header.Append(runFonts_header);
                                Text text_header = new Text();
                                text_header.Text = headstr;

                                run_header.Append(runProperties_header);
                                run_header.Append(text_header);

                                paragraph_header.Append(run_header);

                                titleCell.Append(tableCellProperties_header);
                                titleCell.Append(paragraph_header);


                                headerRow.Append(titleCell);
                                HeadIndex++;
                            }
                            headerRow.Append(tablePropertyExceptions_header);
                            headerRow.Append(tableRowProperties_header);
                            table1.AppendChild(headerRow);

                            foreach (DataRow subdr in ColsDT.Rows)
                            {
                                TableRow dataRow = new TableRow();
                                TablePropertyExceptions tablePropertyExceptions_data = new TablePropertyExceptions();
                                TableCellMarginDefault tableCellMarginDefault_data = new TableCellMarginDefault();
                                TopMargin topMargin_data = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
                                BottomMargin bottomMargin_data = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
                                tableCellMarginDefault_data.Append(topMargin_data);
                                tableCellMarginDefault_data.Append(bottomMargin_data);
                                tablePropertyExceptions_data.Append(tableCellMarginDefault_data);
                                TableRowProperties tableRowProperties_data = new TableRowProperties();
                                CantSplit cantSplit_data = new CantSplit();
                                tableRowProperties_data.Append(cantSplit_data);

                                int DataIndex = 0;
                                foreach (string header in HeaderArray)
                                {
                                    TableCell dataCell = new TableCell();

                                    TableCellProperties tableCellProperties_data = new TableCellProperties();
                                    TableCellWidth tableCellWidth_data = new TableCellWidth() { Width = ColWidths[DataIndex].ToString(), Type = TableWidthUnitValues.Dxa };

                                    TableCellBorders tableCellBorders_data = new TableCellBorders();
                                    TopBorder topBorder_data = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                                    BottomBorder bottomBorder_data = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                                    RightBorder rightBorder_data = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
                                    LeftBorder leftBorder_data = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

                                    tableCellBorders_data.Append(topBorder_data);
                                    tableCellBorders_data.Append(bottomBorder_data);
                                    tableCellBorders_data.Append(rightBorder_data);
                                    tableCellBorders_data.Append(leftBorder_data);
                                    Shading shading_data = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                                    tableCellProperties_data.Append(tableCellWidth_data);
                                    tableCellProperties_data.Append(tableCellBorders_data);
                                    tableCellProperties_data.Append(shading_data);

                                    Paragraph paragraph_data = new Paragraph() { RsidParagraphAddition = "007557D9", RsidRunAdditionDefault = "007557D9" };

                                    Run run_data = new Run();

                                    RunProperties runProperties_data = new RunProperties();
                                    RunFonts runFonts_data = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

                                    runProperties_data.Append(runFonts_data);
                                    Text text_data = new Text();
                                    text_data.Text = subdr[header].ToString();

                                    run_data.Append(runProperties_data);
                                    run_data.Append(text_data);

                                    paragraph_data.Append(run_data);

                                    dataCell.Append(tableCellProperties_data);
                                    dataCell.Append(paragraph_data);

                                    dataRow.Append(dataCell);
                                    DataIndex++;
                                }
                                dataRow.Append(tablePropertyExceptions_data);
                                dataRow.Append(tableRowProperties_data);

                                table1.AppendChild(dataRow);
                            }
                            Console.WriteLine(string.Format("{0}\t{1}\t\t\t\t{2}", j, TabName, TabDescription));
                            body.Append(table1);
                            Paragraph p1 = new Paragraph(new Run(new Text("")));
                            body.Append(p1);
                        }
                    }
                }
            }
        }
    }
}