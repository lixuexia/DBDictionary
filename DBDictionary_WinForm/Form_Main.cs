using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DBDictionary_WinForm
{
    public partial class Form_Main : Form
    {
        private DataTable VDT = new DataTable();
        private string DefaultOutPath = System.Configuration.ConfigurationManager.AppSettings["DefaultOutPath"] ?? "D:/DBDictionary";
        public Form_Main()
        {
            InitializeComponent();
            VDT.Columns.Add("ColId");
            VDT.Columns.Add("ColName");
            VDT.Columns.Add("ColLength");
            VDT.Columns.Add("ColPrimaryKey");
            VDT.Columns.Add("ColNull");
            VDT.Columns.Add("ColType");
            VDT.Columns.Add("ColDefaultVal");
            VDT.Columns.Add("ColDesc");

            this.GV_TabInfo.DataSource = VDT;

            if (ConfigurationManager.ConnectionStrings["DBConn"] != null)
            {
                this.BOX_ConnectionString.Text = ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString;
            }
        }

        private void BTN_GetDbInfo_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(BOX_ConnectionString.Text))
            {
                MessageBox.Show("请填写链接字符串");
                return;
            }
            DataTable TabDT = GetTables();
            if (TabDT != null && TabDT.Rows.Count > 0)
            {
                DBStructure.DBConn = this.BOX_ConnectionString.Text;
                this.CLB_Tab.Items.Clear();
                foreach (DataRow dr in TabDT.Rows)
                {
                    this.CLB_Tab.Items.Add(dr[0].ToString(), false);
                }
            }
        }

        private DataTable GetTables()
        {
            string sql_GetStruct =
@"select t.name from sysobjects t left outer join sys.extended_properties d 
on t.id=d.major_id and d.minor_id=0
where t.xtype='U' order by t.name asc";

            SqlConnection STRUCTConn = new SqlConnection(this.BOX_ConnectionString.Text);
            SqlCommand STRUCTCmd = new SqlCommand(sql_GetStruct, STRUCTConn);
            SqlDataAdapter STRUCTAdp = new SqlDataAdapter(STRUCTCmd);
            DataTable SDT = new DataTable();
            STRUCTAdp.Fill(SDT);
            STRUCTConn.Close();
            STRUCTConn.Dispose();
            return SDT;
        }

        private void BTN_Generate_Click(object sender, EventArgs e)
        {
            string FileName = DefaultOutPath + @"/N数据字典(" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ").docx";
            this.textBox1.Text = FileName;
            this.CurrCount = 0;
            this.TabCount = 0;
            List<string> TabNames = new List<string>();
            if (this.CLB_Tab.CheckedItems != null && this.CLB_Tab.CheckedItems.Count > 0)
            {
                foreach (var item in this.CLB_Tab.CheckedItems)
                {
                    TabNames.Add(item.ToString());
                }
                TabCount = TabNames.Count;
            }
            CreateWordFile(FileName, TabNames);
        }

        private void BTN_View_Click(object sender, EventArgs e)
        {
            if (this.CLB_Tab.SelectedItems != null && this.CLB_Tab.SelectedItems.Count > 0)
            {
                VDT.Rows.Clear();
                string SelTabName = this.CLB_Tab.SelectedItems[0].ToString();
                string sql_GetStruct =
@"select distinct a.colid ColId,a.name ColName,a.length ColLength,CASE WHEN a.PK=1 THEN '是' ELSE '否' END AS ColPrimaryKey,
CASE WHEN a.isnullable=1 THEN '能' ELSE '否' END ColNull,b.value as ColDesc,c.name as ColType,CASE WHEN comm.text IS null THEN 'NULL' ELSE comm.text END as ColDefaultVal from 
(
select id,colid,name,xtype,length,colstat,autoval,isnullable,COLUMNPROPERTY(a.id,a.name,'IsIdentity') as IsIdentity,cdefault,
(SELECT count(*) FROM sysobjects WHERE (name in (SELECT name FROM sysindexes WHERE (id = a.id) AND
(indid in (SELECT indid FROM sysindexkeys WHERE (id = a.id) AND (colid in (SELECT colid FROM syscolumns WHERE (id = a.id) AND (name = a.name))))))) AND (xtype = 'PK')) as PK
from syscolumns as a where name<>'rowguid' and id in(select id from sysobjects where xtype='U' and name=@TabName)
) as a 
left outer join sys.extended_properties as b on (a.id=b.major_id and a.colid=b.minor_id)
left outer join systypes as c on (a.xtype=c.xtype and c.xtype=c.xusertype)
left outer join syscomments as comm on a.cdefault = comm.id 
where b.class_desc ='OBJECT_OR_COLUMN' or b.class_desc is null order by ColId ASC";

                SqlConnection STRUCTConn = new SqlConnection(this.BOX_ConnectionString.Text);
                SqlCommand STRUCTCmd = new SqlCommand(sql_GetStruct, STRUCTConn);
                STRUCTCmd.Parameters.Add("@TabName", SqlDbType.NVarChar).Value = SelTabName;
                SqlDataAdapter STRUCTAdp = new SqlDataAdapter(STRUCTCmd);
                DataTable SDT = new DataTable();
                STRUCTAdp.Fill(SDT);
                STRUCTConn.Close();
                STRUCTConn.Dispose();
                VDT.Clear();
                if (SDT != null && SDT.Rows.Count > 0)
                {
                    foreach (DataRow sdr in SDT.Rows)
                    {
                        DataRow vdr = VDT.NewRow();
                        foreach (DataColumn dc in VDT.Columns)
                        {
                            vdr[dc.ColumnName] = sdr[dc.ColumnName];
                        }
                        VDT.Rows.Add(vdr);
                    }
                }
            }
        }

        public void CreateWordFile(string FileName, List<string> Tabs)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(FileName, WordprocessingDocumentType.Document))
            {
                //文档主对象
                MainDocumentPart mdp = doc.AddMainDocumentPart();
                mdp.Document = new Document();
                Body body = mdp.Document.AppendChild(new Body());
                //获取表数据
                DataTable TabDT = DBStructure.GetTables(Tabs);
                if (TabDT != null && TabDT.Rows.Count > 0)
                {
                    //Console.WriteLine("共计： " + TabDT.Rows.Count.ToString() + " 个表");
                    //Console.WriteLine("序号\t表名\t\t\t\t表描述");
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
                            //Console.WriteLine(string.Format("{0}\t{1}\t\t\t\t{2}", j, TabName, TabDescription));
                            CurrCount++;
                            ShowProgress(CurrCount, TabCount);
                            body.Append(table1);
                            Paragraph p1 = new Paragraph(new Run(new Text("")));
                            body.Append(p1);
                        }
                    }
                }
            }
        }

        private int TabCount = 0;
        private int CurrCount = 0;

        public delegate void ShowProgressHandler(int CurrCount, int TabCount);
        public ShowProgressHandler ShowProgressMethod = null;
        public void ShowProgress(int CurrCount, int TabCount)
        {
            if (this.PB_Generate.InvokeRequired)
            {
                if (ShowProgressMethod == null)
                {
                    ShowProgressMethod = new ShowProgressHandler(SetPro);
                }
                ShowProgressMethod.Invoke(CurrCount, TabCount);
            }
            else
            {
                SetPro(CurrCount, TabCount);
            }
        }

        public void SetPro(int CurrCount, int TabCount)
        {
            this.PB_Generate.Maximum = TabCount;
            this.PB_Generate.Value = CurrCount;
            if (CurrCount == TabCount)
            {
                MessageBox.Show("处理完成");
            }
        }

        private void BTN_SelAll_Click(object sender, EventArgs e)
        {
            if (this.CLB_Tab.Items == null || this.CLB_Tab.Items.Count == 0)
            {
                return;
            }
            for (int i = 0; i < this.CLB_Tab.Items.Count; i++)
            {
                this.CLB_Tab.SetItemChecked(i, true);
            }
        }
    }
}