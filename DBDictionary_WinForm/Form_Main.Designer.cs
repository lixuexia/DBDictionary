namespace DBDictionary_WinForm
{
    partial class Form_Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.BOX_ConnectionString = new System.Windows.Forms.TextBox();
            this.BTN_GetDbInfo = new System.Windows.Forms.Button();
            this.BTN_Generate = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.CLB_Tab = new System.Windows.Forms.CheckedListBox();
            this.PB_Generate = new System.Windows.Forms.ProgressBar();
            this.GV_TabInfo = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BTN_View = new System.Windows.Forms.Button();
            this.BTN_SelAll = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.GV_TabInfo)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "链接字符串";
            // 
            // BOX_ConnectionString
            // 
            this.BOX_ConnectionString.Location = new System.Drawing.Point(83, 6);
            this.BOX_ConnectionString.Name = "BOX_ConnectionString";
            this.BOX_ConnectionString.Size = new System.Drawing.Size(751, 21);
            this.BOX_ConnectionString.TabIndex = 1;
            // 
            // BTN_GetDbInfo
            // 
            this.BTN_GetDbInfo.Location = new System.Drawing.Point(840, 4);
            this.BTN_GetDbInfo.Name = "BTN_GetDbInfo";
            this.BTN_GetDbInfo.Size = new System.Drawing.Size(75, 23);
            this.BTN_GetDbInfo.TabIndex = 2;
            this.BTN_GetDbInfo.Text = "连接";
            this.BTN_GetDbInfo.UseVisualStyleBackColor = true;
            this.BTN_GetDbInfo.Click += new System.EventHandler(this.BTN_GetDbInfo_Click);
            // 
            // BTN_Generate
            // 
            this.BTN_Generate.Location = new System.Drawing.Point(921, 4);
            this.BTN_Generate.Name = "BTN_Generate";
            this.BTN_Generate.Size = new System.Drawing.Size(75, 23);
            this.BTN_Generate.TabIndex = 3;
            this.BTN_Generate.Text = "生成";
            this.BTN_Generate.UseVisualStyleBackColor = true;
            this.BTN_Generate.Click += new System.EventHandler(this.BTN_Generate_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "保 存 目 录";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(83, 44);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(913, 21);
            this.textBox1.TabIndex = 5;
            // 
            // CLB_Tab
            // 
            this.CLB_Tab.CheckOnClick = true;
            this.CLB_Tab.FormattingEnabled = true;
            this.CLB_Tab.Location = new System.Drawing.Point(14, 103);
            this.CLB_Tab.Name = "CLB_Tab";
            this.CLB_Tab.Size = new System.Drawing.Size(183, 580);
            this.CLB_Tab.TabIndex = 6;
            // 
            // PB_Generate
            // 
            this.PB_Generate.Location = new System.Drawing.Point(203, 71);
            this.PB_Generate.Name = "PB_Generate";
            this.PB_Generate.Size = new System.Drawing.Size(793, 23);
            this.PB_Generate.TabIndex = 7;
            // 
            // GV_TabInfo
            // 
            this.GV_TabInfo.AllowUserToAddRows = false;
            this.GV_TabInfo.AllowUserToDeleteRows = false;
            this.GV_TabInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.GV_TabInfo.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column7,
            this.Column8});
            this.GV_TabInfo.Location = new System.Drawing.Point(203, 129);
            this.GV_TabInfo.Name = "GV_TabInfo";
            this.GV_TabInfo.ReadOnly = true;
            this.GV_TabInfo.RowHeadersVisible = false;
            this.GV_TabInfo.RowTemplate.Height = 23;
            this.GV_TabInfo.Size = new System.Drawing.Size(793, 550);
            this.GV_TabInfo.TabIndex = 8;
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "ColId";
            this.Column1.FillWeight = 75F;
            this.Column1.Frozen = true;
            this.Column1.HeaderText = "序号";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 75;
            // 
            // Column2
            // 
            this.Column2.DataPropertyName = "ColName";
            this.Column2.Frozen = true;
            this.Column2.HeaderText = "名称";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            // 
            // Column3
            // 
            this.Column3.DataPropertyName = "ColType";
            this.Column3.Frozen = true;
            this.Column3.HeaderText = "类型";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            // 
            // Column4
            // 
            this.Column4.DataPropertyName = "ColLength";
            this.Column4.FillWeight = 75F;
            this.Column4.Frozen = true;
            this.Column4.HeaderText = "长度";
            this.Column4.Name = "Column4";
            this.Column4.ReadOnly = true;
            this.Column4.Width = 75;
            // 
            // Column5
            // 
            this.Column5.DataPropertyName = "ColNull";
            this.Column5.Frozen = true;
            this.Column5.HeaderText = "能否为空";
            this.Column5.Name = "Column5";
            this.Column5.ReadOnly = true;
            // 
            // Column6
            // 
            this.Column6.DataPropertyName = "ColPrimaryKey";
            this.Column6.Frozen = true;
            this.Column6.HeaderText = "是否主键";
            this.Column6.Name = "Column6";
            this.Column6.ReadOnly = true;
            // 
            // Column7
            // 
            this.Column7.DataPropertyName = "ColDefaultVal";
            this.Column7.HeaderText = "默认值";
            this.Column7.Name = "Column7";
            this.Column7.ReadOnly = true;
            // 
            // Column8
            // 
            this.Column8.DataPropertyName = "ColDesc";
            this.Column8.HeaderText = "描述";
            this.Column8.Name = "Column8";
            this.Column8.ReadOnly = true;
            // 
            // BTN_View
            // 
            this.BTN_View.Location = new System.Drawing.Point(203, 100);
            this.BTN_View.Name = "BTN_View";
            this.BTN_View.Size = new System.Drawing.Size(75, 23);
            this.BTN_View.TabIndex = 9;
            this.BTN_View.Text = "查看表结构";
            this.BTN_View.UseVisualStyleBackColor = true;
            this.BTN_View.Click += new System.EventHandler(this.BTN_View_Click);
            // 
            // BTN_SelAll
            // 
            this.BTN_SelAll.Location = new System.Drawing.Point(14, 74);
            this.BTN_SelAll.Name = "BTN_SelAll";
            this.BTN_SelAll.Size = new System.Drawing.Size(75, 23);
            this.BTN_SelAll.TabIndex = 10;
            this.BTN_SelAll.Text = "全选";
            this.BTN_SelAll.UseVisualStyleBackColor = true;
            this.BTN_SelAll.Click += new System.EventHandler(this.BTN_SelAll_Click);
            // 
            // Form_Main
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1008, 691);
            this.Controls.Add(this.BTN_SelAll);
            this.Controls.Add(this.BTN_View);
            this.Controls.Add(this.GV_TabInfo);
            this.Controls.Add(this.PB_Generate);
            this.Controls.Add(this.CLB_Tab);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BTN_Generate);
            this.Controls.Add(this.BTN_GetDbInfo);
            this.Controls.Add(this.BOX_ConnectionString);
            this.Controls.Add(this.label1);
            this.MaximumSize = new System.Drawing.Size(1024, 730);
            this.MinimumSize = new System.Drawing.Size(1024, 730);
            this.Name = "Form_Main";
            this.Text = "数据字典生成";
            ((System.ComponentModel.ISupportInitialize)(this.GV_TabInfo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox BOX_ConnectionString;
        private System.Windows.Forms.Button BTN_GetDbInfo;
        private System.Windows.Forms.Button BTN_Generate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckedListBox CLB_Tab;
        private System.Windows.Forms.ProgressBar PB_Generate;
        private System.Windows.Forms.DataGridView GV_TabInfo;
        private System.Windows.Forms.Button BTN_View;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.Button BTN_SelAll;
    }
}

