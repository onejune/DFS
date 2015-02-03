namespace DFSystem
{
    partial class SelectColumnsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnReturn = new System.Windows.Forms.Button();
            this.btnEscape = new System.Windows.Forms.Button();
            this.columnListBox = new System.Windows.Forms.CheckedListBox();
            this.groupColumn = new System.Windows.Forms.CheckedListBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.chkListDeviceType = new System.Windows.Forms.CheckedListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cb3 = new System.Windows.Forms.CheckBox();
            this.cb2 = new System.Windows.Forms.CheckBox();
            this.cb1 = new System.Windows.Forms.CheckBox();
            this.cmb3 = new System.Windows.Forms.ComboBox();
            this.cmb2 = new System.Windows.Forms.ComboBox();
            this.cmb1 = new System.Windows.Forms.ComboBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.cb211 = new System.Windows.Forms.CheckBox();
            this.cbOrgName = new System.Windows.Forms.CheckBox();
            this.cb985 = new System.Windows.Forms.CheckBox();
            this.cbGroupAgain = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(18, 30);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "下一步";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnReturn
            // 
            this.btnReturn.Location = new System.Drawing.Point(18, 78);
            this.btnReturn.Name = "btnReturn";
            this.btnReturn.Size = new System.Drawing.Size(75, 23);
            this.btnReturn.TabIndex = 1;
            this.btnReturn.Text = "上一步";
            this.btnReturn.UseVisualStyleBackColor = true;
            this.btnReturn.Click += new System.EventHandler(this.btnReturn_Click);
            // 
            // btnEscape
            // 
            this.btnEscape.Location = new System.Drawing.Point(18, 127);
            this.btnEscape.Name = "btnEscape";
            this.btnEscape.Size = new System.Drawing.Size(75, 23);
            this.btnEscape.TabIndex = 2;
            this.btnEscape.Text = "退出";
            this.btnEscape.UseVisualStyleBackColor = true;
            this.btnEscape.Click += new System.EventHandler(this.btnEscape_Click);
            // 
            // columnListBox
            // 
            this.columnListBox.CheckOnClick = true;
            this.columnListBox.FormattingEnabled = true;
            this.columnListBox.Location = new System.Drawing.Point(8, 23);
            this.columnListBox.Name = "columnListBox";
            this.columnListBox.Size = new System.Drawing.Size(153, 388);
            this.columnListBox.TabIndex = 4;
            this.columnListBox.ThreeDCheckBoxes = true;
            this.toolTip1.SetToolTip(this.columnListBox, "当前显示的内容为上一步选择的excel数据表中的所有列名，\r\n通过勾选决定您需要在最终的word文件中显示哪些列！最多\r\n只能选择7列！“仪器分类大类”和“仪器分" +
                    "类中类”这两列\r\n将合并为一列！");
            // 
            // groupColumn
            // 
            this.groupColumn.FormattingEnabled = true;
            this.groupColumn.Items.AddRange(new object[] {
            "区域",
            "仪器类型",
            "隶属关系",
            "基地类型"});
            this.groupColumn.Location = new System.Drawing.Point(7, 23);
            this.groupColumn.Name = "groupColumn";
            this.groupColumn.Size = new System.Drawing.Size(155, 116);
            this.groupColumn.TabIndex = 5;
            this.toolTip1.SetToolTip(this.groupColumn, "请选择分类依据：\r\n1.按区域分，将会按\"所在省\"进行分类汇总！\r\n2.按仪器类型分，将会按照\"仪器分类大类\"进行分类汇总！\r\n");
            this.groupColumn.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.groupColumn_ItemCheck);
            this.groupColumn.SelectedIndexChanged += new System.EventHandler(this.groupColumn_SelectedIndexChanged);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(6, 48);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(138, 21);
            this.textBox1.TabIndex = 9;
            this.toolTip1.SetToolTip(this.textBox1, "设置按“原值”筛选的数值范围，两者都不输入意为选择全部数据！\r\n例如：\r\nfrom=10000，to=100000，表示选择原值为10000-100000的数据！" +
                    "\r\nfrom=空，to=500000，表示选择原值小于等于500000的数据！\r\nfrom=500000，to=空，表示选择大于等于500000的数据！\r\n\r\n" +
                    "");
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(8, 115);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(136, 21);
            this.textBox2.TabIndex = 10;
            this.toolTip1.SetToolTip(this.textBox2, "设置按“原值”筛选的数值范围，两者都不输入意为选择全部数据！\r\n例如：\r\nfrom=10000，to=100000，表示选择原值为10000-100000的数据！" +
                    "\r\nfrom=空，to=500000，表示选择原值小于等于500000的数据！\r\nfrom=500000，to=空，表示选择大于等于500000的数据！\r\n\r\n" +
                    "");
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Location = new System.Drawing.Point(384, 32);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(172, 170);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "原值筛选";
            this.toolTip1.SetToolTip(this.groupBox1, "设置按“原值”筛选的数值范围，两者都不输入意为选择全部数据！\r\n例如：\r\nfrom=10000，to=100000，表示选择原值为10000-100000的数据！" +
                    "\r\nfrom=空，to=500000，表示选择原值小于等于500000的数据！\r\nfrom=500000，to=空，表示选择大于等于500000的数据！\r\n");
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(23, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "To:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "From:";
            // 
            // toolTip1
            // 
            this.toolTip1.AutoPopDelay = 8000;
            this.toolTip1.InitialDelay = 500;
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ReshowDelay = 100;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // chkListDeviceType
            // 
            this.chkListDeviceType.CheckOnClick = true;
            this.chkListDeviceType.FormattingEnabled = true;
            this.chkListDeviceType.Items.AddRange(new object[] {
            "大气探测仪器",
            "地球探测仪器",
            "电子测量仪器",
            "分析仪器",
            "工艺实验设备",
            "海洋仪器",
            "核仪器",
            "激光器",
            "计量仪器",
            "计算机及其配套设备",
            "特种检测仪器",
            "天文仪器",
            "物理性能测试仪器",
            "医学诊断仪器",
            "其他仪器"});
            this.chkListDeviceType.Location = new System.Drawing.Point(6, 20);
            this.chkListDeviceType.Name = "chkListDeviceType";
            this.chkListDeviceType.Size = new System.Drawing.Size(159, 196);
            this.chkListDeviceType.TabIndex = 12;
            this.chkListDeviceType.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkListDeviceType_ItemCheck);
            this.chkListDeviceType.SelectedIndexChanged += new System.EventHandler(this.chkListDeviceType_SelectedIndexChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkListDeviceType);
            this.groupBox2.Location = new System.Drawing.Point(192, 222);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(172, 231);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "仪器类型";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.groupColumn);
            this.groupBox3.Location = new System.Drawing.Point(195, 32);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(169, 170);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "分组依据";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.cb3);
            this.groupBox4.Controls.Add(this.cb2);
            this.groupBox4.Controls.Add(this.cb1);
            this.groupBox4.Controls.Add(this.cmb3);
            this.groupBox4.Controls.Add(this.cmb2);
            this.groupBox4.Controls.Add(this.cmb1);
            this.groupBox4.Location = new System.Drawing.Point(384, 221);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(172, 232);
            this.groupBox4.TabIndex = 13;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "共享模式以及年服务机时";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(43, 137);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 12);
            this.label5.TabIndex = 18;
            this.label5.Text = "或者";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(43, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 17;
            this.label2.Text = "或者";
            // 
            // cb3
            // 
            this.cb3.AutoSize = true;
            this.cb3.Location = new System.Drawing.Point(113, 165);
            this.cb3.Name = "cb3";
            this.cb3.Size = new System.Drawing.Size(54, 16);
            this.cb3.TabIndex = 16;
            this.cb3.Text = "大于0";
            this.cb3.UseVisualStyleBackColor = true;
            // 
            // cb2
            // 
            this.cb2.AutoSize = true;
            this.cb2.Location = new System.Drawing.Point(112, 101);
            this.cb2.Name = "cb2";
            this.cb2.Size = new System.Drawing.Size(54, 16);
            this.cb2.TabIndex = 16;
            this.cb2.Text = "大于0";
            this.cb2.UseVisualStyleBackColor = true;
            // 
            // cb1
            // 
            this.cb1.AutoSize = true;
            this.cb1.Location = new System.Drawing.Point(114, 38);
            this.cb1.Name = "cb1";
            this.cb1.Size = new System.Drawing.Size(54, 16);
            this.cb1.TabIndex = 16;
            this.cb1.Text = "大于0";
            this.cb1.UseVisualStyleBackColor = true;
            // 
            // cmb3
            // 
            this.cmb3.FormattingEnabled = true;
            this.cmb3.Items.AddRange(new object[] {
            "不共享",
            "外部共享",
            "内部共享"});
            this.cmb3.Location = new System.Drawing.Point(8, 163);
            this.cmb3.Name = "cmb3";
            this.cmb3.Size = new System.Drawing.Size(94, 20);
            this.cmb3.TabIndex = 2;
            // 
            // cmb2
            // 
            this.cmb2.FormattingEnabled = true;
            this.cmb2.Items.AddRange(new object[] {
            "不共享",
            "外部共享",
            "内部共享"});
            this.cmb2.Location = new System.Drawing.Point(7, 101);
            this.cmb2.Name = "cmb2";
            this.cmb2.Size = new System.Drawing.Size(94, 20);
            this.cmb2.TabIndex = 1;
            // 
            // cmb1
            // 
            this.cmb1.FormattingEnabled = true;
            this.cmb1.Items.AddRange(new object[] {
            "不共享",
            "外部共享",
            "内部共享"});
            this.cmb1.Location = new System.Drawing.Point(8, 37);
            this.cmb1.Name = "cmb1";
            this.cmb1.Size = new System.Drawing.Size(94, 20);
            this.cmb1.TabIndex = 0;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnReturn);
            this.groupBox5.Controls.Add(this.btnStart);
            this.groupBox5.Controls.Add(this.btnEscape);
            this.groupBox5.Location = new System.Drawing.Point(811, 286);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(113, 167);
            this.groupBox5.TabIndex = 16;
            this.groupBox5.TabStop = false;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.columnListBox);
            this.groupBox6.Location = new System.Drawing.Point(4, 32);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(172, 421);
            this.groupBox6.TabIndex = 17;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "列名";
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.cb211);
            this.groupBox7.Controls.Add(this.cbOrgName);
            this.groupBox7.Controls.Add(this.cb985);
            this.groupBox7.Controls.Add(this.cbGroupAgain);
            this.groupBox7.Location = new System.Drawing.Point(574, 32);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(182, 170);
            this.groupBox7.TabIndex = 18;
            this.groupBox7.TabStop = false;
            // 
            // cb211
            // 
            this.cb211.AutoSize = true;
            this.cb211.Location = new System.Drawing.Point(24, 65);
            this.cb211.Name = "cb211";
            this.cb211.Size = new System.Drawing.Size(102, 16);
            this.cb211.TabIndex = 19;
            this.cb211.Text = "只显示211学校";
            this.cb211.UseVisualStyleBackColor = true;
            this.cb211.CheckedChanged += new System.EventHandler(this.cb211_CheckedChanged);
            // 
            // cbOrgName
            // 
            this.cbOrgName.AutoSize = true;
            this.cbOrgName.Location = new System.Drawing.Point(24, 131);
            this.cbOrgName.Name = "cbOrgName";
            this.cbOrgName.Size = new System.Drawing.Size(108, 16);
            this.cbOrgName.TabIndex = 20;
            this.cbOrgName.Text = "按单位名称分组";
            this.cbOrgName.UseVisualStyleBackColor = true;
            this.cbOrgName.CheckedChanged += new System.EventHandler(this.cbOrgName_CheckedChanged);
            // 
            // cb985
            // 
            this.cb985.AutoSize = true;
            this.cb985.Location = new System.Drawing.Point(24, 97);
            this.cb985.Name = "cb985";
            this.cb985.Size = new System.Drawing.Size(102, 16);
            this.cb985.TabIndex = 20;
            this.cb985.Text = "只显示985学校";
            this.cb985.UseVisualStyleBackColor = true;
            this.cb985.CheckedChanged += new System.EventHandler(this.cb985_CheckedChanged);
            // 
            // cbGroupAgain
            // 
            this.cbGroupAgain.AutoSize = true;
            this.cbGroupAgain.Location = new System.Drawing.Point(24, 31);
            this.cbGroupAgain.Name = "cbGroupAgain";
            this.cbGroupAgain.Size = new System.Drawing.Size(132, 16);
            this.cbGroupAgain.TabIndex = 0;
            this.cbGroupAgain.Text = "对分析仪器再次分组";
            this.cbGroupAgain.UseVisualStyleBackColor = true;
            this.cbGroupAgain.CheckedChanged += new System.EventHandler(this.cbGroupAgain_CheckedChanged);
            // 
            // SelectColumnsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(949, 483);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "SelectColumnsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "数据筛选";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SelectColumnsForm_FormClosing);
            this.Load += new System.EventHandler(this.SelectColumnsForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnReturn;
        private System.Windows.Forms.Button btnEscape;
        private System.Windows.Forms.CheckedListBox columnListBox;
        private System.Windows.Forms.CheckedListBox groupColumn;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckedListBox chkListDeviceType;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox cb3;
        private System.Windows.Forms.CheckBox cb2;
        private System.Windows.Forms.CheckBox cb1;
        private System.Windows.Forms.ComboBox cmb3;
        private System.Windows.Forms.ComboBox cmb2;
        private System.Windows.Forms.ComboBox cmb1;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.CheckBox cbGroupAgain;
        private System.Windows.Forms.CheckBox cb211;
        private System.Windows.Forms.CheckBox cb985;
        private System.Windows.Forms.CheckBox cbOrgName;
    }
}