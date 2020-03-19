namespace FiTools
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.zhengfu = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.button7 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.xiaoshuwei = this.Factory.CreateRibbonEditBox();
            this.R_num = this.Factory.CreateRibbonButton();
            this.R_Set = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.cal = this.Factory.CreateRibbonButton();
            this.ntbook = this.Factory.CreateRibbonButton();
            this.About = this.Factory.CreateRibbonGroup();
            this.about0 = this.Factory.CreateRibbonButton();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.Rup_num = this.Factory.CreateRibbonButton();
            this.Rup_set = this.Factory.CreateRibbonButton();
            this.Rdown_num = this.Factory.CreateRibbonButton();
            this.Rdown_set = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.About.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.About);
            this.tab1.Label = "FiTools";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.zhengfu);
            this.group1.Label = "算术运算";
            this.group1.Name = "group1";
            // 
            // editBox1
            // 
            this.editBox1.Image = ((System.Drawing.Image)(resources.GetObject("editBox1.Image")));
            this.editBox1.Label = "换算";
            this.editBox1.Name = "editBox1";
            this.editBox1.ShowImage = true;
            this.editBox1.SuperTip = "默认10000，可修改，记得回车";
            this.editBox1.Text = "10000";
            this.editBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_TextChanged);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "乘以一万";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "除以一万";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "带公式乘";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "带公式除";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // zhengfu
            // 
            this.zhengfu.Image = ((System.Drawing.Image)(resources.GetObject("zhengfu.Image")));
            this.zhengfu.Label = "正负转换";
            this.zhengfu.Name = "zhengfu";
            this.zhengfu.ShowImage = true;
            this.zhengfu.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zhengfu_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button6);
            this.group2.Label = "异常处理";
            this.group2.Name = "group2";
            // 
            // button5
            // 
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "零->#NA";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "#NA->零";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.editBox2);
            this.group3.Items.Add(this.editBox3);
            this.group3.Items.Add(this.button7);
            this.group3.Label = "删除指定行";
            this.group3.Name = "group3";
            // 
            // editBox2
            // 
            this.editBox2.Image = ((System.Drawing.Image)(resources.GetObject("editBox2.Image")));
            this.editBox2.Label = "指定列";
            this.editBox2.Name = "editBox2";
            this.editBox2.ShowImage = true;
            this.editBox2.SuperTip = "默认B，可修改，记得回车";
            this.editBox2.Text = "B";
            this.editBox2.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox2_TextChanged);
            // 
            // editBox3
            // 
            this.editBox3.Image = ((System.Drawing.Image)(resources.GetObject("editBox3.Image")));
            this.editBox3.Label = "指定值";
            this.editBox3.Name = "editBox3";
            this.editBox3.ShowImage = true;
            this.editBox3.SuperTip = "默认空值，可修改，记得回车";
            this.editBox3.Text = null;
            this.editBox3.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox3_TextChanged);
            // 
            // button7
            // 
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "删除指定行";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.xiaoshuwei);
            this.group4.Items.Add(this.R_num);
            this.group4.Items.Add(this.R_Set);
            this.group4.Items.Add(this.Rup_num);
            this.group4.Items.Add(this.Rup_set);
            this.group4.Items.Add(this.Rdown_num);
            this.group4.Items.Add(this.Rdown_set);
            this.group4.Label = "小数位";
            this.group4.Name = "group4";
            // 
            // xiaoshuwei
            // 
            this.xiaoshuwei.Image = ((System.Drawing.Image)(resources.GetObject("xiaoshuwei.Image")));
            this.xiaoshuwei.Label = "小数位";
            this.xiaoshuwei.Name = "xiaoshuwei";
            this.xiaoshuwei.ShowImage = true;
            this.xiaoshuwei.SuperTip = "默认2，可修改，记得回车";
            this.xiaoshuwei.Text = "2";
            this.xiaoshuwei.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.xiaoshuwei_TextChanged);
            // 
            // R_num
            // 
            this.R_num.Image = ((System.Drawing.Image)(resources.GetObject("R_num.Image")));
            this.R_num.Label = "Round一下";
            this.R_num.Name = "R_num";
            this.R_num.ShowImage = true;
            this.R_num.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.R_num_Click);
            // 
            // R_Set
            // 
            this.R_Set.Image = ((System.Drawing.Image)(resources.GetObject("R_Set.Image")));
            this.R_Set.Label = "Round公式";
            this.R_Set.Name = "R_Set";
            this.R_Set.ShowImage = true;
            this.R_Set.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.R_Set_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.cal);
            this.group5.Items.Add(this.ntbook);
            this.group5.Label = "其他";
            this.group5.Name = "group5";
            // 
            // cal
            // 
            this.cal.Image = ((System.Drawing.Image)(resources.GetObject("cal.Image")));
            this.cal.Label = "计算器";
            this.cal.Name = "cal";
            this.cal.ShowImage = true;
            this.cal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cal_Click);
            // 
            // ntbook
            // 
            this.ntbook.Image = ((System.Drawing.Image)(resources.GetObject("ntbook.Image")));
            this.ntbook.Label = "记事本";
            this.ntbook.Name = "ntbook";
            this.ntbook.ShowImage = true;
            this.ntbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ntbook_Click);
            // 
            // About
            // 
            this.About.Items.Add(this.about0);
            this.About.Label = "关于";
            this.About.Name = "About";
            // 
            // about0
            // 
            this.about0.Image = ((System.Drawing.Image)(resources.GetObject("about0.Image")));
            this.about0.Label = "关于";
            this.about0.Name = "about0";
            this.about0.ShowImage = true;
            this.about0.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.about0_Click);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            // 
            // Rup_num
            // 
            this.Rup_num.Image = ((System.Drawing.Image)(resources.GetObject("Rup_num.Image")));
            this.Rup_num.Label = "RoundUp一下";
            this.Rup_num.Name = "Rup_num";
            this.Rup_num.ShowImage = true;
            this.Rup_num.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Rup_num_Click);
            // 
            // Rup_set
            // 
            this.Rup_set.Image = ((System.Drawing.Image)(resources.GetObject("Rup_set.Image")));
            this.Rup_set.Label = "RoundUp公式";
            this.Rup_set.Name = "Rup_set";
            this.Rup_set.ShowImage = true;
            this.Rup_set.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Rup_set_Click);
            // 
            // Rdown_num
            // 
            this.Rdown_num.Image = ((System.Drawing.Image)(resources.GetObject("Rdown_num.Image")));
            this.Rdown_num.Label = "RoundDown一下";
            this.Rdown_num.Name = "Rdown_num";
            this.Rdown_num.ShowImage = true;
            this.Rdown_num.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Rdown_num_Click);
            // 
            // Rdown_set
            // 
            this.Rdown_set.Image = ((System.Drawing.Image)(resources.GetObject("Rdown_set.Image")));
            this.Rdown_set.Label = "RoundDown公式";
            this.Rdown_set.Name = "Rdown_set";
            this.Rdown_set.ShowImage = true;
            this.Rdown_set.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Rdown_set_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.About.ResumeLayout(false);
            this.About.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton zhengfu;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox xiaoshuwei;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton R_num;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton R_Set;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ntbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup About;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton about0;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Rup_num;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Rup_set;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Rdown_num;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Rdown_set;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
