
namespace 皮皮助手
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtInput = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button2 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.txtSerial = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtNum = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.lblPage = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.txtEDGE = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnEDGE = new System.Windows.Forms.Button();
            this.lblDh = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.chkAuto = new System.Windows.Forms.CheckBox();
            this.button4 = new System.Windows.Forms.Button();
            this.alphaFormTransformer1 = new AlphaForm.AlphaFormTransformer();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.alphaFormMarker1 = new AlphaForm.AlphaFormMarker();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.alphaFormTransformer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Location = new System.Drawing.Point(326, 71);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(50, 25);
            this.button1.TabIndex = 0;
            this.button1.Text = "选择";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Location = new System.Drawing.Point(3, 3);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(438, 131);
            this.textBox1.TabIndex = 1;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Microsoft Yunyang Online (Natural) - Chinese (Mainland)",
            "Microsoft HiuMaan Online (Natural) - Chinese (Hong Kong)",
            "Microsoft Xiaoxiao Online (Natural) - Chinese (Mainland)",
            "Microsoft Yaoyao - Chinese (Simplified, PRC)",
            "Microsoft Huihui - Chinese (Simplified, PRC)"});
            this.comboBox1.Location = new System.Drawing.Point(85, 42);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(235, 20);
            this.comboBox1.TabIndex = 2;
            this.comboBox1.Text = "Microsoft Yunyang Online (Natural) - Chinese (Mainland)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "音色选择";
            // 
            // txtInput
            // 
            this.txtInput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtInput.Location = new System.Drawing.Point(3, 3);
            this.txtInput.Multiline = true;
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(438, 131);
            this.txtInput.TabIndex = 4;
            this.txtInput.Text = "请输入朗读的内容。";
            this.txtInput.Click += new System.EventHandler(this.txtInput_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(25, 134);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(452, 163);
            this.tabControl1.TabIndex = 5;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.txtInput);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(444, 137);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "文本阅读";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(353, 87);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(72, 32);
            this.button2.TabIndex = 6;
            this.button2.Text = "朗读";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panel2);
            this.tabPage2.Controls.Add(this.panel1);
            this.tabPage2.Controls.Add(this.textBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(444, 137);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "PPT日志";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Location = new System.Drawing.Point(438, 18);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(200, 50);
            this.panel2.TabIndex = 16;
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(595, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(10, 10);
            this.panel1.TabIndex = 15;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.txtSerial);
            this.tabPage3.Controls.Add(this.label3);
            this.tabPage3.Controls.Add(this.txtNum);
            this.tabPage3.Controls.Add(this.label5);
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.button3);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(444, 137);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "软件注册";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // txtSerial
            // 
            this.txtSerial.Location = new System.Drawing.Point(103, 75);
            this.txtSerial.Multiline = true;
            this.txtSerial.Name = "txtSerial";
            this.txtSerial.Size = new System.Drawing.Size(188, 20);
            this.txtSerial.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(37, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 2;
            this.label3.Tag = "21";
            this.label3.Text = "序列号：";
            // 
            // txtNum
            // 
            this.txtNum.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.txtNum.Location = new System.Drawing.Point(103, 35);
            this.txtNum.Multiline = true;
            this.txtNum.Name = "txtNum";
            this.txtNum.ReadOnly = true;
            this.txtNum.Size = new System.Drawing.Size(188, 20);
            this.txtNum.TabIndex = 6;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(37, 38);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 3;
            this.label5.Tag = "20";
            this.label5.Text = "机器码：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(26, 23);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 12);
            this.label6.TabIndex = 4;
            this.label6.Tag = "19";
            this.label6.Visible = false;
            // 
            // button3
            // 
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Location = new System.Drawing.Point(331, 51);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(91, 23);
            this.button3.TabIndex = 0;
            this.button3.Text = "注册正式版";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(35, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "PPT文件";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtPath
            // 
            this.txtPath.AllowDrop = true;
            this.txtPath.Location = new System.Drawing.Point(86, 74);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(234, 21);
            this.txtPath.TabIndex = 7;
            this.txtPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
            this.txtPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            // 
            // lblPage
            // 
            this.lblPage.AutoSize = true;
            this.lblPage.Location = new System.Drawing.Point(326, 45);
            this.lblPage.Name = "lblPage";
            this.lblPage.Size = new System.Drawing.Size(71, 12);
            this.lblPage.TabIndex = 11;
            this.lblPage.Text = "当前PPT 0页";
            this.lblPage.Visible = false;
            // 
            // button5
            // 
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button5.Location = new System.Drawing.Point(419, 70);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(50, 25);
            this.button5.TabIndex = 13;
            this.button5.Text = "拆分";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click_1);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(327, 130);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(146, 23);
            this.progressBar1.TabIndex = 15;
            this.progressBar1.Visible = false;
            // 
            // txtEDGE
            // 
            this.txtEDGE.Location = new System.Drawing.Point(86, 107);
            this.txtEDGE.Name = "txtEDGE";
            this.txtEDGE.ReadOnly = true;
            this.txtEDGE.Size = new System.Drawing.Size(234, 21);
            this.txtEDGE.TabIndex = 17;
            this.txtEDGE.Text = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 112);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 18;
            this.label4.Text = "EDGE路径";
            // 
            // btnEDGE
            // 
            this.btnEDGE.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEDGE.Location = new System.Drawing.Point(326, 106);
            this.btnEDGE.Name = "btnEDGE";
            this.btnEDGE.Size = new System.Drawing.Size(50, 25);
            this.btnEDGE.TabIndex = 19;
            this.btnEDGE.Text = "选择";
            this.btnEDGE.UseVisualStyleBackColor = true;
            this.btnEDGE.Click += new System.EventHandler(this.btnEDGE_Click);
            // 
            // lblDh
            // 
            this.lblDh.AutoSize = true;
            this.lblDh.Location = new System.Drawing.Point(568, 32);
            this.lblDh.Name = "lblDh";
            this.lblDh.Size = new System.Drawing.Size(65, 12);
            this.lblDh.TabIndex = 17;
            this.lblDh.Text = "          ";
            // 
            // timer1
            // 
            this.timer1.Interval = 5000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // chkAuto
            // 
            this.chkAuto.AutoSize = true;
            this.chkAuto.Location = new System.Drawing.Point(405, 44);
            this.chkAuto.Name = "chkAuto";
            this.chkAuto.Size = new System.Drawing.Size(72, 16);
            this.chkAuto.TabIndex = 20;
            this.chkAuto.Text = "自动播放";
            this.chkAuto.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button4.Location = new System.Drawing.Point(419, 104);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(50, 24);
            this.button4.TabIndex = 21;
            this.button4.Text = "保存";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // alphaFormTransformer1
            // 
            this.alphaFormTransformer1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("alphaFormTransformer1.BackgroundImage")));
            this.alphaFormTransformer1.Controls.Add(this.linkLabel1);
            this.alphaFormTransformer1.Controls.Add(this.progressBar1);
            this.alphaFormTransformer1.Controls.Add(this.alphaFormMarker1);
            this.alphaFormTransformer1.Controls.Add(this.button4);
            this.alphaFormTransformer1.Controls.Add(this.chkAuto);
            this.alphaFormTransformer1.Controls.Add(this.tabControl1);
            this.alphaFormTransformer1.Controls.Add(this.btnEDGE);
            this.alphaFormTransformer1.Controls.Add(this.button1);
            this.alphaFormTransformer1.Controls.Add(this.lblPage);
            this.alphaFormTransformer1.Controls.Add(this.button5);
            this.alphaFormTransformer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.alphaFormTransformer1.DragSleep = ((uint)(30u));
            this.alphaFormTransformer1.Location = new System.Drawing.Point(0, 0);
            this.alphaFormTransformer1.Name = "alphaFormTransformer1";
            this.alphaFormTransformer1.Size = new System.Drawing.Size(503, 309);
            this.alphaFormTransformer1.TabIndex = 22;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.linkLabel1.Location = new System.Drawing.Point(457, 21);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(16, 16);
            this.linkLabel1.TabIndex = 22;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "X";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // alphaFormMarker1
            // 
            this.alphaFormMarker1.FillBorder = ((uint)(4u));
            this.alphaFormMarker1.Location = new System.Drawing.Point(305, 20);
            this.alphaFormMarker1.Name = "alphaFormMarker1";
            this.alphaFormMarker1.Size = new System.Drawing.Size(15, 16);
            this.alphaFormMarker1.TabIndex = 0;
            this.alphaFormMarker1.Load += new System.EventHandler(this.alphaFormMarker1_Load);
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(503, 309);
            this.Controls.Add(this.lblDh);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtEDGE);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.alphaFormTransformer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "皮皮助手";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.alphaFormTransformer1.ResumeLayout(false);
            this.alphaFormTransformer1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Label lblPage;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox txtEDGE;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnEDGE;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblDh;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.CheckBox chkAuto;
        private System.Windows.Forms.Button button4;
        private AlphaForm.AlphaFormTransformer alphaFormTransformer1;
        private AlphaForm.AlphaFormMarker alphaFormMarker1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TextBox txtSerial;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtNum;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button3;
    }
}

