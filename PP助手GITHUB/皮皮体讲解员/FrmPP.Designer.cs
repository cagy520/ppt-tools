
namespace 皮皮助手
{
    partial class FrmPP
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPP));
            this.alphaFormTransformer1 = new AlphaForm.AlphaFormTransformer();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.音色选择ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.男声ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.女声xiaoxiaoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.alphaFormMarker1 = new AlphaForm.AlphaFormMarker();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.ttMsg = new System.Windows.Forms.ToolTip(this.components);
            this.alphaFormTransformer1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // alphaFormTransformer1
            // 
            this.alphaFormTransformer1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.alphaFormTransformer1.CanDrag = true;
            this.alphaFormTransformer1.ContextMenuStrip = this.contextMenuStrip1;
            this.alphaFormTransformer1.Controls.Add(this.pictureBox1);
            this.alphaFormTransformer1.Controls.Add(this.alphaFormMarker1);
            this.alphaFormTransformer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.alphaFormTransformer1.DragSleep = ((uint)(30u));
            this.alphaFormTransformer1.Location = new System.Drawing.Point(0, 0);
            this.alphaFormTransformer1.Name = "alphaFormTransformer1";
            this.alphaFormTransformer1.Size = new System.Drawing.Size(210, 238);
            this.alphaFormTransformer1.TabIndex = 1;
            this.alphaFormTransformer1.MouseEnter += new System.EventHandler(this.alphaFormTransformer1_MouseEnter);
            this.alphaFormTransformer1.MouseLeave += new System.EventHandler(this.alphaFormTransformer1_MouseLeave);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.音色选择ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(125, 48);
            // 
            // 音色选择ToolStripMenuItem
            // 
            this.音色选择ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.男声ToolStripMenuItem,
            this.女声xiaoxiaoToolStripMenuItem});
            this.音色选择ToolStripMenuItem.Name = "音色选择ToolStripMenuItem";
            this.音色选择ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.音色选择ToolStripMenuItem.Text = "音色选择";
            // 
            // 男声ToolStripMenuItem
            // 
            this.男声ToolStripMenuItem.Name = "男声ToolStripMenuItem";
            this.男声ToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.男声ToolStripMenuItem.Text = "男声(Yunyang)";
            this.男声ToolStripMenuItem.Click += new System.EventHandler(this.男声ToolStripMenuItem_Click);
            // 
            // 女声xiaoxiaoToolStripMenuItem
            // 
            this.女声xiaoxiaoToolStripMenuItem.Name = "女声xiaoxiaoToolStripMenuItem";
            this.女声xiaoxiaoToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.女声xiaoxiaoToolStripMenuItem.Text = "女声(Xiaoxiao)";
            this.女声xiaoxiaoToolStripMenuItem.Click += new System.EventHandler(this.女声xiaoxiaoToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox1.Location = new System.Drawing.Point(69, 142);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(17, 18);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Tag = "点击展开设置";
            this.pictureBox1.Click += new System.EventHandler(this.DrawerButton_Click);
            // 
            // alphaFormMarker1
            // 
            this.alphaFormMarker1.FillBorder = ((uint)(4u));
            this.alphaFormMarker1.Location = new System.Drawing.Point(93, 98);
            this.alphaFormMarker1.Name = "alphaFormMarker1";
            this.alphaFormMarker1.Size = new System.Drawing.Size(17, 17);
            this.alphaFormMarker1.TabIndex = 0;
            // 
            // timer1
            // 
            this.timer1.Interval = 4500;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // timer2
            // 
            this.timer2.Enabled = true;
            this.timer2.Interval = 2000;
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // ttMsg
            // 
            this.ttMsg.AutomaticDelay = 1000;
            this.ttMsg.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ttMsg.IsBalloon = true;
            this.ttMsg.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // FrmPP
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(210, 238);
            this.Controls.Add(this.alphaFormTransformer1);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmPP";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "皮皮助手";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.FrmPP_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.FrmPP_DragEnter);
            this.MouseEnter += new System.EventHandler(this.FrmPP_MouseEnter);
            this.MouseLeave += new System.EventHandler(this.FrmPP_MouseLeave);
            this.alphaFormTransformer1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AlphaForm.AlphaFormTransformer alphaFormTransformer1;
        private AlphaForm.AlphaFormMarker alphaFormMarker1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 音色选择ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 男声ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 女声xiaoxiaoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.ToolTip ttMsg;
    }
}