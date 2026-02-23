
using SKAI_MCUP;

namespace SKAIChips_Verification
{
    partial class MainForm
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuStrip_Main = new System.Windows.Forms.MenuStrip();
            this.fileFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.eXITXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.registerControllerRToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.setupSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.instrumentToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip_Main = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel_Main = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar_Main = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.vectorDownloaderVToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.hciControllerHToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip_Main.SuspendLayout();
            this.statusStrip_Main.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip_Main
            // 
            this.menuStrip_Main.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip_Main.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileFToolStripMenuItem,
            this.toolTToolStripMenuItem,
            this.setupSToolStripMenuItem});
            this.menuStrip_Main.Location = new System.Drawing.Point(0, 0);
            this.menuStrip_Main.Name = "menuStrip_Main";
            this.menuStrip_Main.Size = new System.Drawing.Size(482, 28);
            this.menuStrip_Main.TabIndex = 0;
            this.menuStrip_Main.Text = "menuStrip1";
            // 
            // fileFToolStripMenuItem
            // 
            this.fileFToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.eXITXToolStripMenuItem});
            this.fileFToolStripMenuItem.Name = "fileFToolStripMenuItem";
            this.fileFToolStripMenuItem.Size = new System.Drawing.Size(63, 24);
            this.fileFToolStripMenuItem.Text = "File(&F)";
            // 
            // eXITXToolStripMenuItem
            // 
            this.eXITXToolStripMenuItem.Name = "eXITXToolStripMenuItem";
            this.eXITXToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.eXITXToolStripMenuItem.Text = "EXIT(&X)";
            this.eXITXToolStripMenuItem.Click += new System.EventHandler(this.eXITXToolStripMenuItem_Click);
            // 
            // toolTToolStripMenuItem
            // 
            this.toolTToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.registerControllerRToolStripMenuItem,
            this.toolStripSeparator1,
            this.hciControllerHToolStripMenuItem,
            this.vectorDownloaderVToolStripMenuItem});
            this.toolTToolStripMenuItem.Name = "toolTToolStripMenuItem";
            this.toolTToolStripMenuItem.Size = new System.Drawing.Size(71, 24);
            this.toolTToolStripMenuItem.Text = "Tool(&T)";
            // 
            // registerControllerRToolStripMenuItem
            // 
            this.registerControllerRToolStripMenuItem.Name = "registerControllerRToolStripMenuItem";
            this.registerControllerRToolStripMenuItem.Size = new System.Drawing.Size(244, 26);
            this.registerControllerRToolStripMenuItem.Text = "Register Controller(&R)";
            this.registerControllerRToolStripMenuItem.Click += new System.EventHandler(this.registerControllerRToolStripMenuItem_Click);
            // 
            // setupSToolStripMenuItem
            // 
            this.setupSToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.instrumentToolStripMenuItem});
            this.setupSToolStripMenuItem.Name = "setupSToolStripMenuItem";
            this.setupSToolStripMenuItem.Size = new System.Drawing.Size(80, 24);
            this.setupSToolStripMenuItem.Text = "Setup(&S)";
            // 
            // instrumentToolStripMenuItem
            // 
            this.instrumentToolStripMenuItem.Name = "instrumentToolStripMenuItem";
            this.instrumentToolStripMenuItem.Size = new System.Drawing.Size(179, 26);
            this.instrumentToolStripMenuItem.Text = "Instrument(&I)";
            this.instrumentToolStripMenuItem.Click += new System.EventHandler(this.instrumentToolStripMenuItem_Click);
            // 
            // statusStrip_Main
            // 
            this.statusStrip_Main.BackColor = System.Drawing.Color.Coral;
            this.statusStrip_Main.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip_Main.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel_Main,
            this.toolStripStatusLabel2,
            this.toolStripProgressBar_Main});
            this.statusStrip_Main.Location = new System.Drawing.Point(0, 427);
            this.statusStrip_Main.Name = "statusStrip_Main";
            this.statusStrip_Main.Size = new System.Drawing.Size(482, 26);
            this.statusStrip_Main.TabIndex = 1;
            this.statusStrip_Main.Text = "statusStrip1";
            // 
            // toolStripStatusLabel_Main
            // 
            this.toolStripStatusLabel_Main.ForeColor = System.Drawing.Color.White;
            this.toolStripStatusLabel_Main.Name = "toolStripStatusLabel_Main";
            this.toolStripStatusLabel_Main.Size = new System.Drawing.Size(67, 20);
            this.toolStripStatusLabel_Main.Text = "Idle";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(148, 20);
            this.toolStripStatusLabel2.Spring = true;
            // 
            // toolStripProgressBar_Main
            // 
            this.toolStripProgressBar_Main.Name = "toolStripProgressBar_Main";
            this.toolStripProgressBar_Main.Size = new System.Drawing.Size(250, 18);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(241, 6);
            // 
            // hciControllerHToolStripMenuItem
            // 
            this.hciControllerHToolStripMenuItem.Name = "hciControllerHToolStripMenuItem";
            this.hciControllerHToolStripMenuItem.Size = new System.Drawing.Size(244, 26);
            this.hciControllerHToolStripMenuItem.Text = "HCI Controller(&H)";
            this.hciControllerHToolStripMenuItem.Click += new System.EventHandler(this.hciControllerHToolStripMenuItem_Click);
            // 
            // vectorDownloaderVToolStripMenuItem
            // 
            this.vectorDownloaderVToolStripMenuItem.Name = "vectorDownloaderVToolStripMenuItem";
            this.vectorDownloaderVToolStripMenuItem.Size = new System.Drawing.Size(244, 26);
            this.vectorDownloaderVToolStripMenuItem.Text = "Vector Downloader(&V)";
            this.vectorDownloaderVToolStripMenuItem.Click += new System.EventHandler(this.vectorDownloaderVToolStripMenuItem_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(482, 453);
            this.Controls.Add(this.statusStrip_Main);
            this.Controls.Add(this.menuStrip_Main);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip_Main;
            this.Margin = new System.Windows.Forms.Padding(2, 4, 2, 4);
            this.Name = "MainForm";
            this.Text = "SKAIChips";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuStrip_Main.ResumeLayout(false);
            this.menuStrip_Main.PerformLayout();
            this.statusStrip_Main.ResumeLayout(false);
            this.statusStrip_Main.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip_Main;
        private System.Windows.Forms.ToolStripMenuItem fileFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolTToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem setupSToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip_Main;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel_Main;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar_Main;
        private System.Windows.Forms.ToolStripMenuItem eXITXToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem registerControllerRToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem instrumentToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem vectorDownloaderVToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem hciControllerHToolStripMenuItem;

    }
}

