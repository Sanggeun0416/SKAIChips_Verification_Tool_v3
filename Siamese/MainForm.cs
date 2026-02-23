using SKAICHIPS_Verificationin;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SKAIChips_Verification
{
    public partial class MainForm : Form
    {
        #region Properties & Fields
        private RegContForm RegContForm { get; set; }

        public static string AppName { get; } = "SKAIChips_Verification";
        public static string Version { get; } = "v3.1.4 [Confidential]";

        public static string AppPath { get; } = System.IO.Directory.GetCurrentDirectory();

        public JLcLib.IniFile IniFile { get; set; }
        #endregion Properties & Fields

        #region Constructor & Form Event Handlersm
        public MainForm()
        {
            InitializeComponent();

            this.Text = AppName + " " + Version;

            JLcLib.Comn.FT_Device.FindAvailableDevices();

            IniFile = new JLcLib.IniFile(AppPath + "\\" + AppName + ".ini");
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            JLcLib.Instrument.InstrumentForm.MakeInstrumentList(IniFile);

            ReadSettingFIle();

            ShowRegContForm();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            WriteSettingFile();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
        }
        #endregion Constructor & Form Event Handlers

        #region Menu Click Event Handlers
        private void eXITXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void registerControllerRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsFormAvailable("RegContForm"))
            {
                ShowRegContForm();
            }
            else
            {
                RegContForm.Focus();
            }
        }

        private void instrumentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            JLcLib.Custom.InstrumentForm InsForm = new JLcLib.Custom.InstrumentForm(IniFile);
            InsForm.Show();
        }

        private void vectorDownloaderVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VectorDownForm vectorDownloaderForm = new VectorDownForm();
            vectorDownloaderForm.MdiParent = this; vectorDownloaderForm.IniFile = this.IniFile; vectorDownloaderForm.ProgressBar = toolStripProgressBar_Main.ProgressBar; vectorDownloaderForm.Show();
        }

        private void hciControllerHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HCIControlForm HciContForm = new HCIControlForm();
            HciContForm.MdiParent = this; HciContForm.Show();
        }
        #endregion Menu Click Event Handlers

        #region Helper Methods
        public bool IsFormAvailable(string formName)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Name == formName)
                {
                    return true;
                }
            }
            return false;
        }

        private void ReadSettingFIle()
        {
            IniFile.Read(this);
        }

        private void WriteSettingFile()
        {
            IniFile.Write(this);
        }

        private void ShowRegContForm()
        {
            RegContForm = new RegContForm()
            {
                MdiParent = this,
                ProgressBar = toolStripProgressBar_Main.ProgressBar,
                IniFile = this.IniFile,
                StatusBar = toolStripStatusLabel_Main
            };
            RegContForm.Show();
        }
        #endregion Helper Methods
    }

    public class TimedMessageBox : Form
    {
        private DialogResult result = DialogResult.None; private Timer timer; private Label label; private Button btnYes, btnNo;
        public static DialogResult Show(string text, string caption, int timeoutMs)
        {
            using (TimedMessageBox msgBox = new TimedMessageBox(text, caption, timeoutMs))
            {
                msgBox.StartPosition = FormStartPosition.CenterParent; msgBox.ShowDialog(); return msgBox.result;
            }
        }

        private TimedMessageBox(string text, string caption, int timeoutMs)
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Text = caption;
            this.Width = 400;
            this.Height = 150;
            this.Font = new Font("Segoe UI", 10);

            label = new Label()
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 60,
                TextAlign = ContentAlignment.MiddleCenter
            };

            btnYes = new Button() { Text = "Yes", DialogResult = DialogResult.Yes, Left = 100, Width = 80, Top = 70 };
            btnNo = new Button() { Text = "No", DialogResult = DialogResult.No, Left = 200, Width = 80, Top = 70 };

            btnYes.Click += (s, e) => { result = DialogResult.Yes; timer.Stop(); Close(); };
            btnNo.Click += (s, e) => { result = DialogResult.No; timer.Stop(); Close(); };

            Controls.Add(label);
            Controls.Add(btnYes);
            Controls.Add(btnNo);

            timer = new Timer { Interval = timeoutMs };
            timer.Tick += (s, e) =>
            {
                timer.Stop(); result = DialogResult.Yes; Close();
            };
            timer.Start();
        }
    }
}