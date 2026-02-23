using FTD2XX_NET;
using JLcLib.Comn;
using JLcLib.Custom;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace SKAIChips_Verification
{
    public partial class RegContForm : Form
    {
        public enum ChipTypes
        {
            Unknown,

            /* ABOV */
            ABOV_Toscana,
            Aladdin,

            /* SKAIChips SG MCU */
            EQUUS,
            TWS,

            /* SKAIChips MCUP */
            Barcelona,
            Toscana,
            Segovia,
            Columbus,
            Oasis,
            OASIS,

            /* SKAIChips Y-Flash */
            SLM4781A,
            STD1402Q,

            /* SKAIChips SCP1501_N1C */
            IRIS_N1C,
            IRIS_R4,
            IRIS_R5,
            IRIS_N2,

            /* SKAIChips IP with SPI */
            Lyon,
            Salus8,
            IRIS2,
            Chicago,

            /* GCT chips */
            GRF7255RF,
            GRF7255IF,

            /* Celfras */
            CWP2000,

            /* TI */
            TMP117,

            /* Hyundai Mobis */
            Santorini,

            MaxChipTypes,
        }

        public delegate void LogControlFunc();
        public delegate void CommandControlFunc(string Command);
        public delegate bool MakeRegisterFunc(string Name, string[,] Data);
        public delegate void ButtonClickFunc();

        private System.Diagnostics.Stopwatch RegStopwatch = null;
        private System.Threading.Thread LogThread = null;
        private System.Threading.ManualResetEvent ComnEvent = new System.Threading.ManualResetEvent(true);
        private volatile bool IsLogRunning = false;

        private LogControlFunc RunLogFunc = null;
        private CommandControlFunc SendCommand = null;

        public ProgressBar ProgressBar { get; set; }
        public ToolStripStatusLabel StatusBar { get; set; }
        public JLcLib.IniFile IniFile { get; set; }
        public WireComnTypes ComnType { get; set; } = WireComnTypes.I2C;
        public IComn iComn { get; set; }
        public JLcLib.Custom.I2C I2C { get; set; } = new JLcLib.Custom.I2C();
        public JLcLib.Custom.I2C I2C1 { get; set; } = new JLcLib.Custom.I2C();
        public JLcLib.Custom.I2C I2C2 { get; set; } = new JLcLib.Custom.I2C();
        public JLcLib.Custom.I2C I2C3 { get; set; } = new JLcLib.Custom.I2C();
        public JLcLib.Custom.I2C I2C4 { get; set; } = new JLcLib.Custom.I2C();
        public JLcLib.Custom.I2C I2C5 { get; set; } = new JLcLib.Custom.I2C();
        public JLcLib.Custom.MPSSE MPSSE { get; set; } = new JLcLib.Custom.MPSSE();
        public JLcLib.Custom.SPI SPI { get; set; } = new JLcLib.Custom.SPI();
        public Serial Serial { get; set; } = new Serial();
        public FTDI fTDI { get; set; } = new FTDI();
        public JLcLib.Chip.RegisterManager RegMgr { get; set; }
        public JLcLib.Excel.ExcelManager xlMgr { get; set; }
        public string SelectedExcelFile { get; set; } = "";
        public List<string> SelectedSheets { get; set; } = new List<string>();
        public string SelectedScriptFile { get; private set; } = "";
        public string FirmwareName { get; set; } = "";
        public string SelectedLogFile { get; private set; } = "";
        public ChipTypes ChipType { get; set; } = ChipTypes.Unknown;
        public ChipControl ChipCtrl { get; set; } = null;

        /* Variable for chip specific control */
        public ComboBox ComboBox_TestItems { get; set; }
        public TextBox TextBox_TestArgument { get; set; }
        public Button[] ChipCtrlButtons { get; set; } = new Button[12];
        public TextBox[] ChipCtrlTextboxes { get; set; } = new TextBox[12];
        public JLcLib.CommandTextBox CommandTextBox { get; set; } = null;
        public JLcLib.LogRichTextBox LogRichTextBox { get; set; } = null;

        public RegContForm()
        {
            InitializeComponent();

            RegMgr = new JLcLib.Chip.RegisterManager(treeView_RegMap);
            xlMgr = new JLcLib.Excel.ExcelManager();

            RegStopwatch = new System.Diagnostics.Stopwatch();
            RegStopwatch.Start();

            LogThread = new System.Threading.Thread(LogThreadFunc);
            LogThread.Start();

            CommandTextBox = new JLcLib.CommandTextBox(textBox_RegCommand);
            CommandTextBox.PressEnterKey += CommandTextBox_PressEnterKey;
            LogRichTextBox = new JLcLib.LogRichTextBox(richTextBox_LogCtrl);
            LogRichTextBox.SetContextMenu();

            ComboBox_TestItems = comboBox_TestItems;
            TextBox_TestArgument = textBox_TestArgument;
        }

        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                Form prompt = new Form()
                {
                    Width = 360,
                    Height = 150,
                    Text = caption,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    StartPosition = FormStartPosition.CenterParent
                };

                Label label = new Label() { Left = 20, Top = 20, Text = text, Width = 300 };
                TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 300 };
                Button buttonOk = new Button() { Text = "OK", Left = 240, Width = 80, Top = 80 };

                buttonOk.Click += (sender, e) => { prompt.DialogResult = DialogResult.OK; prompt.Close(); };

                prompt.Controls.Add(label);
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(buttonOk);
                prompt.AcceptButton = buttonOk;

                return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }

        private void toolStripMenuItem_SearchReg_Click(object sender, EventArgs e)
        {
            string keyword = Prompt.ShowDialog("Enter the name of the register to search for.", "Register search");
            if (string.IsNullOrWhiteSpace(keyword)) return;

            List<TreeNode> matchedNodes = new List<TreeNode>();
            foreach (TreeNode node in treeView_RegMap.Nodes)
                matchedNodes.AddRange(FindAllMatchingNodes(node, keyword.ToUpper()));

            if (matchedNodes.Count == 0)
            {
                MessageBox.Show("No registers found with keyword: " + keyword, "Search Result");
                return;
            }

            if (matchedNodes.Count == 1)
            {
                treeView_RegMap.SelectedNode = matchedNodes[0];
                matchedNodes[0].EnsureVisible();
                return;
            }

            using (Form resultsForm = new Form())
            {
                resultsForm.Text = "Select Register";
                resultsForm.StartPosition = FormStartPosition.CenterParent;

                ListBox listBox = new ListBox()
                {
                    Font = new Font("Segoe UI", 10),
                    IntegralHeight = false,
                    Dock = DockStyle.Fill
                };

                int maxCharLen = 0;
                for (int i = 0; i < matchedNodes.Count; i++)
                {
                    string itemText = $"{(i + 1).ToString("D2")}. {matchedNodes[i].FullPath}";
                    listBox.Items.Add(itemText);
                    if (itemText.Length > maxCharLen)
                        maxCharLen = itemText.Length;
                }

                int maxPixelWidth = 0;
                using (Graphics g = treeView_RegMap.CreateGraphics())
                {
                    foreach (var item in listBox.Items)
                    {
                        int width = (int)g.MeasureString(item.ToString(), listBox.Font).Width;
                        if (width > maxPixelWidth)
                            maxPixelWidth = width;
                    }
                }

                int maxAllowedWidth = 900;
                listBox.Width = Math.Min(maxPixelWidth + 40, maxAllowedWidth);
                resultsForm.Width = listBox.Width + 30;


                int visibleItems = Math.Min(matchedNodes.Count, 12);
                int rowHeight = TextRenderer.MeasureText("Sample", listBox.Font).Height + 4;
                int listHeight = rowHeight * visibleItems + 20;

                int avgCharWidth = TextRenderer.MeasureText("0", listBox.Font).Width;
                int listWidth = avgCharWidth * Math.Min(maxCharLen + 4, 100);

                listBox.Height = listHeight;
                listBox.Width = listWidth;

                resultsForm.Width = listBox.Width + 30;
                resultsForm.Height = listBox.Height + 50;

                listBox.DoubleClick += (s, args) =>
                {
                    int index = listBox.SelectedIndex;
                    if (index >= 0)
                    {
                        TreeNode selected = matchedNodes[index];
                        treeView_RegMap.SelectedNode = selected;
                        selected.EnsureVisible();
                        resultsForm.Close();
                    }
                };

                resultsForm.Controls.Add(listBox);
                resultsForm.ShowDialog(this);
            }
        }

        private List<TreeNode> FindAllMatchingNodes(TreeNode root, string keyword)
        {
            List<TreeNode> result = new List<TreeNode>();
            if (root.Text.ToUpper().Contains(keyword))
                result.Add(root);

            foreach (TreeNode child in root.Nodes)
                result.AddRange(FindAllMatchingNodes(child, keyword));

            return result;
        }

        private TreeNode FindNodeRecursive(TreeNode root, string keyword)
        {
            if (root.Text.ToUpper().Contains(keyword))
                return root;

            foreach (TreeNode child in root.Nodes)
            {
                TreeNode result = FindNodeRecursive(child, keyword);
                if (result != null)
                    return result;
            }
            return null;
        }

        private void RegContForm_Load(object sender, EventArgs e)
        {
            ReadSettingFile();
            CommandTextBox.ReadSettingFile(IniFile);

            SetCommunicationUI();
            SetRegisterMapUI();

            RegMgr.SetRegControlGroupBox(groupBox_RegControl);
            RegMgr.SetRegControlButtonEnable(false);
            RegMgr.LogDataGridView = dataGridView_RegLog;

            groupBox_ChipSpecific.Enabled = false;
            tableLayoutPanel_RegCommand.Enabled = false;

            this.WindowState = FormWindowState.Maximized;

            this.treeView_RegMap.AfterSelect += TreeView_RegMap_AfterSelect_DisableRW;
        }

        private static readonly HashSet<uint> DisabledReadAddresses = new HashSet<uint>
        {

        };

        private static readonly HashSet<uint> DisabledWriteAddresses = new HashSet<uint>
        {
            // EF_RDATA0~7
            0x418, 0x419, 0x41A, 0x41B, 0x41C, 0x41D, 0x41E, 0x41F,
            // Device ID
            0x4FC,
            // REGRDATA0
            0x423,
            // READ_FLAG0
            0x4FE, 0x4FF
        };

        private bool ShouldDisableRead(uint addr)
        {
            return DisabledReadAddresses.Contains(addr);
        }

        private bool ShouldDisableWrite(uint addr)
        {
            return DisabledWriteAddresses.Contains(addr);
        }

        private void TreeView_RegMap_AfterSelect_DisableRW(object sender, TreeViewEventArgs e)
        {
            try
            {
                if (ChipType != ChipTypes.Chicago) return;
                if (SPI.IsOpen == false)    return;

                var readBtn = RegMgr?.ReadButton;
                var writeBtn = RegMgr?.WriteButton;
                var readAllBtn = RegMgr?.ReadAllButton;
                var writeAllBtn = RegMgr?.WriteAllButton;
                if (readBtn == null || writeBtn == null) return;

                var sel = RegMgr.SelectedRegister;
                if (sel == null)
                {
                    readBtn.Enabled = true;
                    writeBtn.Enabled = true;
                    return;
                }

                bool disableWrite = ShouldDisableWrite(sel.Address);
                writeBtn.Enabled = !disableWrite;
                writeAllBtn.Enabled = !disableWrite;

                bool disableRead = ShouldDisableRead(sel.Address);
                readBtn.Enabled = !disableRead;
                readAllBtn.Enabled = !disableRead;
            }
            catch { }
        }

        private void RegContForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            WriteSettingFile();
            CommandTextBox.WriteSettingFile(IniFile);

            StopLogThread();
            if (iComn != null && iComn.IsOpen)
                iComn.Close();
            LogThread.Abort();
            while (LogThread.IsAlive)
                ;
        }

        private void RegContForm_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void ReadSettingFile()
        {
            float fVar = 0;
            string sVar;

            if (IniFile == null)
                return;
            IniFile.Read(this);
            if (IniFile.Read(Name, "RegMapSplitDistance", ref fVar) == true)
                splitContainer_RegMap.SplitterDistance = (int)(fVar * splitContainer_RegMap.Size.Height);

            ReadCommunicationSettingFile();
            ReadRegisterMapSettingFile();
            ReadRegisterScriptSettingFile();

            sVar = IniFile.Read(Name, "RegLogFileName");
            if (sVar != null && sVar != "")
                SelectedLogFile = sVar;
        }

        private void WriteSettingFile()
        {
            if (IniFile == null)
                return;

            IniFile.Write(this);
            float fVar = (float)splitContainer_RegMap.SplitterDistance / (float)splitContainer_RegMap.Size.Height;
            IniFile.Write(Name, "RegMapSplitDistance", fVar.ToString("F2"));

            WriteCommunicationSettingFile();
            WriteRegisterMapSettingFile();
            WriteRegisterScriptSettingFile();
            WriteChipSpecificationSettingFile();

            IniFile.Write(Name, "RegLogFileName", SelectedLogFile);
        }

        #region Communication control group methods
        private void SetCommunicationUI()
        {
            comboBox_ComnTypes.Items.Clear();
            for (int i = 0; i < (int)WireComnTypes.MaxComn; i++)
                comboBox_ComnTypes.Items.Add(((WireComnTypes)i).ToString());
            comboBox_ComnTypes.SelectedIndex = (int)ComnType;
        }

        private void ReadCommunicationSettingFile()
        {
            int iVar = 0;

            if (IniFile.Read(Name, "ComnType", ref iVar))
                ComnType = (WireComnTypes)iVar;

            I2C.ReadSettingFile(IniFile, Name);
            SPI.ReadSettingFile(IniFile, Name);
            Serial.ReadSettingFile(IniFile, Name);
        }

        private void WriteCommunicationSettingFile()
        {
            IniFile.Write(Name, "ComnType", ((int)ComnType).ToString());
            I2C.WriteSettingFile(IniFile, Name);
            SPI.WriteSettingFile(IniFile, Name);
            Serial.WriteSettingFile(IniFile, Name);
        }

        private IComn GetCommunication()
        {
            IComn comn = null;

            switch (ComnType)
            {
                case WireComnTypes.I2C:
                    comn = I2C;
                    break;
                case WireComnTypes.SPI:
                    comn = SPI;
                    break;
                case WireComnTypes.Serial:
                    comn = Serial;
                    break;
            }
            return comn;
        }

        private void CheckCommunication()
        {
            if (iComn.IsAvailable())
            {
                label_ComnStatus.Text = iComn.DeviceName + " can be used.";
                label_ComnStatus.ForeColor = Color.DarkSlateBlue;
            }
            else
            {
                label_ComnStatus.Text = iComn.DeviceName + " cannot be used. Click \'Setup\'";
                label_ComnStatus.ForeColor = Color.OrangeRed;
            }
        }

        private bool IsConnection()
        {
            if ((iComn == null) || (iComn.IsOpen == false))
            {
                LogRichTextBox.WriteLine("Communication is not established!!", Color.Coral, LogRichTextBox.RichTextBox.BackColor);
                return false;
            }
            return true;
        }

        private void comboBox_ComnTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComnType = (WireComnTypes)comboBox_ComnTypes.SelectedIndex;
            iComn = GetCommunication();
            CheckCommunication();
        }

        private void button_Connection_Click(object sender, EventArgs e)
        {
            if (iComn == null)
                iComn = GetCommunication();

            if (iComn.IsOpen == false)
            {
                if (OpenFtdiByDescription())
                {
                    fTDI.SetBitMode(0x00, FTDI.FT_BIT_MODES.FT_BIT_MODE_RESET);
                }
                fTDI.Close();

                if (true)
                {
                    iComn.Open();
                }
                if (iComn.IsOpen)
                {
                    button_Connection.Text = "Disconnect";
                    label_ComnStatus.Text = iComn.StatusMessage;
                    label_ComnStatus.ForeColor = Color.ForestGreen;
                    RegMgr.SetRegControlButtonEnable(true);
                    groupBox_ChipSpecific.Enabled = true;
                    tableLayoutPanel_RegCommand.Enabled = true;

                    if (ChipType == ChipTypes.Chicago)
                    {
                        MPSSE.GPIOL_SetPins(251, 255, Protected: false);
                        MPSSE.SendCommand();
                    }
                }
                else
                {
                    button_Connection.Text = "Connect";
                    CheckCommunication();
                    RegMgr.SetRegControlButtonEnable(false);
                    groupBox_ChipSpecific.Enabled = false;
                    tableLayoutPanel_RegCommand.Enabled = false;
                }
            }
            else
            {
                StopLogThread();
                iComn.Close();
                button_Connection.Text = "Connect";
                CheckCommunication();
                RegMgr.SetRegControlButtonEnable(false);
                groupBox_ChipSpecific.Enabled = false;
                tableLayoutPanel_RegCommand.Enabled = false;

                if (OpenFtdiByDescription())
                {
                    fTDI.SetBitMode(0xFF, FTDI.FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG);
                    byte[] lowData = new byte[] { 0x00 };
                    uint bytesWritten = 0;
                    fTDI.Write(lowData, lowData.Length, ref bytesWritten);
                }
                fTDI.Close();
            }
        }

        private void button_ComnSetup_Click(object sender, EventArgs e)
        {
            switch (ComnType)
            {
                case WireComnTypes.I2C:
                    JLcLib.Custom.WireComnForm.Show(I2C);
                    break;
                case WireComnTypes.SPI:
                    JLcLib.Custom.WireComnForm.Show(SPI);
                    break;
                case WireComnTypes.Serial:
                    JLcLib.Custom.WireComnForm.Show(Serial);
                    break;
            }
            WriteCommunicationSettingFile();
            if (!iComn.IsOpen)
                CheckCommunication();
        }
        #endregion Communication control group methods

        #region Register map excel file control group methods
        private void ReadRegisterMapSettingFile()
        {
            string sVar;
            sVar = IniFile.Read(Name, "RegisterMapFileName");
            if (sVar != null && sVar != "")
                SelectedExcelFile = sVar;

            ReadRegisterMapSheetsNames();
        }

        private void ReadRegisterMapSheetsNames()
        {
            string sVar;
            string FileName = System.IO.Path.GetFileNameWithoutExtension(SelectedExcelFile);

            sVar = IniFile.Read(Name, "RegisterSheetsNames_in_" + FileName);
            if (sVar != null && sVar != "")
            {
                string[] Sheets = sVar.Split(new char[] { ',' });
                SelectedSheets.Clear();
                foreach (string Name in Sheets)
                    SelectedSheets.Add(Name);
            }
        }

        private void WriteRegisterMapSettingFile()
        {
            if (xlMgr.File.SelectedName != null)
                IniFile.Write(Name, "RegisterMapFileName", xlMgr.File.SelectedName);

            WriteRegisterMapSheetsNames();
        }

        private void WriteRegisterMapSheetsNames()
        {
            string FileName = System.IO.Path.GetFileNameWithoutExtension(SelectedExcelFile);
            if (SelectedSheets.Count > 0)
            {
                string Sheets = "";
                for (int i = 0; i < SelectedSheets.Count; i++)
                {
                    Sheets += SelectedSheets[i];
                    if (i < SelectedSheets.Count - 1)
                        Sheets += ",";
                }
                IniFile.Write(Name, "RegisterSheetsNames_in_" + FileName, Sheets);
            }
        }

        private void SetRegisterMapFilesToComboBox()
        {
            comboBox_SelectMapFIle.Items.Clear();
            if (xlMgr.File.Count > 0)
            {
                int Index = 0;
                for (int i = 0; i < xlMgr.File.Count; i++)
                {
                    comboBox_SelectMapFIle.Items.Add(xlMgr.File.Names[i]);
                    if (xlMgr.File.Names[i].CompareTo(System.IO.Path.GetFileName(SelectedExcelFile)) == 0)
                        Index = i;
                }
                comboBox_SelectMapFIle.SelectedIndex = Index;
            }
        }

        private void SetRegisterMapSheetsToComboBox()
        {
            if (xlMgr.Sheet.Count > 0)
            {
                int Index = 0;
                comboBox_SelectMapSheet.Items.Clear();
                RegMgr.Clear();
                for (int i = 0; i < xlMgr.Sheet.Count; i++)
                {
                    comboBox_SelectMapSheet.Items.Add(xlMgr.Sheet.Names[i]);
                    foreach (string Name in SelectedSheets)
                    {
                        if (Name.CompareTo(xlMgr.Sheet.Names[i]) == 0)
                        {
                            Index = i;
                        }
                    }
                }
                comboBox_SelectMapSheet.SelectedIndex = Index;
            }
        }

        private void SetRegisterMapUI()
        {
            SetRegisterMapFilesToComboBox();
        }

        private void AddRegisterMapToTreeview(string SheetName)
        {
            xlMgr.Sheet.Select(SheetName);

            string[,] Data = xlMgr.Cell.ReadAll();
            if (Data == null)
                return;

            if (ChipCtrl != null)
                RegMgr.AddRegisterGroup(ChipCtrl.MakeRegisterGroup(SheetName, Data));
        }

        private void comboBox_SelectMapFIle_SelectedIndexChanged(object sender, EventArgs e)
        {
            int Index = comboBox_SelectMapFIle.SelectedIndex;

            if (Index >= 0)
            {
                SelectedExcelFile = xlMgr.File.FullNames[Index];
                xlMgr.File.Select(Index + 1);
                ChipType = ChipTypes.Unknown;
                for (int i = 0; i < (int)ChipTypes.MaxChipTypes; i++)
                {
                    if (JLcLib.StringCtrl.CompareIn(((ChipTypes)i).ToString(), (string)comboBox_SelectMapFIle.SelectedItem) >= 0)
                    {
                        ChipType = (ChipTypes)i;
                        comboBox_TestItems.Items.Clear();
                        ReadRegisterMapSheetsNames();
                        SetChipSpecificationFunc(ChipType);
                        break;
                    }
                }
                label_RegMapName.Text = ChipType.ToString() + " Chip Register Map";

                SetRegisterMapSheetsToComboBox();

                if (SPI != null)
                    SPI.SetChipType(ChipType);
            }
        }

        private void comboBox_SelectMapFIle_DropDown(object sender, EventArgs e)
        {
            if (comboBox_SelectMapFIle.Items.Count != xlMgr.File.Count)
                SetRegisterMapFilesToComboBox();
        }

        private void button_AddRegTree_Click(object sender, EventArgs e)
        {
            int FileIndex = comboBox_SelectMapFIle.SelectedIndex;
            int SheetIndex = comboBox_SelectMapSheet.SelectedIndex;

            if (FileIndex >= 0 && SheetIndex >= 0)
            {
                xlMgr.File.Select(FileIndex + 1);
                string Name = (string)comboBox_SelectMapSheet.SelectedItem;
                SelectedSheets.Add(Name);
                WriteRegisterMapSheetsNames();
                AddRegisterMapToTreeview(Name);
            }
        }

        private void button_ClearRegTree_Click(object sender, EventArgs e)
        {
            RegMgr.Clear();
            SelectedSheets.Clear();
            WriteRegisterMapSheetsNames();
        }

        private void button_OpenMapFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileDialog = new OpenFileDialog())
            {
                if (SelectedExcelFile == "")
                    fileDialog.InitialDirectory = MainForm.AppPath;
                else
                    fileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(SelectedExcelFile);

                fileDialog.Filter = "Map File (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    SelectedExcelFile = fileDialog.FileName;
                    xlMgr.File.Open(SelectedExcelFile);
                    SetRegisterMapFilesToComboBox();
                }
            }
        }
        #endregion Register map excel file control group methods

        #region Register script control group methods
        private void ReadRegisterScriptSettingFile()
        {
            string sVar;
            sVar = IniFile.Read(Name, "RegisterScriptFileName");
            if (sVar != null && sVar != "")
            {
                SelectedScriptFile = sVar;
                label_RegScriptFileName.Text = "Script: " + System.IO.Path.GetFileName(SelectedScriptFile);
            }
        }

        private void WriteRegisterScriptSettingFile()
        {
            IniFile.Write(Name, "RegisterScriptFileName", SelectedScriptFile);
        }

        private void button_LoadScript_Click(object sender, EventArgs e)
        {
            RegMgr.LoadScriptFile(SelectedScriptFile);
        }

        private void button_ImportScript_Click(object sender, EventArgs e)
        {
            string InitPath;
            if (SelectedScriptFile == "")
                InitPath = System.IO.Directory.GetCurrentDirectory();
            else
                InitPath = System.IO.Path.GetDirectoryName(SelectedScriptFile);
            SelectedScriptFile = RegMgr.ImportScriptFile(InitPath);
            label_RegScriptFileName.Text = "Script: " + System.IO.Path.GetFileName(SelectedScriptFile);
        }

        private void button_ExportScript_Click(object sender, EventArgs e)
        {
            string InitPath;
            if (SelectedScriptFile == "")
                InitPath = System.IO.Directory.GetCurrentDirectory();
            else
                InitPath = System.IO.Path.GetDirectoryName(SelectedScriptFile);
            SelectedScriptFile = RegMgr.ExportScriptFile(InitPath);
            label_RegScriptFileName.Text = System.IO.Path.GetFileName(SelectedScriptFile);
        }
        #endregion Register script control group methods

        #region Register log control group methods
        private void button_ClearRegLog_Click(object sender, EventArgs e)
        {
            RegMgr.ClearLog();
        }

        private void button_RunRegLog_Click(object sender, EventArgs e)
        {
            if (iComn.IsOpen)
                RegMgr.RunLog(SelectedLogFile);
        }

        private void button_OpenRegLog_Click(object sender, EventArgs e)
        {
            string InitPath;
            if (SelectedLogFile == "")
                InitPath = System.IO.Directory.GetCurrentDirectory();
            else
                InitPath = System.IO.Path.GetDirectoryName(SelectedLogFile);
            SelectedLogFile = RegMgr.OpenLogFile(InitPath);
        }

        private void button_SaveRegLog_Click(object sender, EventArgs e)
        {
            string InitPath;
            if (SelectedLogFile == "")
                InitPath = System.IO.Directory.GetCurrentDirectory();
            else
                InitPath = System.IO.Path.GetDirectoryName(SelectedLogFile);
            SelectedLogFile = RegMgr.SaveLogFile(InitPath);
        }
        #endregion Register log control group methods

        #region Chip debugging group methods
        private void LogThreadFunc()
        {
            while (true)
            {
                if (IsLogRunning)
                {
                    ComnEvent.WaitOne(500);
                    ComnEvent.Reset();

                    RunLogFunc?.Invoke();

                    ComnEvent.Set();
                    System.Threading.Thread.Sleep(20);
                }
                else
                    System.Threading.Thread.Sleep(200);
            }
        }

        private void button_SendRegCommand_Click(object sender, EventArgs e)
        {
            SendCommand?.Invoke(textBox_RegCommand.Text);
        }

        private void button_RunLog_Click(object sender, EventArgs e)
        {
            if (ChipCtrl == null)
            {
                LogRichTextBox.WriteLine("There is no selected chip!!", Color.Tomato, LogRichTextBox.RichTextBox.BackColor);
                return;
            }
            button_RunLog.Invoke(new MethodInvoker(delegate ()
            {
                if (ChipCtrl.CheckConnectionForLog())
                {
                    if (IsLogRunning == false)
                    {
                        IsLogRunning = true;
                        button_RunLog.Text = "Stop Log";
                    }
                    else
                    {
                        IsLogRunning = false;
                        button_RunLog.Text = "Run Log";
                    }
                }
            }));
        }

        public void StopLogThread()
        {
            if (IsLogRunning)
            {
                button_RunLog_Click(button_RunLog, EventArgs.Empty);
                ComnEvent.WaitOne(500);
            }
        }

        public void StartLogThread()
        {
            if (!IsLogRunning)
                button_RunLog_Click(button_RunLog, EventArgs.Empty);
        }
        #endregion Chip debugging group methods

        #region Chip specific control group methods
        private void SetChipSpecificControlUI()
        {
            int Width = 68, Height = 20;
            int xPos, yPos;
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */
            for (int i = 0; i < ChipCtrlButtons.Length; i++)
            {
                ChipCtrlButtons[i]?.Dispose();
                ChipCtrlTextboxes[i]?.Dispose();

                Button b = ChipCtrlButtons[i] = new Button();
                TextBox t = ChipCtrlTextboxes[i] = new TextBox();

                groupBox_ChipSpecific.Controls.Add(b);
                groupBox_ChipSpecific.Controls.Add(t);
                b.Size = new Size(Width, Height + 1);
                t.Size = new Size(Width, Height);

                xPos = 5 + (Width + 3) * (i % 4);
                yPos = 15 + 25 + (Height + 3) * (i / 4);
                b.Location = new Point(xPos, yPos);
                t.Location = new Point(xPos, yPos);

                b.Visible = false;
                t.Visible = false;
            }
        }

        private void ReadChipSpecificationSettingFile()
        {
            for (int i = 0; i < ChipCtrlTextboxes.Length; i++)
            {
                string sVar = IniFile.Read(Name, ChipType.ToString() + "_Control_Text_" + i.ToString());
                ChipCtrlTextboxes[i].Text = sVar;
            }
        }

        private void WriteChipSpecificationSettingFile()
        {
            for (int i = 0; i < ChipCtrlTextboxes.Length; i++)
            {
                if (ChipCtrlTextboxes[i] != null)
                    IniFile.Write(Name, ChipType.ToString() + "_Control_Text_" + i.ToString(), ChipCtrlTextboxes[i].Text);
            }
        }

        private void SetChipSpecificationFunc(ChipTypes Type)
        {
            switch (Type)
            {
                case ChipTypes.EQUUS:
                    ChipCtrl = new SKAI_SG.EQUUS(this);
                    break;
                case ChipTypes.TWS:
                case ChipTypes.CWP2000:
                    ChipCtrl = new SKAI_SG.TWS(this);
                    break;
                case ChipTypes.Barcelona:
                    ChipCtrl = new SKAI_MCUP.Barcelona(this);
                    break;
                case ChipTypes.Toscana:
                case ChipTypes.Segovia:
                    ChipCtrl = new SKAI_MCUP.Toscana(this);
                    break;
                case ChipTypes.Columbus:
                    ChipCtrl = new SKAI_MCUP.Columbus(this);
                    break;
                case ChipTypes.Oasis:
                    ChipCtrl = new SKAI_MCUP.Oasis(this);
                    break;
                case ChipTypes.OASIS:
                    ChipCtrl = new SKAI_MCUP.Oasis(this);
                    break;
                case ChipTypes.SLM4781A:
                    ChipCtrl = new SKAI_YFLASH.SLM4781A(this);
                    break;
                case ChipTypes.STD1402Q:
                    ChipCtrl = new SKAI_YFLASH.STD1402Q(this);
                    break;
                case ChipTypes.IRIS_N1C:
                    ChipCtrl = new SKAI_IRIS.SCP1501_N1C(this);
                    break;
                case ChipTypes.IRIS_R4:
                    ChipCtrl = new SKAI_IRIS.SCP1501_R4(this);
                    break;
                case ChipTypes.IRIS_R5:
                    ChipCtrl = new SKAI_IRIS.SCP1501_R5(this);
                    break;
                case ChipTypes.IRIS_N2:
                    ChipCtrl = new SKAI_IRIS.SCP1501_N2(this);
                    break;
                case ChipTypes.Lyon:
                    ChipCtrl = new SKAI_SPI.Lyon(this);
                    break;
                case ChipTypes.Salus8:
                    ChipCtrl = new SKAI_SPI.Salus8(this);
                    break;
                case ChipTypes.IRIS2:
                    ChipCtrl = new SKAI_SPI.IRIS2(this);
                    break;
                case ChipTypes.Chicago:
                    ChipCtrl = new SKAI_SPI.Chicago(this);
                    break;
                case ChipTypes.GRF7255RF:
                case ChipTypes.GRF7255IF:
                    ChipCtrl = new GCT.GRF7255(this);
                    break;

                case ChipTypes.ABOV_Toscana:
                    ChipCtrl = new ABOV.ABOV_Toscana(this);
                    break;
                case ChipTypes.Aladdin:
                    ChipCtrl = new ABOV.Aladdin(this);
                    break;
                case ChipTypes.TMP117:
                    ChipCtrl = new TI.TMP117(this);
                    break;

                case ChipTypes.Santorini:
                    ChipCtrl = new HD_MOBIS.Santorini(this);
                    break;
                default:
                    ChipCtrl = null;
                    RunLogFunc = null;
                    SendCommand = null;
                    break;
            }
            if (ChipCtrl != null)
            {
                RunLogFunc = ChipCtrl.RunLog;
                SendCommand = ChipCtrl.SendCommand;

                SetChipSpecificControlUI();
                ChipCtrl.SetChipSpecificUI();
                ReadChipSpecificationSettingFile();
            }
        }

        private void CommandTextBox_PressEnterKey(object sender, JLcLib.CommandTextBox.CommandTextBoxEventArg e)
        {
            SendCommand?.Invoke(e.Command);
        }

        private void button_RunChipSpecTest_Click(object sender, EventArgs e)
        {
            if (ChipCtrl != null)
                ChipCtrl.RunTest(comboBox_TestItems.SelectedIndex, textBox_TestArgument.Text);
        }
        #endregion Chip specific control group methods

        private bool OpenFtdiByDescription(string targetDescription = "UM232H")
        {
            FTDI ftdi = new FTDI();

            uint deviceCount = 0;
            if (ftdi.GetNumberOfDevices(ref deviceCount) != FTDI.FT_STATUS.FT_OK || deviceCount == 0)
            {
                return false;
            }

            FTDI.FT_DEVICE_INFO_NODE[] deviceList = new FTDI.FT_DEVICE_INFO_NODE[deviceCount];
            if (ftdi.GetDeviceList(deviceList) != FTDI.FT_STATUS.FT_OK)
            {
                return false;
            }

            foreach (var device in deviceList)
            {
                if (device.Description.Contains(targetDescription))
                {
                    FTDI.FT_STATUS status = ftdi.OpenBySerialNumber(device.SerialNumber);
                    if (status == FTDI.FT_STATUS.FT_OK)
                    {
                        fTDI = ftdi;
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            return false;
        }

    }

    public abstract class ChipControl
    {
        protected RegContForm Parent { get; }
        protected IComn iComn { get; }
        protected JLcLib.LogRichTextBox Log { get; set; }
        protected ProgressBar ProgressBar { get; set; }
        protected Label StatusLabel { get; set; }
        protected ToolStripStatusLabel StatusBar { get; set; }

        protected ComboBox ComboBox_TestItems { get; set; }
        protected TextBox TextBox_TestArgument { get; set; }

        protected string LogMessage;
        protected RegContForm.ChipTypes ChipType = RegContForm.ChipTypes.Unknown;
        protected int FirmwareSize;
        protected string FirmwareName;
        protected byte[] FirmwareData;
        protected int ReadFirmwareSize;
        protected byte[] ReadFirmwareData;

        protected BackgroundWorker backgroundWorker = null;

        protected delegate void CallbackFunc();

        public bool Status = false;
        public byte[] CalibrationData { get; set; }

        public ChipControl(RegContForm form)
        {
            Parent = form;

            iComn = Parent.iComn;
            Log = Parent.LogRichTextBox;
            ProgressBar = Parent.ProgressBar;
            StatusBar = Parent.StatusBar;
            ChipType = Parent.ChipType;
            FirmwareName = Parent.FirmwareName;
            ComboBox_TestItems = Parent.ComboBox_TestItems;
            TextBox_TestArgument = Parent.TextBox_TestArgument;
        }

        protected bool GetFirmwareFileName()
        {
            using (OpenFileDialog FileDlg = new OpenFileDialog())
            {
                FileDlg.Filter = "FW File (*.bin,*.hex)|*.bin;*.hex|All files (*.*)|*.*";
                FirmwareName = Parent.IniFile.Read(Parent.Name, "FirmwareName");
                if (FirmwareName == "")
                    FileDlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
                else
                    FileDlg.InitialDirectory = System.IO.Path.GetDirectoryName(FirmwareName);

                if (FileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    FirmwareName = FileDlg.FileName;
                    Parent.IniFile.Write(Parent.Name, "FirmwareName", FirmwareName);
                    return true;
                }
            }
            return false;
        }

        protected void RunBackgroudFunc(CallbackFunc Func)
        {
            if (backgroundWorker != null && backgroundWorker.IsBusy == true)
                return;
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerSupportsCancellation = false;
            backgroundWorker.WorkerReportsProgress = false;
            backgroundWorker.DoWork += new DoWorkEventHandler(delegate { Func?.Invoke(); });
            backgroundWorker.RunWorkerAsync();
        }

        protected void RunBackgroudFunc(CallbackFunc DoWorkFunc, CallbackFunc CompletedFunc)
        {
            if (backgroundWorker != null && backgroundWorker.IsBusy == true)
                return;
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerSupportsCancellation = false;
            backgroundWorker.WorkerReportsProgress = false;
            backgroundWorker.DoWork += new DoWorkEventHandler(delegate { DoWorkFunc?.Invoke(); });
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(delegate { CompletedFunc?.Invoke(); });
            backgroundWorker.RunWorkerAsync();
        }

        protected void SetText_ChipCtrlTextboxes(int Index, string Text)
        {
            if (Index < Parent.ChipCtrlTextboxes.Length)
            {
                Parent.ChipCtrlTextboxes[Index].Invoke((new MethodInvoker(delegate ()
                {
                    Parent.ChipCtrlTextboxes[Index].Text = Text;
                })));
            }
        }

        protected string GetText_ChipCtrlTextboxes(int Index)
        {
            string Text = "";
            if (Index < Parent.ChipCtrlTextboxes.Length)
            {
                Parent.ChipCtrlTextboxes[Index].Invoke((new MethodInvoker(delegate ()
                {
                    Text = Parent.ChipCtrlTextboxes[Index].Text;
                })));
            }
            return Text;
        }

        protected int GetInt_ChipCtrlTextboxes(int Index)
        {
            string Text = "";
            int Val = 0;

            if (Index < Parent.ChipCtrlTextboxes.Length)
            {
                Parent.ChipCtrlTextboxes[Index].Invoke((new MethodInvoker(delegate ()
                {
                    Text = Parent.ChipCtrlTextboxes[Index].Text;
                    try { Val = int.Parse(Text, System.Globalization.NumberStyles.Number); }
                    catch { Parent.ChipCtrlTextboxes[Index].Text = Val.ToString(); }
                })));
            }
            return Val;
        }

        protected double GetDouble_ChipCtrlTextboxes(int Index)
        {
            string Text = "";
            double Val = 0F;

            if (Index < Parent.ChipCtrlTextboxes.Length)
            {
                Parent.ChipCtrlTextboxes[Index].Invoke((new MethodInvoker(delegate ()
                {
                    Text = Parent.ChipCtrlTextboxes[Index].Text;
                    try { Val = double.Parse(Text, System.Globalization.NumberStyles.Number); }
                    catch { Parent.ChipCtrlTextboxes[Index].Text = Val.ToString(); }
                })));
            }
            return Val;
        }

        protected string GetFirmwareDataText(byte[] FirmwareData)
        {
            string sData = "";

            for (int i = 0; i < FirmwareData.Length; i++)
            {
                if ((i % 16) == 0)
                    sData += i.ToString("X4") + ": ";
                sData += FirmwareData[i].ToString("X2") + " ";
                if ((i % 16) == 15)
                    sData += "\n";
            }
            return sData;
        }

        protected void WriteComparedFirmwareDataToLogRichTextBox()
        {
            string sData = "";
            int Length;
            int ErrorCount = 0;

            if (FirmwareData == null && ReadFirmwareData == null)
                return;

            Length = (FirmwareData.Length < ReadFirmwareData.Length) ? FirmwareData.Length : ReadFirmwareData.Length;

            for (int Addr = 0; Addr < Length; Addr += 16)
            {
                int Count = 0;
                sData += Addr.ToString("X4") + ": ";

                for (int i = 0; i < 16; i++)
                {
                    if ((Addr + i) >= Length)
                        break;
                    if (FirmwareData[Addr + i] == ReadFirmwareData[Addr + i])
                        sData += ReadFirmwareData[Addr + i].ToString("X2") + " ";
                    else
                    {
                        Log.Write(sData);
                        Log.Write(ReadFirmwareData[Addr + i].ToString("X2"), Color.Red, Log.RichTextBox.BackColor);
                        Log.Write(" ", Color.Black, Log.RichTextBox.BackColor);
                        sData = "";
                        ErrorCount++;
                    }
                    Count++;
                }
                string Spliter = "";
                for (int i = 0; i < 16 - Count; i++)
                    Spliter += "   ";
                sData += Spliter + "| ";
                for (int i = 0; i < 16; i++)
                {
                    if ((Addr + i) >= Length)
                        break;
                    if (FirmwareData[Addr + i] == ReadFirmwareData[Addr + i])
                        sData += FirmwareData[Addr + i].ToString("X2") + " ";
                    else
                    {
                        Log.Write(sData);
                        Log.Write(FirmwareData[Addr + i].ToString("X2"), Color.Red, Log.RichTextBox.BackColor);
                        Log.Write(" ", Color.Black, Log.RichTextBox.BackColor);
                        sData = "";
                    }
                }
                sData += "\n";
                if (ErrorCount > 100)
                    break;
            }
            Log.Write(sData);
        }

        abstract public JLcLib.Chip.RegisterGroup MakeRegisterGroup(string GroupName, string[,] RegData);

        abstract public void SetChipSpecificUI();

        protected void DownloadFirmware(CallbackFunc DownloadFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to download Firmware...");
            Log.WriteLine("→ " + System.IO.Path.GetFileName(FirmwareName));

            RunBackgroudFunc(DownloadFunc, DownloadComplete);
        }

        private void DownloadComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to download.\n", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to download.\n", Color.Coral, Log.RichTextBox.BackColor);

            TextBox_TestArgument.Text = FirmwareSize.ToString();
        }

        protected void EraseFirmware(CallbackFunc EraseFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to erase NV memory...");

            RunBackgroudFunc(EraseFunc, EraseComplete);
        }

        private void EraseComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to erase.\n", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to erase.\n", Color.Coral, Log.RichTextBox.BackColor);
        }

        protected void DumpFirmware(CallbackFunc DumpFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to dump Firmware...");

            RunBackgroudFunc(DumpFunc, DumpComplete);
        }

        private void DumpComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to Dump.\n", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to Dump.\n", Color.Coral, Log.RichTextBox.BackColor);
        }

        protected void VerifyPatternFlash(CallbackFunc VerifyFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to Verify Flash with Pattern...");

            RunBackgroudFunc(VerifyFunc, VerifyPatternFlashComplete);
        }

        private void VerifyPatternFlashComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to Flash Verify with Pattern.\n", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to Flash Verify with Pattern.\n", Color.Coral, Log.RichTextBox.BackColor);

            TextBox_TestArgument.Text = FirmwareSize.ToString();
        }

        protected void WriteCalibrationData(CallbackFunc WriteFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to write calibration data...");

            RunBackgroudFunc(WriteFunc, WriteCalibrationDataComplete);
        }

        private void WriteCalibrationDataComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to write.", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to write.", Color.Coral, Log.RichTextBox.BackColor);

            SetText_ChipCtrlTextboxes(9, FirmwareSize.ToString());
        }

        protected void RemoveCalibrationData(CallbackFunc RemoveFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to remove calibration data...");

            RunBackgroudFunc(RemoveFunc, RemoveCalibrationDataComplete);
        }

        private void RemoveCalibrationDataComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to remove.", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to remove.", Color.Coral, Log.RichTextBox.BackColor);
        }

        protected void ReadCalibrationData(CallbackFunc ReadFunc)
        {
            Parent.StopLogThread();
            Log.WriteLine("Start to read calibration data...");

            RunBackgroudFunc(ReadFunc, ReadCalibrationDataComplete);
        }

        private void ReadCalibrationDataComplete()
        {
            if (Status)
                Log.WriteLine("Succeed to read.", Color.ForestGreen, Log.RichTextBox.BackColor);
            else
                Log.WriteLine("Failed to read.", Color.Coral, Log.RichTextBox.BackColor);
        }

        protected uint GetInt16fromRegItems(JLcLib.Chip.RegisterItem msb8, JLcLib.Chip.RegisterItem lsb8)
        {
            msb8.Read();
            lsb8.Read();
            return ((msb8.Value & 0xFF) << 8) | (lsb8.Value & 0xFF);
        }

        protected void SetInt16byRegItems(JLcLib.Chip.RegisterItem msb8, JLcLib.Chip.RegisterItem lsb8, uint Value)
        {
            msb8.Value = (Value >> 8) & 0xFF;
            lsb8.Value = Value & 0xFF;
            msb8.Write();
            lsb8.Write();
        }

        abstract public bool CheckConnectionForLog();

        abstract public void RunLog();

        abstract public void SendCommand(string Command);

        abstract public void RunTest(int TestItemIndex, string Arg);
    }
}
