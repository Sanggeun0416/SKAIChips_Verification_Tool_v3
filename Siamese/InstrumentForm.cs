using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using JLcLib.Instrument;

namespace JLcLib.Custom
{
    public class InstrumentForm : Form
    {
        private IContainer components = null;

        private TableLayoutPanel tableLayoutPanel_Instrument;

        private GroupBox groupBox_InsSetting;

        private GroupBox groupBox_InsCommand;

        private ComboBox comboBox_InsTypes;

        private Button button_AddInstrument;

        private Button button_RemoveInstrument;

        private DataGridView dataGridView_InsList;

        private Button button_InsDown;

        private Button button_InsUp;

        private TextBox textBox_InsType;

        private TextBox textBox_InsCommand;

        private Button button_SendInsCommand;

        private RichTextBox richTextBox_InsCommandLog;

        private DataGridViewTextBoxColumn Column_InsType;

        private DataGridViewCheckBoxColumn Column_InsEnabled;

        private DataGridViewTextBoxColumn Column_Address;

        private DataGridViewButtonColumn Column_InsTest;

        private DataGridViewTextBoxColumn Column_InsName;

        private Button button_InsScreenCapture;

        private Button button_ClearInsLog;

        private CommandTextBox CommandTextBox { get; set; }

        private LogRichTextBox LogRichTextBox { get; set; }

        public static List<InsInformation> InsInfoList { get; set; } = new List<InsInformation>();


        public IniFile IniFile { get; set; }

        public int SelectedInsIndex { get; private set; } = -1;


        public InsInformation SelectedIns { get; private set; } = null;


        public InstrumentForm(IniFile iniFile)
        {
            InitializeComponent();
            IniFile = iniFile;
            CommandTextBox = new CommandTextBox(textBox_InsCommand);
            CommandTextBox.PressEnterKey += CommandTextBox_PressEnterKey;
            LogRichTextBox = new LogRichTextBox(richTextBox_InsCommandLog);
            LogRichTextBox.SetContextMenu();
        }

        private void InstrumentForm_Load(object sender, EventArgs e)
        {
            ReadSettingFile();
            CommandTextBox.ReadSettingFile(IniFile);
            SetInsTypesComboBox();
            int width = dataGridView_InsList.Size.Width;
            dataGridView_InsList.Columns[0].Width = 120;
            dataGridView_InsList.Columns[1].Width = 24;
            dataGridView_InsList.Columns[2].Width = 200;
            dataGridView_InsList.Columns[3].Width = 40;
            dataGridView_InsList.Columns[4].Width = width - 120 - 24 - 200 - 40 - 20;
            dataGridView_InsList.RowTemplate.Height = 20;
            SetInstrumentListToDataGridView();
            SetSelectedInstrument(SelectedInsIndex);
            if (SelectedInsIndex >= 0)
            {
                dataGridView_InsList.CurrentCell = dataGridView_InsList.Rows[SelectedInsIndex].Cells[0];
            }
        }

        private void InstrumentForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            WriteSettingFile();
            CommandTextBox.WriteSettingFile(IniFile);
        }

        private void InstrumentForm_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void CommandTextBox_PressEnterKey(object sender, CommandTextBox.CommandTextBoxEventArg e)
        {
            if (SelectedIns == null)
            {
                return;
            }

            SCPI sCPI = new SCPI(SelectedIns.Type);
            if (sCPI.Open())
            {
                string text = e.Command + "\n";
                if (e.Command.Contains("?"))
                {
                    text += sCPI.WriteAndReadString(e.Command, 1000);
                }
                else
                {
                    sCPI.Write(e.Command);
                }

                sCPI.Close();
                LogRichTextBox.Write(text);
            }
        }

        private void ReadSettingFile()
        {
            int Value = 0;
            int num = 0;
            if (IniFile.Read(base.Name, "SelectedInstrument", ref Value))
            {
                SelectedInsIndex = Value;
            }

            InsInfoList.Clear();
            if (IniFile.Read(base.Name, "NumInstrument", ref Value))
            {
                num = Value;
            }

            for (int i = 0; i < num; i++)
            {
                string text = IniFile.Read(base.Name, "InstrumentType_" + i);
                if (text == null || text == "")
                {
                    continue;
                }

                for (int j = 0; j < (int)InstrumentTypes.NumInsTypes; j++)
                {
                    string text2 = text;
                    InstrumentTypes instrumentTypes = (InstrumentTypes)j;
                    if (text2 == instrumentTypes.ToString())
                    {
                        bool valid = false;
                        string address = "";
                        InstrumentTypes type = (InstrumentTypes)j;
                        if (IniFile.Read(base.Name, "InstrumentValid_" + i, ref Value))
                        {
                            valid = Value != 0;
                        }

                        text = IniFile.Read(base.Name, "InstrumentAddress_" + i);
                        if (text != null && text != "")
                        {
                            address = text;
                        }

                        InsInfoList.Add(new InsInformation(type, valid, address));
                        break;
                    }
                }
            }

            if (SelectedInsIndex >= 0 && SelectedInsIndex < InsInfoList.Count)
            {
                SelectedIns = InsInfoList[SelectedInsIndex];
            }
        }

        private void WriteSettingFile()
        {
            IniFile.Write(base.Name, "SelectedInstrument", SelectedInsIndex.ToString());
            IniFile.Write(base.Name, "NumInstrument", InsInfoList.Count.ToString());
            for (int i = 0; i < InsInfoList.Count; i++)
            {
                IniFile.Write(base.Name, "InstrumentType_" + i, InsInfoList[i].Type.ToString());
                IniFile.Write(base.Name, "InstrumentValid_" + i, (!InsInfoList[i].Valid) ? "0" : "1");
                IniFile.Write(base.Name, "InstrumentAddress_" + i, InsInfoList[i].Address);
            }
        }

        public static void MakeInstrumentList(IniFile iniFile)
        {
            int Value = 0;
            int num = 0;
            string section = "InstrumentForm";
            InsInfoList.Clear();
            if (iniFile.Read(section, "NumInstrument", ref Value))
            {
                num = Value;
            }

            for (int i = 0; i < num; i++)
            {
                string text = iniFile.Read(section, "InstrumentType_" + i);
                if (text == null || text == "")
                {
                    continue;
                }

                for (int j = 0; j < (int)InstrumentTypes.NumInsTypes; j++)
                {
                    string text2 = text;
                    InstrumentTypes instrumentTypes = (InstrumentTypes)j;
                    if (text2 == instrumentTypes.ToString())
                    {
                        bool valid = false;
                        string address = "";
                        InstrumentTypes type = (InstrumentTypes)j;
                        if (iniFile.Read(section, "InstrumentValid_" + i, ref Value))
                        {
                            valid = Value != 0;
                        }

                        text = iniFile.Read(section, "InstrumentAddress_" + i);
                        if (text != null && text != "")
                        {
                            address = text;
                        }

                        InsInfoList.Add(new InsInformation(type, valid, address));
                        break;
                    }
                }
            }
        }

        private void SetInstrumentListToDataGridView()
        {
            dataGridView_InsList.CellValueChanged -= dataGridView_InsList_CellValueChanged;
            dataGridView_InsList.Rows.Clear();
            for (int i = 0; i < InsInfoList.Count; i++)
            {
                dataGridView_InsList.Rows.Add();
                dataGridView_InsList.Rows[i].Cells[0].Value = InsInfoList[i].Type.ToString();
                dataGridView_InsList.Rows[i].Cells[1].Value = InsInfoList[i].Valid;
                dataGridView_InsList.Rows[i].Cells[2].Value = InsInfoList[i].Address;
                dataGridView_InsList.Rows[i].Cells[3].Value = "Gets";
            }

            dataGridView_InsList.CellValueChanged += dataGridView_InsList_CellValueChanged;
        }

        private void SetInsTypesComboBox()
        {
            comboBox_InsTypes.SelectedIndex = -1;
            comboBox_InsTypes.Items.Clear();
            for (int i = 0; i < (int)InstrumentTypes.NumInsTypes; i++)
            {
                InstrumentTypes instrumentTypes = (InstrumentTypes)i;
                bool flag = true;
                foreach (InsInformation insInfo in InsInfoList)
                {
                    if (insInfo.Type == instrumentTypes)
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {
                    comboBox_InsTypes.Items.Add(instrumentTypes.ToString());
                }
            }

            if (comboBox_InsTypes.Items.Count > 0)
            {
                comboBox_InsTypes.SelectedIndex = 0;
            }
        }

        private InstrumentTypes GetInstrumentTypeFromInsTypesComboBox()
        {
            InstrumentTypes result = InstrumentTypes.NumInsTypes;
            for (int i = 0; i < (int)InstrumentTypes.NumInsTypes; i++)
            {
                InstrumentTypes instrumentTypes = (InstrumentTypes)i;
                if (instrumentTypes.ToString() == (string)comboBox_InsTypes.SelectedItem)
                {
                    result = (InstrumentTypes)i;
                }
            }

            return result;
        }

        private void SetSelectedInstrument(int Index)
        {
            if (Index >= 0 && Index < InsInfoList.Count)
            {
                SelectedInsIndex = Index;
                SelectedIns = InsInfoList[SelectedInsIndex];
                textBox_InsType.Text = SelectedIns.Type.ToString();
            }
            else
            {
                SelectedInsIndex = -1;
                SelectedIns = null;
                textBox_InsType.Text = "";
            }
        }

        private void dataGridView_InsList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                SetSelectedInstrument(e.RowIndex);
            }
        }

        private void dataGridView_InsList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            SelectedInsIndex = e.RowIndex;
            SelectedIns = InsInfoList[e.RowIndex];
            if (dataGridView_InsList.CurrentCell.GetType() == typeof(DataGridViewCheckBoxCell))
            {
                if (dataGridView_InsList.CurrentCell.IsInEditMode && dataGridView_InsList.IsCurrentCellDirty)
                {
                    dataGridView_InsList.EndEdit();
                }
            }
            else if (dataGridView_InsList.CurrentCell.GetType() == typeof(DataGridViewButtonCell))
            {
                SCPI sCPI = new SCPI(SelectedIns.Type);
                DataGridViewCell dataGridViewCell = dataGridView_InsList.Rows[e.RowIndex].Cells[4];
                if (sCPI.Open())
                {
                    SelectedIns.Name = sCPI.GetInstrumentName();
                    sCPI.Close();
                    dataGridViewCell.Value = SelectedIns.Name;
                    dataGridViewCell.Style.ForeColor = Color.ForestGreen;
                }
                else
                {
                    dataGridViewCell.Value = "Check \"Enabled\" or \"VISA address\"!!";
                    dataGridViewCell.Style.ForeColor = Color.Coral;
                }
            }
        }

        private void dataGridView_InsList_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex == 1)
                {
                    InsInfoList[e.RowIndex].Valid = (bool)dataGridView_InsList.CurrentCell.Value;
                }
                else if (e.ColumnIndex == 2)
                {
                    InsInfoList[e.RowIndex].Address = (string)dataGridView_InsList.CurrentCell.Value;
                }
            }
        }

        private void button_AddInstrument_Click(object sender, EventArgs e)
        {
            InstrumentTypes instrumentTypeFromInsTypesComboBox = GetInstrumentTypeFromInsTypesComboBox();
            string text = "";
            for (int i = 0; i < (int)InstrumentTypes.NumInsTypes; i++)
            {
                if (instrumentTypeFromInsTypesComboBox.ToString() == IniFile.Read(base.Name, "InstrumentType_" + i))
                {
                    text = IniFile.Read(base.Name, "InstrumentAddress_" + i);
                    break;
                }
            }

            if (text == "")
            {
                text = "GPIB0::";
                switch (instrumentTypeFromInsTypesComboBox)
                {
                    case InstrumentTypes.SpectrumAnalyzer:
                        text += "18::INSTR";
                        break;
                    case InstrumentTypes.OscilloScope0:
                    case InstrumentTypes.OscilloScope1:
                        text += "7::INSTR";
                        break;
                    case InstrumentTypes.PowerSupply0:
                    case InstrumentTypes.PowerSupply1:
                    case InstrumentTypes.PowerSupply2:
                    case InstrumentTypes.PowerSupply3:
                        text += "5::INSTR";
                        break;
                    default:
                        text += "1::INSTR";
                        break;
                }
            }

            if (instrumentTypeFromInsTypesComboBox != InstrumentTypes.NumInsTypes)
            {
                InsInfoList.Add(new InsInformation(instrumentTypeFromInsTypesComboBox, Valid: false, text));
                SetInstrumentListToDataGridView();
                SetInsTypesComboBox();
            }
        }

        private void button_RemoveInstrument_Click(object sender, EventArgs e)
        {
            if (dataGridView_InsList.CurrentRow != null)
            {
                int index = dataGridView_InsList.CurrentRow.Index;
                int columnIndex = dataGridView_InsList.CurrentCell.ColumnIndex;
                int num = dataGridView_InsList.CurrentCell.RowIndex;
                InsInfoList.RemoveAt(index);
                SetInstrumentListToDataGridView();
                SetInsTypesComboBox();
                if (num > dataGridView_InsList.Rows.Count - 1)
                {
                    num = dataGridView_InsList.Rows.Count - 1;
                }

                if (dataGridView_InsList.Rows.Count > 0)
                {
                    dataGridView_InsList.CurrentCell = dataGridView_InsList.Rows[num].Cells[columnIndex];
                    SetSelectedInstrument(num);
                }
            }
        }

        private void button_InsUp_Click(object sender, EventArgs e)
        {
            if (dataGridView_InsList.CurrentRow != null)
            {
                int index = dataGridView_InsList.CurrentRow.Index;
                InsInformation item = InsInfoList[index];
                int columnIndex = dataGridView_InsList.CurrentCell.ColumnIndex;
                int rowIndex = dataGridView_InsList.CurrentCell.RowIndex;
                if (rowIndex > 0)
                {
                    InsInfoList.RemoveAt(index);
                    InsInfoList.Insert(index - 1, item);
                    SetInstrumentListToDataGridView();
                    dataGridView_InsList.CurrentCell = dataGridView_InsList.Rows[rowIndex - 1].Cells[columnIndex];
                    SetSelectedInstrument(rowIndex - 1);
                }
            }
        }

        private void button_InsDown_Click(object sender, EventArgs e)
        {
            if (dataGridView_InsList.CurrentRow != null)
            {
                int index = dataGridView_InsList.CurrentRow.Index;
                InsInformation item = InsInfoList[index];
                int columnIndex = dataGridView_InsList.CurrentCell.ColumnIndex;
                int rowIndex = dataGridView_InsList.CurrentCell.RowIndex;
                if (rowIndex < InsInfoList.Count - 1)
                {
                    InsInfoList.RemoveAt(index);
                    InsInfoList.Insert(index + 1, item);
                    SetInstrumentListToDataGridView();
                    dataGridView_InsList.CurrentCell = dataGridView_InsList.Rows[rowIndex + 1].Cells[columnIndex];
                    SetSelectedInstrument(rowIndex + 1);
                }
            }
        }

        private void button_SendInsCommand_Click(object sender, EventArgs e)
        {
            if (SelectedIns == null)
            {
                return;
            }

            SCPI sCPI = new SCPI(SelectedIns.Type);
            if (sCPI.Open())
            {
                string text = textBox_InsCommand.Text;
                string text2 = text + "\n";
                if (text.Contains("?"))
                {
                    text2 += sCPI.WriteAndReadString(text, 1000);
                }
                else
                {
                    sCPI.Write(text);
                }

                sCPI.Close();
                LogRichTextBox.Write(text2);
            }
        }

        private void button_InsScreenCapture_Click(object sender, EventArgs e)
        {
            if (SelectedIns == null)
            {
                return;
            }

            SCPI sCPI = new SCPI(SelectedIns.Type);
            if (!sCPI.Open())
            {
                return;
            }

            byte[] array = null;
            switch (SelectedIns.Type)
            {
                case InstrumentTypes.SpectrumAnalyzer:
                    {
                        string text = sCPI.WriteAndReadString(":MMEMory:CDIRectory?");
                        string text2 = "\\JL_TempScreenCapture.png";
                        string text3 = "";
                        string text4 = text;
                        for (int i = 0; i < text4.Length; i++)
                        {
                            char c = text4[i];
                            if (c != '"' && c != '\n')
                            {
                                text3 += c;
                            }
                        }

                        text3 += text2;
                        sCPI.Write(":MMEMory:STORe:SCReen \"" + text3 + "\"");
                        sCPI.Write("*WAI");
                        array = sCPI.WriteAndReadBytes(":MMEM:DATA? \"" + text3 + "\"", 30000);
                        sCPI.Write("*WAI");
                        sCPI.Write(":MMEM:DEL \"" + text3 + "\"");
                        sCPI.Write("*CLS");
                        break;
                    }
                case InstrumentTypes.OscilloScope0:
                case InstrumentTypes.OscilloScope1:
                    array = sCPI.WriteAndReadBytes(":HCOPY:SDUMp:DATA?");
                    break;
            }

            if (array != null)
            {
                MemoryStream stream = new MemoryStream(array);
                Bitmap image = (Bitmap)Image.FromStream(stream);
                Clipboard.Clear();
                Clipboard.SetImage(image);
                richTextBox_InsCommandLog.Paste();
            }

            sCPI.Close();
        }

        private void button_ClearInsLog_Click(object sender, EventArgs e)
        {
            richTextBox_InsCommandLog.Clear();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle = new System.Windows.Forms.DataGridViewCellStyle();
            tableLayoutPanel_Instrument = new System.Windows.Forms.TableLayoutPanel();
            groupBox_InsSetting = new System.Windows.Forms.GroupBox();
            button_InsDown = new System.Windows.Forms.Button();
            button_InsUp = new System.Windows.Forms.Button();
            dataGridView_InsList = new System.Windows.Forms.DataGridView();
            Column_InsType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Column_InsEnabled = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            Column_Address = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Column_InsTest = new System.Windows.Forms.DataGridViewButtonColumn();
            Column_InsName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            button_RemoveInstrument = new System.Windows.Forms.Button();
            button_AddInstrument = new System.Windows.Forms.Button();
            comboBox_InsTypes = new System.Windows.Forms.ComboBox();
            groupBox_InsCommand = new System.Windows.Forms.GroupBox();
            button_ClearInsLog = new System.Windows.Forms.Button();
            button_InsScreenCapture = new System.Windows.Forms.Button();
            richTextBox_InsCommandLog = new System.Windows.Forms.RichTextBox();
            button_SendInsCommand = new System.Windows.Forms.Button();
            textBox_InsCommand = new System.Windows.Forms.TextBox();
            textBox_InsType = new System.Windows.Forms.TextBox();
            tableLayoutPanel_Instrument.SuspendLayout();
            groupBox_InsSetting.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView_InsList).BeginInit();
            groupBox_InsCommand.SuspendLayout();
            SuspendLayout();
            tableLayoutPanel_Instrument.ColumnCount = 1;
            tableLayoutPanel_Instrument.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
            tableLayoutPanel_Instrument.Controls.Add(groupBox_InsSetting, 0, 0);
            tableLayoutPanel_Instrument.Controls.Add(groupBox_InsCommand, 0, 1);
            tableLayoutPanel_Instrument.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutPanel_Instrument.Location = new System.Drawing.Point(0, 0);
            tableLayoutPanel_Instrument.Name = "tableLayoutPanel_Instrument";
            tableLayoutPanel_Instrument.RowCount = 2;
            tableLayoutPanel_Instrument.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 240f));
            tableLayoutPanel_Instrument.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
            tableLayoutPanel_Instrument.Size = new System.Drawing.Size(828, 499);
            tableLayoutPanel_Instrument.TabIndex = 0;
            groupBox_InsSetting.Controls.Add(button_InsDown);
            groupBox_InsSetting.Controls.Add(button_InsUp);
            groupBox_InsSetting.Controls.Add(dataGridView_InsList);
            groupBox_InsSetting.Controls.Add(button_RemoveInstrument);
            groupBox_InsSetting.Controls.Add(button_AddInstrument);
            groupBox_InsSetting.Controls.Add(comboBox_InsTypes);
            groupBox_InsSetting.Dock = System.Windows.Forms.DockStyle.Fill;
            groupBox_InsSetting.Location = new System.Drawing.Point(3, 3);
            groupBox_InsSetting.Name = "groupBox_InsSetting";
            groupBox_InsSetting.Size = new System.Drawing.Size(822, 234);
            groupBox_InsSetting.TabIndex = 0;
            groupBox_InsSetting.TabStop = false;
            groupBox_InsSetting.Text = "Instrument Setting";
            button_InsDown.Location = new System.Drawing.Point(746, 20);
            button_InsDown.Name = "button_InsDown";
            button_InsDown.Size = new System.Drawing.Size(70, 25);
            button_InsDown.TabIndex = 5;
            button_InsDown.Text = "Down";
            button_InsDown.UseVisualStyleBackColor = true;
            button_InsDown.Click += new System.EventHandler(button_InsDown_Click);
            button_InsUp.Location = new System.Drawing.Point(670, 20);
            button_InsUp.Name = "button_InsUp";
            button_InsUp.Size = new System.Drawing.Size(70, 25);
            button_InsUp.TabIndex = 4;
            button_InsUp.Text = "Up";
            button_InsUp.UseVisualStyleBackColor = true;
            button_InsUp.Click += new System.EventHandler(button_InsUp_Click);
            dataGridView_InsList.AllowUserToAddRows = false;
            dataGridView_InsList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView_InsList.Columns.AddRange(Column_InsType, Column_InsEnabled, Column_Address, Column_InsTest, Column_InsName);
            dataGridView_InsList.Location = new System.Drawing.Point(6, 49);
            dataGridView_InsList.Name = "dataGridView_InsList";
            dataGridView_InsList.RowHeadersVisible = false;
            dataGridView_InsList.RowHeadersWidth = 50;
            dataGridViewCellStyle.Font = new System.Drawing.Font("굴림", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 129);
            dataGridView_InsList.RowsDefaultCellStyle = dataGridViewCellStyle;
            dataGridView_InsList.RowTemplate.Height = 20;
            dataGridView_InsList.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            dataGridView_InsList.Size = new System.Drawing.Size(810, 180);
            dataGridView_InsList.TabIndex = 3;
            dataGridView_InsList.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView_InsList_CellClick);
            dataGridView_InsList.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView_InsList_CellContentClick);
            dataGridView_InsList.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView_InsList_CellValueChanged);
            Column_InsType.HeaderText = "Type";
            Column_InsType.MinimumWidth = 10;
            Column_InsType.Name = "Column_InsType";
            Column_InsType.ReadOnly = true;
            Column_InsType.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            Column_InsType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            Column_InsType.Width = 125;
            Column_InsEnabled.HeaderText = "En";
            Column_InsEnabled.MinimumWidth = 10;
            Column_InsEnabled.Name = "Column_InsEnabled";
            Column_InsEnabled.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            Column_InsEnabled.Width = 30;
            Column_Address.HeaderText = "VISA Address";
            Column_Address.MinimumWidth = 125;
            Column_Address.Name = "Column_Address";
            Column_Address.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            Column_Address.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            Column_Address.Width = 250;
            Column_InsTest.HeaderText = "Test";
            Column_InsTest.MinimumWidth = 10;
            Column_InsTest.Name = "Column_InsTest";
            Column_InsTest.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            Column_InsTest.Text = "";
            Column_InsTest.Width = 50;
            Column_InsName.HeaderText = "Name";
            Column_InsName.MinimumWidth = 10;
            Column_InsName.Name = "Column_InsName";
            Column_InsName.ReadOnly = true;
            Column_InsName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            Column_InsName.Width = 325;
            button_RemoveInstrument.Location = new System.Drawing.Point(386, 20);
            button_RemoveInstrument.Name = "button_RemoveInstrument";
            button_RemoveInstrument.Size = new System.Drawing.Size(100, 25);
            button_RemoveInstrument.TabIndex = 2;
            button_RemoveInstrument.Text = "Remove";
            button_RemoveInstrument.UseVisualStyleBackColor = true;
            button_RemoveInstrument.Click += new System.EventHandler(button_RemoveInstrument_Click);
            button_AddInstrument.Location = new System.Drawing.Point(280, 20);
            button_AddInstrument.Name = "button_AddInstrument";
            button_AddInstrument.Size = new System.Drawing.Size(100, 25);
            button_AddInstrument.TabIndex = 1;
            button_AddInstrument.Text = "Add";
            button_AddInstrument.UseVisualStyleBackColor = true;
            button_AddInstrument.Click += new System.EventHandler(button_AddInstrument_Click);
            comboBox_InsTypes.FormattingEnabled = true;
            comboBox_InsTypes.Location = new System.Drawing.Point(6, 20);
            comboBox_InsTypes.Name = "comboBox_InsTypes";
            comboBox_InsTypes.Size = new System.Drawing.Size(250, 23);
            comboBox_InsTypes.TabIndex = 0;
            groupBox_InsCommand.Controls.Add(button_ClearInsLog);
            groupBox_InsCommand.Controls.Add(button_InsScreenCapture);
            groupBox_InsCommand.Controls.Add(richTextBox_InsCommandLog);
            groupBox_InsCommand.Controls.Add(button_SendInsCommand);
            groupBox_InsCommand.Controls.Add(textBox_InsCommand);
            groupBox_InsCommand.Controls.Add(textBox_InsType);
            groupBox_InsCommand.Dock = System.Windows.Forms.DockStyle.Fill;
            groupBox_InsCommand.Location = new System.Drawing.Point(3, 243);
            groupBox_InsCommand.Name = "groupBox_InsCommand";
            groupBox_InsCommand.Size = new System.Drawing.Size(822, 253);
            groupBox_InsCommand.TabIndex = 1;
            groupBox_InsCommand.TabStop = false;
            groupBox_InsCommand.Text = "Instrument Command";
            button_ClearInsLog.Location = new System.Drawing.Point(696, 217);
            button_ClearInsLog.Name = "button_ClearInsLog";
            button_ClearInsLog.Size = new System.Drawing.Size(120, 30);
            button_ClearInsLog.TabIndex = 5;
            button_ClearInsLog.Text = "Clear Log";
            button_ClearInsLog.UseVisualStyleBackColor = true;
            button_ClearInsLog.Click += new System.EventHandler(button_ClearInsLog_Click);
            button_InsScreenCapture.Location = new System.Drawing.Point(696, 51);
            button_InsScreenCapture.Name = "button_InsScreenCapture";
            button_InsScreenCapture.Size = new System.Drawing.Size(120, 30);
            button_InsScreenCapture.TabIndex = 4;
            button_InsScreenCapture.Text = "Capture";
            button_InsScreenCapture.UseVisualStyleBackColor = true;
            button_InsScreenCapture.Click += new System.EventHandler(button_InsScreenCapture_Click);
            richTextBox_InsCommandLog.Location = new System.Drawing.Point(6, 51);
            richTextBox_InsCommandLog.Name = "richTextBox_InsCommandLog";
            richTextBox_InsCommandLog.Size = new System.Drawing.Size(684, 196);
            richTextBox_InsCommandLog.TabIndex = 3;
            richTextBox_InsCommandLog.Text = "";
            button_SendInsCommand.Location = new System.Drawing.Point(165, 20);
            button_SendInsCommand.Name = "button_SendInsCommand";
            button_SendInsCommand.Size = new System.Drawing.Size(140, 25);
            button_SendInsCommand.TabIndex = 2;
            button_SendInsCommand.Text = "Send Command";
            button_SendInsCommand.UseVisualStyleBackColor = true;
            button_SendInsCommand.Click += new System.EventHandler(button_SendInsCommand_Click);
            textBox_InsCommand.Location = new System.Drawing.Point(311, 20);
            textBox_InsCommand.Name = "textBox_InsCommand";
            textBox_InsCommand.Size = new System.Drawing.Size(505, 25);
            textBox_InsCommand.TabIndex = 1;
            textBox_InsType.Location = new System.Drawing.Point(9, 20);
            textBox_InsType.Name = "textBox_InsType";
            textBox_InsType.ReadOnly = true;
            textBox_InsType.Size = new System.Drawing.Size(150, 25);
            textBox_InsType.TabIndex = 0;
            base.AutoScaleDimensions = new System.Drawing.SizeF(120f, 120f);
            base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            base.ClientSize = new System.Drawing.Size(828, 499);
            base.Controls.Add(tableLayoutPanel_Instrument);
            base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            base.Name = "InstrumentForm";
            Text = "Instrument Setup";
            base.FormClosing += new System.Windows.Forms.FormClosingEventHandler(InstrumentForm_FormClosing);
            base.FormClosed += new System.Windows.Forms.FormClosedEventHandler(InstrumentForm_FormClosed);
            base.Load += new System.EventHandler(InstrumentForm_Load);
            tableLayoutPanel_Instrument.ResumeLayout(false);
            groupBox_InsSetting.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView_InsList).EndInit();
            groupBox_InsCommand.ResumeLayout(false);
            groupBox_InsCommand.PerformLayout();
            ResumeLayout(false);
        }
    }
}
