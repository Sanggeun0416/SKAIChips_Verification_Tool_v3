
namespace SKAIChips_Verification
{
    partial class RegContForm
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
            this.tableLayoutPanel_Base = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel_Ctrl = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox_Comn = new System.Windows.Forms.GroupBox();
            this.button_ComnSetup = new System.Windows.Forms.Button();
            this.button_Connection = new System.Windows.Forms.Button();
            this.label_ComnStatus = new System.Windows.Forms.Label();
            this.comboBox_ComnTypes = new System.Windows.Forms.ComboBox();
            this.groupBox_RegMapControl = new System.Windows.Forms.GroupBox();
            this.button_ClearRegTree = new System.Windows.Forms.Button();
            this.button_AddRegTree = new System.Windows.Forms.Button();
            this.button_OpenMapFile = new System.Windows.Forms.Button();
            this.comboBox_SelectMapSheet = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox_SelectMapFIle = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox_RegControl = new System.Windows.Forms.GroupBox();
            this.groupBox_RegScript = new System.Windows.Forms.GroupBox();
            this.button_ExportScript = new System.Windows.Forms.Button();
            this.button_ImportScript = new System.Windows.Forms.Button();
            this.button_LoadScript = new System.Windows.Forms.Button();
            this.label_RegScriptFileName = new System.Windows.Forms.Label();
            this.groupBox_ChipSpecific = new System.Windows.Forms.GroupBox();
            this.button_RunChipSpecTest = new System.Windows.Forms.Button();
            this.textBox_TestArgument = new System.Windows.Forms.TextBox();
            this.comboBox_TestItems = new System.Windows.Forms.ComboBox();
            this.panel_RegLog = new System.Windows.Forms.Panel();
            this.button_OpenRegLog = new System.Windows.Forms.Button();
            this.button_RunRegLog = new System.Windows.Forms.Button();
            this.button_SaveRegLog = new System.Windows.Forms.Button();
            this.button_ClearRegLog = new System.Windows.Forms.Button();
            this.dataGridView_RegLog = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel_RegMap = new System.Windows.Forms.TableLayoutPanel();
            this.label_RegMapName = new System.Windows.Forms.Label();
            this.splitContainer_RegMap = new System.Windows.Forms.SplitContainer();
            this.treeView_RegMap = new System.Windows.Forms.TreeView();
            this.contextMenu_RegTree = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem_SearchReg = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel_LogCtrl = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel_RegCommand = new System.Windows.Forms.TableLayoutPanel();
            this.textBox_RegCommand = new System.Windows.Forms.TextBox();
            this.button_SendRegCommand = new System.Windows.Forms.Button();
            this.button_RunLog = new System.Windows.Forms.Button();
            this.richTextBox_LogCtrl = new System.Windows.Forms.RichTextBox();
            this.tableLayoutPanel_Base.SuspendLayout();
            this.tableLayoutPanel_Ctrl.SuspendLayout();
            this.groupBox_Comn.SuspendLayout();
            this.groupBox_RegMapControl.SuspendLayout();
            this.groupBox_RegScript.SuspendLayout();
            this.groupBox_ChipSpecific.SuspendLayout();
            this.panel_RegLog.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_RegLog)).BeginInit();
            this.tableLayoutPanel_RegMap.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_RegMap)).BeginInit();
            this.splitContainer_RegMap.Panel1.SuspendLayout();
            this.splitContainer_RegMap.Panel2.SuspendLayout();
            this.splitContainer_RegMap.SuspendLayout();
            this.contextMenu_RegTree.SuspendLayout();
            this.tableLayoutPanel_LogCtrl.SuspendLayout();
            this.tableLayoutPanel_RegCommand.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel_Base
            // 
            this.tableLayoutPanel_Base.ColumnCount = 2;
            this.tableLayoutPanel_Base.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Base.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 300F));
            this.tableLayoutPanel_Base.Controls.Add(this.tableLayoutPanel_Ctrl, 1, 0);
            this.tableLayoutPanel_Base.Controls.Add(this.tableLayoutPanel_RegMap, 0, 0);
            this.tableLayoutPanel_Base.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_Base.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel_Base.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel_Base.Name = "tableLayoutPanel_Base";
            this.tableLayoutPanel_Base.RowCount = 1;
            this.tableLayoutPanel_Base.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Base.Size = new System.Drawing.Size(626, 522);
            this.tableLayoutPanel_Base.TabIndex = 0;
            // 
            // tableLayoutPanel_Ctrl
            // 
            this.tableLayoutPanel_Ctrl.ColumnCount = 1;
            this.tableLayoutPanel_Ctrl.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Ctrl.Controls.Add(this.groupBox_Comn, 0, 0);
            this.tableLayoutPanel_Ctrl.Controls.Add(this.groupBox_RegMapControl, 0, 1);
            this.tableLayoutPanel_Ctrl.Controls.Add(this.groupBox_RegControl, 0, 2);
            this.tableLayoutPanel_Ctrl.Controls.Add(this.groupBox_RegScript, 0, 3);
            this.tableLayoutPanel_Ctrl.Controls.Add(this.groupBox_ChipSpecific, 0, 5);
            this.tableLayoutPanel_Ctrl.Controls.Add(this.panel_RegLog, 0, 4);
            this.tableLayoutPanel_Ctrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_Ctrl.Location = new System.Drawing.Point(328, 2);
            this.tableLayoutPanel_Ctrl.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel_Ctrl.Name = "tableLayoutPanel_Ctrl";
            this.tableLayoutPanel_Ctrl.RowCount = 6;
            this.tableLayoutPanel_Ctrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64F));
            this.tableLayoutPanel_Ctrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 96F));
            this.tableLayoutPanel_Ctrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 116F));
            this.tableLayoutPanel_Ctrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64F));
            this.tableLayoutPanel_Ctrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Ctrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 116F));
            this.tableLayoutPanel_Ctrl.Size = new System.Drawing.Size(296, 518);
            this.tableLayoutPanel_Ctrl.TabIndex = 0;
            // 
            // groupBox_Comn
            // 
            this.groupBox_Comn.Controls.Add(this.button_ComnSetup);
            this.groupBox_Comn.Controls.Add(this.button_Connection);
            this.groupBox_Comn.Controls.Add(this.label_ComnStatus);
            this.groupBox_Comn.Controls.Add(this.comboBox_ComnTypes);
            this.groupBox_Comn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_Comn.Location = new System.Drawing.Point(2, 3);
            this.groupBox_Comn.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.groupBox_Comn.Name = "groupBox_Comn";
            this.groupBox_Comn.Padding = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.groupBox_Comn.Size = new System.Drawing.Size(292, 58);
            this.groupBox_Comn.TabIndex = 0;
            this.groupBox_Comn.TabStop = false;
            this.groupBox_Comn.Text = "Communication";
            // 
            // button_ComnSetup
            // 
            this.button_ComnSetup.Location = new System.Drawing.Point(230, 14);
            this.button_ComnSetup.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button_ComnSetup.Name = "button_ComnSetup";
            this.button_ComnSetup.Size = new System.Drawing.Size(56, 24);
            this.button_ComnSetup.TabIndex = 3;
            this.button_ComnSetup.Text = "Setup";
            this.button_ComnSetup.UseVisualStyleBackColor = true;
            this.button_ComnSetup.Click += new System.EventHandler(this.button_ComnSetup_Click);
            // 
            // button_Connection
            // 
            this.button_Connection.Location = new System.Drawing.Point(129, 14);
            this.button_Connection.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button_Connection.Name = "button_Connection";
            this.button_Connection.Size = new System.Drawing.Size(96, 24);
            this.button_Connection.TabIndex = 2;
            this.button_Connection.Text = "Connect";
            this.button_Connection.UseVisualStyleBackColor = true;
            this.button_Connection.Click += new System.EventHandler(this.button_Connection_Click);
            // 
            // label_ComnStatus
            // 
            this.label_ComnStatus.AutoSize = true;
            this.label_ComnStatus.Location = new System.Drawing.Point(5, 40);
            this.label_ComnStatus.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label_ComnStatus.Name = "label_ComnStatus";
            this.label_ComnStatus.Size = new System.Drawing.Size(198, 12);
            this.label_ComnStatus.TabIndex = 1;
            this.label_ComnStatus.Text = "Communication status information";
            // 
            // comboBox_ComnTypes
            // 
            this.comboBox_ComnTypes.FormattingEnabled = true;
            this.comboBox_ComnTypes.Location = new System.Drawing.Point(5, 16);
            this.comboBox_ComnTypes.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.comboBox_ComnTypes.Name = "comboBox_ComnTypes";
            this.comboBox_ComnTypes.Size = new System.Drawing.Size(97, 20);
            this.comboBox_ComnTypes.TabIndex = 0;
            this.comboBox_ComnTypes.SelectedIndexChanged += new System.EventHandler(this.comboBox_ComnTypes_SelectedIndexChanged);
            // 
            // groupBox_RegMapControl
            // 
            this.groupBox_RegMapControl.Controls.Add(this.button_ClearRegTree);
            this.groupBox_RegMapControl.Controls.Add(this.button_AddRegTree);
            this.groupBox_RegMapControl.Controls.Add(this.button_OpenMapFile);
            this.groupBox_RegMapControl.Controls.Add(this.comboBox_SelectMapSheet);
            this.groupBox_RegMapControl.Controls.Add(this.label2);
            this.groupBox_RegMapControl.Controls.Add(this.comboBox_SelectMapFIle);
            this.groupBox_RegMapControl.Controls.Add(this.label1);
            this.groupBox_RegMapControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_RegMapControl.Location = new System.Drawing.Point(2, 66);
            this.groupBox_RegMapControl.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox_RegMapControl.Name = "groupBox_RegMapControl";
            this.groupBox_RegMapControl.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox_RegMapControl.Size = new System.Drawing.Size(292, 92);
            this.groupBox_RegMapControl.TabIndex = 1;
            this.groupBox_RegMapControl.TabStop = false;
            this.groupBox_RegMapControl.Text = "Register Map File Control";
            // 
            // button_ClearRegTree
            // 
            this.button_ClearRegTree.Location = new System.Drawing.Point(114, 62);
            this.button_ClearRegTree.Margin = new System.Windows.Forms.Padding(2);
            this.button_ClearRegTree.Name = "button_ClearRegTree";
            this.button_ClearRegTree.Size = new System.Drawing.Size(104, 24);
            this.button_ClearRegTree.TabIndex = 6;
            this.button_ClearRegTree.Text = "Clear Reg tree";
            this.button_ClearRegTree.UseVisualStyleBackColor = true;
            this.button_ClearRegTree.Click += new System.EventHandler(this.button_ClearRegTree_Click);
            // 
            // button_AddRegTree
            // 
            this.button_AddRegTree.Location = new System.Drawing.Point(5, 62);
            this.button_AddRegTree.Margin = new System.Windows.Forms.Padding(2);
            this.button_AddRegTree.Name = "button_AddRegTree";
            this.button_AddRegTree.Size = new System.Drawing.Size(104, 24);
            this.button_AddRegTree.TabIndex = 5;
            this.button_AddRegTree.Text = "Add to Reg tree";
            this.button_AddRegTree.UseVisualStyleBackColor = true;
            this.button_AddRegTree.Click += new System.EventHandler(this.button_AddRegTree_Click);
            // 
            // button_OpenMapFile
            // 
            this.button_OpenMapFile.Location = new System.Drawing.Point(230, 62);
            this.button_OpenMapFile.Margin = new System.Windows.Forms.Padding(2);
            this.button_OpenMapFile.Name = "button_OpenMapFile";
            this.button_OpenMapFile.Size = new System.Drawing.Size(56, 24);
            this.button_OpenMapFile.TabIndex = 4;
            this.button_OpenMapFile.Text = "Open";
            this.button_OpenMapFile.UseVisualStyleBackColor = true;
            this.button_OpenMapFile.Click += new System.EventHandler(this.button_OpenMapFile_Click);
            // 
            // comboBox_SelectMapSheet
            // 
            this.comboBox_SelectMapSheet.FormattingEnabled = true;
            this.comboBox_SelectMapSheet.Location = new System.Drawing.Point(126, 39);
            this.comboBox_SelectMapSheet.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox_SelectMapSheet.Name = "comboBox_SelectMapSheet";
            this.comboBox_SelectMapSheet.Size = new System.Drawing.Size(161, 20);
            this.comboBox_SelectMapSheet.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 42);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "2. Select sheet:";
            // 
            // comboBox_SelectMapFIle
            // 
            this.comboBox_SelectMapFIle.FormattingEnabled = true;
            this.comboBox_SelectMapFIle.Location = new System.Drawing.Point(86, 16);
            this.comboBox_SelectMapFIle.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox_SelectMapFIle.Name = "comboBox_SelectMapFIle";
            this.comboBox_SelectMapFIle.Size = new System.Drawing.Size(201, 20);
            this.comboBox_SelectMapFIle.TabIndex = 1;
            this.comboBox_SelectMapFIle.DropDown += new System.EventHandler(this.comboBox_SelectMapFIle_DropDown);
            this.comboBox_SelectMapFIle.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectMapFIle_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 18);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "1. Select file:";
            // 
            // groupBox_RegControl
            // 
            this.groupBox_RegControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_RegControl.Location = new System.Drawing.Point(2, 162);
            this.groupBox_RegControl.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox_RegControl.Name = "groupBox_RegControl";
            this.groupBox_RegControl.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox_RegControl.Size = new System.Drawing.Size(292, 112);
            this.groupBox_RegControl.TabIndex = 2;
            this.groupBox_RegControl.TabStop = false;
            this.groupBox_RegControl.Text = "Register Control";
            // 
            // groupBox_RegScript
            // 
            this.groupBox_RegScript.Controls.Add(this.button_ExportScript);
            this.groupBox_RegScript.Controls.Add(this.button_ImportScript);
            this.groupBox_RegScript.Controls.Add(this.button_LoadScript);
            this.groupBox_RegScript.Controls.Add(this.label_RegScriptFileName);
            this.groupBox_RegScript.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_RegScript.Location = new System.Drawing.Point(2, 278);
            this.groupBox_RegScript.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox_RegScript.Name = "groupBox_RegScript";
            this.groupBox_RegScript.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox_RegScript.Size = new System.Drawing.Size(292, 60);
            this.groupBox_RegScript.TabIndex = 3;
            this.groupBox_RegScript.TabStop = false;
            this.groupBox_RegScript.Text = "Register Script";
            // 
            // button_ExportScript
            // 
            this.button_ExportScript.Location = new System.Drawing.Point(206, 14);
            this.button_ExportScript.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button_ExportScript.Name = "button_ExportScript";
            this.button_ExportScript.Size = new System.Drawing.Size(80, 24);
            this.button_ExportScript.TabIndex = 6;
            this.button_ExportScript.Text = "Export";
            this.button_ExportScript.UseVisualStyleBackColor = true;
            this.button_ExportScript.Click += new System.EventHandler(this.button_ExportScript_Click);
            // 
            // button_ImportScript
            // 
            this.button_ImportScript.Location = new System.Drawing.Point(121, 14);
            this.button_ImportScript.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button_ImportScript.Name = "button_ImportScript";
            this.button_ImportScript.Size = new System.Drawing.Size(80, 24);
            this.button_ImportScript.TabIndex = 5;
            this.button_ImportScript.Text = "Import";
            this.button_ImportScript.UseVisualStyleBackColor = true;
            this.button_ImportScript.Click += new System.EventHandler(this.button_ImportScript_Click);
            // 
            // button_LoadScript
            // 
            this.button_LoadScript.Location = new System.Drawing.Point(5, 14);
            this.button_LoadScript.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button_LoadScript.Name = "button_LoadScript";
            this.button_LoadScript.Size = new System.Drawing.Size(96, 24);
            this.button_LoadScript.TabIndex = 4;
            this.button_LoadScript.Text = "Load";
            this.button_LoadScript.UseVisualStyleBackColor = true;
            this.button_LoadScript.Click += new System.EventHandler(this.button_LoadScript_Click);
            // 
            // label_RegScriptFileName
            // 
            this.label_RegScriptFileName.AutoSize = true;
            this.label_RegScriptFileName.Location = new System.Drawing.Point(5, 40);
            this.label_RegScriptFileName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label_RegScriptFileName.Name = "label_RegScriptFileName";
            this.label_RegScriptFileName.Size = new System.Drawing.Size(142, 12);
            this.label_RegScriptFileName.TabIndex = 2;
            this.label_RegScriptFileName.Text = "Register script file name";
            // 
            // groupBox_ChipSpecific
            // 
            this.groupBox_ChipSpecific.Controls.Add(this.button_RunChipSpecTest);
            this.groupBox_ChipSpecific.Controls.Add(this.textBox_TestArgument);
            this.groupBox_ChipSpecific.Controls.Add(this.comboBox_TestItems);
            this.groupBox_ChipSpecific.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_ChipSpecific.Location = new System.Drawing.Point(2, 404);
            this.groupBox_ChipSpecific.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox_ChipSpecific.Name = "groupBox_ChipSpecific";
            this.groupBox_ChipSpecific.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox_ChipSpecific.Size = new System.Drawing.Size(292, 112);
            this.groupBox_ChipSpecific.TabIndex = 4;
            this.groupBox_ChipSpecific.TabStop = false;
            this.groupBox_ChipSpecific.Text = "Chip Specific Control";
            // 
            // button_RunChipSpecTest
            // 
            this.button_RunChipSpecTest.Location = new System.Drawing.Point(218, 14);
            this.button_RunChipSpecTest.Margin = new System.Windows.Forms.Padding(2);
            this.button_RunChipSpecTest.Name = "button_RunChipSpecTest";
            this.button_RunChipSpecTest.Size = new System.Drawing.Size(68, 21);
            this.button_RunChipSpecTest.TabIndex = 8;
            this.button_RunChipSpecTest.Text = "Run Test";
            this.button_RunChipSpecTest.UseVisualStyleBackColor = true;
            this.button_RunChipSpecTest.Click += new System.EventHandler(this.button_RunChipSpecTest_Click);
            // 
            // textBox_TestArgument
            // 
            this.textBox_TestArgument.Location = new System.Drawing.Point(130, 14);
            this.textBox_TestArgument.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_TestArgument.Name = "textBox_TestArgument";
            this.textBox_TestArgument.Size = new System.Drawing.Size(86, 21);
            this.textBox_TestArgument.TabIndex = 7;
            // 
            // comboBox_TestItems
            // 
            this.comboBox_TestItems.FormattingEnabled = true;
            this.comboBox_TestItems.Location = new System.Drawing.Point(5, 14);
            this.comboBox_TestItems.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox_TestItems.Name = "comboBox_TestItems";
            this.comboBox_TestItems.Size = new System.Drawing.Size(121, 20);
            this.comboBox_TestItems.TabIndex = 6;
            // 
            // panel_RegLog
            // 
            this.panel_RegLog.Controls.Add(this.button_OpenRegLog);
            this.panel_RegLog.Controls.Add(this.button_RunRegLog);
            this.panel_RegLog.Controls.Add(this.button_SaveRegLog);
            this.panel_RegLog.Controls.Add(this.button_ClearRegLog);
            this.panel_RegLog.Controls.Add(this.dataGridView_RegLog);
            this.panel_RegLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel_RegLog.Location = new System.Drawing.Point(0, 340);
            this.panel_RegLog.Margin = new System.Windows.Forms.Padding(0);
            this.panel_RegLog.Name = "panel_RegLog";
            this.panel_RegLog.Size = new System.Drawing.Size(296, 62);
            this.panel_RegLog.TabIndex = 5;
            // 
            // button_OpenRegLog
            // 
            this.button_OpenRegLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_OpenRegLog.Location = new System.Drawing.Point(158, 40);
            this.button_OpenRegLog.Margin = new System.Windows.Forms.Padding(2);
            this.button_OpenRegLog.Name = "button_OpenRegLog";
            this.button_OpenRegLog.Size = new System.Drawing.Size(64, 20);
            this.button_OpenRegLog.TabIndex = 4;
            this.button_OpenRegLog.Text = "Open";
            this.button_OpenRegLog.UseVisualStyleBackColor = true;
            this.button_OpenRegLog.Click += new System.EventHandler(this.button_OpenRegLog_Click);
            // 
            // button_RunRegLog
            // 
            this.button_RunRegLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_RunRegLog.Location = new System.Drawing.Point(81, 40);
            this.button_RunRegLog.Margin = new System.Windows.Forms.Padding(2);
            this.button_RunRegLog.Name = "button_RunRegLog";
            this.button_RunRegLog.Size = new System.Drawing.Size(72, 20);
            this.button_RunRegLog.TabIndex = 3;
            this.button_RunRegLog.Text = "Run Reg";
            this.button_RunRegLog.UseVisualStyleBackColor = true;
            this.button_RunRegLog.Click += new System.EventHandler(this.button_RunRegLog_Click);
            // 
            // button_SaveRegLog
            // 
            this.button_SaveRegLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_SaveRegLog.Location = new System.Drawing.Point(227, 40);
            this.button_SaveRegLog.Margin = new System.Windows.Forms.Padding(2);
            this.button_SaveRegLog.Name = "button_SaveRegLog";
            this.button_SaveRegLog.Size = new System.Drawing.Size(64, 20);
            this.button_SaveRegLog.TabIndex = 2;
            this.button_SaveRegLog.Text = "Save";
            this.button_SaveRegLog.UseVisualStyleBackColor = true;
            this.button_SaveRegLog.Click += new System.EventHandler(this.button_SaveRegLog_Click);
            // 
            // button_ClearRegLog
            // 
            this.button_ClearRegLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_ClearRegLog.Location = new System.Drawing.Point(4, 40);
            this.button_ClearRegLog.Margin = new System.Windows.Forms.Padding(2);
            this.button_ClearRegLog.Name = "button_ClearRegLog";
            this.button_ClearRegLog.Size = new System.Drawing.Size(72, 20);
            this.button_ClearRegLog.TabIndex = 1;
            this.button_ClearRegLog.Text = "Clear";
            this.button_ClearRegLog.UseVisualStyleBackColor = true;
            this.button_ClearRegLog.Click += new System.EventHandler(this.button_ClearRegLog_Click);
            // 
            // dataGridView_RegLog
            // 
            this.dataGridView_RegLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.dataGridView_RegLog.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_RegLog.Location = new System.Drawing.Point(3, 4);
            this.dataGridView_RegLog.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView_RegLog.Name = "dataGridView_RegLog";
            this.dataGridView_RegLog.RowHeadersWidth = 51;
            this.dataGridView_RegLog.RowTemplate.Height = 27;
            this.dataGridView_RegLog.Size = new System.Drawing.Size(290, 31);
            this.dataGridView_RegLog.TabIndex = 0;
            // 
            // tableLayoutPanel_RegMap
            // 
            this.tableLayoutPanel_RegMap.ColumnCount = 1;
            this.tableLayoutPanel_RegMap.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_RegMap.Controls.Add(this.label_RegMapName, 0, 0);
            this.tableLayoutPanel_RegMap.Controls.Add(this.splitContainer_RegMap, 0, 1);
            this.tableLayoutPanel_RegMap.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_RegMap.Location = new System.Drawing.Point(2, 2);
            this.tableLayoutPanel_RegMap.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel_RegMap.Name = "tableLayoutPanel_RegMap";
            this.tableLayoutPanel_RegMap.RowCount = 2;
            this.tableLayoutPanel_RegMap.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 16F));
            this.tableLayoutPanel_RegMap.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_RegMap.Size = new System.Drawing.Size(322, 518);
            this.tableLayoutPanel_RegMap.TabIndex = 1;
            // 
            // label_RegMapName
            // 
            this.label_RegMapName.AutoSize = true;
            this.label_RegMapName.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label_RegMapName.ForeColor = System.Drawing.Color.DarkOrange;
            this.label_RegMapName.Location = new System.Drawing.Point(2, 2);
            this.label_RegMapName.Margin = new System.Windows.Forms.Padding(2);
            this.label_RegMapName.Name = "label_RegMapName";
            this.label_RegMapName.Size = new System.Drawing.Size(102, 12);
            this.label_RegMapName.TabIndex = 0;
            this.label_RegMapName.Text = "Register Name";
            // 
            // splitContainer_RegMap
            // 
            this.splitContainer_RegMap.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer_RegMap.Location = new System.Drawing.Point(2, 18);
            this.splitContainer_RegMap.Margin = new System.Windows.Forms.Padding(2);
            this.splitContainer_RegMap.Name = "splitContainer_RegMap";
            this.splitContainer_RegMap.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer_RegMap.Panel1
            // 
            this.splitContainer_RegMap.Panel1.Controls.Add(this.treeView_RegMap);
            // 
            // splitContainer_RegMap.Panel2
            // 
            this.splitContainer_RegMap.Panel2.Controls.Add(this.tableLayoutPanel_LogCtrl);
            this.splitContainer_RegMap.Size = new System.Drawing.Size(318, 498);
            this.splitContainer_RegMap.SplitterDistance = 305;
            this.splitContainer_RegMap.TabIndex = 1;
            // 
            // treeView_RegMap
            // 
            this.treeView_RegMap.ContextMenuStrip = this.contextMenu_RegTree;
            this.treeView_RegMap.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView_RegMap.LineColor = System.Drawing.Color.BlanchedAlmond;
            this.treeView_RegMap.Location = new System.Drawing.Point(0, 0);
            this.treeView_RegMap.Margin = new System.Windows.Forms.Padding(2);
            this.treeView_RegMap.Name = "treeView_RegMap";
            this.treeView_RegMap.Size = new System.Drawing.Size(318, 305);
            this.treeView_RegMap.TabIndex = 0;
            // 
            // contextMenu_RegTree
            // 
            this.contextMenu_RegTree.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem_SearchReg});
            this.contextMenu_RegTree.Name = "contextMenu_RegTree";
            this.contextMenu_RegTree.Size = new System.Drawing.Size(111, 26);
            // 
            // toolStripMenuItem_SearchReg
            // 
            this.toolStripMenuItem_SearchReg.Name = "toolStripMenuItem_SearchReg";
            this.toolStripMenuItem_SearchReg.Size = new System.Drawing.Size(110, 22);
            this.toolStripMenuItem_SearchReg.Text = "Search";
            this.toolStripMenuItem_SearchReg.Click += new System.EventHandler(this.toolStripMenuItem_SearchReg_Click);
            // 
            // tableLayoutPanel_LogCtrl
            // 
            this.tableLayoutPanel_LogCtrl.ColumnCount = 1;
            this.tableLayoutPanel_LogCtrl.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_LogCtrl.Controls.Add(this.tableLayoutPanel_RegCommand, 0, 0);
            this.tableLayoutPanel_LogCtrl.Controls.Add(this.richTextBox_LogCtrl, 0, 1);
            this.tableLayoutPanel_LogCtrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_LogCtrl.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel_LogCtrl.Name = "tableLayoutPanel_LogCtrl";
            this.tableLayoutPanel_LogCtrl.RowCount = 2;
            this.tableLayoutPanel_LogCtrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel_LogCtrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_LogCtrl.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 16F));
            this.tableLayoutPanel_LogCtrl.Size = new System.Drawing.Size(318, 189);
            this.tableLayoutPanel_LogCtrl.TabIndex = 0;
            // 
            // tableLayoutPanel_RegCommand
            // 
            this.tableLayoutPanel_RegCommand.ColumnCount = 3;
            this.tableLayoutPanel_RegCommand.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_RegCommand.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120F));
            this.tableLayoutPanel_RegCommand.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.tableLayoutPanel_RegCommand.Controls.Add(this.textBox_RegCommand, 0, 0);
            this.tableLayoutPanel_RegCommand.Controls.Add(this.button_SendRegCommand, 1, 0);
            this.tableLayoutPanel_RegCommand.Controls.Add(this.button_RunLog, 2, 0);
            this.tableLayoutPanel_RegCommand.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_RegCommand.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel_RegCommand.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel_RegCommand.Name = "tableLayoutPanel_RegCommand";
            this.tableLayoutPanel_RegCommand.RowCount = 1;
            this.tableLayoutPanel_RegCommand.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_RegCommand.Size = new System.Drawing.Size(318, 20);
            this.tableLayoutPanel_RegCommand.TabIndex = 0;
            // 
            // textBox_RegCommand
            // 
            this.textBox_RegCommand.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_RegCommand.Location = new System.Drawing.Point(0, 0);
            this.textBox_RegCommand.Margin = new System.Windows.Forms.Padding(0, 0, 2, 0);
            this.textBox_RegCommand.Name = "textBox_RegCommand";
            this.textBox_RegCommand.Size = new System.Drawing.Size(116, 21);
            this.textBox_RegCommand.TabIndex = 0;
            // 
            // button_SendRegCommand
            // 
            this.button_SendRegCommand.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_SendRegCommand.Location = new System.Drawing.Point(120, 0);
            this.button_SendRegCommand.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.button_SendRegCommand.Name = "button_SendRegCommand";
            this.button_SendRegCommand.Size = new System.Drawing.Size(116, 20);
            this.button_SendRegCommand.TabIndex = 1;
            this.button_SendRegCommand.Text = "Send Command";
            this.button_SendRegCommand.UseVisualStyleBackColor = true;
            this.button_SendRegCommand.Click += new System.EventHandler(this.button_SendRegCommand_Click);
            // 
            // button_RunLog
            // 
            this.button_RunLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_RunLog.Location = new System.Drawing.Point(240, 0);
            this.button_RunLog.Margin = new System.Windows.Forms.Padding(2, 0, 0, 0);
            this.button_RunLog.Name = "button_RunLog";
            this.button_RunLog.Size = new System.Drawing.Size(78, 20);
            this.button_RunLog.TabIndex = 2;
            this.button_RunLog.Text = "Run Log";
            this.button_RunLog.UseVisualStyleBackColor = true;
            this.button_RunLog.Click += new System.EventHandler(this.button_RunLog_Click);
            // 
            // richTextBox_LogCtrl
            // 
            this.richTextBox_LogCtrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox_LogCtrl.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox_LogCtrl.Location = new System.Drawing.Point(0, 22);
            this.richTextBox_LogCtrl.Margin = new System.Windows.Forms.Padding(0, 2, 0, 0);
            this.richTextBox_LogCtrl.Name = "richTextBox_LogCtrl";
            this.richTextBox_LogCtrl.Size = new System.Drawing.Size(318, 167);
            this.richTextBox_LogCtrl.TabIndex = 1;
            this.richTextBox_LogCtrl.Text = "";
            // 
            // RegContForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(626, 522);
            this.Controls.Add(this.tableLayoutPanel_Base);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "RegContForm";
            this.Text = "RegContForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.RegContForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.RegContForm_FormClosed);
            this.Load += new System.EventHandler(this.RegContForm_Load);
            this.tableLayoutPanel_Base.ResumeLayout(false);
            this.tableLayoutPanel_Ctrl.ResumeLayout(false);
            this.groupBox_Comn.ResumeLayout(false);
            this.groupBox_Comn.PerformLayout();
            this.groupBox_RegMapControl.ResumeLayout(false);
            this.groupBox_RegMapControl.PerformLayout();
            this.groupBox_RegScript.ResumeLayout(false);
            this.groupBox_RegScript.PerformLayout();
            this.groupBox_ChipSpecific.ResumeLayout(false);
            this.groupBox_ChipSpecific.PerformLayout();
            this.panel_RegLog.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_RegLog)).EndInit();
            this.tableLayoutPanel_RegMap.ResumeLayout(false);
            this.tableLayoutPanel_RegMap.PerformLayout();
            this.splitContainer_RegMap.Panel1.ResumeLayout(false);
            this.splitContainer_RegMap.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_RegMap)).EndInit();
            this.splitContainer_RegMap.ResumeLayout(false);
            this.contextMenu_RegTree.ResumeLayout(false);
            this.tableLayoutPanel_LogCtrl.ResumeLayout(false);
            this.tableLayoutPanel_RegCommand.ResumeLayout(false);
            this.tableLayoutPanel_RegCommand.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_Base;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_Ctrl;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_RegMap;
        private System.Windows.Forms.Label label_RegMapName;
        private System.Windows.Forms.GroupBox groupBox_Comn;
        private System.Windows.Forms.Label label_ComnStatus;
        private System.Windows.Forms.ComboBox comboBox_ComnTypes;
        private System.Windows.Forms.Button button_ComnSetup;
        private System.Windows.Forms.Button button_Connection;
        private System.Windows.Forms.SplitContainer splitContainer_RegMap;
        private System.Windows.Forms.TreeView treeView_RegMap;
        private System.Windows.Forms.GroupBox groupBox_RegMapControl;
        private System.Windows.Forms.ComboBox comboBox_SelectMapSheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox_SelectMapFIle;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_OpenMapFile;
        private System.Windows.Forms.Button button_ClearRegTree;
        private System.Windows.Forms.Button button_AddRegTree;
        private System.Windows.Forms.GroupBox groupBox_RegControl;
        private System.Windows.Forms.GroupBox groupBox_RegScript;
        private System.Windows.Forms.Button button_LoadScript;
        private System.Windows.Forms.Label label_RegScriptFileName;
        private System.Windows.Forms.Button button_ImportScript;
        private System.Windows.Forms.Button button_ExportScript;
        private System.Windows.Forms.GroupBox groupBox_ChipSpecific;
        private System.Windows.Forms.Panel panel_RegLog;
        private System.Windows.Forms.DataGridView dataGridView_RegLog;
        private System.Windows.Forms.Button button_SaveRegLog;
        private System.Windows.Forms.Button button_ClearRegLog;
        private System.Windows.Forms.Button button_OpenRegLog;
        private System.Windows.Forms.Button button_RunRegLog;
        private System.Windows.Forms.Button button_RunChipSpecTest;
        private System.Windows.Forms.TextBox textBox_TestArgument;
        private System.Windows.Forms.ComboBox comboBox_TestItems;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_LogCtrl;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_RegCommand;
        private System.Windows.Forms.TextBox textBox_RegCommand;
        private System.Windows.Forms.Button button_SendRegCommand;
        private System.Windows.Forms.RichTextBox richTextBox_LogCtrl;
        private System.Windows.Forms.Button button_RunLog;
        private System.Windows.Forms.ContextMenuStrip contextMenu_RegTree;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem_SearchReg;
    }
}