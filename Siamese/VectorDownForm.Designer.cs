namespace SKAIChips_Verification
{
    partial class VectorDownForm
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.tableLayoutPanel_Base = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel_View = new System.Windows.Forms.TableLayoutPanel();
            this.chart_Spectrum = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.chart_IQ_TimeDomain = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.tableLayoutPanel_Control = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox_VectorSetting = new System.Windows.Forms.GroupBox();
            this.textBox_ToneFreq2_kHz = new System.Windows.Forms.TextBox();
            this.checkBox_Tone2 = new System.Windows.Forms.CheckBox();
            this.checkBox_Tone1 = new System.Windows.Forms.CheckBox();
            this.button_ExportVectorFile = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_VectorLength = new System.Windows.Forms.TextBox();
            this.button_GenVector = new System.Windows.Forms.Button();
            this.textBox_BitRate_bps = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_BitPattern = new System.Windows.Forms.TextBox();
            this.textBox_ToneFreq1_kHz = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_FFT_Size = new System.Windows.Forms.TextBox();
            this.textBox_SamplingFreq_kHz = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox_Download = new System.Windows.Forms.GroupBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox_ALC_Start = new System.Windows.Forms.TextBox();
            this.textBox_ALC_Stop = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.textBox_VectorPeriod = new System.Windows.Forms.TextBox();
            this.textBox_DownloadSamplingFreq_kHz = new System.Windows.Forms.TextBox();
            this.button_ImportVectorFile = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBox_SelectSigGen = new System.Windows.Forms.ComboBox();
            this.button_Download = new System.Windows.Forms.Button();
            this.tableLayoutPanel_Base.SuspendLayout();
            this.tableLayoutPanel_View.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart_Spectrum)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart_IQ_TimeDomain)).BeginInit();
            this.tableLayoutPanel_Control.SuspendLayout();
            this.groupBox_VectorSetting.SuspendLayout();
            this.groupBox_Download.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel_Base
            // 
            this.tableLayoutPanel_Base.ColumnCount = 2;
            this.tableLayoutPanel_Base.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Base.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 210F));
            this.tableLayoutPanel_Base.Controls.Add(this.tableLayoutPanel_View, 0, 0);
            this.tableLayoutPanel_Base.Controls.Add(this.tableLayoutPanel_Control, 1, 0);
            this.tableLayoutPanel_Base.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_Base.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel_Base.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel_Base.Name = "tableLayoutPanel_Base";
            this.tableLayoutPanel_Base.RowCount = 1;
            this.tableLayoutPanel_Base.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Base.Size = new System.Drawing.Size(684, 461);
            this.tableLayoutPanel_Base.TabIndex = 0;
            // 
            // tableLayoutPanel_View
            // 
            this.tableLayoutPanel_View.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel_View.ColumnCount = 1;
            this.tableLayoutPanel_View.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_View.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 16F));
            this.tableLayoutPanel_View.Controls.Add(this.chart_Spectrum, 0, 0);
            this.tableLayoutPanel_View.Controls.Add(this.chart_IQ_TimeDomain, 0, 1);
            this.tableLayoutPanel_View.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_View.Location = new System.Drawing.Point(2, 2);
            this.tableLayoutPanel_View.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel_View.Name = "tableLayoutPanel_View";
            this.tableLayoutPanel_View.RowCount = 2;
            this.tableLayoutPanel_View.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_View.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 160F));
            this.tableLayoutPanel_View.Size = new System.Drawing.Size(470, 457);
            this.tableLayoutPanel_View.TabIndex = 0;
            // 
            // chart_Spectrum
            // 
            chartArea1.Name = "ChartArea1";
            this.chart_Spectrum.ChartAreas.Add(chartArea1);
            this.chart_Spectrum.Dock = System.Windows.Forms.DockStyle.Fill;
            legend1.Name = "Legend1";
            this.chart_Spectrum.Legends.Add(legend1);
            this.chart_Spectrum.Location = new System.Drawing.Point(4, 4);
            this.chart_Spectrum.Name = "chart_Spectrum";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart_Spectrum.Series.Add(series1);
            this.chart_Spectrum.Size = new System.Drawing.Size(462, 288);
            this.chart_Spectrum.TabIndex = 0;
            this.chart_Spectrum.Text = "chart1";
            // 
            // chart_IQ_TimeDomain
            // 
            chartArea2.Name = "ChartArea1";
            this.chart_IQ_TimeDomain.ChartAreas.Add(chartArea2);
            this.chart_IQ_TimeDomain.Dock = System.Windows.Forms.DockStyle.Fill;
            legend2.Name = "Legend1";
            this.chart_IQ_TimeDomain.Legends.Add(legend2);
            this.chart_IQ_TimeDomain.Location = new System.Drawing.Point(4, 299);
            this.chart_IQ_TimeDomain.Name = "chart_IQ_TimeDomain";
            series2.ChartArea = "ChartArea1";
            series2.Legend = "Legend1";
            series2.Name = "Series1";
            this.chart_IQ_TimeDomain.Series.Add(series2);
            this.chart_IQ_TimeDomain.Size = new System.Drawing.Size(462, 154);
            this.chart_IQ_TimeDomain.SuppressExceptions = true;
            this.chart_IQ_TimeDomain.TabIndex = 1;
            this.chart_IQ_TimeDomain.Text = "chart1";
            // 
            // tableLayoutPanel_Control
            // 
            this.tableLayoutPanel_Control.ColumnCount = 1;
            this.tableLayoutPanel_Control.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Control.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 16F));
            this.tableLayoutPanel_Control.Controls.Add(this.groupBox_VectorSetting, 0, 0);
            this.tableLayoutPanel_Control.Controls.Add(this.groupBox_Download, 0, 1);
            this.tableLayoutPanel_Control.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_Control.Location = new System.Drawing.Point(476, 2);
            this.tableLayoutPanel_Control.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel_Control.Name = "tableLayoutPanel_Control";
            this.tableLayoutPanel_Control.RowCount = 7;
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 241F));
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 154F));
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22F));
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel_Control.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_Control.Size = new System.Drawing.Size(206, 457);
            this.tableLayoutPanel_Control.TabIndex = 1;
            // 
            // groupBox_VectorSetting
            // 
            this.groupBox_VectorSetting.Controls.Add(this.textBox_ToneFreq2_kHz);
            this.groupBox_VectorSetting.Controls.Add(this.checkBox_Tone2);
            this.groupBox_VectorSetting.Controls.Add(this.checkBox_Tone1);
            this.groupBox_VectorSetting.Controls.Add(this.button_ExportVectorFile);
            this.groupBox_VectorSetting.Controls.Add(this.label6);
            this.groupBox_VectorSetting.Controls.Add(this.textBox_VectorLength);
            this.groupBox_VectorSetting.Controls.Add(this.button_GenVector);
            this.groupBox_VectorSetting.Controls.Add(this.textBox_BitRate_bps);
            this.groupBox_VectorSetting.Controls.Add(this.label5);
            this.groupBox_VectorSetting.Controls.Add(this.label3);
            this.groupBox_VectorSetting.Controls.Add(this.textBox_BitPattern);
            this.groupBox_VectorSetting.Controls.Add(this.textBox_ToneFreq1_kHz);
            this.groupBox_VectorSetting.Controls.Add(this.label2);
            this.groupBox_VectorSetting.Controls.Add(this.textBox_FFT_Size);
            this.groupBox_VectorSetting.Controls.Add(this.textBox_SamplingFreq_kHz);
            this.groupBox_VectorSetting.Controls.Add(this.label1);
            this.groupBox_VectorSetting.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_VectorSetting.Location = new System.Drawing.Point(2, 2);
            this.groupBox_VectorSetting.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox_VectorSetting.Name = "groupBox_VectorSetting";
            this.groupBox_VectorSetting.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox_VectorSetting.Size = new System.Drawing.Size(202, 237);
            this.groupBox_VectorSetting.TabIndex = 0;
            this.groupBox_VectorSetting.TabStop = false;
            this.groupBox_VectorSetting.Text = "Generate Vector Signal";
            // 
            // textBox_ToneFreq2_kHz
            // 
            this.textBox_ToneFreq2_kHz.Location = new System.Drawing.Point(125, 70);
            this.textBox_ToneFreq2_kHz.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_ToneFreq2_kHz.Name = "textBox_ToneFreq2_kHz";
            this.textBox_ToneFreq2_kHz.Size = new System.Drawing.Size(73, 21);
            this.textBox_ToneFreq2_kHz.TabIndex = 18;
            this.textBox_ToneFreq2_kHz.Text = "-1000";
            // 
            // checkBox_Tone2
            // 
            this.checkBox_Tone2.AutoSize = true;
            this.checkBox_Tone2.Location = new System.Drawing.Point(4, 73);
            this.checkBox_Tone2.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox_Tone2.Name = "checkBox_Tone2";
            this.checkBox_Tone2.Size = new System.Drawing.Size(123, 16);
            this.checkBox_Tone2.TabIndex = 17;
            this.checkBox_Tone2.Text = "Tone2 Freq(kHz):";
            this.checkBox_Tone2.UseVisualStyleBackColor = true;
            // 
            // checkBox_Tone1
            // 
            this.checkBox_Tone1.AutoSize = true;
            this.checkBox_Tone1.Checked = true;
            this.checkBox_Tone1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Tone1.Location = new System.Drawing.Point(4, 47);
            this.checkBox_Tone1.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox_Tone1.Name = "checkBox_Tone1";
            this.checkBox_Tone1.Size = new System.Drawing.Size(123, 16);
            this.checkBox_Tone1.TabIndex = 15;
            this.checkBox_Tone1.Text = "Tone1 Freq(kHz):";
            this.checkBox_Tone1.UseVisualStyleBackColor = true;
            // 
            // button_ExportVectorFile
            // 
            this.button_ExportVectorFile.Location = new System.Drawing.Point(5, 210);
            this.button_ExportVectorFile.Name = "button_ExportVectorFile";
            this.button_ExportVectorFile.Size = new System.Drawing.Size(70, 21);
            this.button_ExportVectorFile.TabIndex = 14;
            this.button_ExportVectorFile.Text = "Export";
            this.button_ExportVectorFile.UseVisualStyleBackColor = true;
            this.button_ExportVectorFile.Click += new System.EventHandler(this.button_ExportVectorFile_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(3, 158);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(87, 12);
            this.label6.TabIndex = 13;
            this.label6.Text = "Vector Length:";
            // 
            // textBox_VectorLength
            // 
            this.textBox_VectorLength.Location = new System.Drawing.Point(118, 155);
            this.textBox_VectorLength.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_VectorLength.Name = "textBox_VectorLength";
            this.textBox_VectorLength.ReadOnly = true;
            this.textBox_VectorLength.Size = new System.Drawing.Size(80, 21);
            this.textBox_VectorLength.TabIndex = 12;
            // 
            // button_GenVector
            // 
            this.button_GenVector.Location = new System.Drawing.Point(107, 210);
            this.button_GenVector.Name = "button_GenVector";
            this.button_GenVector.Size = new System.Drawing.Size(90, 21);
            this.button_GenVector.TabIndex = 11;
            this.button_GenVector.Text = "Generate";
            this.button_GenVector.UseVisualStyleBackColor = true;
            this.button_GenVector.Click += new System.EventHandler(this.button_GenVector_Click);
            // 
            // textBox_BitRate_bps
            // 
            this.textBox_BitRate_bps.Location = new System.Drawing.Point(118, 125);
            this.textBox_BitRate_bps.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_BitRate_bps.Name = "textBox_BitRate_bps";
            this.textBox_BitRate_bps.Size = new System.Drawing.Size(80, 21);
            this.textBox_BitRate_bps.TabIndex = 10;
            this.textBox_BitRate_bps.Text = "8163";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 126);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(87, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "Bit Rate (bps):";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 103);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "Bit pattern:";
            // 
            // textBox_BitPattern
            // 
            this.textBox_BitPattern.Location = new System.Drawing.Point(118, 100);
            this.textBox_BitPattern.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_BitPattern.Name = "textBox_BitPattern";
            this.textBox_BitPattern.Size = new System.Drawing.Size(80, 21);
            this.textBox_BitPattern.TabIndex = 6;
            this.textBox_BitPattern.Text = "10101010";
            // 
            // textBox_ToneFreq1_kHz
            // 
            this.textBox_ToneFreq1_kHz.Location = new System.Drawing.Point(125, 45);
            this.textBox_ToneFreq1_kHz.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_ToneFreq1_kHz.Name = "textBox_ToneFreq1_kHz";
            this.textBox_ToneFreq1_kHz.Size = new System.Drawing.Size(73, 21);
            this.textBox_ToneFreq1_kHz.TabIndex = 5;
            this.textBox_ToneFreq1_kHz.Text = "1000";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 183);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "FFT Size(^2):";
            // 
            // textBox_FFT_Size
            // 
            this.textBox_FFT_Size.Location = new System.Drawing.Point(118, 180);
            this.textBox_FFT_Size.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_FFT_Size.Name = "textBox_FFT_Size";
            this.textBox_FFT_Size.ReadOnly = true;
            this.textBox_FFT_Size.Size = new System.Drawing.Size(80, 21);
            this.textBox_FFT_Size.TabIndex = 2;
            // 
            // textBox_SamplingFreq_kHz
            // 
            this.textBox_SamplingFreq_kHz.Location = new System.Drawing.Point(128, 18);
            this.textBox_SamplingFreq_kHz.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_SamplingFreq_kHz.Name = "textBox_SamplingFreq_kHz";
            this.textBox_SamplingFreq_kHz.Size = new System.Drawing.Size(70, 21);
            this.textBox_SamplingFreq_kHz.TabIndex = 1;
            this.textBox_SamplingFreq_kHz.Text = "8000";
            this.textBox_SamplingFreq_kHz.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_SamplingFreq_kHz_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Sampling Freq(kHz):";
            // 
            // groupBox_Download
            // 
            this.groupBox_Download.Controls.Add(this.label12);
            this.groupBox_Download.Controls.Add(this.label11);
            this.groupBox_Download.Controls.Add(this.textBox_ALC_Start);
            this.groupBox_Download.Controls.Add(this.textBox_ALC_Stop);
            this.groupBox_Download.Controls.Add(this.label10);
            this.groupBox_Download.Controls.Add(this.textBox_VectorPeriod);
            this.groupBox_Download.Controls.Add(this.textBox_DownloadSamplingFreq_kHz);
            this.groupBox_Download.Controls.Add(this.button_ImportVectorFile);
            this.groupBox_Download.Controls.Add(this.label9);
            this.groupBox_Download.Controls.Add(this.label8);
            this.groupBox_Download.Controls.Add(this.comboBox_SelectSigGen);
            this.groupBox_Download.Controls.Add(this.button_Download);
            this.groupBox_Download.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_Download.Location = new System.Drawing.Point(3, 244);
            this.groupBox_Download.Name = "groupBox_Download";
            this.groupBox_Download.Size = new System.Drawing.Size(200, 148);
            this.groupBox_Download.TabIndex = 1;
            this.groupBox_Download.TabStop = false;
            this.groupBox_Download.Text = "Vector down load";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(129, 98);
            this.label12.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(14, 12);
            this.label12.TabIndex = 26;
            this.label12.Text = "~";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(3, 98);
            this.label11.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(69, 12);
            this.label11.TabIndex = 25;
            this.label11.Text = "ALC range:";
            // 
            // textBox_ALC_Start
            // 
            this.textBox_ALC_Start.Location = new System.Drawing.Point(76, 95);
            this.textBox_ALC_Start.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_ALC_Start.Name = "textBox_ALC_Start";
            this.textBox_ALC_Start.Size = new System.Drawing.Size(50, 21);
            this.textBox_ALC_Start.TabIndex = 24;
            this.textBox_ALC_Start.Text = "32";
            // 
            // textBox_ALC_Stop
            // 
            this.textBox_ALC_Stop.Location = new System.Drawing.Point(147, 95);
            this.textBox_ALC_Stop.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_ALC_Stop.Name = "textBox_ALC_Stop";
            this.textBox_ALC_Stop.Size = new System.Drawing.Size(50, 21);
            this.textBox_ALC_Stop.TabIndex = 23;
            this.textBox_ALC_Stop.Text = "128";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(2, 73);
            this.label10.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(127, 12);
            this.label10.TabIndex = 22;
            this.label10.Text = "Vector Period(msec):";
            // 
            // textBox_VectorPeriod
            // 
            this.textBox_VectorPeriod.Location = new System.Drawing.Point(127, 70);
            this.textBox_VectorPeriod.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_VectorPeriod.Name = "textBox_VectorPeriod";
            this.textBox_VectorPeriod.ReadOnly = true;
            this.textBox_VectorPeriod.Size = new System.Drawing.Size(70, 21);
            this.textBox_VectorPeriod.TabIndex = 21;
            // 
            // textBox_DownloadSamplingFreq_kHz
            // 
            this.textBox_DownloadSamplingFreq_kHz.Location = new System.Drawing.Point(127, 45);
            this.textBox_DownloadSamplingFreq_kHz.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_DownloadSamplingFreq_kHz.Name = "textBox_DownloadSamplingFreq_kHz";
            this.textBox_DownloadSamplingFreq_kHz.Size = new System.Drawing.Size(70, 21);
            this.textBox_DownloadSamplingFreq_kHz.TabIndex = 20;
            this.textBox_DownloadSamplingFreq_kHz.Text = "4000";
            // 
            // button_ImportVectorFile
            // 
            this.button_ImportVectorFile.Location = new System.Drawing.Point(3, 121);
            this.button_ImportVectorFile.Name = "button_ImportVectorFile";
            this.button_ImportVectorFile.Size = new System.Drawing.Size(90, 21);
            this.button_ImportVectorFile.TabIndex = 19;
            this.button_ImportVectorFile.Text = "Import";
            this.button_ImportVectorFile.UseVisualStyleBackColor = true;
            this.button_ImportVectorFile.Click += new System.EventHandler(this.button_ImportVectorFile_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(3, 51);
            this.label9.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(122, 12);
            this.label9.TabIndex = 19;
            this.label9.Text = "Sampling Freq(kHz):";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(3, 23);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(89, 12);
            this.label8.TabIndex = 14;
            this.label8.Text = "Select SigGen:";
            // 
            // comboBox_SelectSigGen
            // 
            this.comboBox_SelectSigGen.FormattingEnabled = true;
            this.comboBox_SelectSigGen.Location = new System.Drawing.Point(97, 20);
            this.comboBox_SelectSigGen.Name = "comboBox_SelectSigGen";
            this.comboBox_SelectSigGen.Size = new System.Drawing.Size(100, 20);
            this.comboBox_SelectSigGen.TabIndex = 1;
            this.comboBox_SelectSigGen.DropDown += new System.EventHandler(this.comboBox_SelectSigGen_DropDown);
            // 
            // button_Download
            // 
            this.button_Download.Location = new System.Drawing.Point(110, 121);
            this.button_Download.Name = "button_Download";
            this.button_Download.Size = new System.Drawing.Size(90, 21);
            this.button_Download.TabIndex = 0;
            this.button_Download.Text = "Download";
            this.button_Download.UseVisualStyleBackColor = true;
            this.button_Download.Click += new System.EventHandler(this.button_Download_Click);
            // 
            // VectorDownForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(684, 461);
            this.Controls.Add(this.tableLayoutPanel_Base);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "VectorDownForm";
            this.Text = "VectorDownloaderForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.VectorDownForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.VectorDownForm_FormClosed);
            this.Load += new System.EventHandler(this.VectorDownForm_Load);
            this.tableLayoutPanel_Base.ResumeLayout(false);
            this.tableLayoutPanel_View.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart_Spectrum)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart_IQ_TimeDomain)).EndInit();
            this.tableLayoutPanel_Control.ResumeLayout(false);
            this.groupBox_VectorSetting.ResumeLayout(false);
            this.groupBox_VectorSetting.PerformLayout();
            this.groupBox_Download.ResumeLayout(false);
            this.groupBox_Download.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_Base;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_View;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_Control;
        private System.Windows.Forms.GroupBox groupBox_VectorSetting;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_SamplingFreq_kHz;
        private System.Windows.Forms.TextBox textBox_FFT_Size;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_ToneFreq1_kHz;
        private System.Windows.Forms.TextBox textBox_BitPattern;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox_BitRate_bps;
        private System.Windows.Forms.Button button_GenVector;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_VectorLength;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart_Spectrum;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart_IQ_TimeDomain;
        private System.Windows.Forms.Button button_ExportVectorFile;
        private System.Windows.Forms.TextBox textBox_ToneFreq2_kHz;
        private System.Windows.Forms.CheckBox checkBox_Tone2;
        private System.Windows.Forms.CheckBox checkBox_Tone1;
        private System.Windows.Forms.GroupBox groupBox_Download;
        private System.Windows.Forms.Button button_Download;
        private System.Windows.Forms.ComboBox comboBox_SelectSigGen;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button_ImportVectorFile;
        private System.Windows.Forms.TextBox textBox_DownloadSamplingFreq_kHz;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textBox_VectorPeriod;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBox_ALC_Start;
        private System.Windows.Forms.TextBox textBox_ALC_Stop;
    }
}