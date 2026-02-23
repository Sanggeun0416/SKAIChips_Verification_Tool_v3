using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SKAIChips_Verification
{
    public partial class VectorDownForm : Form
    {
        public ProgressBar ProgressBar { get; set; }
        public JLcLib.IniFile IniFile { get; set; }
        private VectorSignal Vector { get; set; }

        private string IQFileName;
        
        public VectorDownForm()
        {
            InitializeComponent();
        }

        private void VectorDownForm_Load(object sender, EventArgs e)
        {
            ReadSettingFile();
            SetInstrument();
        }

        private void VectorDownForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            WriteSettingFile();
        }

        private void VectorDownForm_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void ReadSettingFile()
        {
            string sVar;
            if (IniFile == null)
                return;
            IniFile.Read(this);
            sVar = IniFile.Read(Name, "IQ_File_Name");
            if (sVar != null && sVar != "")
                IQFileName = sVar;
        }

        private void WriteSettingFile()
        {
            if (IniFile == null)
                return;

            IniFile.Write(this);
            IniFile.Write(Name, "IQ_File_Name", IQFileName);
        }

        private void textBox_SamplingFreq_kHz_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                GetSamplingFreq();
        }

        private double GetSamplingFreq()
        {
            double Fs_kHz = 8000;
            try { Fs_kHz = double.Parse(textBox_SamplingFreq_kHz.Text); }
            catch { textBox_SamplingFreq_kHz.Text = Fs_kHz.ToString(); };

            return Fs_kHz;
        }
        private double GetTone1Freq()
        {
            double Tone1_kHz = 1000;
            try { Tone1_kHz = double.Parse(textBox_ToneFreq1_kHz.Text); }
            catch { textBox_ToneFreq1_kHz.Text = Tone1_kHz.ToString(); };
            return Tone1_kHz;
        }
        private double GetTone2Freq()
        {
            double Tone2_kHz = -1000;
            try { Tone2_kHz = double.Parse(textBox_ToneFreq2_kHz.Text); }
            catch { textBox_ToneFreq2_kHz.Text = Tone2_kHz.ToString(); };
            return Tone2_kHz;
        }
        private double GetBitRate()
        {
            double BitRate_bps = 8163;
            try { BitRate_bps = double.Parse(textBox_BitRate_bps.Text); }
            catch { textBox_BitRate_bps.Text = BitRate_bps.ToString(); };
            return BitRate_bps;
        }
        private int[] GetBitPattern()
        {
            List<int> BitPattern = new List<int>();
            BitPattern.Clear();
            for (int i = 0; i < textBox_BitPattern.Text.Length; i++)
            {
                if (textBox_BitPattern.Text[i] == '0')
                    BitPattern.Add(0);
                else
                    BitPattern.Add(1);
            }
            string StrPattern = "";
            for (int i = 0; i < BitPattern.Count; i++)
            {
                if (BitPattern[i] == 0)
                    StrPattern += "0";
                else
                    StrPattern += "1";
            }
            textBox_BitPattern.Text = StrPattern;

            return BitPattern.ToArray();
        }

        private void button_GenVector_Click(object sender, EventArgs e)
        {
            double I, Q;
            double Fs_kHz = GetSamplingFreq();
            double Ftone1_kHz = GetTone1Freq();
            double Ftone2_kHz = GetTone2Freq();
            double BitRate_bps = GetBitRate();
            int[] BitPattern = GetBitPattern();

            double Ts_nsec = 1E6 / Fs_kHz;
            double Tpatten_nsec = 1E9 / BitRate_bps;
            int Size = (int)(Tpatten_nsec / Ts_nsec);
            List<double> IValues = new List<double>();
            List<double> QValues = new List<double>();

            int VectorLength = (int)(Tpatten_nsec * BitPattern.Length / Ts_nsec);

            textBox_DownloadSamplingFreq_kHz.Text = Fs_kHz.ToString();
            textBox_VectorPeriod.Text = (Tpatten_nsec * BitPattern.Length / 1E6).ToString("F3");

            IValues.Clear();
            QValues.Clear();
            for (int i = 0; i < VectorLength; i++)
            {
                I = Q = 0F;
                if (checkBox_Tone1.Checked)
                {
                    I += Math.Cos(2 * Math.PI * Ftone1_kHz * 1000 * (i * Ts_nsec) / 1E9);
                    Q += Math.Sin(2 * Math.PI * Ftone1_kHz * 1000 * (i * Ts_nsec) / 1E9);
                }
                if (checkBox_Tone2.Checked)
                {
                    I += Math.Cos(2 * Math.PI * Ftone2_kHz * 1000 * (i * Ts_nsec) / 1E9);
                    Q += Math.Sin(2 * Math.PI * Ftone2_kHz * 1000 * (i * Ts_nsec) / 1E9);
                }
                IValues.Add(I);
                QValues.Add(Q);
            }
            for (int i = 0; i < BitPattern.Length; i++)
            {
                for (int j = 0; j < Size; j++)
                {
                    if (BitPattern[i] == 0)
                        IValues[i * Size + j] = QValues[i * Size + j] = 0;
                }
            }
            Vector = new VectorSignal(IValues.ToArray(), QValues.ToArray(), Fs_kHz);
            Vector.Name = "TestTone_" + Fs_kHz.ToString("F2") + "kHz_Pattern_" + textBox_BitPattern.Text;
            Vector.RunFFT();
            textBox_VectorLength.Text = Vector.Length.ToString();
            textBox_FFT_Size.Text = Vector.FFT_Size.ToString();

            SetTimeDomainIQ(Vector.TimeValues.ToArray(), Vector.IValues.ToArray(), Vector.QValues.ToArray());
            double[] xFreq, yPow;
            Vector.MakeSpecturmSignal(Vector.FFT_Size / 1024 + 1, out xFreq, out yPow);
            SetSpectrum(xFreq, yPow);
            //SetSpectrum(Vector.Frequencies, Vector.PowerValues);
            SetChartSetting();
        }

        private void SetInstrument()
        {
            comboBox_SelectSigGen.Items.Clear();
            foreach (JLcLib.Instrument.InsInformation Ins in JLcLib.Instrument.InstrumentForm.InsInfoList)
            {
                if (Ins.Valid == true)
                {
                    switch (Ins.Type)
                    {
                        case JLcLib.Instrument.InstrumentTypes.SignalGenerator0:
                        case JLcLib.Instrument.InstrumentTypes.SignalGenerator1:
                        case JLcLib.Instrument.InstrumentTypes.SignalGenerator2:
                            comboBox_SelectSigGen.Items.Add(Ins.Type.ToString());
                            break;
                        default:
                            break;
                    }
                }
            }
            if (comboBox_SelectSigGen.Items.Count > 0)
                comboBox_SelectSigGen.SelectedIndex = 0;
        }

        private void comboBox_SelectSigGen_DropDown(object sender, EventArgs e)
        {
            SetInstrument();
        }

        private void button_ExportVectorFile_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog FileDlg = new SaveFileDialog())
            {
                FileDlg.Filter = "Text File (*.txt)|*.txt|All files (*.*)|*.*";
                if (IQFileName == "")
                    FileDlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
                else
                    FileDlg.InitialDirectory = System.IO.Path.GetDirectoryName(IQFileName);

                if (FileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    IQFileName = FileDlg.FileName;

                    System.IO.FileStream fs = new System.IO.FileStream(IQFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs);

                    if (Vector != null && Vector.Length > 0)
                    {
                        for (int i = 0; i < Vector.Length; i++)
                            sw.WriteLine(Vector.IValues[i].ToString("F8") + "," + Vector.QValues[i].ToString("F8"));
                            //sw.WriteLine(Vector.IValues[i].ToString("F8") + "\t" + Vector.QValues[i].ToString("F8"));
                    }
                    if (sw != null) sw.Close();
                    if (fs != null) fs.Close();
                }
            }
        }

        private void button_Download_Click(object sender, EventArgs e)
        {
            if (Vector == null || Vector.Length == 0)
                return;

            double Fs_kHz = 4000;
            try { Fs_kHz = double.Parse(textBox_DownloadSamplingFreq_kHz.Text); }
            catch { textBox_DownloadSamplingFreq_kHz.Text = Fs_kHz.ToString(); };

            Vector.SetSamplingFreq(Fs_kHz);
            Vector.Normalize(VectorSignal.eNormalize.NORMALIZED_14BIT);
            DownloadWaveform();
        }

        private void button_ImportVectorFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog FileDlg = new OpenFileDialog())
            {
                FileDlg.Filter = "Waveform File (*.txt;*.dat)|*.txt;*.dat|All files (*.*)|*.*";
                if (IQFileName == "")
                    FileDlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
                else
                    FileDlg.InitialDirectory = System.IO.Path.GetDirectoryName(IQFileName);

                if (FileDlg.ShowDialog() == DialogResult.OK)
                {
                    IQFileName = FileDlg.FileName;
                    
                    System.IO.FileStream fs = new System.IO.FileStream(FileDlg.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                    System.IO.StreamReader sr = new System.IO.StreamReader(fs);
                    List<double> IValues = new List<double>();
                    List<double> QValues = new List<double>();
                    IValues.Clear();
                    QValues.Clear();

                    for (int i = 0; ; i++)
                    {
                        if (sr.EndOfStream)
                            break;
                        string[] sIQValues = sr.ReadLine().Split(new char[] { ' ', ',', '\t' });
                        if (sIQValues.Length >= 2)
                        {
                            IValues.Add(double.Parse(sIQValues[0], System.Globalization.NumberStyles.Number));
                            QValues.Add(double.Parse(sIQValues[1], System.Globalization.NumberStyles.Number));
                        }
                    }
                    sr.Close();
                    fs.Close();

                    double Fs_kHz = 4000;
                    try { Fs_kHz = double.Parse(textBox_DownloadSamplingFreq_kHz.Text); }
                    catch { textBox_DownloadSamplingFreq_kHz.Text = Fs_kHz.ToString(); };
                    Vector = new VectorSignal(IValues.ToArray(), QValues.ToArray(), Fs_kHz);
                    Vector.Name = System.IO.Path.GetFileNameWithoutExtension(FileDlg.FileName);
                    Vector.RunFFT();
                    
                    SetTimeDomainIQ(Vector.TimeValues.ToArray(), Vector.IValues.ToArray(), Vector.QValues.ToArray());
                    SetSpectrum(Vector.Frequencies, Vector.PowerValues);
                    SetChartSetting();
                }
            }
        }

        private bool DownloadWaveform()
        {
            string VectorName = Vector.Name;
            int ALC_Start = int.Parse(this.textBox_ALC_Start.Text, System.Globalization.NumberStyles.Integer);
            int ALC_Stop = int.Parse(this.textBox_ALC_Stop.Text, System.Globalization.NumberStyles.Integer);
            JLcLib.Instrument.InstrumentTypes InsType;

            InsType = (JLcLib.Instrument.InstrumentTypes)(comboBox_SelectSigGen.SelectedIndex + (int)JLcLib.Instrument.InstrumentTypes.SignalGenerator0);
            JLcLib.Instrument.SCPI SigGen = new JLcLib.Instrument.SCPI(InsType);

            // limit ALC range
            if (ALC_Start <= 0)
            {
                ALC_Start = 1;
                textBox_ALC_Start.Text = ALC_Start.ToString();
            }
            if (ALC_Stop > Vector.Length)
            {
                ALC_Stop = Vector.Length;
                textBox_ALC_Stop.Text = ALC_Stop.ToString();
            }

            // Find for signal generator that is downloaded waveform 
            SigGen.Open();
            if (SigGen.IsOpen == false)
            {
                MessageBox.Show("Faild to connection to " + SigGen.InsInfo.Type.ToString() + ".\nPlease, check cable and address.\nAddress can be found \"Keysight IO libries suite\"");
                return false;
            }
            // Download waveform
            SigGen.Write(":SOUR:RAD:ARB 0"); // Trun off ARB
            int MemSize = Vector.Length * 2 * 2;
            SigGen.WaitAllOperationsComplete();
            SigGen.Write(":MMEM:DATA \"WFM1:" + VectorName + "\",#" + (MemSize.ToString()).Length.ToString() + MemSize.ToString(), Vector.IQBytes);
            SigGen.WaitAllOperationsComplete();
            SigGen.Write(":SOUR:RAD:ARB:WAV \"WFM1:" + VectorName + "\"");
            // Vector setup
            SigGen.Write(":SOUR:RAD:ARB:SCL:RATE " + Vector.Fs_kHz.ToString() + " kHz"); // Set ARB clock
            int MarkNum = 1;
            int FirstPoint = 1;
            int LastPoint = Vector.Length / 64;
            int SkipCount = 0;
            SigGen.Write(":SOUR:RAD:ARB:MARK:CLE:ALL \"" + VectorName + "\"," + MarkNum.ToString());
            SigGen.Write(":SOUR:RAD:ARB:MARK:SET \"" + VectorName + "\"," + MarkNum.ToString() + "," + FirstPoint.ToString() + "," + LastPoint.ToString() + "," + SkipCount.ToString());
            MarkNum = 4;
            FirstPoint = ALC_Start;
            LastPoint = ALC_Stop;
            SigGen.Write(":SOUR:RAD:ARB:MARK:CLE:ALL \"" + VectorName + "\"," + MarkNum.ToString());
            SigGen.Write(":SOUR:RAD:ARB:MARK:SET \"" + VectorName + "\"," + MarkNum.ToString() + "," + FirstPoint.ToString() + "," + LastPoint.ToString() + "," + SkipCount.ToString());

            SigGen.Write(":SOUR:RAD:ARB:MDES:ALCH M4"); // NONE, M1, M2, M3, M4 (ex M4 means that ALC hold period is defined marker 4 range)
            SigGen.Write(":SOUR:RAD:ARB:TRIG:TYPE CONT"); // CONT, SING, GATE, SADV
            SigGen.Write(":SOUR:RAD:ARB:TRIG:TYPE:CONT FREE"); // FREE, TRIG, RES
            SigGen.Write(":SOUR:RAD:ARB:HEAD:SAVE"); // Save Test vector hearder

            // Vector on
            SigGen.Write(":SOUR:RAD:ARB 1");

            SigGen.Close();

            return true;
        }

        void SetChartSetting()
        {
            chart_Spectrum.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            chart_Spectrum.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            chart_Spectrum.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_Spectrum.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_IQ_TimeDomain.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            //chart_IQ_TimeDomain.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            chart_IQ_TimeDomain.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            //chart_IQ_TimeDomain.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
        }

        void SetTimeDomainIQ(double[] xTime, double[] yIValues, double[] yQValues)
        {
            Series Iseries = new Series("I");
            Series Qseries = new Series("Q");

            Iseries.ChartType = SeriesChartType.FastLine;
            Qseries.ChartType = SeriesChartType.FastLine;
            Iseries.Points.DataBindXY(xTime, yIValues);
            Qseries.Points.DataBindXY(xTime, yQValues);
            
            chart_IQ_TimeDomain.Series.Clear();
            chart_IQ_TimeDomain.Series.Add(Iseries);
            chart_IQ_TimeDomain.Series.Add(Qseries);
            chart_IQ_TimeDomain.ChartAreas[0].AxisX.Minimum = xTime.Min();
            chart_IQ_TimeDomain.ChartAreas[0].AxisX.Maximum = xTime.Max();
            chart_IQ_TimeDomain.ChartAreas[0].AxisY.Minimum = Math.Min(yIValues.Min(), yQValues.Min());
            chart_IQ_TimeDomain.ChartAreas[0].AxisY.Maximum = Math.Max(yIValues.Max(), yQValues.Max());
        }
        
        void SetSpectrum(double[] xFrequencies, double[] yPowerValues)
        {
            Series IQseries = new Series("IQ");

            IQseries.ChartType = SeriesChartType.FastLine;
            IQseries.Points.DataBindXY(xFrequencies, yPowerValues);
            
            chart_Spectrum.Series.Clear();
            chart_Spectrum.Series.Add(IQseries);
            chart_Spectrum.ChartAreas[0].AxisX.Minimum = -Vector.Fs_kHz / 2;
            chart_Spectrum.ChartAreas[0].AxisX.Maximum = Vector.Fs_kHz / 2;
            chart_Spectrum.ChartAreas[0].AxisX.Interval = Vector.Fs_kHz / 10;
            int yMin = (int)Math.Round(yPowerValues.Min());
            int yMax = (int)Math.Round(yPowerValues.Max());
            yMin = ((yMin - 5) / 10 - 1) * 10;
            yMax = ((yMax + 5) / 10 + 0) * 10;
            //chart_Spectrum.ChartAreas[0].AxisY.Minimum = Math.Round(yPowerValues.Min());
            //chart_Spectrum.ChartAreas[0].AxisY.Maximum = Math.Round(yPowerValues.Max());
            chart_Spectrum.ChartAreas[0].AxisY.Minimum = yMin;
            chart_Spectrum.ChartAreas[0].AxisY.Maximum = yMax;
            chart_Spectrum.ChartAreas[0].AxisY.Interval = (yMax - yMin) / 10;
        }
    }

    public class VectorSignal
    {
        public enum eNormalize
        {
            NORMALIZED_NONE,
            NORMALIZED_6BIT = (1 << 6),
            NORMALIZED_7BIT = (1 << 7),
            NORMALIZED_8BIT = (1 << 8),
            NORMALIZED_9BIT = (1 << 9),
            NORMALIZED_10BIT = (1 << 10),
            NORMALIZED_11BIT = (1 << 11),
            NORMALIZED_12BIT = (1 << 12),
            NORMALIZED_13BIT = (1 << 13),
            NORMALIZED_14BIT = (1 << 14),
            NORMALIZED_15BIT = (1 << 15),
            NORMALIZED_16BIT = (1 << 16),
        }
        public int Length { get { return IValues.Count; } }
        public int FFT_Size { get; private set; }
        /// <summary>
        /// Sampling frequency in kHz
        /// </summary>
        public double Fs_kHz { get; set; }
        /// <summary>
        /// Sampling interval in nsec (1E6 / Fs_kHz)
        /// </summary>
        public double Ts_nsec { get; set; }
        public double Fi_kHz { get; private set; }
        public string Name { get; set; }
        public List<double> IValues { get; private set; } = new List<double>();
        public List<double> QValues { get; private set; } = new List<double>();
        /// <summary>
        /// IQ bytes data for Signal generator. IQBytes is generated after the Normalize() function is executed.
        /// </summary>
        public byte[] IQBytes { get; private set; }
        /// <summary>
        /// Time-axis values in usec
        /// </summary>
        public List<double> TimeValues { get; private set; } = new List<double>();
        /// <summary>
        /// IQ complex signal, It is limited to power of 2 for FFT
        /// </summary>
        public JLcLib.Math.Complex[] IQValues;
        /// <summary>
        /// frequency axis values. these are created after FFT.
        /// </summary>
        public double[] Frequencies; // { get; private set; }
        /// <summary>
        /// Maginitude values. these are created after FFT.
        /// </summary>
        public double[] PowerValues; // { get; private set; }

        public VectorSignal(double[] IValues, double[] QValues, double Fs_kHz)
        {
            int Length = Math.Min(IValues.Length, QValues.Length);

            this.IValues.Clear();
            this.QValues.Clear();

            for (int i = 0; i < Length; i++)
            {
                this.IValues.Add(IValues[i]);
                this.QValues.Add(QValues[i]);
            }
            for (int i = 30; i >= 0; i--)
            {
                FFT_Size = (1 << i);
                if (FFT_Size <= Length)
                    break;
            }
            //MakeComplexSignal();

            SetSamplingFreq(Fs_kHz);
        }

        public void SetSamplingFreq(double Fs_kHz)
        {
            this.Fs_kHz = Fs_kHz;
            Ts_nsec = 1E6 / Fs_kHz;

            TimeValues.Clear();
            for (int i = 0; i < Length; i++)
                TimeValues.Add(i * Ts_nsec / 1000F);
        }

        public bool Normalize(eNormalize Normalize)
        {
            if (Normalize == eNormalize.NORMALIZED_NONE || IValues.Count == 0)
                return false;
        
            double Seed, MaxSeed = 0;
            for (int i = 0; i < Length; i++)
            {
                Seed = (Math.Abs(IValues[i]) >= Math.Abs(QValues[i])) ? Math.Abs(IValues[i]) : Math.Abs(QValues[i]);
                if (Seed > MaxSeed)
                    MaxSeed = Seed;
            }
            double Normalizedvalue = (int)Normalize * Math.Sqrt(2.0) / MaxSeed;
            IQBytes = new byte[Length * 4];
            for (int i = 0; i < Length; i++)
            {
                IValues[i] = Math.Round(IValues[i] * Normalizedvalue);
                QValues[i] = Math.Round(QValues[i] * Normalizedvalue);

                IQBytes[i * 4 + 0] = (byte)(((int)IValues[i] >> 8) & 0xff);
                IQBytes[i * 4 + 1] = (byte)(((int)IValues[i] >> 0) & 0xff);
                IQBytes[i * 4 + 2] = (byte)(((int)QValues[i] >> 8) & 0xff);
                IQBytes[i * 4 + 3] = (byte)(((int)QValues[i] >> 0) & 0xff);
            }
            return true;
        }

        public void RunFFT()
        {
            IQValues = JLcLib.Math.FourierTransform.MakeComplexSignal(IValues.ToArray(), QValues.ToArray(), JLcLib.Math.FourierTransform.WindowType.Rectangular);
            Array.Resize(ref IQValues, FFT_Size);
            JLcLib.Math.FourierTransform.DFT(IQValues, JLcLib.Math.FourierTransform.Direction.Forward);
            //JLcLib.Math.FourierTransform.FFT(IQValues, JLcLib.Math.FourierTransform.Direction.Forward);
            
            for (int i = 0; i < IQValues.Length; i++)
            {
                if (IQValues[i].Re == 0 && IQValues[i].Im == 0)
                    IQValues[i].Re = IQValues[i].Im = 1E-64;
            }

            PowerValues = new double[FFT_Size + 1];
            for (int i = 0; i < PowerValues.Length / 2; i++)
                PowerValues[PowerValues.Length / 2 + 1 + i] += IQValues[1 + i].SquaredMagnitude;
            for (int i = 0; i < PowerValues.Length / 2; i++)
                PowerValues[PowerValues.Length / 2 - 1 - i] += IQValues[FFT_Size - 1 - i].SquaredMagnitude;
            PowerValues[PowerValues.Length / 2] += IQValues[0].SquaredMagnitude;

            Fi_kHz = Fs_kHz / FFT_Size;
            Frequencies = new double[FFT_Size + 1];
            Frequencies[0] = -Fi_kHz * (Frequencies.Length / 2);
            for (int i = 1; i < Frequencies.Length; i++)
                Frequencies[i] = Frequencies[i - 1] + Fi_kHz;
        }

        public double MakeSpecturmSignal(int NumMerges, out double[] xFrequencies, out double[] yPowerValues)
        {
            int MergeSize = NumMerges * 2 - 1;
            double RBW_kHz = (Fs_kHz / FFT_Size) * MergeSize;
            int Len = ((FFT_Size / 2) - (MergeSize / 2)) / MergeSize * 2 + 1;
            int Start = (PowerValues.Length - Len * MergeSize) / 2;

            xFrequencies = new double[Len];
            yPowerValues = new double[Len];

            for (int i = 0; i < Len; i++)
            {
                for (int j = 0; j < MergeSize; j++)
                    yPowerValues[i] += PowerValues[Start + i * MergeSize + j];
                yPowerValues[i] = 10 * Math.Log10(yPowerValues[i]);
                xFrequencies[i] = Frequencies[Start + i * MergeSize + MergeSize / 2];
            }
            return RBW_kHz;
        }
    }
        
}
