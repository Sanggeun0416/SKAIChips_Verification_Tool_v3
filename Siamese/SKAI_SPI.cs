using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using JLcLib.Chip;
using JLcLib.Comn;
using JLcLib.Custom;
using SKAIChips_Verification;
using static System.Net.Mime.MediaTypeNames;
using static SKAIChips_Verification.RegContForm;

namespace SKAI_SPI
{
    public class Lyon : ChipControl
    {
        public enum TEST_ITEMS_MANUAL { JUWON, NUM_TEST_ITEMS }
        public enum TEST_ITEMS_AUTO { NUM_TEST_ITEMS }
        public enum COMBOBOX_ITEMS { MANUAL, AUTO }

        private JLcLib.Custom.SPI SPI = null;
        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();

        private bool IsSerialReceivedData = false;
        private bool IsRunCal = false;

        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter1 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter2 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter3 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI SpectrumAnalyzer = null;
        JLcLib.Instrument.SCPI TempChamber = null;

        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.MANUAL;

        public Lyon(RegContForm form) : base(form)
        {
            SPI = Parent.SPI;
            Serial.ReadSettingFile(form.IniFile, "Lyon");
            Serial.DataReceived += Serial_DataReceived;

            CalibrationData = new byte[256];

            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e) => IsSerialReceivedData = true;

        private void WriteRegister(uint Address, uint Data)
        {
            var bytes = new byte[3];
            const byte rwFlag = 1;
            bytes[0] = (byte)((rwFlag << 7) | (Address & 0x7F));
            bytes[1] = (byte)((Data >> 8) & 0xFF);
            bytes[2] = (byte)(Data & 0xFF);
            SPI.WriteBytes(bytes, 3, true);
        }

        private uint ReadRegister(uint Address)
        {
            const byte rwFlag = 0;
            var tx = new[] { (byte)((rwFlag << 7) | (Address & 0x7F)) };
            var rx = SPI.WriteAndReadBytes(tx, 1, 2);
            return (uint)((rx[0] << 8) | rx[1]);
        }

        public override JLcLib.Chip.RegisterGroup MakeRegisterGroup(string groupName, string[,] regData)
        {
            var rg = new JLcLib.Chip.RegisterGroup(groupName, WriteRegister, ReadRegister);

            for (int xStart = 0; xStart < 3; xStart++)
            {
                for (int row = 0; row < regData.GetLength(0); row++)
                {
                    if ((row + 2 < regData.GetLength(0)) &&
                        regData[row, xStart] == "Bit" &&
                        regData[row + 1, xStart] == "Name" &&
                        regData[row + 2, xStart] == "Default")
                    {
                        string strAddr = regData[row - 1, xStart + 1];
                        if (strAddr != null && strAddr.StartsWith("0x")) strAddr = strAddr.Substring(2);

                        uint address;
                        if (uint.TryParse(strAddr, System.Globalization.NumberStyles.HexNumber, null, out address))
                        {
                            var regName = regData[row - 1, xStart + 2];
                            var reg = rg.Registers.Add(regName, address);

                            for (int col = xStart + 1; col < regData.GetLength(1); col++)
                            {
                                if (regData[row + 2, col] == "X" || regData[row + 2, col] == "-" ||
                                    regData[row + 2, col] == null || regData[row + 1, col] == null)
                                    continue;

                                var itemName = regData[row + 1, col];
                                int upper = int.Parse(regData[row, col]);
                                int lower = upper;
                                uint val = (uint)(regData[row + 2, col] == "0" ? 0 : 1);

                                for (int x = col + 1; x < regData.GetLength(1); x++)
                                {
                                    if (regData[row + 1, x] != null) break;
                                    if (regData[row, x] != null)
                                    {
                                        lower = int.Parse(regData[row, x]);
                                        val = (val << 1) | (uint)(regData[row + 2, x] == "0" ? 0 : 1);
                                    }
                                }

                                string desc = null;
                                for (int y = row; y < regData.GetLength(0); y++)
                                {
                                    if (regData[y, xStart] == itemName)
                                    {
                                        desc = regData[y, xStart + 1];
                                        for (int dr = y + 1; dr < regData.GetLength(0); dr++)
                                        {
                                            if (regData[dr, xStart] != null) break;
                                            string line = "";
                                            if (regData[dr, xStart + 3] != null && regData[dr, xStart + 4] != null)
                                                line = "\n" + regData[dr, xStart + 3] + "=" + regData[dr, xStart + 4];
                                            else if (regData[dr, xStart + 4] != null)
                                                line = "\n" + regData[dr, xStart + 4];
                                            else if (regData[dr, xStart + 3] != null)
                                                line = "\n" + regData[dr, xStart + 3] + "=";
                                            desc += line;
                                        }
                                        break;
                                    }
                                }
                                reg.Items.Add(itemName, upper, lower, val, desc);
                            }
                        }
                    }
                }
            }
            return rg;
        }

        public override void SetChipSpecificUI()
        {
            Parent.ChipCtrlTextboxes[0].ReadOnly = true;
            Parent.ChipCtrlTextboxes[0].Visible = true;
            Parent.ChipCtrlTextboxes[0].Text = Serial.Config.PortName;
            Parent.ChipCtrlTextboxes[0].Size = new System.Drawing.Size(
                Parent.ChipCtrlTextboxes[0].Size.Width * 2 + 3,
                Parent.ChipCtrlTextboxes[0].Size.Height);

            Parent.ChipCtrlButtons[2].Text = "Connect";
            Parent.ChipCtrlButtons[2].Visible = true;
            Parent.ChipCtrlButtons[2].Click += SerialConnect_Click;

            Parent.ChipCtrlButtons[3].Text = "Set UART";
            Parent.ChipCtrlButtons[3].Visible = true;
            Parent.ChipCtrlButtons[3].Click += SerialSetting_Click;

            Parent.ChipCtrlButtons[8].Text = "Manual";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;

            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_Auto_Test_Items;
        }

        private void SerialConnect_Click(object sender, EventArgs e)
        {
            var b = sender as Button;
            if (!Serial.IsOpen)
            {
                Serial.Open();
                if (Serial.IsOpen)
                {
                    b.Text = "Disconn";
                    Parent.ChipCtrlTextboxes[0].Text = $"{Serial.Config.PortName}-{Serial.Config.BaudRate}";
                }
            }
            else
            {
                Serial.Close();
                if (!Serial.IsOpen) b.Text = "Connect";
            }
        }

        private void SerialSetting_Click(object sender, EventArgs e)
        {
            JLcLib.Comn.WireComnForm.Show(Serial);
            Serial.WriteSettingFile(Parent.IniFile, "Lyon");
        }

        private void Change_To_Manual_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.MANUAL;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Change_To_Auto_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.AUTO;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_AUTO.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_AUTO)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        public override bool CheckConnectionForLog() => Serial != null && Serial.IsOpen;

        public override void RunLog()
        {
            if (!Serial.IsOpen || Serial.RcvQueue.Count == 0) return;

            for (int i = 0; i < Serial.RcvQueue.Count; i++)
            {
                byte b = Serial.RcvQueue.Get();
                if ((b > 0x00 && b < 0x80) && b != '\0')
                {
                    LogMessage += (char)b;
                    if (b == '\n')
                    {
                        Log.WriteWithElapsedTime(LogMessage);
                        Log.WriteLine("\n");
                        LogMessage = "";
                    }
                }
            }
        }

        public override void SendCommand(string command)
        {
            command += "\r\n";
            var buf = Encoding.ASCII.GetBytes(command);
            Serial.WriteBytes(buf, buf.Length, true);
        }

        public override void RunTest(int testItemIndex, string arg)
        {
            int iVal;
            if (!int.TryParse(arg, System.Globalization.NumberStyles.Number, null, out iVal))
                iVal = 0;

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.MANUAL:
                    switch ((TEST_ITEMS_MANUAL)testItemIndex)
                    {
                        case TEST_ITEMS_MANUAL.JUWON:
                            if (string.IsNullOrEmpty(arg))
                                Log.WriteLine("시작할 값을 텍스트상자에 10진수로 적어주세요. (0~511)");
                            else
                            {
                                Log.WriteLine("Start : " + arg);
                                Write_Register_for_juwon(iVal);
                            }
                            break;
                    }
                    break;

                case COMBOBOX_ITEMS.AUTO:
                    break;
            }
        }

        private void Write_Register_for_juwon(int iVal)
        {
            var ZCD_DELAY_TRIM_SEL = Parent.RegMgr.GetRegisterItem("ZCD_DELAY_TRIM_SEL<8:0>");
            for (int i = iVal; i >= 0; i--)
            {
                ZCD_DELAY_TRIM_SEL.Value = (uint)i;
                ZCD_DELAY_TRIM_SEL.Write();
            }
        }
    }

    public class Salus8 : ChipControl
    {
        #region Variable and declaration

        public enum TEST_ITEMS_MANUAL
        {
            MANUAL_TEST,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            Start_Temp_Test,
            Stop_Temp_Test,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
        }

        private JLcLib.Custom.SPI SPI = null;

        public bool Enable_TempChamber = true;

        private Serial Serial { get; set; } = new Serial();

        JLcLib.Custom.SCPI PowerSupply0 = null;
        JLcLib.Custom.SCPI PowerSupply1 = null;
        JLcLib.Custom.SCPI DigitalMultimeter0 = null;
        JLcLib.Custom.SCPI DigitalMultimeter1 = null;
        JLcLib.Custom.SCPI DigitalMultimeter2 = null;
        JLcLib.Custom.SCPI DigitalMultimeter3 = null;
        JLcLib.Custom.SCPI OscilloScope0 = null;
        JLcLib.Custom.SCPI SourceAnalyzer = null;
        JLcLib.Custom.SCPI WaveformGenerator = null;
        JLcLib.Custom.SCPI TempChamber = null;

        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.MANUAL;

        #endregion Variable and declaration

        public Salus8(RegContForm form) : base(form)
        {
            SPI = Parent.SPI;

            Serial.ReadSettingFile(form.IniFile, "Salus8");

            CalibrationData = new byte[256];

            /* UI의 테스트 항목 콤보박스를 초기화합니다. */
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] Bytes = new byte[3];
            byte rwFlag = 1; // 쓰기 플래그

            Bytes[0] = (byte)((rwFlag << 7) | (Address & 0x7F));
            Bytes[1] = (byte)((Data >> 8) & 0xff);
            Bytes[2] = (byte)((Data >> 0) & 0xff);

            SPI.WriteBytes(Bytes, 3, true);
        }

        private uint ReadRegister(uint Address)
        {
            byte rwFlag = 0; // 읽기 플래그
            uint Data = 0xFFFF;
            byte[] Bytes = new byte[1];
            byte[] Buff = new byte[2];

            Bytes[0] = (byte)((rwFlag << 7) | (Address & 0x7F));

            Buff = SPI.WriteAndReadBytes(Bytes, 1, 2);

            Data = (uint)((Buff[0] << 8) | (Buff[1] << 0));

            return Data;
        }

        public override JLcLib.Chip.RegisterGroup MakeRegisterGroup(string GroupName, string[,] RegData)
        {
            JLcLib.Chip.RegisterGroup rg = new JLcLib.Chip.RegisterGroup(GroupName, WriteRegister, ReadRegister);

            for (int xStart = 0; xStart < 3; xStart++)
            {
                for (int Row = 0; Row < RegData.GetLength(0); Row++)
                {
                    if ((Row + 2 < RegData.GetLength(0)) && (RegData[Row, xStart] == "Bit") && (RegData[Row + 1, xStart] == "Name") && (RegData[Row + 2, xStart] == "Default"))
                    {
                        string RegName = null;
                        uint Address = 0;

                        string StrAddr = RegData[Row - 1, xStart + 1];
                        if (StrAddr != null && StrAddr.StartsWith("0x"))
                            StrAddr = StrAddr.Substring(2);

                        if (uint.TryParse(StrAddr, System.Globalization.NumberStyles.HexNumber, null, out Address))
                        {
                            RegName = RegData[Row - 1, xStart + 2];
                            JLcLib.Chip.Register reg = rg.Registers.Add(RegName, Address);

                            for (int Column = xStart + 1; Column < RegData.GetLength(1); Column++)
                            {
                                if (!((RegData[Row + 2, Column] == "X") || (RegData[Row + 2, Column] == "-") || (RegData[Row + 2, Column] == null) || (RegData[Row + 1, Column] == null)))
                                {
                                    string ItemName = null, ItemDesc = null;
                                    int UpperBit = 0, LowerBit = 0;
                                    uint ItemValue = 0;

                                    ItemName = RegData[Row + 1, Column];
                                    UpperBit = int.Parse(RegData[Row, Column]);
                                    LowerBit = int.Parse(RegData[Row, Column]);
                                    ItemValue = (uint)((RegData[Row + 2, Column] == "0") ? 0 : 1);

                                    for (int x = Column + 1; x < RegData.GetLength(1); x++)
                                    {
                                        if (RegData[Row + 1, x] == null)
                                        {
                                            if (RegData[Row, x] != null)
                                            {
                                                LowerBit = int.Parse(RegData[Row, x]);
                                                ItemValue = (ItemValue << 1) | (uint)((RegData[Row + 2, x] == "0") ? 0 : 1);
                                            }
                                        }
                                        else break;
                                    }

                                    for (int y = Row; y < RegData.GetLength(0); y++)
                                    {
                                        if (RegData[y, xStart] == ItemName)
                                        {
                                            ItemDesc = RegData[y, xStart + 1];

                                            for (int desc_row = y + 1; desc_row < RegData.GetLength(0); desc_row++)
                                            {
                                                if (RegData[desc_row, xStart] == null)
                                                {
                                                    string lineDesc = "";
                                                    if (RegData[desc_row, xStart + 3] != null && RegData[desc_row, xStart + 4] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 3] + "=" + RegData[desc_row, xStart + 4];
                                                    else if (RegData[desc_row, xStart + 4] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 4];
                                                    else if (RegData[desc_row, xStart + 3] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 3] + "=";

                                                    ItemDesc += lineDesc;
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }
                                            break; // 해당 ItemName에 대한 설명 파싱 완료
                                        }
                                    }
                                    reg.Items.Add(ItemName, UpperBit, LowerBit, ItemValue, ItemDesc);
                                }
                            }
                        }
                    }
                }
            }
            return rg;
        }

        public override void SetChipSpecificUI()
        {
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */

            Parent.ChipCtrlButtons[8].Text = "TEMP_O";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Toogle_Enable_TempChamber;

            Parent.ChipCtrlButtons[9].Text = "RS_SPI";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Run_Reset_SPI;

            Parent.ChipCtrlButtons[10].Text = "Manual";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Change_To_Manual_Test_Items;

            Parent.ChipCtrlButtons[11].Text = "AUTO";
            Parent.ChipCtrlButtons[11].Visible = true;
            Parent.ChipCtrlButtons[11].Click += Change_To_Auto_Test_Items;
        }

        private void SerialConnect_Click(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (Serial.IsOpen == false)
            {
                Serial.Open();
                if (Serial.IsOpen)
                {
                    b.Text = "Disconn";
                    Parent.ChipCtrlTextboxes[0].Text = Serial.Config.PortName + "-" + Serial.Config.BaudRate.ToString();
                }
            }
            else
            {
                Serial.Close();
                if (!Serial.IsOpen)
                    b.Text = "Connect";
            }
        }

        private void SerialSetting_Click(object sender, EventArgs e)
        {
            JLcLib.Comn.WireComnForm.Show(Serial);
            Serial.WriteSettingFile(Parent.IniFile, "Salus8");
        }

        private void Change_To_Manual_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.MANUAL;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Change_To_Auto_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.AUTO;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_AUTO.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_AUTO)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Toogle_Enable_TempChamber(object sender, EventArgs e)
        {
            if (Parent.ChipCtrlButtons[8].Text == "TEMP_O")
            {
                Enable_TempChamber = false;
                Parent.ChipCtrlButtons[8].Text = "TEMP_X";
            }
            else
            {
                Enable_TempChamber = true;
                Parent.ChipCtrlButtons[8].Text = "TEMP_O";
            }
        }

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
            if (Serial.IsOpen && (Serial.RcvQueue.Count != 0))
            {
                for (int i = 0; i < Serial.RcvQueue.Count; i++)
                {
                    byte b = Serial.RcvQueue.Get();
                    if ((b > 0x00 && b < 0x80) && b != '\0')
                    {
                        LogMessage += (char)b;
                        if (b == '\n')
                        {
                            Log.WriteWithElapsedTime(LogMessage);
                            Log.WriteLine("\n");
                            LogMessage = "";
                        }
                    }
                }
            }
        }

        public override void SendCommand(string Command)
        {
            Command += "\r\n";
            byte[] Var = Encoding.ASCII.GetBytes(Command);
            Serial.WriteBytes(Var, Var.Length, true);
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal;
            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.MANUAL:
                    switch ((TEST_ITEMS_MANUAL)TestItemIndex)
                    {
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((TEST_ITEMS_AUTO)TestItemIndex)
                    {
                        case TEST_ITEMS_AUTO.Start_Temp_Test:
                            Start_TEMP_TEST(iVal);
                            break;
                        case TEST_ITEMS_AUTO.Stop_Temp_Test:
                            Stop_TEMP_TEST();
                            break;
                        default:
                            break;
                    }
                    break;
                default:
                    break;
            }
        }

        #region Function for Chip Test
        private void Run_Reset_SPI(object sender, EventArgs e)
        {
            Reset_SPI();
        }

        private void Reset_SPI() => SPI.Toggle_GPIO3();
        #endregion

        #region Function for Temp Test (POR + PLL)
        private BackgroundWorker temptestWorker;
        private TimeSpan lastCycleDuration = TimeSpan.Zero;

        private void Check_Instrument()
        {
            foreach (JLcLib.Custom.InsInformation Ins in JLcLib.Custom.InstrumentForm.InsInfoList)
            {
                switch (Ins.Type)
                {
                    case JLcLib.Custom.InstrumentTypes.PowerSupply0:
                        if (PowerSupply0 == null) PowerSupply0 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!PowerSupply0.IsOpen) PowerSupply0.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.PowerSupply1:
                        if (PowerSupply1 == null) PowerSupply1 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!PowerSupply1.IsOpen) PowerSupply1.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.DigitalMultimeter0:
                        if (DigitalMultimeter0 == null) DigitalMultimeter0 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!DigitalMultimeter0.IsOpen) DigitalMultimeter0.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.DigitalMultimeter1:
                        if (DigitalMultimeter1 == null) DigitalMultimeter1 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!DigitalMultimeter1.IsOpen) DigitalMultimeter1.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.DigitalMultimeter2:
                        if (DigitalMultimeter2 == null) DigitalMultimeter2 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!DigitalMultimeter2.IsOpen) DigitalMultimeter2.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.DigitalMultimeter3:
                        if (DigitalMultimeter3 == null) DigitalMultimeter3 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!DigitalMultimeter3.IsOpen) DigitalMultimeter3.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.OscilloScope0:
                        if (OscilloScope0 == null) OscilloScope0 = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!OscilloScope0.IsOpen) OscilloScope0.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.SourceAnalyzer:
                        if (SourceAnalyzer == null) SourceAnalyzer = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!SourceAnalyzer.IsOpen) SourceAnalyzer.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.TempChamber:
                        if (TempChamber == null) TempChamber = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!TempChamber.IsOpen) TempChamber.Open();
                        break;
                    case JLcLib.Custom.InstrumentTypes.WaveformGenerator:
                        if (WaveformGenerator == null) WaveformGenerator = new JLcLib.Custom.SCPI(Ins.Type);
                        if (!WaveformGenerator.IsOpen) WaveformGenerator.Open();
                        break;
                }
            }
        }

        private void SafeSleep(TimeSpan duration, Func<bool> isCancelled)
        {
            int slice = 100;
            int total = (int)duration.TotalMilliseconds;
            int elapsed = 0;

            while (elapsed < total)
            {
                if (isCancelled != null && isCancelled())
                    throw new OperationCanceledException();

                int wait = Math.Min(slice, total - elapsed);
                Thread.Sleep(wait);
                elapsed += wait;
            }
        }

        public void Start_TEMP_TEST(int initialRowOffset)
        {
            if (temptestWorker != null && temptestWorker.IsBusy)
            {
                Log.WriteLine("자동화 테스트가 이미 실행 중입니다.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            temptestWorker = new BackgroundWorker
            {
                WorkerSupportsCancellation = true,
                WorkerReportsProgress = true
            };

            temptestWorker.DoWork += TEMP_TEST_Worker_DoWork;

            temptestWorker.ProgressChanged += (s, e) =>
            {
                try
                {
                    StatusBar?.GetCurrentParent()?.Invoke(new MethodInvoker(delegate
                    {
                        StatusBar.Text = $"{e.UserState as string}";
                    }));
                }
                catch
                {

                }
            };

            temptestWorker.RunWorkerCompleted += (s, e) =>
            {
                try
                {
                    if (e.Cancelled || e.Error is OperationCanceledException)
                        Log.WriteLine("TEMP_TEST가 사용자에 의해 취소되었습니다.", Color.Orange, Parent.BackColor);
                    else if (e.Error != null)
                        Log.WriteLine($"TEMP_TEST 중 오류 발생: {e.Error.Message}", Color.Red, Parent.BackColor);
                    else
                        Log.WriteLine("TEMP_TEST 시퀀스가 종료되었습니다.", Color.Green, Parent.BackColor);

                    StatusBar?.GetCurrentParent()?.Invoke(new MethodInvoker(delegate
                    {
                        StatusBar.Text = "Ready.";
                    }));
                }
                catch
                {

                }
            };

            Log.WriteLine($"TEMP_TEST 시작 (Initial Row Offset: {initialRowOffset})...", Color.Cyan, Parent.BackColor);

            try
            {
                StatusBar?.GetCurrentParent()?.Invoke(new MethodInvoker(delegate
                {
                    StatusBar.Text = "TEMP_TEST 시작…";
                }));
            }
            catch
            {

            }

            temptestWorker.RunWorkerAsync(initialRowOffset);
        }

        public void Stop_TEMP_TEST()
        {
            if (temptestWorker != null && temptestWorker.IsBusy)
            {
                temptestWorker.CancelAsync();
                Log.WriteLine("TEMP_TEST 중지 요청을 보냈습니다...", Color.Orange, Parent.BackColor);

                try
                {
                    StatusBar?.GetCurrentParent()?.Invoke(new MethodInvoker(delegate
                    {
                        StatusBar.Text = "TEMP_TEST 중지 요청 중...";
                    }));
                }
                catch
                {

                }
            }
            else
            {
                Log.WriteLine("현재 실행 중인 TEMP_TEST가 없습니다.", Color.OrangeRed, Parent.BackColor);
            }
        }

        private void TEMP_TEST_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = (BackgroundWorker)sender;
            int rowOffset = (int)e.Argument;

            try
            {
                while (!worker.CancellationPending)
                {
                    worker.ReportProgress(0, $"Chip {rowOffset + 1} 테스트 시작...");

                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    Execute_TEMP_TEST_Single(worker, rowOffset);
                    sw.Stop();

                    lastCycleDuration = sw.Elapsed;

                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }

                    var secs = lastCycleDuration.TotalSeconds.ToString("F1", System.Globalization.CultureInfo.InvariantCulture);
                    worker.ReportProgress(100, $"Chip {rowOffset + 1} 완료 ({secs}s).");

                    DialogResult result = DialogResult.None;
                    try
                    {
                        Parent.Invoke(new MethodInvoker(delegate
                        {
                            string msg =
                                $"Chip Num : {rowOffset + 1} 테스트 완료.\r\n" +
                                $"마지막 사이클 소요 시간: {secs} 초\r\n" +
                                $"Next Chip Num : {rowOffset + 2}\r\n" +
                                $"다음 테스트를 이어서 진행할까요?";
                            result = MessageBox.Show(msg, "TEMP_TEST", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        }));
                    }
                    catch
                    {
                        e.Cancel = true;
                        break;
                    }

                    if (result == DialogResult.No)
                    {
                        e.Cancel = true;
                        break;
                    }

                    rowOffset++;
                }
            }
            catch (OperationCanceledException)
            {
                e.Cancel = true;
            }
        }

        private void Execute_TEMP_TEST_Single(BackgroundWorker worker, int colOffset)
        {
            try
            {
                Check_Instrument();

                double dmm_v, vdd_v, dmm_a;
                double[] volts = new double[3];
                var inv = CultureInfo.InvariantCulture;

                int FailCount = new int();

                for (int ch = 1; ch <= 3; ch++)
                {
                    PowerSupply0.Write($"OUTP OFF, (@{ch})");
                }
                PowerSupply1.Write("OUTP OFF, (@2)");
                WaveformGenerator.Write($"OUTP1 OFF");

                #region NT_TEST @ 23.5℃

                #region POR_TEST
                var ok = SetAndWaitTemp(
                    target: 23.5,
                    tolerance: 0.1,
                    pollInterval: TimeSpan.FromMilliseconds(10000),
                    stableHold: TimeSpan.FromSeconds(60),
                    timeout: TimeSpan.FromMinutes(120),
                    soak: TimeSpan.FromMinutes(3),
                    isCancelled: () => worker.CancellationPending);
                if (!ok) return;

                Parent.xlMgr.Sheet.Select("POR_NT");
                Parent.xlMgr.Sheet.Activate();

                PowerSupply1.Write("SOUR:VOLT:SLEW:IMM 25, (@2)");
                PowerSupply1.Write("SOUR:VOLT:LEV:IMM:AMPL 0, (@2)");
                PowerSupply1.Write("OUTP ON, (@2)");

                for (int i = 0; i < 17; i++)
                {
                    if (worker.CancellationPending) return;

                    string vdd_str = Parent.xlMgr.Cell.Read(2, 6 + i);
                    if (!double.TryParse(vdd_str, NumberStyles.Float, inv, out vdd_v))
                        throw new FormatException($"Excel {2 + colOffset}행, {6 + i}열 전압값이 잘못되었습니다: '{vdd_str}'");

                    worker.ReportProgress(i * 100 / 17, $"VDD {vdd_v.ToString("F3", inv)} V...");
                    PowerSupply1.Write($"SOUR:VOLT:LEV:IMM:AMPL {vdd_v.ToString(inv)}, (@2)");

                    SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);

                    worker.ReportProgress(i * 100 / 17, "Measure POR_OUT...");
                    dmm_v = double.Parse(DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?"), inv);

                    Parent.xlMgr.Cell.Write(4 + colOffset, 6 + i, dmm_v.ToString("F4", inv));
                }

                PowerSupply1.Write("OUTP OFF, (@2)");

                SafeSleep(TimeSpan.FromSeconds(1), () => temptestWorker?.CancellationPending == true);

                #endregion POR_TEST

                #region PLL_TEST
                worker.ReportProgress(40, "PLL Testing...");
                Parent.xlMgr.Sheet.Select("PLL_NV_NT");
                Parent.xlMgr.Sheet.Activate();

                volts = new double[] { 5.0, 1.2, 0.75 };

                Set_PowerSupply0(volts, 1);

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 7, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 8, dmm_a.ToString("F3", inv));

                Set_WaveformGenerator(12.5, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 10, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 11, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 12, dmm_v.ToString("F3", inv));

                string carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                string[] carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, "NA");
                }

                double m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 14, m4y.ToString("F3", inv));

                double m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 15, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(440, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 17, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 18, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 19, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 21, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 22, m5y.ToString("F3", inv));

                WaveformGenerator.Write($"OUTP1 OFF");

                Set_WaveformGenerator(25, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(110, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 24, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 25, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 26, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 28, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 29, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }


                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 31, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 32, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 33, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 35, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 36, m5y.ToString("F3", inv));

                for (int ch = 1; ch <= 3; ch++)
                {
                    PowerSupply0.Write($"OUTP OFF, (@{ch})");
                }
                PowerSupply1.Write("OUTP OFF, (@2)");
                WaveformGenerator.Write($"OUTP1 OFF");

                Log.WriteLine("Test Complete - NT.", Color.LightGreen, Parent.BackColor);
                worker.ReportProgress(100, "NT 단계 완료. LT 단계로 진행합니다.");
                #endregion PLL_TEST

                #endregion NT_TEST @ 23.5℃

                #region LT_TEST @ -40℃

                #region POR_TEST
                ok = SetAndWaitTemp(
                    target: -40,
                    tolerance: 0.1,
                    pollInterval: TimeSpan.FromMilliseconds(10000),
                    stableHold: TimeSpan.FromSeconds(60),
                    timeout: TimeSpan.FromMinutes(120),
                    soak: TimeSpan.FromMinutes(3),
                    isCancelled: () => worker.CancellationPending);
                if (!ok) return;

                Parent.xlMgr.Sheet.Select("POR_LT");
                Parent.xlMgr.Sheet.Activate();

                PowerSupply1.Write("SOUR:VOLT:SLEW:IMM 25, (@2)");
                PowerSupply1.Write("SOUR:VOLT:LEV:IMM:AMPL 0, (@2)");
                PowerSupply1.Write("OUTP ON, (@2)");

                for (int i = 0; i < 17; i++)
                {
                    if (worker.CancellationPending) return;

                    string vdd_str = Parent.xlMgr.Cell.Read(2, 6 + i);
                    if (!double.TryParse(vdd_str, NumberStyles.Float, inv, out vdd_v))
                        throw new FormatException($"Excel {2 + colOffset}행, {6 + i}열 전압값이 잘못되었습니다: '{vdd_str}'");

                    worker.ReportProgress(i * 100 / 17, $"VDD {vdd_v.ToString("F3", inv)} V...");
                    PowerSupply1.Write($"SOUR:VOLT:LEV:IMM:AMPL {vdd_v.ToString(inv)}, (@2)");

                    SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);

                    worker.ReportProgress(i * 100 / 17, "Measure POR_OUT...");
                    dmm_v = double.Parse(DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?"), inv);

                    Parent.xlMgr.Cell.Write(4 + colOffset, 6 + i, dmm_v.ToString("F4", inv));
                }

                PowerSupply1.Write("OUTP OFF, (@2)");

                SafeSleep(TimeSpan.FromSeconds(1), () => temptestWorker?.CancellationPending == true);
                #endregion POR_TEST

                #region PLL_TEST
                worker.ReportProgress(40, "PLL Testing...");
                Parent.xlMgr.Sheet.Select("PLL_LV_LT");
                Parent.xlMgr.Sheet.Activate();

                volts = new double[] { 5.0, 1.08, 0.675 };

                Set_PowerSupply0(volts, 1);
                Reset_SPI();

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 7, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 8, dmm_a.ToString("F3", inv));

                Set_WaveformGenerator(12.5, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 10, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 11, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 12, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 14, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 15, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(440, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 17, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 18, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 19, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 21, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 22, m5y.ToString("F3", inv));

                WaveformGenerator.Write($"OUTP1 OFF");

                Set_WaveformGenerator(25, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(110, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 24, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 25, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 26, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 28, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 29, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 31, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 32, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 33, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 35, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 36, m5y.ToString("F3", inv));

                for (int ch = 1; ch <= 3; ch++)
                {
                    PowerSupply0.Write($"OUTP OFF, (@{ch})");
                }
                PowerSupply1.Write("OUTP OFF, (@2)");
                WaveformGenerator.Write($"OUTP1 OFF");

                Parent.xlMgr.Sheet.Select("PLL_HV_LT");
                Parent.xlMgr.Sheet.Activate();

                volts = new double[] { 5.0, 1.32, 0.9825 };

                Set_PowerSupply0(volts, 1);
                Reset_SPI();

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 7, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 8, dmm_a.ToString("F3", inv));

                Set_WaveformGenerator(12.5, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 10, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 11, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 12, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 14, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 15, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(440, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 17, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 18, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 19, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 21, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 22, m5y.ToString("F3", inv));

                WaveformGenerator.Write($"OUTP1 OFF");

                Set_WaveformGenerator(25, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(110, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 24, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 25, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 26, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 28, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 29, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 31, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 32, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 33, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 35, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 36, m5y.ToString("F3", inv));

                for (int ch = 1; ch <= 3; ch++)
                {
                    PowerSupply0.Write($"OUTP OFF, (@{ch})");
                }
                PowerSupply1.Write("OUTP OFF, (@2)");
                WaveformGenerator.Write($"OUTP1 OFF");

                Log.WriteLine("Test Complete - LT.", Color.LightGreen, Parent.BackColor);
                worker.ReportProgress(100, "LT 단계 완료. HT 단계로 진행합니다.");
                #endregion PLL_TEST

                #endregion LT_TEST @ -40℃

                #region HT_TEST @ 85℃

                #region POR_TEST
                ok = SetAndWaitTemp(
                    target: 85,
                    tolerance: 0.1,
                    pollInterval: TimeSpan.FromMilliseconds(10000),
                    stableHold: TimeSpan.FromSeconds(60),
                    timeout: TimeSpan.FromMinutes(120),
                    soak: TimeSpan.FromMinutes(3),
                    isCancelled: () => worker.CancellationPending);
                if (!ok) return;

                Parent.xlMgr.Sheet.Select("POR_HT");
                Parent.xlMgr.Sheet.Activate();

                PowerSupply1.Write("SOUR:VOLT:SLEW:IMM 25, (@2)");
                PowerSupply1.Write("SOUR:VOLT:LEV:IMM:AMPL 0, (@2)");
                PowerSupply1.Write("OUTP ON, (@2)");

                for (int i = 0; i < 17; i++)
                {
                    if (worker.CancellationPending) return;

                    string vdd_str = Parent.xlMgr.Cell.Read(2, 6 + i);
                    if (!double.TryParse(vdd_str, NumberStyles.Float, inv, out vdd_v))
                        throw new FormatException($"Excel {2 + colOffset}행, {6 + i}열 전압값이 잘못되었습니다: '{vdd_str}'");

                    worker.ReportProgress(i * 100 / 17, $"VDD {vdd_v.ToString("F3", inv)} V...");
                    PowerSupply1.Write($"SOUR:VOLT:LEV:IMM:AMPL {vdd_v.ToString(inv)}, (@2)");

                    SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);

                    worker.ReportProgress(i * 100 / 17, "Measure POR_OUT...");
                    dmm_v = double.Parse(DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?"), inv);

                    Parent.xlMgr.Cell.Write(4 + colOffset, 6 + i, dmm_v.ToString("F4", inv));
                }

                PowerSupply1.Write("OUTP OFF, (@2)");

                SafeSleep(TimeSpan.FromSeconds(1), () => temptestWorker?.CancellationPending == true);
                #endregion POR_TEST

                #region PLL_TEST
                worker.ReportProgress(40, "PLL Testing...");
                Parent.xlMgr.Sheet.Select("PLL_LV_HT");
                Parent.xlMgr.Sheet.Activate();

                volts = new double[] { 5.0, 1.08, 0.675 };

                Set_PowerSupply0(volts, 1);
                Reset_SPI();

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 7, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 8, dmm_a.ToString("F3", inv));

                Set_WaveformGenerator(12.5, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 10, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 11, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 12, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 14, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 15, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(440, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 17, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 18, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 19, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 21, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 22, m5y.ToString("F3", inv));

                WaveformGenerator.Write($"OUTP1 OFF");

                Set_WaveformGenerator(25, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(110, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 24, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 25, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 26, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 28, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 29, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 31, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 32, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 33, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 35, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 36, m5y.ToString("F3", inv));

                for (int ch = 1; ch <= 3; ch++)
                {
                    PowerSupply0.Write($"OUTP OFF, (@{ch})");
                }
                PowerSupply1.Write("OUTP OFF, (@2)");
                WaveformGenerator.Write($"OUTP1 OFF");

                Parent.xlMgr.Sheet.Select("PLL_HV_HT");
                Parent.xlMgr.Sheet.Activate();

                volts = new double[] { 5.0, 1.32, 0.9825 };

                Set_PowerSupply0(volts, 1);
                Reset_SPI();

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 7, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 8, dmm_a.ToString("F3", inv));

                Set_WaveformGenerator(12.5, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 10, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 11, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 12, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 13, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 14, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 15, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(440, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 17, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 18, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 19, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 20, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 21, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 22, m5y.ToString("F3", inv));

                WaveformGenerator.Write($"OUTP1 OFF");

                Set_WaveformGenerator(25, 0.75, 0.375, 50);
                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(110, 1, 5))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 24, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 25, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 26, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 27, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 28, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 29, m5y.ToString("F3", inv));

                FailCount = 0;
                while (true)
                {
                    if (Set_PLL_Configuration(220, 1, 6))
                    {
                        break;
                    }
                    else
                    {
                        FailCount++;
                        Log.WriteLine($"Fail to Write Register. Fail Count: {FailCount}", Color.OrangeRed, Parent.BackColor);
                        if (FailCount >= 100)
                        {
                            throw new FormatException($"Register Write가 정상적으로 이루어지지 않습니다.");
                        }
                        for (int ch = 1; ch <= 3; ch++)
                        {
                            PowerSupply0.Write($"OUTP OFF, (@{ch})");
                        }
                        SafeSleep(TimeSpan.FromSeconds(1), () => worker?.CancellationPending == true);
                        Set_PowerSupply0(volts, 1);
                    }
                }

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 31, dmm_a.ToString("F3", inv));

                dmm_a = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)"), inv) * 1000;
                Parent.xlMgr.Cell.Write(4 + colOffset, 32, dmm_a.ToString("F3", inv));

                dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 33, dmm_v.ToString("F3", inv));

                carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
                carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (carrTokens.Length >= 2)
                {
                    double carrHz = double.Parse(carrTokens[0], inv);
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, (carrHz / 1e6).ToString("F6", inv));
                }
                else
                {
                    Parent.xlMgr.Cell.Write(4 + colOffset, 34, "NA");
                }

                m4y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK4:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 35, m4y.ToString("F3", inv));

                m5y = double.Parse(SourceAnalyzer.WriteAndReadString(":CALC:PN1:TRAC1:MARK5:Y?"), inv);
                Parent.xlMgr.Cell.Write(4 + colOffset, 36, m5y.ToString("F3", inv));

                for (int ch = 1; ch <= 3; ch++)
                {
                    PowerSupply0.Write($"OUTP OFF, (@{ch})");
                }
                PowerSupply1.Write("OUTP OFF, (@2)");
                WaveformGenerator.Write($"OUTP1 OFF");

                Log.WriteLine("Test Complete - HT.", Color.LightGreen, Parent.BackColor);
                worker.ReportProgress(100, "HT 단계 완료. 테스트 마무리...");
                #endregion PLL_TEST

                #endregion HT_TEST @ 85℃
            }
            finally
            {
                string resp = TempChamber.WriteAndReadString("01,TEMP,S" + 23.5.ToString(CultureInfo.InvariantCulture));
            }
        }

        private void Set_PowerSupply0(double[] volts, double limit_i)
        {
            var inv = CultureInfo.InvariantCulture;

            for (int ch = 0; ch < 3; ch++)
            {
                PowerSupply0.Write($"VOLT {volts[ch].ToString(inv)}, (@{ch + 1})");
                PowerSupply0.Write($"CURR {limit_i.ToString(inv)}, (@{ch + 1})");
                PowerSupply0.Write($"OUTP ON, (@{ch + 1})");
            }

            SafeSleep(TimeSpan.FromSeconds(1), () => temptestWorker?.CancellationPending == true);
        }

        private void Set_WaveformGenerator(double freqMHz, double ampVpp, double offsetV, double dutyPer)
        {
            var inv = CultureInfo.InvariantCulture;

            WaveformGenerator.Write("FUNC SQU");
            WaveformGenerator.Write("PHAS 0");
            WaveformGenerator.Write($"FREQ {freqMHz.ToString(inv)}E6");
            WaveformGenerator.Write($"VOLT {ampVpp.ToString(inv)}");
            WaveformGenerator.Write($"VOLT:OFFS {offsetV.ToString(inv)}");
            WaveformGenerator.Write($"FUNC:SQU:DCYC {dutyPer.ToString(inv)}");
            WaveformGenerator.Write("OUTP1 ON");

            SafeSleep(TimeSpan.FromSeconds(1), () => temptestWorker?.CancellationPending == true);
        }

        private bool Set_PLL_Configuration(uint PF_M, uint PF_P, uint PF_S, int burstWrites = 20, TimeSpan? interWriteDelay = null)
        {
            var inv = CultureInfo.InvariantCulture;
            var delay = interWriteDelay ?? TimeSpan.FromMilliseconds(100);

            uint reg02 = 0x0000 | (PF_M & 0x3FF);
            uint reg05 = 0x0000 | ((PF_P & 0x3F) << 10) | ((PF_S & 0x07) << 7);
            uint reg01 = 0xFE00 | (1u << 8);
            double dmm_v, carrHz = new double();

            bool IsWritten = false;

            Reset_SPI();

            for (int i = 0; i < burstWrites; i++)
            {
                if (temptestWorker?.CancellationPending == true)
                    return false;

                WriteRegister(0x02, reg02);
                SafeSleep(delay, () => temptestWorker?.CancellationPending == true);

                WriteRegister(0x05, reg05);
                SafeSleep(delay, () => temptestWorker?.CancellationPending == true);

                WriteRegister(0x01, reg01);
                SafeSleep(delay, () => temptestWorker?.CancellationPending == true);
            }

            dmm_v = double.Parse(DigitalMultimeter1.WriteAndReadString("MEAS:VOLT:DC?"), inv);

            string carr = SourceAnalyzer.WriteAndReadString(":CALC:PN1:DATA:CARR?");
            string[] carrTokens = carr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            if (carrTokens.Length >= 2)
            {
                carrHz = double.Parse(carrTokens[0], inv);
            }

            if ((dmm_v >= 0.720 && dmm_v <= 0.780) && ((carrHz / 1e6) >= 85 && (carrHz / 1e6) <= 86))
            {
                IsWritten = true;
            }

            return IsWritten;
        }

        public bool SetAndWaitTemp(
            double target,
            double tolerance = 0.1,
            TimeSpan? pollInterval = null,
            TimeSpan? stableHold = null,
            TimeSpan? timeout = null,
            TimeSpan? soak = null,
            Func<bool> isCancelled = null)
        {
            var inv = CultureInfo.InvariantCulture;

            if (!Enable_TempChamber)
            {
                Log.WriteLine("Temperature control skipped. (flag = 0)");
                return true;
            }

            var poll = pollInterval ?? TimeSpan.FromSeconds(1);
            var hold = stableHold ?? TimeSpan.FromSeconds(15);
            var to = timeout ?? TimeSpan.FromMinutes(10);
            var soakTime = soak ?? TimeSpan.FromMinutes(5);

            string resp = TempChamber.WriteAndReadString("01,TEMP,S" + target.ToString(inv));
            SafeSleep(TimeSpan.FromSeconds(1), isCancelled);

            resp = TempChamber.WriteAndReadString("TEMP?");
            SafeSleep(TimeSpan.FromMilliseconds(100), isCancelled);

            var sw = Stopwatch.StartNew();
            int stableCount = 0;
            int needStable = Math.Max(1, (int)Math.Ceiling(hold.TotalMilliseconds / poll.TotalMilliseconds));
            double lastReal = double.NaN, lastSet = double.NaN, abs = double.NaN;

            while (true)
            {
                if (isCancelled != null && isCancelled()) return false;

                resp = TempChamber.WriteAndReadString("TEMP?");
                SafeSleep(TimeSpan.FromMilliseconds(100), isCancelled);

                var vals = ParseFour(resp);
                if (vals.ok)
                {
                    lastReal = vals.a;
                    lastSet = vals.b;
                    abs = Math.Round(Math.Abs(lastReal - target), 2);
                    bool within = abs <= tolerance;
                    stableCount = within ? stableCount + 1 : 0;

                    Log.WriteLine(
                        $"RealTemp : {lastReal.ToString("0.00", inv)} | " +
                        $"SetTemp : {lastSet.ToString("0.00", inv)}\n" +
                        $"stableCount : {stableCount.ToString()} | " +
                        $"needStable : {needStable.ToString()}");

                    if (stableCount >= needStable) break;
                }

                if (sw.Elapsed > to)
                    throw new TimeoutException(
                        $"Temperature did not stabilize at {target.ToString(inv)}±{tolerance.ToString(inv)} within {to}.");

                SafeSleep(poll, isCancelled);
            }

            Log.WriteLine($"Stabilization waiting time starts for {soakTime.TotalMinutes.ToString("0", inv)} minutes.");
            SafeSleep(soakTime, isCancelled);

            Log.WriteLine(
                $"Stabilization waiting time is {soakTime.TotalMinutes.ToString("0", inv)} minutes complete.\n" +
                $"RealTemp : {lastReal.ToString("0.000", inv)} | SetTemp : {lastSet.ToString("0.000", inv)}"
            );

            return true;
        }

        private (bool ok, double a, double b, double c, double d) ParseFour(string s)
        {
            var inv = System.Globalization.CultureInfo.InvariantCulture;

            if (string.IsNullOrWhiteSpace(s)) return (false, 0, 0, 0, 0);
            var parts = s.Split(',');
            if (parts.Length < 2) return (false, 0, 0, 0, 0);

            double a = new double(), b = new double(), c = new double(), d = new double();
            bool okA = double.TryParse(parts[0].Trim(), NumberStyles.Float, inv, out a);
            bool okB = double.TryParse(parts[1].Trim(), NumberStyles.Float, inv, out b);
            bool okC = parts.Length > 2 && double.TryParse(parts[2].Trim(), NumberStyles.Float, inv, out c);
            bool okD = parts.Length > 3 && double.TryParse(parts[3].Trim(), NumberStyles.Float, inv, out d);

            if (!okC) c = 0;
            if (!okD) d = 0;

            return (okA && okB, a, b, c, d);
        }
        #endregion Function for Temp Test (POR + PLL)
    }

    public class IRIS2 : ChipControl
    {
        #region Variable and declaration

        public enum TEST_ITEMS_MANUAL
        {
            TEST,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            TEST,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
        }

        private JLcLib.Custom.SPI SPI = null;

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();

        private bool IsSerialReceivedData = false;
        private bool IsRunCal = false;

        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter1 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter2 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter3 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI SpectrumAnalyzer = null;
        JLcLib.Instrument.SCPI TempChamber = null;

        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.MANUAL;

        #endregion Variable and declaration

        public IRIS2(RegContForm form) : base(form)
        {
            SPI = Parent.SPI;

            Serial.ReadSettingFile(form.IniFile, "IRIS2");

            Serial.DataReceived += Serial_DataReceived;

            CalibrationData = new byte[256];

            /* UI의 테스트 항목 콤보박스를 초기화합니다. */
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e)
        {
            IsSerialReceivedData = true;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] Bytes = new byte[3];
            byte rwFlag = 1; // 쓰기 플래그

            Bytes[0] = (byte)((rwFlag << 7) | (Address & 0x7F));
            Bytes[1] = (byte)((Data >> 8) & 0xff);
            Bytes[2] = (byte)((Data >> 0) & 0xff);

            SPI.WriteBytes(Bytes, 3, true);
        }

        private uint ReadRegister(uint Address)
        {
            byte rwFlag = 0; // 읽기 플래그
            uint Data = 0xFFFF;
            byte[] Bytes = new byte[1];
            byte[] Buff = new byte[2];

            Bytes[0] = (byte)((rwFlag << 7) | (Address & 0x7F));

            Buff = SPI.WriteAndReadBytes(Bytes, 1, 2);

            Data = (uint)((Buff[0] << 8) | (Buff[1] << 0));

            return Data;
        }

        public override JLcLib.Chip.RegisterGroup MakeRegisterGroup(string GroupName, string[,] RegData)
        {
            JLcLib.Chip.RegisterGroup rg = new JLcLib.Chip.RegisterGroup(GroupName, WriteRegister, ReadRegister);

            for (int xStart = 0; xStart < 3; xStart++)
            {
                for (int Row = 0; Row < RegData.GetLength(0); Row++)
                {
                    if ((Row + 2 < RegData.GetLength(0)) && (RegData[Row, xStart] == "Bit") && (RegData[Row + 1, xStart] == "Name") && (RegData[Row + 2, xStart] == "Default"))
                    {
                        string RegName = null;
                        uint Address = 0;

                        string StrAddr = RegData[Row - 1, xStart + 1];
                        if (StrAddr != null && StrAddr.StartsWith("0x"))
                            StrAddr = StrAddr.Substring(2);

                        if (uint.TryParse(StrAddr, System.Globalization.NumberStyles.HexNumber, null, out Address))
                        {
                            RegName = RegData[Row - 1, xStart + 2];
                            JLcLib.Chip.Register reg = rg.Registers.Add(RegName, Address);

                            for (int Column = xStart + 1; Column < RegData.GetLength(1); Column++)
                            {
                                if (!((RegData[Row + 2, Column] == "X") || (RegData[Row + 2, Column] == "-") || (RegData[Row + 2, Column] == null) || (RegData[Row + 1, Column] == null)))
                                {
                                    string ItemName = null, ItemDesc = null;
                                    int UpperBit = 0, LowerBit = 0;
                                    uint ItemValue = 0;

                                    ItemName = RegData[Row + 1, Column];
                                    UpperBit = int.Parse(RegData[Row, Column]);
                                    LowerBit = int.Parse(RegData[Row, Column]);
                                    ItemValue = (uint)((RegData[Row + 2, Column] == "0") ? 0 : 1);

                                    for (int x = Column + 1; x < RegData.GetLength(1); x++)
                                    {
                                        if (RegData[Row + 1, x] == null)
                                        {
                                            if (RegData[Row, x] != null)
                                            {
                                                LowerBit = int.Parse(RegData[Row, x]);
                                                ItemValue = (ItemValue << 1) | (uint)((RegData[Row + 2, x] == "0") ? 0 : 1);
                                            }
                                        }
                                        else break;
                                    }

                                    for (int y = Row; y < RegData.GetLength(0); y++)
                                    {
                                        if (RegData[y, xStart] == ItemName)
                                        {
                                            ItemDesc = RegData[y, xStart + 1];

                                            for (int desc_row = y + 1; desc_row < RegData.GetLength(0); desc_row++)
                                            {
                                                if (RegData[desc_row, xStart] == null)
                                                {
                                                    string lineDesc = "";
                                                    if (RegData[desc_row, xStart + 3] != null && RegData[desc_row, xStart + 4] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 3] + "=" + RegData[desc_row, xStart + 4];
                                                    else if (RegData[desc_row, xStart + 4] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 4];
                                                    else if (RegData[desc_row, xStart + 3] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 3] + "=";

                                                    ItemDesc += lineDesc;
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }
                                            break; // 해당 ItemName에 대한 설명 파싱 완료
                                        }
                                    }
                                    reg.Items.Add(ItemName, UpperBit, LowerBit, ItemValue, ItemDesc);
                                }
                            }
                        }
                    }
                }
            }
            return rg;
        }

        public override void SetChipSpecificUI()
        {
            /*
             * 이 공간에 IRIS2 칩에만 필요한 버튼이나 텍스트 박스 등의 UI 컨트롤을
             * 활성화하고 이벤트 핸들러를 할당하는 코드를 작성합니다.
             */
        }

        private void SerialConnect_Click(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (Serial.IsOpen == false)
            {
                Serial.Open();
                if (Serial.IsOpen)
                {
                    b.Text = "Disconn";
                    Parent.ChipCtrlTextboxes[0].Text = Serial.Config.PortName + "-" + Serial.Config.BaudRate.ToString();
                }
            }
            else
            {
                Serial.Close();
                if (!Serial.IsOpen)
                    b.Text = "Connect";
            }
        }

        private void SerialSetting_Click(object sender, EventArgs e)
        {
            JLcLib.Comn.WireComnForm.Show(Serial);
            Serial.WriteSettingFile(Parent.IniFile, "IRIS2");
        }

        private void Change_To_Manual_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.MANUAL;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Change_To_Auto_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.AUTO;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_AUTO.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_AUTO)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
            if (Serial.IsOpen && (Serial.RcvQueue.Count != 0))
            {
                for (int i = 0; i < Serial.RcvQueue.Count; i++)
                {
                    byte b = Serial.RcvQueue.Get();
                    if ((b > 0x00 && b < 0x80) && b != '\0')
                    {
                        LogMessage += (char)b;
                        if (b == '\n')
                        {
                            Log.WriteWithElapsedTime(LogMessage);
                            Log.WriteLine("\n");
                            LogMessage = "";
                        }
                    }
                }
            }
        }

        public override void SendCommand(string Command)
        {
            Command += "\r\n";
            byte[] Var = Encoding.ASCII.GetBytes(Command);
            Serial.WriteBytes(Var, Var.Length, true);
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal;
            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.MANUAL:
                    switch ((TEST_ITEMS_MANUAL)TestItemIndex)
                    {
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((TEST_ITEMS_AUTO)TestItemIndex)
                    {
                        default:
                            break;
                    }
                    break;
                default:
                    break;
            }
        }

        #region Function for Chip Test
        #endregion
    }

    public class Chicago : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS_MANUAL
        {
            Write_SRAM,
            Read_SRAM,
            //Dummy321Bytes,
            VoltageSweep_Start,
            VoltageSweep_Stop,
            //AiP33A09_Config,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            Initial_Test,
            //SET_ANA_125,
            //SET_ANA_150,
            //SET_ANA_PTAT,
            //SET_ANA_OSC,
            //Disable_ANA,
            //RUN_TRIM_125,
            //RUN_TRIM_150,
            //RUN_TRIM_PTAT,
            //RUN_TRIM_OSC,
            //SET_SEG_TEST,
            //Disable_SEG_TEST,
            //RUN_Write_EFUSE,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_LED
        {
            Stop_LED_Effect,
            WaveEffect_Start,
            ScrollText_Start,
            BreathingEffect_Start,
            ClockEffect_Start,
            FireworkEffect_Start,
            //Display_Number,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_CMD
        {
            Update_Display,
            Write_Display,
            TurnOff_Display,
            WakeUp,
            Sleep,
            Reset,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
            LED,
            CMD,
        }

        private JLcLib.Custom.SPI SPI = null;

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();

        private bool IsSerialReceivedData = false;

        private BackgroundWorker animationWorker = null;
        private BackgroundWorker voltageSweepWorker = null;

        #region SCPI Instrument
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter1 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter2 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter3 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI SpectrumAnalyzer = null;
        JLcLib.Instrument.SCPI TempChamber = null;
        #endregion SCPI Instrument

        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.MANUAL;

        #endregion Variable and declaration

        public Chicago(RegContForm form) : base(form)
        {
            SPI = Parent.SPI;
            Serial.ReadSettingFile(form.IniFile, "Chicago");
            Serial.DataReceived += Serial_DataReceived;
            CalibrationData = new byte[256];

            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private readonly Dictionary<char, byte[]> Font = new Dictionary<char, byte[]>()
        {
            {'A', new byte[]{0x7E, 0x11, 0x11, 0x11, 0x7E}},
            {'B', new byte[]{0x7F, 0x49, 0x49, 0x49, 0x36}},
            {'C', new byte[]{0x3E, 0x41, 0x41, 0x41, 0x22}},
            {'D', new byte[]{0x7F, 0x41, 0x41, 0x22, 0x1C}},
            {'E', new byte[]{0x7F, 0x49, 0x49, 0x49, 0x41}},
            {'F', new byte[]{0x7F, 0x09, 0x09, 0x09, 0x01}},
            {'G', new byte[]{0x3E, 0x41, 0x49, 0x49, 0x7A}},
            {'H', new byte[]{0x7F, 0x08, 0x08, 0x08, 0x7F}},
            {'I', new byte[]{0x00, 0x41, 0x7F, 0x41, 0x00}},
            {'J', new byte[]{0x20, 0x40, 0x41, 0x3F, 0x01}},
            {'K', new byte[]{0x7F, 0x08, 0x14, 0x22, 0x41}},
            {'L', new byte[]{0x7F, 0x40, 0x40, 0x40, 0x40}},
            {'M', new byte[]{0x7F, 0x02, 0x04, 0x02, 0x7F}},
            {'N', new byte[]{0x7F, 0x04, 0x08, 0x10, 0x7F}},
            {'O', new byte[]{0x3E, 0x41, 0x41, 0x41, 0x3E}},
            {'P', new byte[]{0x7F, 0x09, 0x09, 0x09, 0x06}},
            {'Q', new byte[]{0x3E, 0x41, 0x51, 0x21, 0x5E}},
            {'R', new byte[]{0x7F, 0x09, 0x19, 0x29, 0x46}},
            {'S', new byte[]{0x46, 0x49, 0x49, 0x49, 0x31}},
            {'T', new byte[]{0x01, 0x01, 0x7F, 0x01, 0x01}},
            {'U', new byte[]{0x3F, 0x40, 0x40, 0x40, 0x3F}},
            {'V', new byte[]{0x1F, 0x20, 0x40, 0x20, 0x1F}},
            {'W', new byte[]{0x3F, 0x40, 0x38, 0x40, 0x3F}},
            {'X', new byte[]{0x63, 0x14, 0x08, 0x14, 0x63}},
            {'Y', new byte[]{0x07, 0x08, 0x70, 0x08, 0x07}},
            {'Z', new byte[]{0x61, 0x51, 0x49, 0x45, 0x43}},
            {'0', new byte[]{0x3E, 0x51, 0x49, 0x45, 0x3E}},
            {'1', new byte[]{0x00, 0x42, 0x7F, 0x40, 0x00}},
            {'2', new byte[]{0x42, 0x61, 0x51, 0x49, 0x46}},
            {'3', new byte[]{0x21, 0x41, 0x45, 0x4B, 0x31}},
            {'4', new byte[]{0x18, 0x14, 0x12, 0x7F, 0x10}},
            {'5', new byte[]{0x27, 0x45, 0x45, 0x45, 0x39}},
            {'6', new byte[]{0x3C, 0x4A, 0x49, 0x49, 0x30}},
            {'7', new byte[]{0x01, 0x71, 0x09, 0x05, 0x03}},
            {'8', new byte[]{0x36, 0x49, 0x49, 0x49, 0x36}},
            {'9', new byte[]{0x06, 0x49, 0x49, 0x29, 0x1E}},
            {':', new byte[]{0x00, 0x24, 0x24, 0x00, 0x00}},
            {' ', new byte[]{0x00, 0x00, 0x00, 0x00, 0x00}},
        };

        private static readonly byte[,] DigitPatterns = new byte[10, 8]
        {
            {0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0x00,0xFF}, // 0xFC
            {0x00,0xFF,0xFF,0x00,0x00,0x00,0x00,0xFF}, // 0x60
            {0xFF,0xFF,0x00,0xFF,0xFF,0x00,0xFF,0xFF}, // 0xDA
            {0xFF,0xFF,0xFF,0xFF,0x00,0x00,0xFF,0xFF}, // 0xF2
            {0x00,0xFF,0xFF,0x00,0x00,0xFF,0xFF,0xFF}, // 0x66
            {0xFF,0x00,0xFF,0xFF,0x00,0xFF,0xFF,0xFF}, // 0xB6
            {0xFF,0x00,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF}, // 0xBE
            {0xFF,0xFF,0xFF,0x00,0x00,0x00,0x00,0xFF}, // 0xE0
            {0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF}, // 0xFE
            {0xFF,0xFF,0xFF,0xFF,0x00,0xFF,0xFF,0xFF}  // 0xF6
        };

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e)
        {
            IsSerialReceivedData = true;
        }

        public override JLcLib.Chip.RegisterGroup MakeRegisterGroup(string GroupName, string[,] RegData)
        {
            JLcLib.Chip.RegisterGroup rg = new JLcLib.Chip.RegisterGroup(GroupName, WriteRegister, ReadRegister);

            for (int xStart = 0; xStart < 3; xStart++)
            {
                for (int Row = 0; Row < RegData.GetLength(0); Row++)
                {
                    if ((Row + 2 < RegData.GetLength(0)) && (RegData[Row, xStart] == "Bit") && (RegData[Row + 1, xStart] == "Name") && (RegData[Row + 2, xStart] == "Default"))
                    {
                        string RegName = null;
                        uint Address = 0;

                        string StrAddr = RegData[Row - 1, xStart + 1];
                        if (StrAddr != null && StrAddr.StartsWith("0x"))
                            StrAddr = StrAddr.Substring(2);

                        if (uint.TryParse(StrAddr, System.Globalization.NumberStyles.HexNumber, null, out Address))
                        {
                            RegName = RegData[Row - 1, xStart + 2];
                            JLcLib.Chip.Register reg = rg.Registers.Add(RegName, Address);

                            for (int Column = xStart + 1; Column < RegData.GetLength(1); Column++)
                            {
                                if (!((RegData[Row + 2, Column] == "X") || (RegData[Row + 2, Column] == "-") || (RegData[Row + 2, Column] == null) || (RegData[Row + 1, Column] == null)))
                                {
                                    string ItemName = null, ItemDesc = null;
                                    int UpperBit = 0, LowerBit = 0;
                                    uint ItemValue = 0;

                                    ItemName = RegData[Row + 1, Column];
                                    UpperBit = int.Parse(RegData[Row, Column]);
                                    LowerBit = int.Parse(RegData[Row, Column]);
                                    ItemValue = (uint)((RegData[Row + 2, Column] == "0") ? 0 : 1);

                                    for (int x = Column + 1; x < RegData.GetLength(1); x++)
                                    {
                                        if (RegData[Row + 1, x] == null)
                                        {
                                            if (RegData[Row, x] != null)
                                            {
                                                LowerBit = int.Parse(RegData[Row, x]);
                                                ItemValue = (ItemValue << 1) | (uint)((RegData[Row + 2, x] == "0") ? 0 : 1);
                                            }
                                        }
                                        else break;
                                    }

                                    for (int y = Row; y < RegData.GetLength(0); y++)
                                    {
                                        if (RegData[y, xStart] == ItemName)
                                        {
                                            ItemDesc = RegData[y, xStart + 1];

                                            for (int desc_row = y + 1; desc_row < RegData.GetLength(0); desc_row++)
                                            {
                                                if (RegData[desc_row, xStart] == null)
                                                {
                                                    string lineDesc = "";
                                                    if (RegData[desc_row, xStart + 3] != null && RegData[desc_row, xStart + 4] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 3] + "=" + RegData[desc_row, xStart + 4];
                                                    else if (RegData[desc_row, xStart + 4] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 4];
                                                    else if (RegData[desc_row, xStart + 3] != null)
                                                        lineDesc = "\n" + RegData[desc_row, xStart + 3] + "=";

                                                    ItemDesc += lineDesc;
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }
                                            break; // 해당 ItemName에 대한 설명 파싱 완료
                                        }
                                    }
                                    reg.Items.Add(ItemName, UpperBit, LowerBit, ItemValue, ItemDesc);
                                }
                            }
                        }
                    }
                }
            }
            return rg;
        }

        public override void SetChipSpecificUI()
        {
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */

            //Parent.ChipCtrlButtons[7].Text = "TEST";
            //Parent.ChipCtrlButtons[7].Visible = true;
            //Parent.ChipCtrlButtons[7].Click += TEST;

            //Parent.ChipCtrlButtons[4].Text = "LED_ON";
            //Parent.ChipCtrlButtons[4].Visible = true;
            //Parent.ChipCtrlButtons[4].Click += RunLEDFunctionTest;

            Parent.ChipCtrlButtons[4].Text = "GH0_H";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;

            Parent.ChipCtrlButtons[6].Text = "LED_ON";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += RunLEDFunctionTest;

            Parent.ChipCtrlButtons[7].Text = "PS_ON";
            Parent.ChipCtrlButtons[7].Visible = true;
            Parent.ChipCtrlButtons[7].Click += Toggle_PowerSupply0_Output;

            Parent.ChipCtrlButtons[8].Text = "MANUAL";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;

            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_AUTO_Test_Items;

            Parent.ChipCtrlButtons[10].Text = "LED";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Change_To_LED_Test_Items;

            Parent.ChipCtrlButtons[11].Text = "CMD";
            Parent.ChipCtrlButtons[11].Visible = true;
            Parent.ChipCtrlButtons[11].Click += Change_To_CMD_Test_Items;
        }

        private void SerialConnect_Click(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (Serial.IsOpen == false)
            {
                Serial.Open();
                if (Serial.IsOpen)
                {
                    b.Text = "Disconn";
                    Parent.ChipCtrlTextboxes[0].Text = Serial.Config.PortName + "-" + Serial.Config.BaudRate.ToString();
                }
            }
            else
            {
                Serial.Close();
                if (!Serial.IsOpen)
                    b.Text = "Connect";
            }
        }

        private void SerialSetting_Click(object sender, EventArgs e)
        {
            JLcLib.Comn.WireComnForm.Show(Serial);
            Serial.WriteSettingFile(Parent.IniFile, "Chicago");
        }

        private void Change_To_Manual_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.MANUAL;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Change_To_AUTO_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.AUTO;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_AUTO.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_AUTO)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Change_To_LED_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.LED;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_LED.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_LED)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Change_To_CMD_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.CMD;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_CMD.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_CMD)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Toggle_PowerSupply0_Output(object sender, EventArgs e)
        {
            if (PowerSupply0 == null)
                PowerSupply0 = new JLcLib.Instrument.SCPI(JLcLib.Instrument.InstrumentTypes.PowerSupply0);
            if (PowerSupply0.IsOpen == false)
                PowerSupply0.Open();

            if (PowerSupply0.IsOpen == false)
            {
                MessageBox.Show("Check PowerSupply0 Connection!");
                return;
            }

            if (Parent.ChipCtrlButtons[7].Text == "PS_ON")
            {
                PowerSupply0.Write("OUTP ON");
                Parent.ChipCtrlButtons[7].Text = "PS_OFF";
            }
            else
            {
                PowerSupply0.Write("OUTP OFF");
                Parent.ChipCtrlButtons[7].Text = "PS_ON";
            }
        }

        private void Toogle_GPIO_GH0(object sender, EventArgs e)
        {
            if (Parent.ChipCtrlButtons[4].Text == "GH0_H")
            {
                SPI.GPIOs[4].Direction = GPIO_Direction.Output;
                SPI.GPIOs[4].State = GPIO_State.Low;
                Parent.ChipCtrlButtons[4].Text = "GH0_L";
            }
            else
            {
                SPI.GPIOs[4].Direction = GPIO_Direction.Output;
                SPI.GPIOs[4].State = GPIO_State.High;
                Parent.ChipCtrlButtons[4].Text = "GH0_H";
            }
        }

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
            if (Serial.IsOpen && (Serial.RcvQueue.Count != 0))
            {
                for (int i = 0; i < Serial.RcvQueue.Count; i++)
                {
                    byte b = Serial.RcvQueue.Get();
                    if ((b > 0x00 && b < 0x80) && b != '\0')
                    {
                        LogMessage += (char)b;
                        if (b == '\n')
                        {
                            Log.WriteWithElapsedTime(LogMessage);
                            Log.WriteLine("\n");
                            LogMessage = "";
                        }
                    }
                }
            }
        }

        public override void SendCommand(string Command)
        {
            Command += "\r\n";
            byte[] Var = Encoding.ASCII.GetBytes(Command);
            Serial.WriteBytes(Var, Var.Length, true);
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            uint iVal;
            try
            {
                iVal = System.Convert.ToUInt32(Arg, 10);
            }
            catch
            {
                iVal = 0;
            }

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.MANUAL:
                    switch ((TEST_ITEMS_MANUAL)TestItemIndex)
                    {
                        case TEST_ITEMS_MANUAL.Write_SRAM:
                            iVal = System.Convert.ToUInt32(Arg, 16);
                            RunSRAMWrite(iVal); break;
                        case TEST_ITEMS_MANUAL.Read_SRAM:
                            iVal = System.Convert.ToUInt32(Arg, 16);
                            RunSRAMRead(iVal); break;
                        //case TEST_ITEMS_MANUAL.Dummy321Bytes:
                        //    Run_TEST_POWER_ON_RESET();
                        //    break;
                        case TEST_ITEMS_MANUAL.VoltageSweep_Start: RunVoltageSweepTest(); break;
                        case TEST_ITEMS_MANUAL.VoltageSweep_Stop: StopVoltageSweep(); break;
                            //case TEST_ITEMS_MANUAL.AiP33A09_Config: WriteConfigurationRegister(); break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((TEST_ITEMS_AUTO)TestItemIndex)
                    {
                        case TEST_ITEMS_AUTO.Initial_Test: Run_INITIAL_TEST_SEQ(iVal); break;
                            //case TEST_ITEMS_AUTO.SET_ANA_125: Set_TEST_ANA_125_REF(); break;
                            //case TEST_ITEMS_AUTO.SET_ANA_150: Set_TEST_ANA_150_REF(); break;
                            //case TEST_ITEMS_AUTO.SET_ANA_PTAT: Set_TEST_ANA_PTAT_OUT(); break;
                            //case TEST_ITEMS_AUTO.SET_ANA_OSC: Set_TEST_ANA_OSC_FREQ(); break;
                            //case TEST_ITEMS_AUTO.Disable_ANA: Disable_TEST_ANA(); break;
                            //case TEST_ITEMS_AUTO.SET_SEG_TEST: Set_SEG_TEST_RES(true); break;
                            //case TEST_ITEMS_AUTO.Disable_SEG_TEST: Set_SEG_TEST_RES(false); break;
                            //case TEST_ITEMS_AUTO.RUN_Write_EFUSE: Run_WriteAndRead_EFUSE(); break;
                            //case TEST_ITEMS_AUTO.RUN_TRIM_125: Run_CAL_125_REF(); break;
                            //case TEST_ITEMS_AUTO.RUN_TRIM_150: Run_CAL_150_REF(); break;
                            //case TEST_ITEMS_AUTO.RUN_TRIM_PTAT: Run_CAL_PTAT_OUT(); break;
                            //case TEST_ITEMS_AUTO.RUN_TRIM_OSC: Run_CAL_OSC_FREQ(); break;
                    }
                    break;
                case COMBOBOX_ITEMS.LED:
                    switch ((TEST_ITEMS_LED)TestItemIndex)
                    {
                        case TEST_ITEMS_LED.WaveEffect_Start: RunWaveEffectTest(); break;
                        case TEST_ITEMS_LED.ScrollText_Start: RunScrollTextTest(); break;
                        case TEST_ITEMS_LED.BreathingEffect_Start: RunBreathingEffectTest(); break;
                        case TEST_ITEMS_LED.ClockEffect_Start: RunClockEffectTest(); break;
                        case TEST_ITEMS_LED.FireworkEffect_Start: RunFireworkEffectTest(); break;
                        case TEST_ITEMS_LED.Stop_LED_Effect: StopLEDMatrixEffect(); break;
                            //case TEST_ITEMS_LED.Display_Number: ShowDigit(); break;
                    }
                    break;
                case COMBOBOX_ITEMS.CMD:
                    switch ((TEST_ITEMS_CMD)TestItemIndex)
                    {
                        case TEST_ITEMS_CMD.Update_Display: WriteCommand(0x04); break;
                        case TEST_ITEMS_CMD.Write_Display: RunWriteLEDData(); break;
                        case TEST_ITEMS_CMD.TurnOff_Display: WriteCommand(0x08); break;
                        case TEST_ITEMS_CMD.WakeUp: WriteCommand(0x0D); break;
                        case TEST_ITEMS_CMD.Sleep:
                            var SLEEP = Parent.RegMgr.GetRegisterItem("SLEEP");
                            SLEEP.Value = 1;
                            SLEEP.Write();
                            WriteCommand(0x0E); 
                            break;
                        case TEST_ITEMS_CMD.Reset: WriteCommand(0xCE); break;
                    }
                    break;
                default:
                    break;
            }
        }

        #region Function for SPI Communication
        private void WriteRegister(uint Address, uint Data)
        {
            byte[] bytes;

            try
            {
                bytes = new byte[7];
                bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC0;
                bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF); // 체크섬
                bytes[4] = (byte)((Address >> 8) & 0xFF); // 주소 상위 바이트
                bytes[5] = (byte)((Address >> 0) & 0xFF); // 주소 하위 바이트
                bytes[6] = (byte)((bytes[4] + bytes[5]) & 0xFF); // 체크섬
                SPI.WriteBytes(bytes, 7, true);

                bytes = new byte[7];
                bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC1;
                bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
                bytes[4] = 0x00;
                bytes[5] = 0x01; // 길이 = 1 Byte
                bytes[6] = (byte)((bytes[4] + bytes[5]) & 0xFF);
                SPI.WriteBytes(bytes, 7, true);

                bytes = new byte[6];
                bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC2;
                bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
                bytes[4] = (byte)(Data & 0xff); // 실제 데이터
                bytes[5] = bytes[4]; // 체크섬
                SPI.WriteBytesForChicago(bytes, 6, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fail to Write: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private uint ReadRegister(uint Address)
        {
            uint Data = 0xFF;
            byte[] bytes;
            byte[] rx;

            try
            {
                switch (Address)
                {
                    case 0x4FE:
                    case 0x4FF:
                        bytes = new byte[4];
                        bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0x07;
                        bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
                        rx = SPI.WriteAndReadBytesForChicago(bytes, 4, 3);

                        if (Address == 0x4FE)
                        {
                            Data = (uint)(rx[0] & 0xff);
                        }
                        else if (Address == 0x4FF)
                        {
                            Data = (uint)(rx[1] & 0xff);
                        }
                        break;
                    default:
                        bytes = new byte[7];
                        bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC0;
                        bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
                        bytes[4] = (byte)((Address >> 8) & 0xFF);
                        bytes[5] = (byte)((Address >> 0) & 0xFF);
                        bytes[6] = (byte)((bytes[4] + bytes[5]) & 0xFF);
                        SPI.WriteBytes(bytes, 7, true);

                        bytes = new byte[7];
                        bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC1;
                        bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
                        bytes[4] = 0x00;
                        bytes[5] = 0x01;
                        bytes[6] = (byte)((bytes[4] + bytes[5]) & 0xFF);
                        SPI.WriteBytes(bytes, 7, true);

                        bytes = new byte[4];
                        bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC3;
                        bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
                        rx = SPI.WriteAndReadBytesForChicago(bytes, 4, 2);

                        Data = (uint)(rx[0] & 0xff);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fail to Read: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return Data;
        }

        private void RunSRAMWrite(uint startAddress)
        {
            const int sramTotalBytes = 320;
            const int sramSegmentsPerGrid = 32;

            using (var hexForm = new HexWriteForm(sramTotalBytes, sramSegmentsPerGrid))
            {
                hexForm.Text = $"SRAM Write (Addr: 0x{startAddress:X4}, Len: {sramTotalBytes})";
                if (hexForm.ShowDialog(Parent) == DialogResult.OK)
                {
                    WriteSRAMWithHexForm(startAddress, (uint)sramTotalBytes, hexForm.ResultData);
                }
            }
        }

        private void RunSRAMRead(uint startAddress)
        {
            const int sramTotalBytes = 320;
            const int sramSegmentsPerGrid = 32;

            byte[] data = ReadSRAMWithHexForm(startAddress, (uint)sramTotalBytes);
            if (data == null)
            {
                MessageBox.Show("데이터 읽기에 실패했습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var hexForm = new HexWriteForm(sramTotalBytes, sramSegmentsPerGrid, data))
            {
                hexForm.Text = $"SRAM Read Result (Addr: 0x{startAddress:X4}, Len: {sramTotalBytes})";
                hexForm.SetReadOnlyMode();
                hexForm.ShowDialog(Parent);
            }
        }

        private void WriteSRAMWithHexForm(uint Address, uint length, byte[] hexData)
        {
            if (hexData == null || hexData.Length < length) return;
            byte[] bytes;
            bytes = new byte[7];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC0;
            bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
            bytes[4] = (byte)((Address >> 8) & 0xFF);
            bytes[5] = (byte)((Address >> 0) & 0xFF);
            bytes[6] = (byte)((bytes[4] + bytes[5]) & 0xFF);
            SPI.WriteBytes(bytes, 7, true);

            bytes = new byte[7];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC1;
            bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
            bytes[4] = (byte)((length >> 8) & 0xFF);
            bytes[5] = (byte)((length >> 0) & 0xFF);
            bytes[6] = (byte)((bytes[4] + bytes[5]) & 0xFF);
            SPI.WriteBytes(bytes, 7, true);

            bytes = new byte[4];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC2;
            bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
            SPI.WriteBytes(bytes, 4, true);

            bytes = new byte[length + 1];
            Array.Copy(hexData, 0, bytes, 0, length);
            byte cksum = 0;
            for (int i = 0; i < length; i++) cksum += bytes[i];
            bytes[bytes.Length - 1] = cksum;
            SPI.WriteBytesForChicago(bytes, bytes.Length, true);
        }

        private byte[] ReadSRAMWithHexForm(uint startAddress, uint length)
        {
            if (length == 0 || length > 320) return null;
            byte[] bytes;
            uint cksum;
            bytes = new byte[7];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC0;
            cksum = (uint)(bytes[0] + bytes[1] + bytes[2]);
            bytes[3] = (byte)(cksum & 0xFF);
            bytes[4] = (byte)((startAddress >> 8) & 0xFF);
            bytes[5] = (byte)(startAddress & 0xFF);
            cksum = (uint)(bytes[4] + bytes[5]);
            bytes[6] = (byte)(cksum & 0xFF);
            SPI.WriteBytes(bytes, 7, true);

            bytes = new byte[7];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC1;
            cksum = (uint)(bytes[0] + bytes[1] + bytes[2]);
            bytes[3] = (byte)(cksum & 0xFF);
            bytes[4] = (byte)((length >> 8) & 0xFF);
            bytes[5] = (byte)(length & 0xFF);
            cksum = (uint)(bytes[4] + bytes[5]);
            bytes[6] = (byte)(cksum & 0xFF);
            SPI.WriteBytes(bytes, 7, true);

            bytes = new byte[4];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = 0xC3;
            cksum = (uint)(bytes[0] + bytes[1] + bytes[2]);
            bytes[3] = (byte)(cksum & 0xFF);

            byte[] rx = SPI.WriteAndReadBytesForChicago(bytes, 4, (int)(length + 1));
            if (rx == null || rx.Length < length) return null;

            byte[] payload = new byte[length];
            Array.Copy(rx, 0, payload, 0, length);
            return payload;
        }

        private void WriteCommand(uint cmd)
        {
            byte[] bytes = new byte[4];
            bytes[0] = 0x5A; bytes[1] = 0xFF; bytes[2] = (byte)(cmd & 0xff);
            bytes[3] = (byte)((bytes[0] + bytes[1] + bytes[2]) & 0xFF);
            SPI.WriteBytesForChicago(bytes, 4, true);
        }
        #endregion Function for SPI Communication

        #region Function for Manual Test
        private void Check_Instrument()
        {
            foreach (JLcLib.Instrument.InsInformation Ins in JLcLib.Instrument.InstrumentForm.InsInfoList)
            {
                switch (Ins.Type)
                {
                    case JLcLib.Instrument.InstrumentTypes.PowerSupply0:
                        if (PowerSupply0 == null)
                            PowerSupply0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (PowerSupply0.IsOpen == false)
                            PowerSupply0.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter0:
                        if (DigitalMultimeter0 == null)
                            DigitalMultimeter0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter0.IsOpen == false)
                            DigitalMultimeter0.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter1:
                        if (DigitalMultimeter1 == null)
                            DigitalMultimeter1 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter1.IsOpen == false)
                            DigitalMultimeter1.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter2:
                        if (DigitalMultimeter2 == null)
                            DigitalMultimeter2 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter2.IsOpen == false)
                            DigitalMultimeter2.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter3:
                        if (DigitalMultimeter3 == null)
                            DigitalMultimeter3 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter3.IsOpen == false)
                            DigitalMultimeter3.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.OscilloScope0:
                        if (OscilloScope0 == null)
                            OscilloScope0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (OscilloScope0.IsOpen == false)
                            OscilloScope0.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.SpectrumAnalyzer:
                        if (SpectrumAnalyzer == null)
                            SpectrumAnalyzer = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (SpectrumAnalyzer.IsOpen == false)
                            SpectrumAnalyzer.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.TempChamber:
                        if (TempChamber == null)
                            TempChamber = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (TempChamber.IsOpen == false)
                            TempChamber.Open();
                        break;
                }
            }
        }

        private readonly Random _random = new Random();

        private int RandomNumber(int min, int max)
        {
            return _random.Next(min, max);
        }

        private void RunVoltageSweepTest()
        {
            if (voltageSweepWorker != null && voltageSweepWorker.IsBusy)
            {
                Log.WriteLine("Voltage sweep is already running.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            try
            {
                double minVoltage = double.Parse(RegContForm.Prompt.ShowDialog("최소 전압 (V):", "Voltage Sweep"));
                double maxVoltage = double.Parse(RegContForm.Prompt.ShowDialog("최대 전압 (V):", "Voltage Sweep"));
                int timeDelay = int.Parse(RegContForm.Prompt.ShowDialog("변화 주기 (ms):", "Voltage Sweep"));

                if (minVoltage >= maxVoltage) throw new Exception("최소 전압은 최대 전압보다 작아야 합니다.");
                if (timeDelay <= 0) throw new Exception("변화 주기는 0보다 커야 합니다.");

                voltageSweepWorker = new BackgroundWorker();
                voltageSweepWorker.WorkerSupportsCancellation = true;

                voltageSweepWorker.DoWork += (sender, e) =>
                {
                    var sweepParams = (Tuple<double, double, int>)e.Argument;
                    ExecuteVoltageSweep(sender as BackgroundWorker, sweepParams.Item1, sweepParams.Item2, sweepParams.Item3);
                };

                voltageSweepWorker.RunWorkerCompleted += (sender, e) => { Log.WriteLine("Voltage sweep stopped.", Color.Green, Parent.BackColor); };

                var startParams = new Tuple<double, double, int>(minVoltage, maxVoltage, timeDelay);
                voltageSweepWorker.RunWorkerAsync(startParams);
                Log.WriteLine("Voltage sweep started.", Color.Green, Parent.BackColor);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"입력 오류: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StopVoltageSweep()
        {
            if (voltageSweepWorker != null && voltageSweepWorker.IsBusy)
            {
                voltageSweepWorker.CancelAsync();
            }
            else
            {
                Log.WriteLine("No voltage sweep is currently running.", Color.OrangeRed, Parent.BackColor);
            }
        }

        private void ExecuteVoltageSweep(BackgroundWorker worker, double minVoltage, double maxVoltage, int delay)
        {
            Check_Instrument();

            bool minFirst = true;

            List<double> voltageSteps = new List<double>();
            for (double v = minVoltage; v <= maxVoltage; v += 0.1)
            {
                voltageSteps.Add(Math.Round(v, 2));
            }

            if (voltageSteps.Count == 0)
            {
                Log.WriteLine("No voltage steps to sweep. Check min/max values.", Color.Red, Parent.BackColor);
                return;
            }

            try
            {
                if (PowerSupply0 == null || !PowerSupply0.IsOpen)
                {
                    Log.WriteLine("PowerSupply0 is not connected.", Color.Red, Parent.BackColor);
                    return;
                }

                PowerSupply0.Write("OUTP ON, (@1, 2, 3)");
                PowerSupply0.Write("CURRent 1, (@1, 2, 3)");

                Log.WriteLine($"Channel 1~3 Voltage sweeping between {minVoltage}V and {maxVoltage}V...", Color.Cyan, Parent.BackColor);

                while (!worker.CancellationPending)
                {
                    //int randomSelect = RandomNumber(0, voltageSteps.Count);

                    //PowerSupply0.Write($"VOLT {voltageSteps[randomSelect]}, (@1)");
                    if (minFirst)
                    {
                        PowerSupply0.Write($"VOLT {minVoltage}, (@1)");
                        minFirst = false;
                    }
                    else
                    {
                        PowerSupply0.Write($"VOLT {maxVoltage}, (@1)");
                        minFirst = true;
                    }
                    if (worker.CancellationPending) break;
                    Thread.Sleep(delay);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLine($"Sweep Error: {ex.Message}", Color.Red, Parent.BackColor);
            }
        }

        private void WriteConfigurationRegister()
        {
            try
            {
                uint i_val = uint.Parse(RegContForm.Prompt.ShowDialog("I[5:0] (0-63):", "Config Write - Byte 1"));

                uint osc_val = uint.Parse(RegContForm.Prompt.ShowDialog("OSC[7] (0 or 1):", "Config Write - Byte 2"));
                uint g_n_val = uint.Parse(RegContForm.Prompt.ShowDialog("G_N[5:2] (0-15):", "Config Write - Byte 2"));
                uint test_val = uint.Parse(RegContForm.Prompt.ShowDialog("TEST[1] (0 or 1):", "Config Write - Byte 2"));
                uint up_val = uint.Parse(RegContForm.Prompt.ShowDialog("UP[0] (0 or 1):", "Config Write - Byte 2"));

                uint tp1_val = uint.Parse(RegContForm.Prompt.ShowDialog("TP1[7] (0 or 1):", "Config Write - Byte 3"));
                uint res_val = uint.Parse(RegContForm.Prompt.ShowDialog("RES[6] (0 or 1):", "Config Write - Byte 3"));
                uint fr_val = uint.Parse(RegContForm.Prompt.ShowDialog("FR1[5:4] (0-3):", "Config Write - Byte 3"));
                uint don_val = uint.Parse(RegContForm.Prompt.ShowDialog("DON[3] (0 or 1):", "Config Write - Byte 3"));
                uint gs_val = uint.Parse(RegContForm.Prompt.ShowDialog("GS[2] (0 or 1):", "Config Write - Byte 3"));
                uint dt_val = uint.Parse(RegContForm.Prompt.ShowDialog("DT[1:0] (0-3):", "Config Write - Byte 3"));

                uint sleep_val = uint.Parse(RegContForm.Prompt.ShowDialog("SLEEP[6] (0 or 1):", "Config Write - Byte 4"));
                uint tp2_val = uint.Parse(RegContForm.Prompt.ShowDialog("TP2[1] (0 or 1):", "Config Write - Byte 4"));

                i_val &= 0x3F;
                osc_val &= 0x01;
                g_n_val &= 0x0F;
                test_val &= 0x01;
                up_val &= 0x01;
                tp1_val &= 0x01;
                res_val &= 0x01;
                fr_val &= 0x03;
                don_val &= 0x01;
                gs_val &= 0x01;
                dt_val &= 0x03;
                sleep_val &= 0x01;
                tp2_val &= 0x01;

                byte[] payload = new byte[4];

                payload[0] = (byte)i_val;
                payload[1] = (byte)((osc_val << 7) | (g_n_val << 2) | (test_val << 1) | up_val);
                payload[2] = (byte)((tp1_val << 7) | (res_val << 6) | (fr_val << 4) | (don_val << 3) | (gs_val << 2) | dt_val);
                payload[3] = (byte)((sleep_val << 6) | (1 << 2) | (tp2_val << 1));

                byte[] finalPacket = new byte[9];
                finalPacket[0] = 0x5A;
                finalPacket[1] = 0xFF;
                finalPacket[2] = 0x01;
                finalPacket[3] = (byte)((finalPacket[0] + finalPacket[1] + finalPacket[2]) & 0xFF);

                Array.Copy(payload, 0, finalPacket, 4, 4);

                finalPacket[8] = (byte)((payload[0] + payload[1] + payload[2] + payload[3]) & 0xFF);

                SPI.WriteBytesForChicago(finalPacket, finalPacket.Length, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"입력 오류 또는 데이터 변환 실패: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion Function for Manual Test

        #region Function for Auto Test
        private void Set_TEST_ANA_125_REF()
        {
            uint reg40C_val = ReadRegister(0x40C);
            uint reg40D_val = ReadRegister(0x40D);

            if (reg40C_val != 0x08)
            {
                WriteRegister(0x40C, 0x08);
            }
            if (((reg40D_val >> 3) & 0x01) != 0x01)
            {
                WriteRegister(0x40D, (reg40D_val | (1 << 3)));
            }
        }

        private void Set_TEST_ANA_150_REF()
        {
            uint reg40C_val = ReadRegister(0x40C);
            uint reg40D_val = ReadRegister(0x40D);

            if (reg40C_val != 0x10)
            {
                WriteRegister(0x40C, 0x10);
            }
            if (((reg40D_val >> 3) & 0x01) != 0x01)
            {
                WriteRegister(0x40D, (reg40D_val | (1 << 3)));
            }
        }

        private void Set_TEST_ANA_PTAT_OUT()
        {
            uint reg40C_val = ReadRegister(0x40C);
            uint reg40D_val = ReadRegister(0x40D);

            if (reg40C_val != 0x04)
            {
                WriteRegister(0x40C, 0x04);
            }
            if (((reg40D_val >> 3) & 0x01) != 0x01)
            {
                WriteRegister(0x40D, (reg40D_val | (1 << 3)));
            }
        }

        private void Set_TEST_ANA_OSC_FREQ()
        {
            uint reg40C_val = ReadRegister(0x40C);
            uint reg40D_val = ReadRegister(0x40D);

            if (reg40C_val != 0x01)
            {
                WriteRegister(0x40C, 0x01);
            }
            if (((reg40D_val >> 3) & 0x01) != 0x01)
            {
                WriteRegister(0x40D, (reg40D_val | (1 << 3)));
            }
        }

        private void Disable_TEST_ANA()
        {
            uint reg40C_val = ReadRegister(0x40C);
            uint reg40D_val = ReadRegister(0x40D);

            if (reg40C_val != 0x00)
            {
                WriteRegister(0x40C, 0x00);
            }
            if (((reg40D_val >> 3) & 0x01) != 0x00)
            {
                WriteRegister(0x40D, (reg40D_val & (0 << 3)));
            }
        }

        private double[] Run_TEST_POWER_ON_RESET()
        {
            if (PowerSupply0 == null || !PowerSupply0.IsOpen)
            {
                MessageBox.Show("Check Connection Status of PowerSupply0", "ERROR");
                return null;
            }

            if (DigitalMultimeter2 == null || !DigitalMultimeter2.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter2", "ERROR");
                return null;
            }

            double[] porDetect = new double[2];

            try
            {
                DigitalMultimeter2.Write("SENS:FUNC 'CURR:DC'");
                DigitalMultimeter2.Write("SENS:CURR:DC:RANGE 1E-3");

                PowerSupply0.Write("VOLT 1.9, (@1)");
                PowerSupply0.Write("CURR 1, (@1)");
                PowerSupply0.Write("OUTP ON, (@1)");

                double temp = new double();

                for (double volt = 2.05; volt < 2.55; volt += 0.05)
                {
                    PowerSupply0.Write($"VOLT {volt}, (@1)");
                    Thread.Sleep(1500);

                    string response = DigitalMultimeter2.WriteAndReadString($"READ?");
                    double measuredValue;

                    if (double.TryParse(response, out measuredValue))
                    {
                        temp = Math.Round(measuredValue * 1_000, 2);

                        if (temp >= 0.15)
                        {
                            porDetect[0] = volt;
                            break;
                        }
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }

                for (double volt = 2.1; volt > 1.7; volt -= 0.05)
                {
                    PowerSupply0.Write($"VOLT {volt}, (@1)");
                    Thread.Sleep(1500);

                    string response = DigitalMultimeter2.WriteAndReadString($"READ?");
                    double measuredValue;

                    if (double.TryParse(response, out measuredValue))
                    {
                        temp = Math.Round(measuredValue * 1_000, 2);

                        if (temp <= 0.15)
                        {
                            porDetect[1] = volt;
                            break;
                        }
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }

            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in POR Test: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                SPI.SetGPIOLHigh();

                PowerSupply0.Write("VOLT 5, (@1)");
                DigitalMultimeter2.Write("SENS:CURR:DC:RANG 0.1");
            }

            return porDetect;
        }

        private double[] Run_CAL_125_REF(double target = 0.9365)
        {
            if (DigitalMultimeter0 == null || !DigitalMultimeter0.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter0", "ERROR");
                return null;
            }

            try
            {
                double[] trimValue = new double[2];
                var TSP_125_TRIM = Parent.RegMgr.GetRegisterItem("TSP_125_TRIM[4:0]");

                int left = 0;
                int right = 31;
                int bestCode = -1;
                double bestDiff = double.MaxValue;
                double dmm_volt = 0.0;
                double diff = 0.0;

                Set_TEST_ANA_125_REF();
                DigitalMultimeter0.Write("SENS:FUNC 'VOLT:DC'");

                while (left <= right)
                {
                    int mid = left + (right - left) / 2;

                    TSP_125_TRIM.Value = (uint)mid;
                    TSP_125_TRIM.Write();
                    Thread.Sleep(100);

                    List<double> measurements = new List<double>();
                    for (int i = 0; i < 3; i++)
                    {
                        string response = DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?");
                        if (double.TryParse(response, out double measuredValue))
                        {
                            measurements.Add(measuredValue);
                        }
                        else
                        {
                            throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                        }
                    }
                    dmm_volt = Math.Round(measurements.Average(), 4);

                    diff = Math.Round(Math.Abs(target - dmm_volt), 4);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestCode = mid;
                    }

                    if (dmm_volt > target)
                        left = mid + 1;
                    else
                        right = mid - 1;
                }

                if (bestCode >= 0)
                {
                    TSP_125_TRIM.Value = (uint)bestCode;
                    TSP_125_TRIM.Write();
                    Thread.Sleep(100);
                }

                List<double> finalMeasurements = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    string response = DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        finalMeasurements.Add(measuredValue);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                dmm_volt = Math.Round(finalMeasurements.Average(), 4);

                trimValue[0] = dmm_volt;
                trimValue[1] = bestCode;
                return trimValue;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_CAL_125_REF: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                DigitalMultimeter0.Write("SENS:FUNC 'VOLT:AC'");
                Disable_TEST_ANA();
            }
        }

        private double[] Run_CAL_150_REF(double target = 1.009)
        {
            if (DigitalMultimeter0 == null || !DigitalMultimeter0.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter0", "ERROR");
                return null;
            }

            try
            {
                double[] trimValue = new double[2];
                var TSP_150_TRIM = Parent.RegMgr.GetRegisterItem("TSP_150_TRIM[4:0]");

                int left = 0;
                int right = 31;
                int bestCode = -1;
                double bestDiff = double.MaxValue;
                double dmm_volt = 0.0;
                double diff = 0.0;

                Set_TEST_ANA_150_REF();
                DigitalMultimeter0.Write("SENS:FUNC 'VOLT:DC'");

                while (left <= right)
                {
                    int mid = left + (right - left) / 2;

                    TSP_150_TRIM.Value = (uint)mid;
                    TSP_150_TRIM.Write();
                    Thread.Sleep(100);

                    List<double> measurements = new List<double>();
                    for (int i = 0; i < 3; i++)
                    {
                        string response = DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?");
                        if (double.TryParse(response, out double measuredValue))
                        {
                            measurements.Add(measuredValue);
                        }
                        else
                        {
                            throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                        }
                    }
                    dmm_volt = Math.Round(measurements.Average(), 4);

                    diff = Math.Round(Math.Abs(target - dmm_volt), 4);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestCode = mid;
                    }

                    if (dmm_volt > target)
                        left = mid + 1;
                    else
                        right = mid - 1;
                }

                if (bestCode >= 0)
                {
                    TSP_150_TRIM.Value = (uint)bestCode;
                    TSP_150_TRIM.Write();
                    Thread.Sleep(100);
                }

                List<double> finalMeasurements = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    string response = DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        finalMeasurements.Add(measuredValue);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                dmm_volt = Math.Round(finalMeasurements.Average(), 4);


                trimValue[0] = dmm_volt;
                trimValue[1] = bestCode;
                return trimValue;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_CAL_150_REF: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                DigitalMultimeter0.Write("SENS:FUNC 'VOLT:AC'");
                Disable_TEST_ANA();
            }
        }

        private double[] Run_CAL_PTAT_OUT(double target = 0.727)
        {
            if (DigitalMultimeter0 == null || !DigitalMultimeter0.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter0", "ERROR");
                return null;
            }

            try
            {
                double[] trimValue = new double[2];
                var VPTAT_OUT_TRIM = Parent.RegMgr.GetRegisterItem("VPTAT_OUT_TRIM[4:0]");

                int OrdToCode(int ord) { return (ord + 16) & 31; }

                double MeasureWithCode(int code)
                {
                    VPTAT_OUT_TRIM.Value = (uint)code;
                    VPTAT_OUT_TRIM.Write();
                    Thread.Sleep(100);

                    List<double> measurements = new List<double>();
                    for (int i = 0; i < 3; i++)
                    {
                        string response = DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?");
                        if (double.TryParse(response, out double measuredValue))
                        {
                            measurements.Add(measuredValue);
                        }
                        else
                        {
                            throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                        }
                    }
                    return Math.Round(measurements.Average(), 4);
                }

                Set_TEST_ANA_PTAT_OUT();
                DigitalMultimeter0.Write("SENS:FUNC 'VOLT:DC'");

                int left = 0, right = 31;
                int bestOrd = -1;
                int bestCode = -1;
                double bestDiff = double.MaxValue;
                double dmm_volt = 0.0, diff = 0.0;

                double vL = MeasureWithCode(OrdToCode(left));
                double vR = MeasureWithCode(OrdToCode(right));
                bool isDescending = vR < vL;

                while (left <= right)
                {
                    int mid = left + (right - left) / 2;
                    dmm_volt = MeasureWithCode(OrdToCode(mid));
                    diff = Math.Round(Math.Abs(target - dmm_volt), 4);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestOrd = mid;
                    }

                    if (isDescending)
                    {
                        if (dmm_volt > target) left = mid + 1;
                        else right = mid - 1;
                    }
                    else
                    {
                        if (dmm_volt < target) left = mid + 1;
                        else right = mid - 1;
                    }
                }

                if (bestOrd >= 0)
                {
                    bestCode = OrdToCode(bestOrd);
                    VPTAT_OUT_TRIM.Value = (uint)bestCode;
                    VPTAT_OUT_TRIM.Write();
                }

                dmm_volt = MeasureWithCode(bestCode);

                trimValue[0] = dmm_volt;
                trimValue[1] = bestCode;
                return trimValue;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_CAL_PTAT_OUT: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                DigitalMultimeter0.Write("SENS:FUNC 'VOLT:AC'");
                Disable_TEST_ANA();
            }
        }

        private double[] Run_CAL_OSC_FREQ(double target = 1.9985)
        {
            if (OscilloScope0 == null || !OscilloScope0.IsOpen)
            {
                MessageBox.Show("Check Connection Status of OscilloScope0", "ERROR");
                return null;
            }

            try
            {
                double[] trimValue = new double[2];
                var RCOSC_BIAS_TRIM = Parent.RegMgr.GetRegisterItem("RC_OSC_BIAS_TRIM[4:0]");

                int left = 0;
                int right = 31;
                int bestCode = -1;
                double bestDiff = double.MaxValue;
                double osc_mhz = 0.0;
                double diff = 0.0;

                Set_TEST_ANA_OSC_FREQ();

                OscilloScope0.Write(":CHAN1:SCAL 2");
                OscilloScope0.Write(":CHAN1:OFFS 0");
                OscilloScope0.Write(":TIM:SCAL 2.00E-7");
                OscilloScope0.Write(":TIM:POS 0");

                while (left <= right)
                {
                    int mid = left + (right - left) / 2;

                    RCOSC_BIAS_TRIM.Value = (uint)mid;
                    RCOSC_BIAS_TRIM.Write();
                    Thread.Sleep(100);

                    List<double> measurements = new List<double>();
                    for (int i = 0; i < 3; i++)
                    {
                        OscilloScope0.Write(":MEAS:FREQ CHAN1");
                        string response = OscilloScope0.WriteAndReadString(":MEAS:FREQ?");
                        if (double.TryParse(response, out double measuredValue))
                        {
                            measurements.Add(measuredValue);
                        }
                        else
                        {
                            throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                        }
                    }
                    osc_mhz = Math.Round(measurements.Average() / 1000000, 4);

                    diff = Math.Round(Math.Abs(target - osc_mhz), 4);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestCode = mid;
                    }

                    if (osc_mhz > target)
                        left = mid + 1;
                    else
                        right = mid - 1;
                }

                if (bestCode >= 0)
                {
                    RCOSC_BIAS_TRIM.Value = (uint)bestCode;
                    RCOSC_BIAS_TRIM.Write();
                    Thread.Sleep(100);
                }

                List<double> finalMeasurements = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    OscilloScope0.Write(":MEAS:FREQ CHAN1");
                    string response = OscilloScope0.WriteAndReadString(":MEAS:FREQ?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        finalMeasurements.Add(measuredValue);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                osc_mhz = Math.Round(finalMeasurements.Average() / 1000000, 4);

                trimValue[0] = osc_mhz;
                trimValue[1] = bestCode;
                return trimValue;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_CAL_OSC_FREQ: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                Disable_TEST_ANA();
            }
        }

        private double[] Run_CAL_REXT_CURRENT(double target = 30)
        {
            if (DigitalMultimeter1 == null || !DigitalMultimeter1.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter1", "ERROR");
                return null;
            }

            var BGR_OUT_TRIM = Parent.RegMgr.GetRegisterItem("BGR_OUT_TRIM[4:0]");
            var SEG_CURR = Parent.RegMgr.GetRegisterItem("I[5:0]");
            var SEG_TEST = Parent.RegMgr.GetRegisterItem("SEG_TEST");
            var RES = Parent.RegMgr.GetRegisterItem("RES");

            try
            {
                double[] trimValue = new double[2];

                SEG_CURR.Value = 63;
                SEG_CURR.Write();
                RES.Value = 1;
                RES.Write();
                SEG_TEST.Value = 1;
                SEG_TEST.Write();
                Thread.Sleep(100);
                PowerSupply0.Write("VOLT 0.6, (@2)");
                PowerSupply0.Write("CURR 1, (@2)");
                PowerSupply0.Write("OUTP ON, (@2)");
                DigitalMultimeter1.Write("CONF:CURR:DC 0.1");

                bool checkConnection = true;
                double temp;
                while (checkConnection)
                {
                    string response = DigitalMultimeter1.WriteAndReadString("READ?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        temp = Math.Round(measuredValue * 1000, 2);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                    if (temp >= 40 || temp <= 20)
                    {
                        if (MessageBox.Show($"Check Segment Connection!\nDigitalMultimeter1 = {temp}", "Segment Channel mismatch", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return null;
                        }
                    }
                    else checkConnection = false;
                }

                int left = 0;
                int right = 31;
                int bestCode = -1;
                double bestDiff = double.MaxValue;
                double current_mA = 0.0;
                double diff = 0.0;

                while (left <= right)
                {
                    int mid = left + (right - left) / 2;

                    BGR_OUT_TRIM.Value = (uint)mid;
                    BGR_OUT_TRIM.Write();
                    Thread.Sleep(100);

                    List<double> measurements = new List<double>();
                    for (int i = 0; i < 3; i++)
                    {
                        string response = DigitalMultimeter1.WriteAndReadString("READ?");
                        if (double.TryParse(response, out double measuredValue))
                        {
                            measurements.Add(measuredValue);
                        }
                        else
                        {
                            throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                        }
                    }
                    current_mA = Math.Round(measurements.Average() * 1000, 2);

                    diff = Math.Round(Math.Abs(target - current_mA), 2);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestCode = mid;
                    }

                    if (current_mA > target)
                        left = mid + 1;
                    else
                        right = mid - 1;
                }

                BGR_OUT_TRIM.Value = (uint)bestCode;
                BGR_OUT_TRIM.Write();
                Thread.Sleep(100);

                List<double> finalMeasurements = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    string response = DigitalMultimeter1.WriteAndReadString("READ?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        finalMeasurements.Add(measuredValue);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                current_mA = Math.Round(finalMeasurements.Average() * 1000, 2);

                trimValue[0] = current_mA;
                trimValue[1] = bestCode;
                return trimValue;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_CAL_REXT_CURRENT: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                PowerSupply0.Write("OUTP OFF, (@2)");
                SEG_TEST.Value = 0;
                SEG_TEST.Write();
                RES.Value = 0;
                RES.Write();
            }
        }

        private double[] Run_CAL_IBIAS_OF_DRV(double target = 30)
        {
            if (DigitalMultimeter1 == null || !DigitalMultimeter1.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter1", "ERROR");
                return null;
            }

            var DRV_IBIAS_TRIM = Parent.RegMgr.GetRegisterItem("DRV_IBIAS_TRIM[4:0]");
            var SEG_CURR = Parent.RegMgr.GetRegisterItem("I[5:0]");
            var SEG_TEST = Parent.RegMgr.GetRegisterItem("SEG_TEST");

            try
            {
                double[] trimValue = new double[2];

                SEG_CURR.Value = 63;
                SEG_CURR.Write();
                SEG_TEST.Value = 1;
                SEG_TEST.Write();
                Thread.Sleep(100);
                PowerSupply0.Write("VOLT 0.6, (@2)");
                PowerSupply0.Write("CURR 1, (@2)");
                PowerSupply0.Write("OUTP ON, (@2)");
                DigitalMultimeter1.Write("CONF:CURR:DC 0.1");

                bool checkConnection = true;
                double temp;
                while (checkConnection)
                {
                    string response = DigitalMultimeter1.WriteAndReadString("READ?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        temp = Math.Round(measuredValue * 1000, 2);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                    if (temp >= 40 || temp <= 20)
                    {
                        if (MessageBox.Show($"Check Segment Connection!\nDigitalMultimeter1 = {temp}", "Segment Channel mismatch", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return null;
                        }
                    }
                    else checkConnection = false;
                }

                int left = 0;
                int right = 31;
                int bestCode = -1;
                double bestDiff = double.MaxValue;
                double current_mA = 0.0;
                double diff = 0.0;

                while (left <= right)
                {
                    int mid = left + (right - left) / 2;

                    DRV_IBIAS_TRIM.Value = (uint)mid;
                    DRV_IBIAS_TRIM.Write();
                    Thread.Sleep(100);

                    List<double> measurements = new List<double>();
                    for (int i = 0; i < 3; i++)
                    {
                        string response = DigitalMultimeter1.WriteAndReadString("READ?");
                        if (double.TryParse(response, out double measuredValue))
                        {
                            measurements.Add(measuredValue);
                        }
                        else
                        {
                            throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                        }
                    }
                    current_mA = Math.Round(measurements.Average() * 1000, 2);

                    diff = Math.Round(Math.Abs(target - current_mA), 2);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestCode = mid;
                    }

                    if (current_mA < target)
                        left = mid + 1;
                    else
                        right = mid - 1;
                }

                DRV_IBIAS_TRIM.Value = (uint)bestCode;
                DRV_IBIAS_TRIM.Write();
                Thread.Sleep(100);

                List<double> finalMeasurements = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    string response = DigitalMultimeter1.WriteAndReadString("READ?");
                    if (double.TryParse(response, out double measuredValue))
                    {
                        finalMeasurements.Add(measuredValue);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                current_mA = Math.Round(finalMeasurements.Average() * 1000, 2);

                trimValue[0] = current_mA;
                trimValue[1] = bestCode;
                return trimValue;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_CAL_IBIAS_OF_DRV: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                PowerSupply0.Write("OUTP OFF, (@2)");
                SEG_TEST.Value = 0;
                SEG_TEST.Write();
            }
        }

        private double[] Run_MEAS_SEG_CURRENT(int xOffset)
        {
            if (DigitalMultimeter1 == null || !DigitalMultimeter1.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter1", "ERROR");
                return null;
            }

            var SEG_CURR = Parent.RegMgr.GetRegisterItem("I[5:0]");
            var SEG_TEST = Parent.RegMgr.GetRegisterItem("SEG_TEST");

            try
            {
                //string message = $"Run Measurement of each Segments Current.\nSEG1 ~ 16 / DigitalMultimeter1 / Current";
                //if (MessageBox.Show(message, "SCH1711 Initial Test Sequence", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                //{
                //    MessageBox.Show("테스트가 사용자에 의해 중단되었습니다.", "테스트 중단", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    return null;
                //}

                Parent.xlMgr.Sheet.Select("Seg_Current");
                Parent.xlMgr.Sheet.Activate();
                Parent.xlMgr.Cell.Write(3 + xOffset, 3, (xOffset + 1).ToString());

                SEG_CURR.Value = 63;
                SEG_CURR.Write();
                SEG_TEST.Value = 1;
                SEG_TEST.Write();
                Thread.Sleep(100);
                PowerSupply0.Write("VOLT 0.6, (@2)");
                PowerSupply0.Write("CURR 1, (@2)");
                PowerSupply0.Write("OUTP ON, (@2)");
                DigitalMultimeter1.Write("CONF:CURR:DC 0.1");

                int[] segCount = { 16, 24, 32 }; 
                double[] segCurrents = new double[segCount[_segMode - 1]];
                for (int i = 0; i < segCurrents.Length; i++)
                {
                    double avgBuff = 0;
                    bool checkConnection = true;

                    if (MessageBox.Show($"SEG{i + 1}를 Probing 후 확인을 눌러주세요.\nSEG{i + 1}", "Segment Channel mismatch", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        while (checkConnection)
                        {
                            string response = DigitalMultimeter1.WriteAndReadString("READ?");
                            double measuredValue;
                            double temp;

                            if (double.TryParse(response, out measuredValue))
                            {
                                temp = Math.Round(measuredValue * 1000, 2);
                            }
                            else
                            {
                                throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                            }

                            if (temp >= 40 || temp <= 20)
                            {
                                if (MessageBox.Show($"Check Segment Connection!\nDigitalMultimeter1 = {temp}", "Segment Channel mismatch", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                                {
                                    return null;
                                }
                            }
                            else
                            {
                                checkConnection = false;
                                avgBuff += temp;
                            }
                        }

                        for (int j = 0; j < 4; j++)
                        {
                            string response = DigitalMultimeter1.WriteAndReadString("READ?");
                            double measuredValue;
                            double current_mA;
                            if (double.TryParse(response, out measuredValue))
                            {
                                current_mA = Math.Round(measuredValue * 1000, 2);
                            }
                            else
                            {
                                throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                            }

                            avgBuff += current_mA;
                        }
                        segCurrents[i] = Math.Round(avgBuff / 5, 2);
                        Parent.xlMgr.Cell.Write(3 + xOffset, 4 + i, segCurrents[i].ToString());
                    }
                    else
                    {
                        return null;
                    }
                }

                double[] minmax = new double[2];
                minmax[0] = segCurrents.Max();
                minmax[1] = segCurrents.Min();

                Parent.xlMgr.Cell.Write(3 + xOffset, 36, minmax[0].ToString());
                Parent.xlMgr.Cell.Write(3 + xOffset, 37, minmax[1].ToString());

                Parent.xlMgr.Sheet.Select("Initial_Test");
                Parent.xlMgr.Sheet.Activate();
                for (int i = 0; i < 2; i++)
                {
                    double mismatchPercentage = Math.Round((minmax[i] - 30) / 30 * 100, 1);
                    string mismatch = $"{mismatchPercentage} ({minmax[i]})";
                    Parent.xlMgr.Cell.Write(19 + i, 7 + xOffset, mismatch);
                }

                return minmax;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_MEAS_SEG_CURRENT: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                PowerSupply0.Write("OUTP OFF, (@2)");
                SEG_TEST.Value = 0;
                SEG_TEST.Write();
            }
        }

        private double[] Run_TEST_STANDBY_CURRENT()
        {
            if (PowerSupply0 == null || !PowerSupply0.IsOpen)
            {
                MessageBox.Show("Check Connection Status of PowerSupply0", "ERROR");
                return null;
            }

            if (DigitalMultimeter2 == null || !DigitalMultimeter2.IsOpen)
            {
                MessageBox.Show("Check Connection Status of DigitalMultimeter2", "ERROR");
                return null;
            }

            double[] standbyCurrent = new double[2];

            try
            {
                var sleep_en = Parent.RegMgr.GetRegisterItem("SLEEP");

                DigitalMultimeter2.Write("SENS:FUNC 'CURR:DC'");
                DigitalMultimeter2.Write("SENS:CURR:DC:RANGE 1E-3");

                PowerSupply0.Write("VOLT 5, (@1)");
                PowerSupply0.Write("CURR 1, (@1)");
                PowerSupply0.Write("OUTP ON, (@1)");
                Thread.Sleep(1000);

                double temp = new double();

                for(int i = 0; i < 5; i++)
                {
                    string response = DigitalMultimeter2.WriteAndReadString($"READ?");
                    double measuredValue;

                    if (double.TryParse(response, out measuredValue))
                    {
                        temp += Math.Round(measuredValue * 1_000, 2);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                standbyCurrent[0] = Math.Round(temp / 5, 2);

                sleep_en.Value = 1;
                sleep_en.Write();
                WriteCommand(0x0E);
                Thread.Sleep(1000);

                DigitalMultimeter2.Write("SENS:CURR:DC:RANGE 1E-5");

                temp = new double();
                for (int i = 0; i < 5; i++)
                {
                    string response = DigitalMultimeter2.WriteAndReadString($"READ?");
                    double measuredValue;

                    if (double.TryParse(response, out measuredValue))
                    {
                        temp += Math.Round(measuredValue * 1_000_000, 2);
                    }
                    else
                    {
                        throw new FormatException($"계측기로부터 유효하지 않은 응답을 받았습니다: '{response}'");
                    }
                }
                standbyCurrent[1] = Math.Round(temp / 5, 2);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Standby Current Test: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                DigitalMultimeter2.Write("SENS:CURR:DC:RANGE:AUTO ON");
                PowerSupply0.Write("OUTP OFF, (@1, 2, 3)");
            }

            return standbyCurrent;
        }

        private void Run_INITIAL_TEST_SEQ(uint num)
        {
            Check_Instrument();

            int xOffset;
            if (num == 0)
                xOffset = 0;
            else
                xOffset = (int)num - 1;

            SPI.GPIOs[4].Direction = GPIO_Direction.Output;
            SPI.GPIOs[4].State = GPIO_State.Low;
            if (Parent.ChipCtrlButtons[4].Text == "GH0_H")
            {
                Parent.ChipCtrlButtons[4].Text = "GH0_L";
            }

            Parent.xlMgr.Sheet.Select("Initial_Test");
            Parent.xlMgr.Sheet.Activate();
            Parent.xlMgr.Cell.Write(2, 7 + xOffset, (xOffset + 1).ToString());

            PowerSupply0.Write("OUTP OFF, (@1, 2, 3)");
            DigitalMultimeter0.Write("SENS:FUNC 'VOLT:AC'");
            Thread.Sleep(1000);

            if (!Select_PKG_TYPE()) throw new OperationCanceledException("사용자에 의해 테스트가 중단되었습니다.");
            string[] pkgType = { "32QFN", "40QFN", "48QFN" };

            while (true)
            {
                Dictionary<string, int> bestCodes = new Dictionary<string, int>();
                double[] trimResult;

                try
                {
                    Parent.xlMgr.Cell.Write(3, 7 + xOffset, pkgType[_segMode - 1]);

                    //if (MessageBox.Show("Run POR Rising/Falling Test.\n\nProbe: N/A\nInstrument: PowerSupply0", "Initial_Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                    //{
                    //    throw new OperationCanceledException("사용자에 의해 테스트가 중단되었습니다.");
                    //}
                    trimResult = Run_TEST_POWER_ON_RESET();
                    if (trimResult == null) throw new Exception("Fail to POR Test");
                    Parent.xlMgr.Cell.Write(5, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(6, 7 + xOffset, trimResult[1].ToString());

                    //if (MessageBox.Show("Run Trim 125/150_REF, VPTAT_OUT.\n\nProbe: CS1\nInstrument: DigitalMultimeter0", "Initial_Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                    //{
                    //    throw new OperationCanceledException("사용자에 의해 테스트가 중단되었습니다.");
                    //}
                    trimResult = Run_CAL_125_REF();
                    if (trimResult == null) throw new Exception("Fail to Trim 125_REF");
                    bestCodes.Add("TSP_125_TRIM", (int)trimResult[1]);
                    Parent.xlMgr.Cell.Write(7, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(8, 7 + xOffset, trimResult[1].ToString());

                    trimResult = Run_CAL_150_REF();
                    if (trimResult == null) throw new Exception("Fail to Trim 150_REF");
                    bestCodes.Add("TSP_150_TRIM", (int)trimResult[1]);
                    Parent.xlMgr.Cell.Write(9, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(10, 7 + xOffset, trimResult[1].ToString());

                    trimResult = Run_CAL_PTAT_OUT();
                    if (trimResult == null) throw new Exception("Fail to Trim PTAT_OUT");
                    bestCodes.Add("VPTAT_OUT_TRIM", (int)trimResult[1]);
                    Parent.xlMgr.Cell.Write(11, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(12, 7 + xOffset, trimResult[1].ToString());

                    if (MessageBox.Show("Run Trim OSC Frequency.\n\nProbe: CS1\nInstrument: OscilloScope0 Port 1", "Initial_Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                    {
                        throw new OperationCanceledException("사용자에 의해 테스트가 중단되었습니다.");
                    }
                    trimResult = Run_CAL_OSC_FREQ();
                    if (trimResult == null) throw new Exception("Fail to Trim OSC");
                    bestCodes.Add("RC_OSC_BIAS_TRIM", (int)trimResult[1]);
                    Parent.xlMgr.Cell.Write(13, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(14, 7 + xOffset, trimResult[1].ToString());

                    //if (MessageBox.Show("Run Trim Segments Current.\n\nProbe: SEG16\nInstrument: DigitalMultimeter1", "Initial_Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                    //{
                    //    throw new OperationCanceledException("사용자에 의해 테스트가 중단되었습니다.");
                    //}
                    trimResult = Run_CAL_REXT_CURRENT();
                    if (trimResult == null) throw new Exception("Fail to Trim REXT Current");
                    bestCodes.Add("BGR_OUT_TRIM", (int)trimResult[1]);
                    Parent.xlMgr.Cell.Write(15, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(16, 7 + xOffset, trimResult[1].ToString());

                    trimResult = Run_CAL_IBIAS_OF_DRV();
                    if (trimResult == null) throw new Exception("Fail to Trim IBIAS of DRV");
                    bestCodes.Add("DRV_IBIAS_TRIM", (int)trimResult[1]);
                    Parent.xlMgr.Cell.Write(17, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(18, 7 + xOffset, trimResult[1].ToString());

                    Log.WriteLine($"[Chip {xOffset + 1}] Trim Codes Saved: {string.Join(", ", bestCodes.Select(kvp => $"{kvp.Key}={kvp.Value}"))}");

                    double[] currentMismatch = Run_MEAS_SEG_CURRENT(xOffset);
                    Thread.Sleep(1000);

                    byte[] wdata = Build_eFUSE_DATA(bestCodes);

                    string confirmationMessage = "Check eFuse Parameters.\n\n";
                    confirmationMessage += "■ Trim Codes\n" + string.Join("\n", bestCodes.Select(kvp => $"  {kvp.Key,-20}: {kvp.Value}")) + "\n\n";
                    confirmationMessage += "■ Fixed Parameters\n";
                    confirmationMessage += $"  {"LsdRefTrim",-20}: {LsdRefTrim} (0b{System.Convert.ToString(LsdRefTrim, 2)})\n";
                    confirmationMessage += $"  {"LodRefTrim",-20}: {LodRefTrim} (0b{System.Convert.ToString(LodRefTrim, 2)})\n";
                    confirmationMessage += $"  {"BgrFeOutTrim",-20}: {BgrFeOutTrim} (0b{System.Convert.ToString(BgrFeOutTrim, 2)})\n";
                    confirmationMessage += $"  {"RcOscCapTrim",-20}: {RcOscCapTrim} (0b{System.Convert.ToString(RcOscCapTrim, 2)})\n";
                    confirmationMessage += $"  {"Crcin",-20}: {Crcin} (0b{System.Convert.ToString(Crcin, 2)})\n";
                    confirmationMessage += $"  {"IMax",-20}: {IMax} (0b{System.Convert.ToString(IMax, 2)})\n";
                    confirmationMessage += $"  {"_segMode",-20}: {_segMode} (0b{System.Convert.ToString(_segMode, 2)})\n";
                    confirmationMessage += $"  {"GridMode",-20}: {GridMode} (0b{System.Convert.ToString(GridMode, 2)})\n";
                    confirmationMessage += $"  {"ChgEn",-20}: {ChgEn} (0b{System.Convert.ToString(ChgEn, 2)})\n";
                    confirmationMessage += $"  {"DchgEn",-20}: {DchgEn} (0b{System.Convert.ToString(DchgEn, 2)})\n";
                    confirmationMessage += $"  {"Res",-20}: {Res} (0b{System.Convert.ToString(Res, 2)})\n";
                    confirmationMessage += $"  {"ProgramFlag",-20}: {ProgramFlag} (0b{System.Convert.ToString(ProgramFlag, 2)})\n";
                    confirmationMessage += $"  {"IReg",-20}: {IReg} (0b{System.Convert.ToString(IReg, 2)})\n";

                    if (MessageBox.Show(confirmationMessage, "eFuse Write", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        for (uint i = 0; i < wdata.Length; i++)
                        {
                            var EF_WDATA = Parent.RegMgr.GetRegisterItem($"EF_WDATA{i}[7:0]");
                            EF_WDATA.Value = wdata[i];
                            EF_WDATA.Write();
                            Thread.Sleep(10);
                        }

                        var EF_PASSCODE = Parent.RegMgr.GetRegisterItem("EF_PASSCODE[7:0]");
                        EF_PASSCODE.Value = 0x38;
                        EF_PASSCODE.Write();
                        Thread.Sleep(100);

                        PowerSupply0.Write("VOLT 5.5, (@1)");
                        Thread.Sleep(500);

                        var EF_PGM = Parent.RegMgr.GetRegisterItem("pgm");
                        EF_PGM.Value = 1;
                        EF_PGM.Write();
                        Thread.Sleep(100);
                        EF_PGM.Value = 0;
                        EF_PGM.Write();
                        Thread.Sleep(100);

                        PowerSupply0.Write("VOLT 5, (@1)");
                        Thread.Sleep(500);

                        var EF_READ = Parent.RegMgr.GetRegisterItem("read");
                        EF_READ.Value = 1;
                        EF_READ.Write();
                        Thread.Sleep(100);
                        EF_READ.Value = 0;
                        EF_READ.Write();
                        Thread.Sleep(100);

                        byte[] readData = new byte[8];
                        bool isMatch = true;

                        for (uint i = 0; i < readData.Length; i++)
                        {
                            var EF_RDATA = Parent.RegMgr.GetRegisterItem($"EF_RDATA{i}[7:0]");
                            EF_RDATA.Read();
                            readData[i] = (byte)EF_RDATA.Value;
                            if (wdata[i] != readData[i])
                            {
                                isMatch = false;
                            }
                            Thread.Sleep(100);
                        }

                        if (isMatch)
                        {
                            Log.WriteLine($"eFuse Verification PASSED.");
                            Parent.xlMgr.Cell.Write(21, 7 + xOffset, "PASS");
                        }
                        else
                        {
                            string writeStr = string.Join(" ", wdata.Select(b => $"0x{b:X2}"));
                            string readStr = string.Join(" ", readData.Select(b => $"0x{b:X2}"));
                            string errorMsg = $"eFuse Verification FAILED.\n\nWritten: {writeStr}\nRead:    {readStr}";
                            Parent.xlMgr.Cell.Write(21, 7 + xOffset, "FAIL");
                            throw new Exception(errorMsg);
                        }
                    }
                    else
                    {
                        throw new OperationCanceledException("eFuse 기록이 사용자에 의해 취소되었습니다.");
                    }

                    PowerSupply0.Write("OUTP OFF, (@1, 2, 3)");
                    Thread.Sleep(1500);

                    trimResult = Run_TEST_STANDBY_CURRENT();
                    if (trimResult == null) throw new Exception("Fail to Test Standby Current.");
                    Parent.xlMgr.Cell.Write(22, 7 + xOffset, trimResult[0].ToString());
                    Parent.xlMgr.Cell.Write(23, 7 + xOffset, trimResult[1].ToString());

                    if (MessageBox.Show($"Chip {xOffset + 1} 테스트 완료. 다음 칩을 진행하시겠습니까?\nNext: Chip {xOffset + 2}", "Initial_Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        xOffset++;
                        Parent.xlMgr.Sheet.Select("Initial_Test");
                        Parent.xlMgr.Sheet.Activate();
                        Parent.xlMgr.Cell.Write(2, 7 + xOffset, (xOffset + 1).ToString());
                    }
                    else
                    {
                        MessageBox.Show("모든 테스트를 종료합니다.", "Initial_Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fail to Test: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
            }
        }

        #region eFuse Parameters
        private const int LsdRefTrim = 0;
        private const int LodRefTrim = 7;
        private const int BgrFeOutTrim = 0;
        private const int RcOscCapTrim = 0;
        private const int Crcin = 0b1110;
        private const int IMax = 0b1111;
        private int _segMode;
        private const int GridMode = 1;
        private const int ChgEn = 1;
        private const int DchgEn = 1;
        private const int Res = 0;
        private const int ProgramFlag = 1;
        private const int IReg = 0;
        #endregion eFuse Parameters

        private bool Select_PKG_TYPE()
        {
            using (var form = new Form())
            {
                form.Text = "Initial_Test";
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Width = 250;
                form.Height = 200;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var label = new Label() { Left = 20, Top = 20, Text = "Select PKG Type:" };
                var rb48 = new RadioButton() { Text = "48QFN", Left = 40, Top = 50, Checked = true };
                var rb40 = new RadioButton() { Text = "40QFN", Left = 40, Top = 75 };
                var rb32 = new RadioButton() { Text = "32QFN", Left = 40, Top = 100 };

                var okButton = new Button() { Text = "확인", Left = 30, Width = 80, Top = 130, DialogResult = DialogResult.OK };
                var cancelButton = new Button() { Text = "취소", Left = 120, Width = 80, Top = 130, DialogResult = DialogResult.Cancel };

                form.Controls.Add(label);
                form.Controls.Add(rb48);
                form.Controls.Add(rb40);
                form.Controls.Add(rb32);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    if (rb48.Checked) this._segMode = 0b11; // 48QFN
                    else if (rb40.Checked) this._segMode = 0b10; // 40QFN
                    else if (rb32.Checked) this._segMode = 0b01; // 32QFN

                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public void Verify_CRC_CALCULATION()
        {
            var manualTrimResults = new Dictionary<string, int>
            {
                { "TSP_125_TRIM", 0 },
                { "TSP_150_TRIM", 0 },
                { "VPTAT_OUT_TRIM", 0 },
                { "BGR_OUT_TRIM", 0b11111 },
                { "DRV_IBIAS_TRIM", 0 },
                { "RC_OSC_BIAS_TRIM", 0 }
            };

            int calculatedCrc = Run_CALCULATE_CRC(manualTrimResults);

            Log.WriteLine("--- Manual CRC Verification ---");
            Log.WriteLine(" [Input Trim Values]");
            foreach (var entry in manualTrimResults)
            {
                Log.WriteLine($"    {entry.Key,-20}: {entry.Value}");
            }
            Log.WriteLine("-----------------------------");
            Log.WriteLine($" [Calculated CRC] : {calculatedCrc} (Binary: {System.Convert.ToString(calculatedCrc, 2).PadLeft(4, '0')}, Hex: 0x{calculatedCrc:X})");
            Log.WriteLine("-----------------------------");
        }

        private int GetBit(int value, int bitPosition)
        {
            return (value >> bitPosition) & 1;
        }

        private int Run_CALCULATE_CRC(Dictionary<string, int> trimResults)
        {
            int tsp125Trim;
            if (!trimResults.TryGetValue("TSP_125_TRIM", out tsp125Trim)) { tsp125Trim = 0; }

            int tsp150Trim;
            if (!trimResults.TryGetValue("TSP_150_TRIM", out tsp150Trim)) { tsp150Trim = 0; }

            int vptatOutTrim;
            if (!trimResults.TryGetValue("VPTAT_OUT_TRIM", out vptatOutTrim)) { vptatOutTrim = 0; }

            int bgrOutTrim;
            if (!trimResults.TryGetValue("BGR_OUT_TRIM", out bgrOutTrim)) { bgrOutTrim = 0; }

            int drvIbiasTrim;
            if (!trimResults.TryGetValue("DRV_IBIAS_TRIM", out drvIbiasTrim)) { drvIbiasTrim = 0; }

            int rcOscBiasTrim;
            if (!trimResults.TryGetValue("RC_OSC_BIAS_TRIM", out rcOscBiasTrim)) { rcOscBiasTrim = 0; }

            int crc3 = 0;
            crc3 ^= GetBit(Crcin, 0) ^ GetBit(Crcin, 1) ^ GetBit(Crcin, 3);
            crc3 ^= GetBit(rcOscBiasTrim, 0) ^ GetBit(rcOscBiasTrim, 1) ^ GetBit(rcOscBiasTrim, 3);
            crc3 ^= GetBit(drvIbiasTrim, 0) ^ GetBit(drvIbiasTrim, 1) ^ GetBit(drvIbiasTrim, 3);
            crc3 ^= GetBit(IReg, 2) ^ GetBit(IReg, 3) ^ GetBit(IReg, 5);
            crc3 ^= GetBit(BgrFeOutTrim, 2) ^ GetBit(BgrFeOutTrim, 3);
            crc3 ^= GetBit(bgrOutTrim, 0) ^ GetBit(bgrOutTrim, 4);
            crc3 ^= GetBit(vptatOutTrim, 0) ^ GetBit(vptatOutTrim, 2);
            crc3 ^= GetBit(tsp150Trim, 1) ^ GetBit(tsp150Trim, 2) ^ GetBit(tsp150Trim, 4);
            crc3 ^= GetBit(tsp125Trim, 3) ^ GetBit(tsp125Trim, 4);
            crc3 ^= GetBit(LodRefTrim, 1) ^ GetBit(LsdRefTrim, 2);
            crc3 ^= GridMode ^ GetBit(IMax, 1) ^ ChgEn ^ GetBit(_segMode, 0) ^ ProgramFlag;

            int crc2 = 0;
            crc2 ^= GetBit(Crcin, 1) ^ GetBit(Crcin, 2) ^ GetBit(Crcin, 3);
            crc2 ^= GetBit(rcOscBiasTrim, 1) ^ GetBit(rcOscBiasTrim, 2) ^ GetBit(rcOscBiasTrim, 3) ^ GetBit(RcOscCapTrim, 1);
            crc2 ^= GetBit(drvIbiasTrim, 1) ^ GetBit(drvIbiasTrim, 2) ^ GetBit(drvIbiasTrim, 3);
            crc2 ^= GetBit(IReg, 1) ^ GetBit(IReg, 3) ^ GetBit(IReg, 4) ^ GetBit(IReg, 5);
            crc2 ^= GetBit(BgrFeOutTrim, 1) ^ GetBit(BgrFeOutTrim, 3) ^ GetBit(BgrFeOutTrim, 4);
            crc2 ^= GetBit(bgrOutTrim, 0) ^ GetBit(bgrOutTrim, 3);
            crc2 ^= GetBit(vptatOutTrim, 0) ^ GetBit(vptatOutTrim, 1) ^ GetBit(vptatOutTrim, 2);
            crc2 ^= GetBit(tsp150Trim, 0) ^ GetBit(tsp150Trim, 2) ^ GetBit(tsp150Trim, 3) ^ GetBit(tsp150Trim, 4);
            crc2 ^= GetBit(tsp125Trim, 2) ^ GetBit(tsp125Trim, 4);
            crc2 ^= GetBit(LodRefTrim, 0) ^ GetBit(LodRefTrim, 1) ^ GetBit(LsdRefTrim, 1);
            crc2 ^= GridMode ^ GetBit(IMax, 0) ^ GetBit(IMax, 1) ^ DchgEn ^ GetBit(_segMode, 0) ^ GetBit(_segMode, 1) ^ ProgramFlag;

            int crc1 = 0;
            crc1 ^= GetBit(Crcin, 2) ^ GetBit(Crcin, 3);
            crc1 ^= GetBit(rcOscBiasTrim, 2) ^ GetBit(rcOscBiasTrim, 3) ^ GetBit(RcOscCapTrim, 0);
            crc1 ^= GetBit(drvIbiasTrim, 2) ^ GetBit(drvIbiasTrim, 3);
            crc1 ^= GetBit(IReg, 0) ^ GetBit(IReg, 4) ^ GetBit(IReg, 5);
            crc1 ^= GetBit(BgrFeOutTrim, 0) ^ GetBit(BgrFeOutTrim, 4);
            crc1 ^= GetBit(bgrOutTrim, 0) ^ GetBit(bgrOutTrim, 2);
            crc1 ^= GetBit(vptatOutTrim, 1) ^ GetBit(vptatOutTrim, 2) ^ GetBit(vptatOutTrim, 4);
            crc1 ^= GetBit(tsp150Trim, 3) ^ GetBit(tsp150Trim, 4);
            crc1 ^= GetBit(tsp125Trim, 1);
            crc1 ^= GetBit(LodRefTrim, 0) ^ GetBit(LodRefTrim, 1) ^ GetBit(LsdRefTrim, 0);
            crc1 ^= GetBit(IMax, 0) ^ GetBit(IMax, 1) ^ GetBit(IMax, 3);
            crc1 ^= GetBit(_segMode, 1) ^ ProgramFlag;

            int crc0 = 0;
            crc0 ^= GetBit(Crcin, 1) ^ GetBit(Crcin, 2);
            crc0 ^= GetBit(rcOscBiasTrim, 1) ^ GetBit(rcOscBiasTrim, 2) ^ GetBit(rcOscBiasTrim, 4);
            crc0 ^= GetBit(drvIbiasTrim, 1) ^ GetBit(drvIbiasTrim, 2) ^ GetBit(drvIbiasTrim, 4);
            crc0 ^= GetBit(IReg, 3) ^ GetBit(IReg, 4) ^ Res;
            crc0 ^= GetBit(BgrFeOutTrim, 3) ^ GetBit(BgrFeOutTrim, 4);
            crc0 ^= GetBit(bgrOutTrim, 1);
            crc0 ^= GetBit(vptatOutTrim, 0) ^ GetBit(vptatOutTrim, 1) ^ GetBit(vptatOutTrim, 3);
            crc0 ^= GetBit(tsp150Trim, 2) ^ GetBit(tsp150Trim, 3);
            crc0 ^= GetBit(tsp125Trim, 0) ^ GetBit(tsp125Trim, 4);
            crc0 ^= GetBit(LodRefTrim, 0) ^ GetBit(LodRefTrim, 2);
            crc0 ^= GridMode ^ GetBit(IMax, 0) ^ GetBit(IMax, 2);
            crc0 ^= GetBit(_segMode, 0) ^ GetBit(_segMode, 1);

            return (crc3 << 3) | (crc2 << 2) | (crc1 << 1) | crc0;
        }

        private byte[] Build_eFUSE_DATA(Dictionary<string, int> trimResults)
        {
            byte[] wdata = new byte[8];

            int tsp125Trim;
            if (!trimResults.TryGetValue("TSP_125_TRIM", out tsp125Trim)) { tsp125Trim = 0; }

            int tsp150Trim;
            if (!trimResults.TryGetValue("TSP_150_TRIM", out tsp150Trim)) { tsp150Trim = 0; }

            int vptatOutTrim;
            if (!trimResults.TryGetValue("VPTAT_OUT_TRIM", out vptatOutTrim)) { vptatOutTrim = 0; }

            int bgrOutTrim;
            if (!trimResults.TryGetValue("BGR_OUT_TRIM", out bgrOutTrim)) { bgrOutTrim = 0; }

            int drvIbiasTrim;
            if (!trimResults.TryGetValue("DRV_IBIAS_TRIM", out drvIbiasTrim)) { drvIbiasTrim = 0; }

            int rcOscBiasTrim;
            if (!trimResults.TryGetValue("RC_OSC_BIAS_TRIM", out rcOscBiasTrim)) { rcOscBiasTrim = 0; }

            int crcValue = Run_CALCULATE_CRC(trimResults);

            wdata[0] = (byte)(
                (GetBit(rcOscBiasTrim, 0) << 0) |
                (GetBit(rcOscBiasTrim, 1) << 1) |
                (GetBit(rcOscBiasTrim, 2) << 2) |
                (GetBit(rcOscBiasTrim, 3) << 3) |
                (GetBit(rcOscBiasTrim, 4) << 4) |
                (GetBit(RcOscCapTrim, 0) << 5) |
                (GetBit(RcOscCapTrim, 1) << 6) |
                (GetBit(drvIbiasTrim, 0) << 7)
            );

            wdata[1] = (byte)(
                (GetBit(drvIbiasTrim, 1) << 0) |
                (GetBit(drvIbiasTrim, 2) << 1) |
                (GetBit(drvIbiasTrim, 3) << 2) |
                (GetBit(drvIbiasTrim, 4) << 3) |
                (GetBit(IReg, 0) << 4) |
                (GetBit(IReg, 1) << 5) |
                (GetBit(IReg, 2) << 6) |
                (GetBit(IReg, 3) << 7)
            );

            wdata[2] = (byte)(
                (GetBit(IReg, 4) << 0) |
                (GetBit(IReg, 5) << 1) |
                (GetBit(Res, 0) << 2) |
                (GetBit(BgrFeOutTrim, 0) << 3) |
                (GetBit(BgrFeOutTrim, 1) << 4) |
                (GetBit(BgrFeOutTrim, 2) << 5) |
                (GetBit(BgrFeOutTrim, 3) << 6) |
                (GetBit(BgrFeOutTrim, 4) << 7)
            );

            wdata[3] = (byte)(
                (GetBit(bgrOutTrim, 0) << 0) |
                (GetBit(bgrOutTrim, 1) << 1) |
                (GetBit(bgrOutTrim, 2) << 2) |
                (GetBit(bgrOutTrim, 3) << 3) |
                (GetBit(bgrOutTrim, 4) << 4) |
                (GetBit(vptatOutTrim, 0) << 5) |
                (GetBit(vptatOutTrim, 1) << 6) |
                (GetBit(vptatOutTrim, 2) << 7)
            );

            wdata[4] = (byte)(
                (GetBit(vptatOutTrim, 3) << 0) |
                (GetBit(vptatOutTrim, 4) << 1) |
                (GetBit(tsp150Trim, 0) << 2) |
                (GetBit(tsp150Trim, 1) << 3) |
                (GetBit(tsp150Trim, 2) << 4) |
                (GetBit(tsp150Trim, 3) << 5) |
                (GetBit(tsp150Trim, 4) << 6) |
                (GetBit(tsp125Trim, 0) << 7)
            );

            wdata[5] = (byte)(
                (GetBit(tsp125Trim, 1) << 0) |
                (GetBit(tsp125Trim, 2) << 1) |
                (GetBit(tsp125Trim, 3) << 2) |
                (GetBit(tsp125Trim, 4) << 3) |
                (GetBit(LodRefTrim, 0) << 4) |
                (GetBit(LodRefTrim, 1) << 5) |
                (GetBit(LodRefTrim, 2) << 6) |
                (GetBit(LsdRefTrim, 0) << 7)
            );

            wdata[6] = (byte)(
                (GetBit(LsdRefTrim, 1) << 0) |
                (GetBit(LsdRefTrim, 2) << 1) |
                (GridMode << 2) |
                (GetBit(IMax, 0) << 3) |
                (GetBit(IMax, 1) << 4) |
                (GetBit(IMax, 2) << 5) |
                (GetBit(IMax, 3) << 6) |
                (DchgEn << 7)
            );

            wdata[7] = (byte)(
                (ChgEn << 0) |
                (GetBit(_segMode, 0) << 1) |
                (GetBit(_segMode, 1) << 2) |
                (ProgramFlag << 3) |
                (GetBit(crcValue, 0) << 4) |
                (GetBit(crcValue, 1) << 5) |
                (GetBit(crcValue, 2) << 6) |
                (GetBit(crcValue, 3) << 7)
            );

            return wdata;
        }
        #endregion Function for Auto Test

        #region Function for LED Test
        static void DelayMicroseconds(int microseconds)
        {
            if (microseconds > 1000)
            {
                Thread.Sleep(microseconds / 1000);
                microseconds %= 1000;
            }

            long ticks = microseconds * (Stopwatch.Frequency / 1000000);
            long start = Stopwatch.GetTimestamp();
            while (Stopwatch.GetTimestamp() - start < ticks)
            {
                Thread.Yield();
            }
        }

        private Tuple<int, int> GetGridDimensions()
        {
            uint reg408_val = ReadRegister(0x408);
            uint reg401_val = ReadRegister(0x401);
            bool is9GridMode = ((reg408_val >> 2) & 0x01) == 1;
            uint g_n = (reg401_val >> 2) & 0x0F;
            int numGrids = 0;
            if (is9GridMode)
            {
                if (g_n <= 7) numGrids = (int)g_n + 1;
                else numGrids = 9;
            }
            else
            {
                if (g_n <= 8) numGrids = (int)g_n + 1;
                else numGrids = 10;
            }
            uint _segMode = reg408_val & 0x03;
            int segmentsPerGrid;
            switch (_segMode)
            {
                case 1: segmentsPerGrid = 16; break;
                case 2: segmentsPerGrid = 24; break;
                case 3: segmentsPerGrid = 32; break;
                default: segmentsPerGrid = 0; break;
            }
            int totalBytes = numGrids * segmentsPerGrid;
            return new Tuple<int, int>(totalBytes, segmentsPerGrid);
        }

        private int CalculateChicagoGridSizeBytes()
        {
            return GetGridDimensions().Item1;
        }

        private void RunWriteLEDData()
        {
            var dimensions = GetGridDimensions();
            int totalBytes = dimensions.Item1;
            int segmentsPerGrid = dimensions.Item2;

            if (totalBytes == 0)
            {
                MessageBox.Show("계산된 LED 개수가 0입니다. (SEG_MODE = 0)", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var hexForm = new HexWriteForm(totalBytes, segmentsPerGrid))
            {
                if (hexForm.ShowDialog(Parent) == DialogResult.OK)
                {
                    WriteLEDDataWithHexForm((uint)totalBytes, hexForm.ResultData);
                }
            }
        }

        private void WriteNormalHeader()
        {
            byte[] header = new byte[6];
            header[0] = 0x00;
            header[1] = 0x00;
            header[2] = 0x5A;
            header[3] = 0xFF;
            header[4] = 0x02;
            header[5] = (byte)((header[2] + header[3] + header[4]) & 0xFF);
            SPI.WriteBytes(header, 6, true);
        }

        private void WriteHeaderWithByteLoss(Random random)
        {
            byte[] header = new byte[6];
            header[0] = 0x00; header[1] = 0x00;
            header[2] = 0x5A; header[3] = 0xFF; header[4] = 0x02;
            header[5] = (byte)((header[2] + header[3] + header[4]) & 0xFF);

            int lengthToSend = random.Next(1, 6);
            //Log.WriteLine($"   -> Type: Header Byte Loss (Sent {lengthToSend}/6 bytes)", Color.DarkMagenta, Parent.BackColor);
            SPI.WriteBytes(header, lengthToSend, true);
        }

        private void WriteHeaderWithExtraByte(Random random)
        {
            byte[] originalHeader = new byte[6];
            originalHeader[0] = 0x00; originalHeader[1] = 0x00;
            originalHeader[2] = 0x5A; originalHeader[3] = 0xFF; originalHeader[4] = 0x02;
            originalHeader[5] = (byte)((originalHeader[2] + originalHeader[3] + originalHeader[4]) & 0xFF);

            int insertionIndex = random.Next(1, 6);
            byte garbageData = (byte)random.Next(256);

            byte[] extraHeader = new byte[7];
            Array.Copy(originalHeader, 0, extraHeader, 0, insertionIndex);
            extraHeader[insertionIndex] = garbageData;
            Array.Copy(originalHeader, insertionIndex, extraHeader, insertionIndex + 1, originalHeader.Length - insertionIndex);

            //Log.WriteLine($"   -> Type: Header Extra Byte (Value: 0x{garbageData:X2} inserted at index {insertionIndex})", Color.DarkMagenta, Parent.BackColor);

            SPI.WriteBytes(extraHeader, extraHeader.Length, true);
        }

        private void WriteHeaderWithCorruption(Random random)
        {
            byte[] corruptedHeader = new byte[6];
            corruptedHeader[0] = 0x00; corruptedHeader[1] = 0x00;
            corruptedHeader[2] = 0x5A; corruptedHeader[3] = 0xFF; corruptedHeader[4] = 0x02;
            corruptedHeader[5] = (byte)((corruptedHeader[2] + corruptedHeader[3] + corruptedHeader[4]) & 0xFF);

            int indexToCorrupt = random.Next(6);
            byte originalByte = corruptedHeader[indexToCorrupt];
            byte corruptedByte = (byte)random.Next(256);
            while (corruptedByte == originalByte)
            {
                corruptedByte = (byte)random.Next(256);
            }
            corruptedHeader[indexToCorrupt] = corruptedByte;

            //Log.WriteLine($"   -> Type: Header Corruption (Index: {indexToCorrupt}, Original: 0x{originalByte:X2}, New: 0x{corruptedByte:X2})", Color.DarkMagenta, Parent.BackColor);
            SPI.WriteBytes(corruptedHeader, corruptedHeader.Length, true);
        }

        private void WriteLEDDataWithHexForm(uint length, byte[] hexData)
        {
            if (length == 0) return;
            if (length > 320) length = 320;
            if (hexData == null || hexData.Length < length) return;

            byte[] header = new byte[4];
            header[0] = 0x5A; header[1] = 0xFF; header[2] = 0x02;
            header[3] = (byte)((header[0] + header[1] + header[2]) & 0xFF);
            SPI.WriteBytes(header, 4, true);

            byte[] bytes = new byte[length + 1];
            Array.Copy(hexData, 0, bytes, 0, length);

            byte cksum = 0;
            for (int i = 0; i < length; i++) cksum += bytes[i];
            bytes[bytes.Length - 1] = cksum;

            SPI.WriteBytesForChicago(bytes, bytes.Length, true);
        }

        private void WriteLEDData_PayloadOnly(uint length, byte[] hexData)
        {
            if (length == 0) return;
            if (length > 320) length = 320;
            if (hexData == null || hexData.Length < length) return;

            byte[] bytes = new byte[length + 1];
            Array.Copy(hexData, 0, bytes, 0, length);

            byte cksum = 0;
            for (int i = 0; i < length; i++) cksum += bytes[i];
            bytes[bytes.Length - 1] = cksum;

            SPI.WriteBytesForChicago(bytes, bytes.Length, true);
        }

        private void WriteLEDDataWithByteLoss_PayloadOnly(uint length, byte[] hexData, Random random)
        {
            if (length < 2)
            {
                WriteLEDDataWithHexForm(length, hexData);
                return;
            }
            if (length > 320) length = 320;
            if (hexData == null || hexData.Length < length) return;

            int lossType = random.Next(2);

            if (lossType == 0)
            {
                //Log.WriteLine("   -> Type: Last Byte Loss (Checksum dropped)", Color.DarkRed, Parent.BackColor);

                byte[] bytes = new byte[length + 1];
                Array.Copy(hexData, 0, bytes, 0, length);

                byte cksum = 0;
                for (int i = 0; i < length; i++) cksum += bytes[i];
                bytes[length] = cksum;

                SPI.WriteBytesForChicago(bytes, bytes.Length - 1, true);
            }
            else
            {
                int indexToDrop = random.Next((int)length);
                //Log.WriteLine("   -> Type: Middle Byte Loss (with original checksum)", Color.DarkRed, Parent.BackColor);

                byte[] lossyPayload = new byte[length - 1];
                if (indexToDrop > 0)
                    Array.Copy(hexData, 0, lossyPayload, 0, indexToDrop);
                if (indexToDrop < length - 1)
                    Array.Copy(hexData, indexToDrop + 1, lossyPayload, indexToDrop, (int)length - 1 - indexToDrop);

                byte[] bytesToSend = new byte[length];
                Array.Copy(lossyPayload, 0, bytesToSend, 0, lossyPayload.Length);

                byte originalCksum = 0;
                for (int i = 0; i < length; i++) originalCksum += hexData[i];

                bytesToSend[bytesToSend.Length - 1] = originalCksum;

                SPI.WriteBytesForChicago(bytesToSend, bytesToSend.Length, true);
            }
        }

        private void WriteLEDDataWithExtraByte_PayloadOnly(uint length, byte[] hexData, Random random)
        {
            if (length == 0) return;
            if (length > 320) length = 320;
            if (hexData == null || hexData.Length < length) return;

            byte[] originalPacket = new byte[length + 1];
            Array.Copy(hexData, 0, originalPacket, 0, length);
            byte originalCksum = 0;
            for (int i = 0; i < length; i++) originalCksum += hexData[i];
            originalPacket[length] = originalCksum;

            int insertionIndex = random.Next(1, (int)length);
            byte garbageData = (byte)random.Next(256);

            byte[] bytesToSend = new byte[originalPacket.Length + 1];
            Array.Copy(originalPacket, 0, bytesToSend, 0, insertionIndex);
            bytesToSend[insertionIndex] = garbageData;
            Array.Copy(originalPacket, insertionIndex, bytesToSend, insertionIndex + 1, originalPacket.Length - insertionIndex);

            //Log.WriteLine($"   -> Type: Payload Extra Byte (Value: 0x{garbageData:X2} inserted at index {insertionIndex})", Color.DarkRed, Parent.BackColor);

            SPI.WriteBytesForChicago(bytesToSend, bytesToSend.Length, true);
        }

        private void WriteLEDDataWithCorruption_PayloadOnly(uint length, byte[] hexData, Random random)
        {
            if (length == 0) return;
            if (length > 320) length = 320;
            if (hexData == null || hexData.Length < length) return;

            byte[] bytesToSend = new byte[length + 1];
            Array.Copy(hexData, 0, bytesToSend, 0, length);

            int indexToCorrupt = random.Next((int)length);

            byte originalByte = bytesToSend[indexToCorrupt];
            byte corruptedByte = (byte)random.Next(256);
            while (corruptedByte == originalByte)
            {
                corruptedByte = (byte)random.Next(256);
            }

            bytesToSend[indexToCorrupt] = corruptedByte;

            byte originalCksum = 0;
            for (int i = 0; i < length; i++) originalCksum += hexData[i];

            bytesToSend[length] = originalCksum;

            //Log.WriteLine($"   -> Type: Payload Corruption (Index: {indexToCorrupt}, Original: 0x{originalByte:X2}, New: 0x{corruptedByte:X2})", Color.DarkRed, Parent.BackColor);
            SPI.WriteBytesForChicago(bytesToSend, bytesToSend.Length, true);
        }

        private void RunLEDFunctionTest(object sender, EventArgs e)
        {
            if (MessageBox.Show("LED 기능 테스트를 시작하시겠습니까?", "테스트 시작", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
            {
                return;
            }

            int chipCount = 0;
            int passCount = 0;
            int failCount = 0;

            try
            {
                while (true)
                {
                    if (PowerSupply0 == null)
                        PowerSupply0 = new JLcLib.Instrument.SCPI(JLcLib.Instrument.InstrumentTypes.PowerSupply0);
                    if (PowerSupply0.IsOpen == false)
                        PowerSupply0.Open();

                    if (PowerSupply0.IsOpen == false)
                    {
                        MessageBox.Show("Check PowerSupply0 Connection!");
                        return;
                    }

                    PowerSupply0.Write("VOLT 5.0, (@1)");
                    PowerSupply0.Write("OUTP ON, (@1)");

                    Thread.Sleep(500);

                    RegisterItem CURRENT = Parent.RegMgr.GetRegisterItem("I[5:0]");
                    RegisterItem GRID_NUM = Parent.RegMgr.GetRegisterItem("G_N[3:0]");
                    RegisterItem UP = Parent.RegMgr.GetRegisterItem("UP");
                    RegisterItem DON = Parent.RegMgr.GetRegisterItem("DON");
                    RegisterItem SEG_MODE = Parent.RegMgr.GetRegisterItem("SEG_MODE[1:0]");
                    RegisterItem TSP_125_TRIM = Parent.RegMgr.GetRegisterItem("TSP_125_TRIM[4:0]");
                    RegisterItem TSP_150_TRIM = Parent.RegMgr.GetRegisterItem("TSP_150_TRIM[4:0]");
                    RegisterItem VPTAT_OUT_TRIM = Parent.RegMgr.GetRegisterItem("VPTAT_OUT_TRIM[4:0]");
                    RegisterItem BGR_OUT_TRIM = Parent.RegMgr.GetRegisterItem("BGR_OUT_TRIM[4:0]");
                    RegisterItem DRV_IBIAS_TRIM = Parent.RegMgr.GetRegisterItem("DRV_IBIAS_TRIM[4:0]");
                    RegisterItem RC_OSC_BIAS_TRIM = Parent.RegMgr.GetRegisterItem("RC_OSC_BIAS_TRIM[4:0]");

                    CURRENT.Value = 63;
                    CURRENT.Write();

                    GRID_NUM.Value = 8;
                    GRID_NUM.Write();

                    UP.Value = 1;
                    UP.Write();

                    DON.Value = 1;
                    DON.Write();

                    int totalBytes = CalculateChicagoGridSizeBytes();
                    if (totalBytes == 0)
                    {
                        MessageBox.Show("계산된 LED 개수가 0입니다. 칩 설정을 확인하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    byte[] Data = new byte[totalBytes + 1];
                    byte cksum = new byte();

                    for (int i = 0; i < totalBytes - 1; i++)
                    {
                        Data[i] = 0xFF;
                        cksum += Data[i];
                    }
                    Data[totalBytes] = (byte)(cksum & 0xFF);

                    WriteLEDDataWithHexForm((uint)totalBytes + 1, Data);

                    SEG_MODE.Read();
                    TSP_125_TRIM.Read();
                    TSP_150_TRIM.Read();
                    VPTAT_OUT_TRIM.Read();
                    BGR_OUT_TRIM.Read();
                    DRV_IBIAS_TRIM.Read();
                    RC_OSC_BIAS_TRIM.Read();

                    Log.WriteLine($"=================");
                    Log.WriteLine($"[Chip #{chipCount + 1}]");
                    Log.WriteLine($"TSP_125_TRIM[4:0] = {TSP_125_TRIM.Value}");
                    Log.WriteLine($"TSP_150_TRIM[4:0] = {TSP_150_TRIM.Value}");
                    Log.WriteLine($"VPTAT_OUT_TRIM[4:0] = {VPTAT_OUT_TRIM.Value}");
                    Log.WriteLine($"BGR_OUT_TRIM[4:0] = {BGR_OUT_TRIM.Value}");
                    Log.WriteLine($"DRV_IBIAS_TRIM[4:0] = {DRV_IBIAS_TRIM.Value}");
                    Log.WriteLine($"RC_OSC_BIAS_TRIM[4:0] = {RC_OSC_BIAS_TRIM.Value}");
                    Log.WriteLine($"SEG_MODE[1:0] = {SEG_MODE.Value}");
                    Log.WriteLine($"=================");

                    chipCount++;

                    string passFailMessage = $"칩 #{chipCount}의 테스트 결과를 판정하십시오.\n\n  - Pass -> Yes\n  - Fail -> No";
                    DialogResult result = MessageBox.Show(passFailMessage, "결과 판정", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        passCount++;
                    }
                    else
                    {
                        failCount++;
                    }

                    PowerSupply0?.Write("OUTP OFF, (@1)");

                    string continueMessage = $"[현재 상태] Pass: {passCount} | Fail: {failCount}\n" +
                                             $"총 {chipCount}개 칩 테스트 완료.\n\n" +
                                             "다음 칩 테스트를 진행하시겠습니까?";

                    if (MessageBox.Show(continueMessage, "진행 여부 확인", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                    {
                        break;
                    }
                }
            }
            finally
            {
                PowerSupply0?.Write("OUTP OFF, (@1)");
                Parent.ChipCtrlButtons[6].Text = "LED_ON";

                string finalSummary = $"테스트가 종료되었습니다.\n\n" +
                                      $"[최종 결과]\n" +
                                      $"- 총 테스트: {chipCount} 개\n" +
                                      $"- Pass: {passCount} 개\n" +
                                      $"- Fail: {failCount} 개";
                MessageBox.Show(finalSummary, "테스트 결과 요약");
            }
        }

        private void StopLEDMatrixEffect()
        {
            if (animationWorker != null && animationWorker.IsBusy)
            {
                animationWorker.CancelAsync();
            }
            else
            {
                Log.WriteLine("No effect is currently running.", Color.OrangeRed, Parent.BackColor);
            }
        }

        private bool CheckLEDEffectPrerequisites()
        {
            uint reg402_val = ReadRegister(0x402);

            uint don = (reg402_val >> 3) & 0x01;

            if (don == 0)
            {
                MessageBox.Show("LED 효과를 시작할 수 없습니다.\n\n→ 0x402: DON = 0.",
                                "Fail to Start LED Effect", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void RunWaveEffectTest()
        {
            if (!CheckLEDEffectPrerequisites()) return;

            if (animationWorker != null && animationWorker.IsBusy)
            {
                Log.WriteLine("Another animation is already running.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            var dimensions = GetGridDimensions();
            int totalBytes = dimensions.Item1;
            int segmentsPerGrid = dimensions.Item2;

            if (totalBytes == 0 || segmentsPerGrid == 0)
            {
                MessageBox.Show("계산된 LED 개수가 0입니다. 칩 설정을 확인하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            animationWorker = new BackgroundWorker
            {
                WorkerSupportsCancellation = true
            };

            animationWorker.DoWork += (sender, e) =>
            {
                BackgroundWorker worker = sender as BackgroundWorker;

                uint frame = 0;
                double WAVE_FREQUENCY = 0.06;
                byte[] ledData = new byte[totalBytes];
                bool UP = (ReadRegister(0x401) & 0x01) == 1;

                while (!worker.CancellationPending)
                {
                    for (int i = 0; i < totalBytes; i++)
                    {
                        int r = i / segmentsPerGrid;
                        int c = i % segmentsPerGrid;

                        double sinInput = (c + r + frame) * WAVE_FREQUENCY;
                        double sinOutput = Math.Sin(sinInput);
                        byte brightness = (byte)(255 - (((sinOutput + 1.0) / 2.0) * 255));
                        if (brightness < 0) brightness = 0;

                        int led_index = r * segmentsPerGrid + ((segmentsPerGrid - 1) - c);

                        if (led_index < totalBytes)
                        {
                            ledData[led_index] = brightness;
                        }
                    }

                    WriteLEDDataWithHexForm((uint)totalBytes, ledData);
                    if (!UP) WriteCommand(0x04);
                    
                    frame += 1;

                    Thread.Sleep(10);
                }
            };

            animationWorker.RunWorkerCompleted += (sender, e) =>
            {
                Log.WriteLine("Wave effect stopped.", Color.Green, Parent.BackColor);
            };

            animationWorker.RunWorkerAsync();
            Log.WriteLine("Wave effect started.", Color.Green, Parent.BackColor);
        }

        private void RunScrollTextTest()
        {
            if (!CheckLEDEffectPrerequisites()) return;

            if (animationWorker != null && animationWorker.IsBusy)
            {
                Log.WriteLine("Another animation is already running.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            DialogResult result = MessageBox.Show("배경을 켜시겠습니까?\n(아니오: 글자 켜기)", "텍스트 표시 방식 선택", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            bool invertDisplay = (result == DialogResult.Yes);

            string inputText = RegContForm.Prompt.ShowDialog("스크롤할 문자를 입력하세요 (최대 20글자):", "Scroll Text");
            if (string.IsNullOrWhiteSpace(inputText)) return;

            if (inputText.Length > 20)
            {
                MessageBox.Show("최대 20글자까지만 입력할 수 있습니다.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            inputText = inputText.ToUpper();

            List<byte> scrollBuffer = new List<byte>();
            foreach (char ch in inputText)
            {
                if (Font.ContainsKey(ch))
                {
                    scrollBuffer.AddRange(Font[ch]);

                    if (ch != ' ')
                    {
                        scrollBuffer.Add(0x00);
                    }
                }
            }
            if (scrollBuffer.Count == 0) return;

            for (int i = 0; i < 32; i++)
            {
                scrollBuffer.Add(0x00);
            }

            int totalBytes = CalculateChicagoGridSizeBytes();
            if (totalBytes != 320)
            {
                MessageBox.Show("해당 Effect는 G1~10 & SEG1~32를 모두 사용합니다. 칩 설정을 확인하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool UP = (ReadRegister(0x401) & 0x01) == 1;

            animationWorker = new BackgroundWorker();
            animationWorker.WorkerSupportsCancellation = true;

            animationWorker.DoWork += (sender, e) =>
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                byte[] ledData = new byte[totalBytes];
                int frame = 0;

                while (!worker.CancellationPending)
                {
                    for (int grid_col = 0; grid_col < 32; grid_col++)
                    {
                        int buffer_index = (grid_col + frame) % scrollBuffer.Count;
                        byte columnData = scrollBuffer[buffer_index];

                        for (int grid_row = 0; grid_row < 10; grid_row++)
                        {
                            int led_index = grid_row * 32 + (31 - grid_col);

                            if (grid_row > 0 && grid_row < 9)
                            {
                                bool isPixelOn = ((columnData >> (grid_row - 1)) & 0x01) == 1;

                                if (invertDisplay) // 배경 켜기 모드
                                {
                                    ledData[led_index] = isPixelOn ? (byte)0 : (byte)255;
                                }
                                else // 글자 켜기 모드
                                {
                                    ledData[led_index] = isPixelOn ? (byte)255 : (byte)0;
                                }
                            }
                            else
                            {
                                ledData[led_index] = invertDisplay ? (byte)255 : (byte)0;
                            }
                        }
                    }

                    WriteLEDDataWithHexForm((uint)totalBytes, ledData);
                    if (!UP) WriteCommand(0x04);

                    Thread.Sleep(100);
                    frame++;
                }
            };

            animationWorker.RunWorkerCompleted += (sender, e) =>
            {
                Log.WriteLine("Scroll text stopped.", Color.Green, Parent.BackColor);
            };

            animationWorker.RunWorkerAsync();
            Log.WriteLine("Scroll text started.", Color.Green, Parent.BackColor);
        }

        private void RunBreathingEffectTest()
        {
            if (!CheckLEDEffectPrerequisites()) return;

            if (animationWorker != null && animationWorker.IsBusy)
            {
                Log.WriteLine("Another animation is already running.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            int totalBytes = CalculateChicagoGridSizeBytes();
            if (totalBytes == 0)
            {
                MessageBox.Show("계산된 LED 개수가 0입니다. 칩 설정을 확인하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool UP = (ReadRegister(0x401) & 0x01) == 1;

            animationWorker = new BackgroundWorker();
            animationWorker.WorkerSupportsCancellation = true;

            animationWorker.DoWork += (sender, e) =>
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                byte[] ledData = new byte[totalBytes];
                uint frame = 0;

                while (!worker.CancellationPending)
                {
                    double sinInput = frame * 0.05;
                    double sinOutput = Math.Sin(sinInput);
                    byte brightness = (byte)(((sinOutput + 1.0) / 2.0) * 255);

                    for (int i = 0; i < totalBytes; i++)
                    {
                        ledData[i] = brightness;
                    }

                    WriteLEDDataWithHexForm((uint)totalBytes, ledData);
                    if (!UP) WriteCommand(0x04);

                    Thread.Sleep(10);
                    frame++;
                }
            };

            animationWorker.RunWorkerCompleted += (sender, e) =>
            {
                Log.WriteLine("Breathing effect stopped.", Color.Green, Parent.BackColor);
            };

            animationWorker.RunWorkerAsync();
            Log.WriteLine("Breathing effect started.", Color.Green, Parent.BackColor);
        }

        private void RunClockEffectTest()
        {
            if (!CheckLEDEffectPrerequisites()) return;

            if (animationWorker != null && animationWorker.IsBusy)
            {
                Log.WriteLine("Another animation is already running.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            int totalBytes = CalculateChicagoGridSizeBytes();
            if (totalBytes != 320)
            {
                MessageBox.Show("해당 Effect는 G1~10 & SEG 1~32를 모두 사용합니다. 칩 설정을 확인하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool UP = (ReadRegister(0x401) & 0x01) == 1 ? true : false;

            animationWorker = new BackgroundWorker();
            animationWorker.WorkerSupportsCancellation = true;

            animationWorker.DoWork += (sender, e) =>
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                byte[] ledData = new byte[totalBytes];

                while (!worker.CancellationPending)
                {
                    string timeString = DateTime.Now.ToString("HH:mm");

                    bool showColon = (DateTime.Now.Millisecond >= 500);

                    List<byte> textBuffer = new List<byte>();
                    foreach (char ch in timeString)
                    {
                        char charToRender = ch;
                        if (ch == ':' && !showColon)
                        {
                            charToRender = ' ';
                        }

                        if (Font.ContainsKey(charToRender))
                        {
                            textBuffer.AddRange(Font[charToRender]);
                            textBuffer.Add(0x00);
                        }
                    }

                    int textWidth = textBuffer.Count;
                    int startCol = (32 - textWidth) / 2;
                    Array.Clear(ledData, 0, ledData.Length);

                    for (int c = 0; c < textWidth; c++)
                    {
                        int physical_col = startCol + c;
                        if (physical_col < 0 || physical_col >= 32) continue;

                        byte columnData = textBuffer[c];

                        for (int grid_row = 0; grid_row < 10; grid_row++)
                        {
                            int led_index = grid_row * 32 + (31 - physical_col);

                            if (grid_row > 0 && grid_row < 9)
                            {
                                if (((columnData >> (grid_row - 1)) & 0x01) == 1)
                                {
                                    ledData[led_index] = 255;
                                }
                            }
                        }
                    }

                    WriteLEDDataWithHexForm((uint)totalBytes, ledData);
                    if (!UP) WriteCommand(0x04);

                    int currentMillisecond = DateTime.Now.Millisecond;
                    int delay = (currentMillisecond < 500) ? (500 - currentMillisecond) : (1000 - currentMillisecond);
                    Thread.Sleep(delay);
                }
            };

            animationWorker.RunWorkerCompleted += (sender, e) =>
            {
                Log.WriteLine("Clock effect stopped.", Color.Green, Parent.BackColor);
            };

            animationWorker.RunWorkerAsync();
            Log.WriteLine("Clock effect started.", Color.Green, Parent.BackColor);
        }

        private void RunFireworkEffectTest()
        {
            //if (!CheckLEDEffectPrerequisites()) return;

            if (animationWorker != null && animationWorker.IsBusy)
            {
                Log.WriteLine("Another animation is already running.", Color.OrangeRed, Parent.BackColor);
                return;
            }

            var I_REG = Parent.RegMgr.GetRegisterItem("I[5:0]");
            var G_N = Parent.RegMgr.GetRegisterItem("G_N[3:0]");
            var UP = Parent.RegMgr.GetRegisterItem("UP");
            var DON = Parent.RegMgr.GetRegisterItem("DON");
            var LV0_RD_EN = Parent.RegMgr.GetRegisterItem("LV0_RD_EN");
            var SLEEP = Parent.RegMgr.GetRegisterItem("SLEEP");
            var SEG_MODE = Parent.RegMgr.GetRegisterItem("SEG_MODE[1:0]");

            I_REG.Value = 63;
            I_REG.Write();

            G_N.Value = 7;
            G_N.Write();

            UP.Value = 1;
            UP.Write();

            LV0_RD_EN.Value = 1;
            LV0_RD_EN.Write();

            SLEEP.Value = 1;
            SLEEP.Write();

            SEG_MODE.Value = 1;
            SEG_MODE.Write();

            int totalBytes = CalculateChicagoGridSizeBytes();
            if (totalBytes <= 0)
            {
                MessageBox.Show("계산된 LED 개수가 0입니다. 칩 설정을 확인하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool isUP = (ReadRegister(0x401) & 0x01) == 1;

            byte[] dummyData = new byte[totalBytes + 1];
            WriteLEDDataWithHexForm((uint)dummyData.Length, dummyData);
            WriteCommand(0x04);

            DON.Value = 1;
            DON.Write();

            animationWorker = new BackgroundWorker();
            animationWorker.WorkerSupportsCancellation = true;

            animationWorker.DoWork += (sender, e) =>
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                byte[] ledData = new byte[totalBytes];

                uint frame = 0;
                uint cycleCount = 0;
                const byte steps = 4;
                const uint idleFrames = 30;
                int scale = 255 / (steps - 1);

                while (!worker.CancellationPending)
                {
                    Array.Clear(ledData, 0, ledData.Length);

                    uint cycleLen = idleFrames + (16u * steps);
                    uint phase = frame % cycleLen;

                    int d = (int)(cycleCount % 100u);
                    int tens = d / 10;
                    int ones = d % 10;

                    int ledsOn =
                        (cycleCount <= 24u) ? 1 :
                        (cycleCount <= 49u) ? 2 :
                        (cycleCount <= 74u) ? 3 : 4;

                    if (phase < idleFrames)
                    {
                        for (uint g = 0; g < 8; g++)
                        {
                            for (uint s = 0; s < 16; s++)
                            {
                                int idx = (int)(g * 16u + s);

                                if (g == 0)
                                {
                                    ledData[idx] = (s < 8) ? DigitPatterns[tens, (int)s] : DigitPatterns[ones, (int)(s - 8)];
                                    continue;
                                }
                                if (g == 1)
                                {
                                    if (s == 0) { ledData[idx] = 0xFF; continue; }
                                    if (s >= 1 && s <= 4) { ledData[idx] = ((int)(s - 1) < ledsOn) ? (byte)0xFF : (byte)0x00; continue; }
                                    if (s == 14 || s == 15) { ledData[idx] = (cycleCount == 100u) ? (byte)0xFF : (byte)0x00; continue; }
                                }
                                if (g == 4 && s < 11) { ledData[idx] = 0xFF; continue; }
                                if (g == 5 && s < 2) { ledData[idx] = 0xFF; continue; }
                                ledData[idx] = 0x00;
                            }
                        }
                    }
                    else
                    {
                        uint act = phase - idleFrames;
                        uint activeS = act / steps;
                        uint activeVal = act % steps;

                        for (uint g = 0; g < 8; g++)
                        {
                            for (uint s = 0; s < 16; s++)
                            {
                                int idx = (int)(g * 16u + s);

                                if (g == 0)
                                {
                                    ledData[idx] = (s < 8) ? DigitPatterns[tens, (int)s] : DigitPatterns[ones, (int)(s - 8)];
                                    continue;
                                }
                                if (g == 1)
                                {
                                    if (s == 0) { ledData[idx] = 0xFF; continue; }
                                    if (s >= 1 && s <= 4) { ledData[idx] = ((int)(s - 1) < ledsOn) ? (byte)0xFF : (byte)0x00; continue; }
                                    if (s == 14 || s == 15) { ledData[idx] = (cycleCount == 100u) ? (byte)0xFF : (byte)0x00; continue; }
                                }
                                if (g == 4 && s < 11) { ledData[idx] = 0xFF; continue; }
                                if (g == 5 && s < 2) { ledData[idx] = 0xFF; continue; }

                                if (s < activeS) ledData[idx] = 0xFF;
                                else if (s == activeS) ledData[idx] = (byte)(activeVal * (uint)scale);
                                else ledData[idx] = 0x00;
                            }
                        }
                    }
                    if (phase == cycleLen - 1)
                        cycleCount = (cycleCount + 1) % 101u;

                    WriteLEDDataWithHexForm((uint)totalBytes, ledData);
                    if (!isUP) WriteCommand(0x04);

                    frame += 1;
                    Thread.Sleep(10);
                }
            };

            animationWorker.RunWorkerCompleted += (sender, e) =>
            {
                Log.WriteLine("Firework effect stopped.", Color.Green, Parent.BackColor);
            };

            animationWorker.RunWorkerAsync();
            Log.WriteLine("Firework effect started.", Color.Green, Parent.BackColor);
        }
        #endregion Function for LED Test
    }
}
