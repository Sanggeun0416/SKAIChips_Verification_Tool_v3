using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using JLcLib.Chip;
using JLcLib.Comn;
using SKAIChips_Verification;
using System.Threading;

namespace ABOV
{
    public class ABOV_Toscana : ChipControl
    {
        public enum TEST_ITEMS
        {
            TEST_BGR_SWEEP,
            TEST_LDO5P0_SWEEP,
            TEST_LDO1P5_SWEEP,
            TEST_LDO3P3_FSK_SWEEP,
            TEST_LDO3P3_ADC_SWEEP,
            TEST_LDO1P8_SWEEP,
            TEST_OSC_SWEEP,
            TEST_MLDO_SWEEP,
            TEST_ADC_VRECT,
            TEST_ADC_VOUT,
            TEST_ADC_IOUT,
            TEST_ADC_VBAT_SNS,
            TEST_ADC_VPTAT,
            TEST_ADC_VNTC,
            TEST_ADC_VBGR,
            TEST_ADC_VLIM,
            TEST_ADC_VREF,
            TEST_ADC_TX_IOUT,
            TEST_ACTIVE_LOAD,

            NUM_TEST_ITEMS,
        }

        public enum ADC
        {
            VRECT,
            VOUT,
            IOUT,
            VBAT_SNS,
            VPTAT,
            VNTC,
            VBGR,
            VLIM,
            VREF,
            VREF_DLOAD,
            I_DLOAD,
        }

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI PowerSupply1 = null;
        JLcLib.Instrument.SCPI Oscilloscope = null;
        JLcLib.Instrument.SCPI DigitalMultimeter = null;
        JLcLib.Instrument.SCPI ElectronicLoad = null;

        public ABOV_Toscana(RegContForm form) : base(form)
        {
            CalibrationData = new byte[256];

            Serial.ReadSettingFile(form.IniFile, "ABOV_Toscana");
            Serial.DataReceived += Serial_DataReceived;

            /* Init test items combo box */
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e)
        {
            IsSerialReceivedData = true;
        }

        #region Chip register control methods
        private void WriteRegister(uint Address, uint Data)
        {
            List<byte> SendBytes = new List<byte>();

            SendBytes.Add(0x02);    // STX
            SendBytes.Add(0x30);    // BCC
            SendBytes.Add(0x30);    // Length_0(MSB)
            SendBytes.Add(0x30);    // Length_1
            SendBytes.Add(0x31);    // Length_2
            SendBytes.Add(0x39);    // Length_3(LSB)
            SendBytes.Add(0x53);    // Command_0
            SendBytes.Add(0x38);    // Command_1
            SendBytes.Add((byte)((Address >> 28) & 0x0f | 0x30));   // Data_0(MSB)
            SendBytes.Add((byte)((Address >> 24) & 0x0f | 0x30));   // Data_1
            SendBytes.Add((byte)((Address >> 20) & 0x0f | 0x30));   // Data_2
            SendBytes.Add((byte)((Address >> 16) & 0x0f | 0x30));   // Data_3
            SendBytes.Add((byte)((Address >> 12) & 0x0f | 0x30));   // Data_4
            SendBytes.Add((byte)((Address >> 8) & 0x0f | 0x30));    // Data_5
            SendBytes.Add((byte)((Address >> 4) & 0x0f | 0x30));    // Data_6
            SendBytes.Add((byte)((Address >> 0) & 0x0f | 0x30));    // Data_7
            SendBytes.Add((byte)((Data >> 28) & 0x0f | 0x30));      // Data_8
            SendBytes.Add((byte)((Data >> 24) & 0x0f | 0x30));      // Data_9
            SendBytes.Add((byte)((Data >> 20) & 0x0f | 0x30));      // Data_10
            SendBytes.Add((byte)((Data >> 16) & 0x0f | 0x30));      // Data_11
            SendBytes.Add((byte)((Data >> 12) & 0x0f | 0x30));      // Data_12
            SendBytes.Add((byte)((Data >> 8) & 0x0f | 0x30));       // Data_13
            SendBytes.Add((byte)((Data >> 4) & 0x0f | 0x30));       // Data_14
            SendBytes.Add((byte)((Data >> 0) & 0x0f | 0x30));       // Data_15
            SendBytes.Add(0x03);    // ETX
            for (int i = 3; i < 25; i++)
            {
                SendBytes[1] ^= SendBytes[i];
            }
            Serial.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
        }

        private uint ReadRegister(uint Address)
        {
            uint Data = 0xFFFFFFFF;
            byte[] RcvData = new byte[22];

            List<byte> SendBytes = new List<byte>();

            SendBytes.Add(0x02);    // STX
            SendBytes.Add(0x30);    // BCC
            SendBytes.Add(0x30);    // Length_0(MSB)
            SendBytes.Add(0x30);    // Length_1
            SendBytes.Add(0x31);    // Length_2
            SendBytes.Add(0x31);    // Length_3(LSB)
            SendBytes.Add(0x53);    // Command_0
            SendBytes.Add(0x37);    // Command_1
            SendBytes.Add((byte)((Address >> 28) & 0x0f | 0x30));   // Data_0(MSB)
            SendBytes.Add((byte)((Address >> 24) & 0x0f | 0x30));   // Data_1
            SendBytes.Add((byte)((Address >> 20) & 0x0f | 0x30));   // Data_2
            SendBytes.Add((byte)((Address >> 16) & 0x0f | 0x30));   // Data_3
            SendBytes.Add((byte)((Address >> 12) & 0x0f | 0x30));   // Data_4
            SendBytes.Add((byte)((Address >> 8) & 0x0f | 0x30));    // Data_5
            SendBytes.Add((byte)((Address >> 4) & 0x0f | 0x30));    // Data_6
            SendBytes.Add((byte)((Address >> 0) & 0x0f | 0x30));    // Data_7(LSB)
            SendBytes.Add(0x03);    // ETX
            for (int i = 3; i < 17; i++)
            {
                SendBytes[1] ^= SendBytes[i];
            }
            Serial.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            System.Threading.Thread.Sleep(30);
            //Log.WriteLine(Serial.RcvQueue.Count.ToString());

            for (int i = 0; i < 22; i++)
            {
                RcvData[i] = Serial.RcvQueue.Get();
            }


            if (RcvData[0] == 0x02 && RcvData[21] == 0x03)
                Data = (uint)(((RcvData[13] & 0xf) << 28) | ((RcvData[14] & 0xf) << 24) | ((RcvData[15] & 0xf) << 20) | ((RcvData[16] & 0xf) << 16) | ((RcvData[17] & 0xf) << 12) | ((RcvData[18] & 0xf) << 8) | ((RcvData[19] & 0xf) << 4) | ((RcvData[20] & 0xf) << 0));

            return Data;
        }
        #endregion Chip register control methods

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
            Parent.ChipCtrlTextboxes[0].ReadOnly = true;
            Parent.ChipCtrlTextboxes[0].Visible = true;
            Parent.ChipCtrlTextboxes[0].Text = Serial.Config.PortName;
            Parent.ChipCtrlTextboxes[0].Size = new System.Drawing.Size(Parent.ChipCtrlTextboxes[0].Size.Width * 2 + 3, Parent.ChipCtrlTextboxes[0].Size.Height);
            Parent.ChipCtrlButtons[2].Text = "Connect";
            Parent.ChipCtrlButtons[2].Visible = true;
            Parent.ChipCtrlButtons[2].Click += SerialConnect_Click;
            Parent.ChipCtrlButtons[3].Text = "Set UART";
            Parent.ChipCtrlButtons[3].Visible = true;
            Parent.ChipCtrlButtons[3].Click += SerialSetting_Click;

            Parent.ChipCtrlButtons[4].Text = "LDO_CAL";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Run_LDO_Calibration_Click;

            Parent.ChipCtrlButtons[5].Text = "SWIPT";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += Run_SWIPT_Click;
        }

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
            /*
            if (Serial.IsOpen && IsSerialReceivedData)
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
                            LogMessage = "";
                        }
                    }
                }
                IsSerialReceivedData = false;
            }
            */
        }

        public override void SendCommand(string Command)
        {
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
            Serial.WriteSettingFile(Parent.IniFile, "ABOV_Toscana");
        }

        private void Run_LDO_Calibration_Click(object sender, EventArgs e)
        {
            TEST_Register_Sweep_BGR();
            TEST_Register_Sweep_LDO5P0();
            TEST_Register_Sweep_LDO1P5();
            TEST_Register_Sweep_LDO3P3_FSK();
            TEST_Register_Sweep_LDO3P3_ADC();
            TEST_Register_Sweep_LDO1P8();
            //TEST_Register_Sweep_OSC();
            TEST_Register_Sweep_MLDO();
            MessageBox.Show("End Trim");
        }

        private void Run_SWIPT_Click(object sender, EventArgs e)
        {
            Random rand = new Random();
            double value;

            for (int i = 0; i < 10; i++)
            {
                value = (double)(rand.Next(220, 290)) / 1000.0;
                Log.WriteLine("CCP : " + value.ToString("F3") + "mW");
                System.Threading.Thread.Sleep(1000);
            }

        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal, Result = 0;

            TEST_ITEMS TestItem = (TEST_ITEMS)TestItemIndex;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch ((TEST_ITEMS)TestItemIndex)
            {
                // GUI test functions
                case TEST_ITEMS.TEST_BGR_SWEEP:
                    TEST_Register_Sweep_BGR();
                    break;
                case TEST_ITEMS.TEST_LDO5P0_SWEEP:
                    TEST_Register_Sweep_LDO5P0();
                    break;
                case TEST_ITEMS.TEST_LDO1P5_SWEEP:
                    TEST_Register_Sweep_LDO1P5();
                    break;
                case TEST_ITEMS.TEST_LDO3P3_FSK_SWEEP:
                    TEST_Register_Sweep_LDO3P3_FSK();
                    break;
                case TEST_ITEMS.TEST_LDO3P3_ADC_SWEEP:
                    TEST_Register_Sweep_LDO3P3_ADC();
                    break;
                case TEST_ITEMS.TEST_LDO1P8_SWEEP:
                    TEST_Register_Sweep_LDO1P8();
                    break;
                case TEST_ITEMS.TEST_OSC_SWEEP:
                    TEST_Register_Sweep_OSC();
                    break;
                case TEST_ITEMS.TEST_MLDO_SWEEP:
                    TEST_Register_Sweep_MLDO();
                    break;
                case TEST_ITEMS.TEST_ADC_VRECT:
                    TEST_ADC_VRECT();
                    break;
                case TEST_ITEMS.TEST_ADC_VOUT:
                    TEST_ADC_VOUT();
                    break;
                case TEST_ITEMS.TEST_ADC_IOUT:
                    TEST_ADC_IOUT();
                    break;
                case TEST_ITEMS.TEST_ADC_VBAT_SNS:
                    TEST_ADC_VBAT_SNS();
                    break;
                case TEST_ITEMS.TEST_ADC_VPTAT:
                    //TEST_ADC_VPTAT();
                    break;
                case TEST_ITEMS.TEST_ADC_VNTC:
                    TEST_ADC_VNTC();
                    break;
                case TEST_ITEMS.TEST_ADC_VBGR:
                    TEST_ADC_VBGR();
                    break;
                case TEST_ITEMS.TEST_ADC_VLIM:
                    TEST_ADC_VLIM();
                    break;
                case TEST_ITEMS.TEST_ADC_VREF:
                    TEST_ADC_VREF();
                    break;
                case TEST_ITEMS.TEST_ADC_TX_IOUT:
                    //TEST_ADC_TX_IOUT();
                    break;
                case TEST_ITEMS.TEST_ACTIVE_LOAD:
                    TEST_ACTIVE_LOAD();
                    break;

                // FW test functions
                default:
                    //Result = RunTest(TestItem, iVal);
                    break;
            }
            Log.WriteLine(TestItem.ToString() + ":" + iVal.ToString() + ":" + Result.ToString());
        }

        #region Chip test methods
        private void AUTO_Test_Instrument()
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
                    case JLcLib.Instrument.InstrumentTypes.PowerSupply1:
                        if (PowerSupply1 == null)
                            PowerSupply1 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (PowerSupply1.IsOpen == false)
                            PowerSupply1.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.OscilloScope0:
                        if (Oscilloscope == null)
                            Oscilloscope = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (Oscilloscope.IsOpen == false)
                            Oscilloscope.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter0:
                        if (DigitalMultimeter == null)
                            DigitalMultimeter = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter.IsOpen == false)
                            DigitalMultimeter.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.ElectronicLoad:
                        if (ElectronicLoad == null)
                            ElectronicLoad = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (ElectronicLoad.IsOpen == false)
                            ElectronicLoad.Open();
                        break;
                }
            }
        }

        private void TEST_Register_Sweep_BGR()
        {
            string sheet_name;
            double bgr_v, diff = 99999.9, target = 1.22;
            uint cal_val = 0;

            RegisterItem PU_TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR");            // 0x400005C8[20:16]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");          // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");    // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");          // 0x400005F4[0]

            MessageBox.Show("DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "BGR_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_BGR");
            Parent.xlMgr.Cell.Write(3, 2, "BGR(V)");

            PU_TRIM_BGR.Read();
            TEST_ANA_SEL.Read(); //0x400005F4 Read

            TEST_ANA_SEL.Value = 0;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1; //0 : PAD is Output Mode 1 : PAD is Input Mode
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            for (uint i = 0; i < 32; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_BGR.Value = i;
                PU_TRIM_BGR.Write();
                System.Threading.Thread.Sleep(100);

                bgr_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, bgr_v.ToString());
                if (Math.Abs(target - bgr_v) < diff)
                {
                    diff = Math.Abs(target - bgr_v);
                    cal_val = i;
                }

                if (bgr_v > 1.29) break;
            }

            PU_TRIM_BGR.Value = cal_val;
            PU_TRIM_BGR.Write();
            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();
        }

        private void TEST_Register_Sweep_LDO5P0()
        {
            string sheet_name;
            double LDO5P0_v, diff = 9999.9, target = 5.0;
            uint cal_val = 0;

            RegisterItem PU_TRIM_LDO5P0 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO5P0");     // 0x400005CC[3:0]

            MessageBox.Show("DigitalMultimeter :LDO5P0(J7_8)");

            AUTO_Test_Instrument();

            sheet_name = "LDO5P0_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_LDO5P0");
            Parent.xlMgr.Cell.Write(3, 2, "LDO5P0(V)");

            PU_TRIM_LDO5P0.Read();

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_LDO5P0.Value = i;
                PU_TRIM_LDO5P0.Write();
                System.Threading.Thread.Sleep(100);

                LDO5P0_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, LDO5P0_v.ToString());

                if (Math.Abs(target - LDO5P0_v) < diff)
                {
                    diff = Math.Abs(target - LDO5P0_v);
                    cal_val = i;
                }
            }
            PU_TRIM_LDO5P0.Value = cal_val;
            PU_TRIM_LDO5P0.Write();
        }

        private void TEST_Register_Sweep_LDO1P5()
        {
            string sheet_name;
            double LDO1P5_v, diff = 9999.9, target = 1.5;
            uint cal_val = 0;

            RegisterItem PU_TRIM_LDO1P5 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO1P5");     // 0x400005D0[25:22]

            MessageBox.Show("DigitalMultimeter :LDO1P5(J7_4)");

            AUTO_Test_Instrument();

            sheet_name = "LDO1P5_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_LDO1P5");
            Parent.xlMgr.Cell.Write(3, 2, "LDO1P5(V)");

            PU_TRIM_LDO1P5.Read();

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_LDO1P5.Value = i;
                PU_TRIM_LDO1P5.Write();
                System.Threading.Thread.Sleep(100);

                LDO1P5_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, LDO1P5_v.ToString());

                if (Math.Abs(target - LDO1P5_v) < diff)
                {
                    diff = Math.Abs(target - LDO1P5_v);
                    cal_val = i;
                }
            }
            PU_TRIM_LDO1P5.Value = cal_val;
            PU_TRIM_LDO1P5.Write();
        }

        private void TEST_Register_Sweep_LDO3P3_FSK()
        {
            string sheet_name;
            double FSK3P3_v, diff = 9999.9, target = 3.3;
            uint cal_val = 0;

            RegisterItem PU_TRIM_FSK_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_FSK_LDO3P3");     // 0x400005CC[28:25]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");     // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");     // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");     // 0x400005F4[0]

            MessageBox.Show("DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "FSK3P3_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_FSK_LDO3P3");
            Parent.xlMgr.Cell.Write(3, 2, "FSK_LDO3P3(V)");

            PU_TRIM_FSK_LDO3P3.Read();
            TEST_ANA_SEL.Read(); //0x400005F4 Read

            TEST_ANA_SEL.Value = 5;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1; //0 : PAD is Output Mode 1 : PAD is Input Mode
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_FSK_LDO3P3.Value = i;
                PU_TRIM_FSK_LDO3P3.Write();
                System.Threading.Thread.Sleep(100);

                FSK3P3_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, FSK3P3_v.ToString());
                if (Math.Abs(target - FSK3P3_v) < diff)
                {
                    diff = Math.Abs(target - FSK3P3_v);
                    cal_val = i;
                }
            }
            PU_TRIM_FSK_LDO3P3.Value = cal_val;
            PU_TRIM_FSK_LDO3P3.Write();
            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();
        }

        private void TEST_Register_Sweep_LDO3P3_ADC()
        {
            string sheet_name;
            double ADC3P3_v, diff = 9999.9, target = 3.3;
            uint cal_val = 0;

            RegisterItem PU_TRIM_ADC_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_ADC_LDO3P3");     // 0x400005CC[19:16]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");     // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");     // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");     // 0x400005F4[0]

            MessageBox.Show("DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "ADC3P3_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_ADC_LDO3P3");
            Parent.xlMgr.Cell.Write(3, 2, "ADC_LDO3P3(V)");

            PU_TRIM_ADC_LDO3P3.Read();
            TEST_ANA_SEL.Read(); //0x400005F4 Read

            TEST_ANA_SEL.Value = 6;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1; //0 : PAD is Output Mode 1 : PAD is Input Mode
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_ADC_LDO3P3.Value = i;
                PU_TRIM_ADC_LDO3P3.Write();
                System.Threading.Thread.Sleep(100);

                ADC3P3_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, ADC3P3_v.ToString());
                if (Math.Abs(target - ADC3P3_v) < diff)
                {
                    diff = Math.Abs(target - ADC3P3_v);
                    cal_val = i;
                }
            }
            PU_TRIM_ADC_LDO3P3.Value = cal_val;
            PU_TRIM_ADC_LDO3P3.Write();
            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();
        }

        private void TEST_Register_Sweep_LDO1P8()
        {
            string sheet_name;
            double LDO1P8_v, diff = 9999.9, target = 1.8;
            uint cal_val = 0;

            RegisterItem PU_TRIM_LDO18 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO18");     // 0x400005D0[12:9]

            MessageBox.Show("DigitalMultimeter :LDO1P8(J7_6)");

            AUTO_Test_Instrument();

            sheet_name = "LDO1P8_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_LDO18");
            Parent.xlMgr.Cell.Write(3, 2, "LDO1P8(V)");

            PU_TRIM_LDO18.Read();

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_LDO18.Value = i;
                PU_TRIM_LDO18.Write();
                System.Threading.Thread.Sleep(100);

                LDO1P8_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, LDO1P8_v.ToString());
                if (Math.Abs(target - LDO1P8_v) < diff)
                {
                    diff = Math.Abs(target - LDO1P8_v);
                    cal_val = i;
                }
            }
            PU_TRIM_LDO18.Value = cal_val;
            PU_TRIM_LDO18.Write();
        }

        private void TEST_Register_Sweep_OSC()
        {
            string sheet_name;
            double OSC_MHz, pre_osc, diff = 9999.9, target = 16.0;
            uint cal_val = 0;

            RegisterItem PU_TRIM_FOSC_16M = Parent.RegMgr.GetRegisterItem("PU_TRIM_FOSC_16M");  // 0x400005C8[31:24]
            RegisterItem TEST_DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL");          // 0x400005F4[11:8]
            RegisterItem TEST_DIG_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN");    // 0x400005F4[2]

            MessageBox.Show("OscilloScope :TEST_DIG(J7_10)");

            AUTO_Test_Instrument();

            sheet_name = "OSC_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_FOSC_16M");
            Parent.xlMgr.Cell.Write(3, 2, "OSC(Hz)");

            PU_TRIM_FOSC_16M.Read();
            TEST_DIG_SEL.Read();
            TEST_DIG_MUX_EN.Read();

            TEST_DIG_SEL.Value = 0;
            TEST_DIG_SEL.Write();

            TEST_DIG_MUX_EN.Value = 1;
            TEST_DIG_SEL.Write();

            pre_osc = 0;
            for (uint i = 0; i < 50; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_FOSC_16M.Value = i;
                PU_TRIM_FOSC_16M.Write();
                System.Threading.Thread.Sleep(100);

                OSC_MHz = double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN1")) / 1000000.0;
                //if ((OSC_MHz - pre_osc) < 0.01) break;
                if (OSC_MHz > 16.7) break;
                pre_osc = OSC_MHz;

                Parent.xlMgr.Cell.Write(3, (int)i + 3, OSC_MHz.ToString("F2"));
                if (Math.Abs(target - OSC_MHz) < diff)
                {
                    diff = Math.Abs(target - OSC_MHz);
                    cal_val = i;
                }
            }
            PU_TRIM_FOSC_16M.Value = cal_val;
            PU_TRIM_FOSC_16M.Write();
            //Register Reset
            TEST_DIG_SEL.Value = 0;
            TEST_DIG_SEL.Write();

            TEST_DIG_MUX_EN.Value = 0;
            TEST_DIG_SEL.Write();
        }

        private void TEST_Register_Sweep_MLDO()
        {
            string sheet_name;
            double MLDO_v, pre_MLDO_v, diff = 9999.9, target = 9.0;
            uint cal_val = 0;

            RegisterItem ML_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET");    // 0x400005D4[31:24]
            RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");              // 0x400005D4[0]

            MessageBox.Show("Vrect : 21V\nDigitalMultimeter :MLDO");

            AUTO_Test_Instrument();

            sheet_name = "MLDO_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "ML_VOUT_SET");
            Parent.xlMgr.Cell.Write(3, 2, "MLDO(V)");

            ML_VOUT_SET.Read();
            ML_VOUT_SET.Value = 40;
            ML_VOUT_SET.Write();
            ML_PEN.Value = 1;
            ML_PEN.Write();
            System.Threading.Thread.Sleep(1000);
#if false
            ML_VOUT_SET.Value = 0;
            ML_VOUT_SET.Write();
            System.Threading.Thread.Sleep(5000);
#endif
            pre_MLDO_v = 0;
            for (uint i = 40; i < 256; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                ML_VOUT_SET.Value = i;
                ML_VOUT_SET.Write();
                System.Threading.Thread.Sleep(100);

                MLDO_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                Parent.xlMgr.Cell.Write(3, (int)i + 3, MLDO_v.ToString());

                if ((MLDO_v - pre_MLDO_v) < 0.05)
                {
                    break;
                }
                pre_MLDO_v = MLDO_v;

                if (Math.Abs(target - MLDO_v) < diff)
                {
                    diff = Math.Abs(target - MLDO_v);
                    cal_val = i;
                }
            }
            ML_VOUT_SET.Value = cal_val;
            ML_VOUT_SET.Write();
        }

        void ADC_Enable(ADC mode)
        {
            RegisterItem ADC_EN = Parent.RegMgr.GetRegisterItem("ADC_EN");          // 0x4000051C[0]
            RegisterItem ADC_PEN = Parent.RegMgr.GetRegisterItem("ADC_PEN");        // 0x4000051C[1]
            RegisterItem ADC_RSTN = Parent.RegMgr.GetRegisterItem("ADC_RSTN");      // 0x4000051C[2]
            RegisterItem ADC_A = Parent.RegMgr.GetRegisterItem("ADC_A");            // 0x4000051C[15:12]

            ADC_EN.Read();

            ADC_EN.Value = 0;
            ADC_PEN.Value = 1;
            ADC_RSTN.Value = 1;
            ADC_A.Value = (uint)mode;
            ADC_EN.Write();
        }

        uint READ_ADC_DATA()
        {
            RegisterItem ADC_DATA = Parent.RegMgr.GetRegisterItem("ADC_DATA");      // 0x40000520[11:0]

            return ADC_DATA.Read();
        }

        void ADC_Disable()
        {
            RegisterItem ADC_EN = Parent.RegMgr.GetRegisterItem("ADC_EN");          // 0x4000051C[0]
            RegisterItem ADC_PEN = Parent.RegMgr.GetRegisterItem("ADC_PEN");        // 0x4000051C[1]
            RegisterItem ADC_RSTN = Parent.RegMgr.GetRegisterItem("ADC_RSTN");      // 0x4000051C[2]

            ADC_EN.Read();

            ADC_EN.Value = 1;
            ADC_PEN.Value = 0;
            ADC_RSTN.Value = 0;
            ADC_EN.Write();
        }

        private void TEST_ADC_VRECT()
        {
            string sheet_name;
            uint adc_code;
            double power, volt_v;

            MessageBox.Show("Power supply(0~25V) : Vrect\nDigitalMultimeter :Vrect");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VRECT_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "Supply(V)");
            Parent.xlMgr.Cell.Write(3, 2, "Vrect(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ADC_Enable(ADC.VRECT);

            PowerSupply0.Write("VOLT 25.0");
            System.Threading.Thread.Sleep(100);

            for (int i = 0; i <= 50; i++)
            {
                power = 25 - (i * 0.5);
                PowerSupply0.Write("VOLT " + power.ToString());
                System.Threading.Thread.Sleep(100);
                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, (int)(power * 2 + 3), power.ToString("F1"));
                Parent.xlMgr.Cell.Write(3, (int)(power * 2 + 3), volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, (int)(power * 2 + 3), adc_code.ToString());
            }

            ADC_Disable();
        }

        private void TEST_ADC_VOUT()
        {
            RegisterItem ML_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET");    // 0x400005D4[31:24]
            RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");              // 0x400005D4[31:24]

            string sheet_name;
            uint adc_code;
            double volt_v;

            MessageBox.Show("Power supply : Vrect 20V\nDigitalMultimeter :Vout");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VOUT_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "ML_VOUT_SET");
            Parent.xlMgr.Cell.Write(3, 2, "VOUT(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ML_PEN.Read();

            ML_PEN.Value = 1;
            ML_PEN.Write();
            System.Threading.Thread.Sleep(500);

            ADC_Enable(ADC.VOUT);

            for (uint i = 10; i < 170; i++)
            {
                ML_VOUT_SET.Value = i;
                ML_VOUT_SET.Write();
                System.Threading.Thread.Sleep(100);

                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));

                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, (int)(i + 3), i.ToString("F1"));
                Parent.xlMgr.Cell.Write(3, (int)(i + 3), volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, (int)(i + 3), adc_code.ToString());
            }

            ADC_Disable();
        }

        private void TEST_ADC_IOUT()
        {
            string sheet_name;
            uint adc_code;
            double Vout, offset_v;
            int y_pos = 0;

            RegisterItem FB = Parent.RegMgr.GetRegisterItem("ML_FB_TRIM");                  // 0x400005D4[7:4]
            RegisterItem VREF = Parent.RegMgr.GetRegisterItem("ML_VREF_TRIM");              // 0x400005D4[21:18]
            RegisterItem ISEN_RVAR = Parent.RegMgr.GetRegisterItem("ML_ISEN_RVAR");         // 0x400005D4[14:13]
            RegisterItem VILIM = Parent.RegMgr.GetRegisterItem("ML_ VILIM_TRIM");           // 0x400005D8[3:0]
            RegisterItem ISEN_VREF = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM");    // 0x400005D8[11:8]

            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");          // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");    // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");          // 0x400005F4[0]

            RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");                      // 0x400005D4[0]

            MessageBox.Show("Power supply(2A) : Vrect(Sweep)\nDigitalMultimeter : VOUT\nLoader : VOUT\n\nSet Vout!! 5V / 9V / 12V / 20V");

            AUTO_Test_Instrument();

            sheet_name = "ADC_Iout_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 3, "ML_FB_TRIM");
            Parent.xlMgr.Cell.Write(3, 3, "ML_VREF_TRIM");
            Parent.xlMgr.Cell.Write(4, 3, "ML_ISEN_RVAR");
            Parent.xlMgr.Cell.Write(5, 3, "ML_VILIM_TRIM");
            Parent.xlMgr.Cell.Write(6, 3, "ML_ISEN_VREF_TRIM");
            Parent.xlMgr.Cell.Write(7, 3, "Iload(mA)");
            Parent.xlMgr.Cell.Write(27, 3, "Iload(mA)");
            Parent.xlMgr.Cell.Write(8, 2, "Iout(code)");
            Parent.xlMgr.Cell.Write(28, 2, "Vout(mV)");
            //Parent.xlMgr.Cell.Write(28, 2, "Iout(mV)");

#if false // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");

            FB.Read();
            VILIM.Read();
            TEST_ANA_SEL.Read();
            ML_PEN.Read();

            TEST_ANA_SEL.Value = 3;
            TEST_ANA_MUX_EN.Value = 1;
            TEST_ANA_PEN.Value = 1;
            TEST_ANA_SEL.Write();

            ADC_Enable(ADC.IOUT);
            for (uint fb = 0; fb < 1; fb++) // 0~15 default 0
            {
                FB.Value = fb;
                FB.Write();
                for (uint vref = 0; vref < 1; vref++) // 0~15 default 0
                {
                    VREF.Value = vref;
                    VREF.Write();
                    for (uint rvar = 0; rvar < 4; rvar++) // 0~3 default 0
                    {
                        ISEN_RVAR.Value = rvar;
                        ISEN_RVAR.Write();
                        for (uint vilim = 0; vilim < 16; vilim++) // 0~15 default 8
                        {
                            VILIM.Value = vilim;
                            VILIM.Write();
                            for (uint isen_vref = 0; isen_vref < 16; isen_vref++) // 0~15 default 0
                            {
                                ISEN_VREF.Value = isen_vref;
                                ISEN_VREF.Write();
                                for (int Iload = 0; Iload <= 2000; Iload += 500) // 0~2000(mA) Electronic Load
                                {
                                    ElectronicLoad.Write("CURR " + (Iload / 1000.0).ToString());
                                    Parent.xlMgr.Cell.Write(2, 4 + (int)(Iload / 50) + y_pos, fb.ToString());
                                    Parent.xlMgr.Cell.Write(3, 4 + (int)(Iload / 50) + y_pos, vref.ToString());
                                    Parent.xlMgr.Cell.Write(4, 4 + (int)(Iload / 50) + y_pos, rvar.ToString());
                                    Parent.xlMgr.Cell.Write(5, 4 + (int)(Iload / 50) + y_pos, vilim.ToString());
                                    Parent.xlMgr.Cell.Write(6, 4 + (int)(Iload / 50) + y_pos, isen_vref.ToString());
                                    Parent.xlMgr.Cell.Write(7, 4 + (int)(Iload / 50) + y_pos, Iload.ToString());
                                    Parent.xlMgr.Cell.Write(27, 4 + (int)(Iload / 50) + y_pos, Iload.ToString());

                                    for (int Vr = 49; Vr <= 55; Vr += 1) // 4.9~5.5(V) Power Supply
                                    {
                                        offset_v = 150; // 0(5V), 5(5.5V), 40(9V), 70(12V), 150(20V)
                                        PowerSupply0.Write("VOLT " + ((Vr + offset_v) / 10.0).ToString());
                                        JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
                                        //if (Vr == 49) JLcLib.Delay.Sleep(1000);
                                        if (Iload == 0)
                                        {
                                            Parent.xlMgr.Cell.Write(8 + (int)(Vr - 49), 3, (((Vr + offset_v) / 10.0).ToString("F2") + "V"));
                                            Parent.xlMgr.Cell.Write(28 + (int)(Vr - 49), 3, (((Vr + offset_v) / 10.0).ToString("F2") + "V"));
                                        }
                                        Vout = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?")) * 1000;
                                        if (Vout < 4.5)
                                        {
                                            break;
                                            ML_PEN.Value = 0;
                                            ML_PEN.Write();
                                            JLcLib.Delay.Sleep(500);
                                            ML_PEN.Value = 1;
                                            ML_PEN.Write();
                                            JLcLib.Delay.Sleep(1000);
                                        }

                                        adc_code = READ_ADC_DATA();

                                        Parent.xlMgr.Cell.Write(8 + (int)(Vr - 49), 4 + (int)(Iload / 50) + y_pos, adc_code.ToString());
                                        Parent.xlMgr.Cell.Write(28 + (int)(Vr - 49), 4 + (int)(Iload / 50) + y_pos, Vout.ToString());
                                    }
;
                                }
                                y_pos += 41;
                            }
                        }
                    }
                }
            }

            ADC_Disable();
            ElectronicLoad.Write("CURR 0");
        }

        private void TEST_ADC_VBAT_SNS()
        {
            string sheet_name;
            uint adc_code;
            double power, volt_v;

            MessageBox.Show("Power supply : Vrect 6.5V, VBAT_SNS(Sweep)(J6_1)\nDigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VBAT_SNS_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "Supply(V)");
            Parent.xlMgr.Cell.Write(3, 2, "VBAT_SNS(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ADC_Enable(ADC.VBAT_SNS);

            for (int i = 0; i <= 55; i++)
            {
                power = i * 0.1;
                PowerSupply0.Write("VOLT " + power.ToString());
                System.Threading.Thread.Sleep(100);
                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, i + 3, power.ToString("F1"));
                Parent.xlMgr.Cell.Write(3, i + 3, volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, i + 3, adc_code.ToString());
            }

            ADC_Disable();
            PowerSupply0.Write("VOLT 0.0");
        }

        private void TEST_ADC_VPTAT() // TBD
        {
            string sheet_name;
            uint adc_code;
            double power, volt_v;

            MessageBox.Show("Power supply : Vrect 6.5V");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VPTAT_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "Supply(V)");
            Parent.xlMgr.Cell.Write(3, 2, "VBAT_SNS(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ADC_Enable(ADC.VBAT_SNS);

            for (int i = 0; i <= 55; i++)
            {
                power = i * 0.1;
                PowerSupply0.Write("VOLT " + power.ToString());
                System.Threading.Thread.Sleep(100);
                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, i + 3, power.ToString("F1"));
                Parent.xlMgr.Cell.Write(3, i + 3, volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, i + 3, adc_code.ToString());
            }

            ADC_Disable();
            PowerSupply0.Write("VOLT 0.0");
        }

        private void TEST_ADC_VNTC()
        {
            string sheet_name;
            uint adc_code;
            double power, volt_v;

            MessageBox.Show("Power supply : Vrect 6.5V, VNTC(Sweep)(J8_2)\nDigitalMultimeter :VNTC(J8_2)");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VNTC_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "Supply(V)");
            Parent.xlMgr.Cell.Write(3, 2, "VNTC(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ADC_Enable(ADC.VNTC);

            for (int i = 0; i <= 33; i++)
            {
                power = i * 0.1;
                PowerSupply0.Write("VOLT " + power.ToString());
                System.Threading.Thread.Sleep(100);
                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, i + 3, power.ToString("F1"));
                Parent.xlMgr.Cell.Write(3, i + 3, volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, i + 3, adc_code.ToString());
            }

            ADC_Disable();
            PowerSupply0.Write("VOLT 0.0");
        }

        private void TEST_ADC_VBGR()
        {
            string sheet_name;
            double volt_v;
            uint adc_code;

            RegisterItem PU_TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR");            // 0x400005C8[20:16]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");          // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");    // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");          // 0x400005F4[0]

            MessageBox.Show("Power supply : Vrect 6.5V, DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VBGR_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "PU_TRIM_BGR");
            Parent.xlMgr.Cell.Write(3, 2, "BGR(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            PU_TRIM_BGR.Read();
            TEST_ANA_SEL.Read(); //0x400005F4 Read

            TEST_ANA_SEL.Value = 0;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1; //0 : PAD is Output Mode 1 : PAD is Input Mode
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            ADC_Enable(ADC.VBGR);

            for (uint i = 0; i < 32; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                PU_TRIM_BGR.Value = i;
                PU_TRIM_BGR.Write();
                System.Threading.Thread.Sleep(100);

                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, (int)(i + 3), i.ToString());
                Parent.xlMgr.Cell.Write(3, (int)(i + 3), volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, (int)(i + 3), adc_code.ToString());
            }

            ADC_Disable();

            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();

            PU_TRIM_BGR.Value = 0;
            PU_TRIM_BGR.Write();
        }

        private void TEST_ADC_VLIM()
        {
            string sheet_name;
            double volt_v;
            uint adc_code;

            RegisterItem ML_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET");            // 0x400005D4[31:24]
            RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");                      // 0x400005D4[0]
            RegisterItem ML_ILIM_SET = Parent.RegMgr.GetRegisterItem("ML_ILIM_SET");            // 0x400005D4[12:8]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");          // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");    // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");          // 0x400005F4[0]

            MessageBox.Show("Power supply : Vrect 6.5V, DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VLIM_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "ML_ILIM_SET");
            Parent.xlMgr.Cell.Write(3, 2, "ML_OCL_VREF(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ML_VOUT_SET.Read();
            TEST_ANA_SEL.Read();

            ML_VOUT_SET.Value = 10;
            ML_VOUT_SET.Write();
            ML_PEN.Value = 1;
            ML_PEN.Write();

            TEST_ANA_SEL.Value = 11;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1;
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            ADC_Enable(ADC.VLIM);

            for (uint i = 0; i < 32; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                ML_ILIM_SET.Value = i;
                ML_ILIM_SET.Write();
                System.Threading.Thread.Sleep(100);

                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, (int)(i + 3), i.ToString());
                Parent.xlMgr.Cell.Write(3, (int)(i + 3), volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, (int)(i + 3), adc_code.ToString());
            }

            ADC_Disable();

            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_SEL.Value = 0;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();
        }

        private void TEST_ADC_VREF()
        {
            string sheet_name;
            double volt_v;
            uint adc_code;

            RegisterItem ML_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET");            // 0x400005D4[31:24]
            RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");                      // 0x400005D4[0]
            RegisterItem ML_VREF_TRIM = Parent.RegMgr.GetRegisterItem("ML_VREF_TRIM");          // 0x400005D4[21:18]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");          // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");    // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");          // 0x400005F4[0]

            MessageBox.Show("Power supply : Vrect 6.5V, DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VREF_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "ML_VREF_TRIM");
            Parent.xlMgr.Cell.Write(3, 2, "ML_VREF(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ML_VOUT_SET.Read();
            TEST_ANA_SEL.Read();

            ML_VOUT_SET.Value = 10;
            ML_VOUT_SET.Write();
            ML_PEN.Value = 1;
            ML_PEN.Write();

            TEST_ANA_SEL.Value = 10;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1;
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            ADC_Enable(ADC.VREF);

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                ML_VREF_TRIM.Value = i;
                ML_VREF_TRIM.Write();
                System.Threading.Thread.Sleep(100);

                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, (int)(i + 3), i.ToString());
                Parent.xlMgr.Cell.Write(3, (int)(i + 3), volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, (int)(i + 3), adc_code.ToString());
            }

            ADC_Disable();

            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_SEL.Value = 0;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();
        }

        private void TEST_ADC_TX_IOUT() // TBD
        {
            string sheet_name;
            double volt_v;
            uint adc_code;

            RegisterItem ML_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET");            // 0x400005D4[31:24]
            RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");                      // 0x400005D4[0]
            RegisterItem ML_VREF_TRIM = Parent.RegMgr.GetRegisterItem("ML_VREF_TRIM");          // 0x400005D4[21:18]
            RegisterItem TEST_ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL");          // 0x400005F4[7:4]
            RegisterItem TEST_ANA_MUX_EN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");    // 0x400005F4[1]
            RegisterItem TEST_ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");          // 0x400005F4[0]

            MessageBox.Show("Power supply : Vrect 6.5V, DigitalMultimeter :VBAT_SNS(J7_1)");

            AUTO_Test_Instrument();

            sheet_name = "ADC_VREF_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(2, 2, "ML_VREF_TRIM");
            Parent.xlMgr.Cell.Write(3, 2, "ML_VREF(V)");
            Parent.xlMgr.Cell.Write(4, 2, "ADC_DATA(Code)");

            ML_VOUT_SET.Read();
            TEST_ANA_SEL.Read();

            ML_VOUT_SET.Value = 10;
            ML_VOUT_SET.Write();
            ML_PEN.Value = 1;
            ML_PEN.Write();

            TEST_ANA_SEL.Value = 10;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 1;
            TEST_ANA_MUX_EN.Write();

            TEST_ANA_PEN.Value = 1;
            TEST_ANA_PEN.Write();

            ADC_Enable(ADC.VREF);

            for (uint i = 0; i < 16; i++)
            {
                Parent.xlMgr.Cell.Write(2, (int)i + 3, i.ToString());

                ML_VREF_TRIM.Value = i;
                ML_VREF_TRIM.Write();
                System.Threading.Thread.Sleep(100);

                volt_v = double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:VOLT:DC?"));
                adc_code = READ_ADC_DATA();
                Parent.xlMgr.Cell.Write(2, (int)(i + 3), i.ToString());
                Parent.xlMgr.Cell.Write(3, (int)(i + 3), volt_v.ToString("F3"));
                Parent.xlMgr.Cell.Write(4, (int)(i + 3), adc_code.ToString());
            }

            ADC_Disable();

            //Register Reset
            TEST_ANA_PEN.Value = 0;
            TEST_ANA_PEN.Write();

            TEST_ANA_SEL.Value = 0;
            TEST_ANA_SEL.Write();

            TEST_ANA_MUX_EN.Value = 0;
            TEST_ANA_MUX_EN.Write();
        }

        private void TEST_ACTIVE_LOAD()
        {
            string sheet_name;
            double Load_i;

            RegisterItem AL_ILOAD_SET = Parent.RegMgr.GetRegisterItem("AL_ILOAD_SET");  // 0x400005EC[9:5]
            RegisterItem AL_BIAS_TRIM = Parent.RegMgr.GetRegisterItem("AL_BIAS_TRIM");  // 0x400005EC[4:1]
            RegisterItem AL_PEN = Parent.RegMgr.GetRegisterItem("AL_PEN");              // 0x400005EC[0:0]

            MessageBox.Show("Power supply : Vrect 6.5V, DigitalMultimeter :VRECT(Measure Current)");

            AUTO_Test_Instrument();

            sheet_name = "ActiveLoad_" + DateTime.Now.ToString("MMddHHmmss");
            Parent.xlMgr.Sheet.Add(sheet_name);

            Parent.xlMgr.Cell.Write(3, 2, "AL_ILOAD_SET");
            Parent.xlMgr.Cell.Write(2, 3, "AL_BIAS_TRIM");

            AL_PEN.Read();
            AL_PEN.Value = 1;
            AL_PEN.Write();

            for (uint i = 0; i < 32; i++)
            {
                Parent.xlMgr.Cell.Write((int)i + 3, 3, i.ToString());
                AL_ILOAD_SET.Value = i;
                AL_ILOAD_SET.Write();
                System.Threading.Thread.Sleep(100);

                for (uint j = 0; j < 16; j++)
                {
                    Parent.xlMgr.Cell.Write(2, (int)j + 4, j.ToString());
                    AL_BIAS_TRIM.Value = j;
                    AL_BIAS_TRIM.Write();
                    Load_i = (double.Parse(DigitalMultimeter.WriteAndReadString("MEAS:CURR:DC?"))) * 1000;
                    Parent.xlMgr.Cell.Write((int)i + 3, (int)j + 4, Load_i.ToString());
                }

            }
            AL_ILOAD_SET.Value = 0;
            AL_ILOAD_SET.Write();

            AL_BIAS_TRIM.Value = 0;
            AL_BIAS_TRIM.Write();

            AL_PEN.Value = 0;
            AL_PEN.Write();
        }
        #endregion
    }

    public class Aladdin : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS_MANUAL
        {
            TEST,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            ReadTest,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
        }

        private JLcLib.Custom.I2C I2C { get; set; }

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        public int SlaveAddress { get; private set; } = 0x3A;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter1 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter2 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter3 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter4 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter5 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI SpectrumAnalyzer = null;
        JLcLib.Instrument.SCPI TempChamber = null;

        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.MANUAL;

        #endregion Variable and declaration

        public Aladdin(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "Aladdin");
            Serial.DataReceived += Serial_DataReceived;
            CalibrationData = new byte[256];

            /* Init test items combo box */
            for (int i = 0; i < (int)TEST_ITEMS_MANUAL.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_MANUAL)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e)
        {
            IsSerialReceivedData = true;
        }

        private byte CalculateParity(byte[] data, int length)
        {
            byte parity = 0x00;
            for (int i = 0; i < 8; i++)
            {
                int ones = 0;
                for (int j = 0; j < length; j++)
                {
                    if (((data[j] >> i) & 1) == 1)
                        ones++;
                }
                if (ones % 2 != 0)
                    parity |= (byte)(1 << i);
            }
            return parity;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            if (!I2C.IsOpen)
                return;

            byte[] Bytes;
            int dataLen = 0;

            if (Address >= 0x10 && Address <= 0x3F)
            {
                Bytes = new byte[2];
                Bytes[0] = (byte)(((Address & 0xF0) | (Data & 0x0F)) & 0xFF);
                Bytes[1] = Bytes[0];
                dataLen = 2;
            }
            else if (Address >= 0x50 && Address <= 0x5F)
            {
                Bytes = new byte[4];
                Bytes[0] = (byte)(Address & 0xFF);
                Bytes[1] = (byte)((Data >> 8) & 0xFF);
                Bytes[2] = (byte)(Data & 0xFF);
                Bytes[3] = CalculateParity(Bytes, 3);
                dataLen = 4;
            }
            else if (Address >= 0x60)
            {
                Bytes = new byte[3];
                Bytes[0] = (byte)(Address & 0xFF);
                Bytes[1] = (byte)(Data & 0xFF);
                Bytes[2] = CalculateParity(Bytes, 2);
                dataLen = 3;
            }
            else
            {
                return;
            }

            I2C.WriteBytes(Bytes, dataLen, true);
        }

        private uint ReadRegister(uint Address)
        {
            if (!I2C.IsOpen)
                return 0;

            byte[] Addr = new byte[] { (byte)Address };
            byte[] Bytes;
            uint Data = 0;

            if (Address == 0x00)
            {
                Bytes = I2C.WriteAndReadBytes(Addr, 1, 5);

                if (Bytes.Length >= 4)
                {
                    Data = (uint)((Bytes[3] << 24) | (Bytes[2] << 16) | (Bytes[1] << 8) | Bytes[0]);
                }
            }
            else if (Address >= 0x60)
            {
                Bytes = I2C.WriteAndReadBytes(Addr, 1, 2);

                if (Bytes.Length >= 1)
                {
                    Data = Bytes[0];
                }
            }

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
            //Parent.ChipCtrlTextboxes[0].ReadOnly = true;
            //Parent.ChipCtrlTextboxes[0].Visible = true;
            //Parent.ChipCtrlTextboxes[0].Text = Serial.Config.PortName;
            //Parent.ChipCtrlTextboxes[0].Size = new System.Drawing.Size(Parent.ChipCtrlTextboxes[0].Size.Width * 2 + 3, Parent.ChipCtrlTextboxes[0].Size.Height);
            //Parent.ChipCtrlButtons[2].Text = "Connect";
            //Parent.ChipCtrlButtons[2].Visible = true;
            //Parent.ChipCtrlButtons[2].Click += SerialConnect_Click;
            //Parent.ChipCtrlButtons[3].Text = "Set UART";
            //Parent.ChipCtrlButtons[3].Visible = true;
            //Parent.ChipCtrlButtons[3].Click += SerialSetting_Click;
            //Parent.ChipCtrlButtons[].Text = "GH0_H";
            //Parent.ChipCtrlButtons[4].Visible = true;
            //Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;

            Parent.ChipCtrlButtons[8].Text = "TEST";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;

            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_Auto_Test_Items;
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
            Serial.WriteSettingFile(Parent.IniFile, "SCP1501_R5");
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

        private void Toogle_GPIO_GH0(object sender, EventArgs e)
        {
            if (Parent.ChipCtrlButtons[4].Text == "GH0_H")
            {
                I2C.GPIOs[4].Direction = GPIO_Direction.Output;
                I2C.GPIOs[4].State = GPIO_State.Low;
                Parent.ChipCtrlButtons[4].Text = "GH0_L";
            }
            else
            {
                I2C.GPIOs[4].Direction = GPIO_Direction.Output;
                I2C.GPIOs[4].State = GPIO_State.High;
                Parent.ChipCtrlButtons[4].Text = "GH0_H";
            }
        }

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
            // if (Serial.IsOpen && IsSerialReceivedData)
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
                // IsSerialReceivedData = false;
            }
        }

        public override void SendCommand(string Command)
        {
            // Test only
            Command += "\r\n";
            byte[] Var = Encoding.ASCII.GetBytes(Command);
            // Log.WriteLine("\nSend : " + Command);
            Serial.WriteBytes(Var, Var.Length, true);
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal;
            uint Result = 0;
            TEST_ITEMS_MANUAL TestItem = (TEST_ITEMS_MANUAL)TestItemIndex;

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
                        case TEST_ITEMS_AUTO.ReadTest:
                            Read_Register_Test((uint)iVal);
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
        #endregion

        #region Function for LAB test
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
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter4:
                        if (DigitalMultimeter4 == null)
                            DigitalMultimeter4 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter4.IsOpen == false)
                            DigitalMultimeter4.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter5:
                        if (DigitalMultimeter5 == null)
                            DigitalMultimeter5 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter5.IsOpen == false)
                            DigitalMultimeter5.Open();
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

        private void Read_Register_Test(uint test)
        {
            uint cnt = new uint();
            uint data = new uint();
            uint address = new uint();
            address = 0x7A;
            if (test <= 0)
            {
                Log.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss.fff")}] Test Done.");
                return;
            }
            else
            {
                cnt = 1;
                for (uint i = cnt; i <= test; i++)
                {
                    data = ReadRegister(address);
                    Log.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss.fff")}] NO.{i} = 0x{address.ToString("X2")} : 0x{data.ToString("X2")}");
                    Thread.Sleep(100);
                }
                Log.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss.fff")}] Test Done.");
            }
        }
        #endregion
    }
}
