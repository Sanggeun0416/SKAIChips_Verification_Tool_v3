using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SKAIChips_Verification;
using JLcLib.Chip;
using JLcLib.Comn;

namespace SKAI_SG
{
    public class EQUUS : ChipControl
    {
        #region Variable and declaration
        public enum FLASH_STATE
        {
            CMD_READY = (1 << 4),
            ADDR_DATA_READY = (1 << 5),
            READY = (1 << 4) | (1 << 5),
            ADDR_LATCHED = (1 << 6),
            DATA_LATCHED = (1 << 7),
            MAT_BUSY = (1 << 8),
        }

        public enum FW_TARGET
        {
            NV_MEM = 0,
            RAM = 1,
        }

        public enum TEST_ITEMS
        {
            ENTER_TEST_PT, // 0x01 (FW command)
            EXIT_TEST_PT, // 0x02

            SET_TEST_VRECT, // 0x04

            LDO_TURNON, // 0x08 (FW command)
            LDO_TURNOFF, // 0x09
            LDO_SET, // 0x0A

            ADC_GET_VRECT, // 0x10,
            ADC_GET_VOUT, // 0x11,
            ADC_GET_IOUT, // 0x12,
            ADC_GET_VBAT_SENS, // 0x13,
            ADC_GET_VPTAT, // 0x14,
            ADC_GET_NTC, // 0x15,
            //ADC_GET_BGR, // 0x16
            //ADC_GET_ILIM, // 0x17
            //ADC_GET_REF, // 0x18
            //ADC_GET_V_DL, // 0x19
            //ADC_GET_I_DL, // 0x1A

            GET_VRECT_mV, // 0x20,
            GET_VOUT_mV, // 0x21,
            GET_IOUT_mA, // 0x22,
            GET_VBAT_SENS_mV, // 0x23,
            GET_VPTAT_mV, // 0x24,
            GET_NTC_mV, // 0x25,
            GET_IOUT_MIN_mA, // 0x26, // test only
            GET_IOUT_MAX_mA, // 0x27, // test only

            /* 0x40 ~ 0x4F : ADC test MIN */
            ADC_GET_VRECT_MIN,
            ADC_GET_VOUT_MIN,
            ADC_GET_IOUT_MIN,
            ADC_GET_VBAT_SENS_MIN,
            ADC_GET_VPTAT_MIN,
            ADC_GET_NTC_MIN,
            //ADC_GET_BGR_MIN,
            //ADC_GET_ILIM_MIN,
            //ADC_GET_REF_MIN,
            //ADC_GET_V_DL_MIN,
            //ADC_GET_I_DL_MIN,

            /* 0x50 ~ 0x5F : ADC test MIN */
            ADC_GET_VRECT_MAX,
            ADC_GET_VOUT_MAX,
            ADC_GET_IOUT_MAX,
            ADC_GET_VBAT_SENS_MAX,
            ADC_GET_VPTAT_MAX,
            ADC_GET_NTC_MAX,
            //ADC_GET_BGR_MAX,
            //ADC_GET_ILIM_MAX,
            //ADC_GET_REF_MAX,
            //ADC_GET_V_DL_MAX,
            //ADC_GET_I_DL_MAX,

            TEST_ISEN, // GUI test function
            TEST_VOUT, // GUI test function
            TEST_VOUT_WPC, // GUI test function
            TEST_VRECT_WPC, // GUI test function
            TEST_ISEN_REG, // GUI test function
            TEST_ACTIVE_LOAD, // GUI test function
            TEST_LOAD_SWEEP, // GUI test function
            TEST_ADC_CODE, // GUI test function


            NUM_TEST_ITEMS,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private FW_TARGET FirmwareTarget { get; set; } = FW_TARGET.NV_MEM;

        public int F1_SlaveAddress { get; private set; } = 0x24;
        public int F2_SlaveAddress { get; } = 0x04;
        public int F3_SlaveAddress { get; } = 0x05;

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        private bool IsRunCal = false;
        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI PowerSupply1 = null;
        JLcLib.Instrument.SCPI Oscilloscope = null;
        JLcLib.Instrument.SCPI ElectronicLoad = null;
        #endregion Variable and declaration

        public EQUUS(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "EQUUS");
            Serial.DataReceived += Serial_DataReceived;
            CalibrationData = new byte[256];

            /* Init test items combo box */
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e)
        {
            IsSerialReceivedData = true;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[2];

            SendData[0] = (byte)(Address & 0xFF);
            SendData[1] = (byte)(Data & 0xFF);

            iComn.WriteBytes(SendData, SendData.Length, true);
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[1];
            byte[] RcvData = new byte[1];

            SendData[0] = (byte)(Address & 0xFF);

            iComn.WriteBytes(SendData, SendData.Length, true);
            RcvData = iComn.ReadBytes(RcvData.Length);

            return RcvData[0];
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

        #region SKAI EQUSS host control methods
        public void SetHostMode()
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56); // MGC
            SendBytes.Add(0x81); // Command (CH_HOST)
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        public void ResetSystem()
        {
            byte[] SendBytes = new byte[2] { 0x56, 0xF0 };

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        int ReadHostMemory(uint Address)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56); // MGC
            SendBytes.Add(0xB0); // Command (Set read address)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            byte[] RcvBytes = I2C.ReadBytes(6);
            int Data = -1;
            if (RcvBytes != null && RcvBytes.Length >= 6 && RcvBytes[0] == 0x56 && RcvBytes[1] == 0x01) // Check MGC and STA
                Data = (RcvBytes[2] << 24) | (RcvBytes[3] << 16) | (RcvBytes[4] << 8) | RcvBytes[5];

            I2C.Config.SlaveAddress = F1_SlaveAddress;
            return Data;
        }

        byte[] ReadMemory(uint Address, int Length)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56); // MGC
            SendBytes.Add(0xB0); // Command (Set write data)
            //SendBytes.Add(0xB0); // Command (Set read address)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            byte[] RcvBytes = I2C.ReadBytes(Length + 2);
#if true
            // if 32bit operation
            List<byte> Data = new List<byte>();
            for (int i = 2; i < RcvBytes.Length; i += 4)
            {
                Data.Add(RcvBytes[3 + i]);
                Data.Add(RcvBytes[2 + i]);
                Data.Add(RcvBytes[1 + i]);
                Data.Add(RcvBytes[0 + i]);
            }
            RcvBytes = Data.ToArray();
            // endif 32bit operation
#endif
            I2C.Config.SlaveAddress = F1_SlaveAddress;

            return RcvBytes;
        }

        void WriteHostMemory(uint Address, uint Data)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56);
            SendBytes.Add(0xA0); // Command (Set write data)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        void WriteHostMemory(int Address, byte[] Data)
        {
            List<byte> SendBytes = new List<byte>();

            if (Data.Length < 4)
                return;

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56);
            SendBytes.Add(0xA0); // Command (Set write data)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            SendBytes.Add(Data[3]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[0]);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        void WriteMemory(uint Address, byte[] Data)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56);
            SendBytes.Add(0xA0); // Command (Set write data)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));

#if false
            // Normal byte operation
            for (int i = 0; i < Data.Length; i++)
                SendBytes.Add(Data[i]);
#else
            // 32bits double word operation
            for (int i = 0; i < Data.Length; i += 4)
            {
                SendBytes.Add(Data[3 + i]);
                SendBytes.Add(Data[2 + i]);
                SendBytes.Add(Data[1 + i]);
                SendBytes.Add(Data[0 + i]);
            }
#endif
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        void CheckRemap()
        {
            int Data = ReadHostMemory(0x4001F000);
            if (Data != 0)
            {
                WriteHostMemory(0x4001F000, 0x0);
                ResetSystem();
            }
        }

        bool WaitFlashState(FLASH_STATE State, int Timeout = 300)
        {
            bool Completed = false;
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();

            sw.Start();
            do
            {
                int Data = ReadHostMemory(0x0FFE0004);
                switch (State)
                {
                    case FLASH_STATE.CMD_READY:
                    case FLASH_STATE.ADDR_DATA_READY:
                        if ((Data & (int)State) == (int)State)
                            Completed = true;
                        break;
                    case FLASH_STATE.READY:
                        if (Data == (int)State)
                            Completed = true;
                        break;
                    case FLASH_STATE.ADDR_LATCHED:
                    case FLASH_STATE.DATA_LATCHED:
                    case FLASH_STATE.MAT_BUSY:
                        if ((Data & (int)State) == 0)
                        {
                            if (State == FLASH_STATE.MAT_BUSY)
                            {
                                if ((Data & 0xFF) == 0x30)
                                    Completed = true;
                            }
                            else
                                Completed = true;
                        }
                        break;
                }
            } while (sw.Elapsed.TotalMilliseconds < Timeout && Completed == false);

            return Completed;
        }
        #endregion SKAI EQUSS host control methods

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

            Parent.ChipCtrlButtons[4].Text = "WR CAL";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += WriteCalibrationData_Click;
            Parent.ChipCtrlButtons[5].Text = "RM CAL";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += RemoveCalibrationData_Click;
            Parent.ChipCtrlButtons[6].Text = "RD CAL";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += ReadCalibrationData_Click;
            Parent.ChipCtrlButtons[7].Text = "Run CAL";
            Parent.ChipCtrlButtons[7].Visible = true;
            Parent.ChipCtrlButtons[7].Click += RunCalibration_Click; ;
            button_RunCal0 = Parent.ChipCtrlButtons[7];

            //Parent.ChipCtrlButtons[7].Text = "DN RAM";
            //Parent.ChipCtrlButtons[7].Visible = true;
            //Parent.ChipCtrlButtons[7].Click += DownloadFirmwareToRAM_Click;

            Parent.ChipCtrlButtons[8].Text = "Down FW";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += DownloadFirmwareToNV_Click;
            Parent.ChipCtrlTextboxes[9].Text = ""; // Firmware size textbox
            Parent.ChipCtrlTextboxes[9].Visible = true;
            Parent.ChipCtrlButtons[10].Text = "Dump FW";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += new EventHandler(delegate { DumpFW(); });
            Parent.ChipCtrlButtons[11].Text = "Erase FW";
            Parent.ChipCtrlButtons[11].Visible = true;
            Parent.ChipCtrlButtons[11].Click += new EventHandler(delegate { EraseFW(); });
        }

        private void SerialSetting_Click(object sender, EventArgs e)
        {
            JLcLib.Comn.WireComnForm.Show(Serial);
            Serial.WriteSettingFile(Parent.IniFile, "EQUUS");
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

        private void WriteCalibrationData_Click(object sender, EventArgs e)
        {
#if true // For LSI
            RegisterItem BGR = Parent.RegMgr.GetRegisterItem("BGR");
            RegisterItem LDO5P0 = Parent.RegMgr.GetRegisterItem("LDO5P0");
            RegisterItem LDO1P5 = Parent.RegMgr.GetRegisterItem("LDO1P5");
            RegisterItem ADC_LDO3P3 = Parent.RegMgr.GetRegisterItem("ADC_LDO3P3");
            RegisterItem FSK_LDO3P3 = Parent.RegMgr.GetRegisterItem("FSK_LDO3P3");
            RegisterItem LDO1P8 = Parent.RegMgr.GetRegisterItem("LDO1P8");
            RegisterItem OSC = Parent.RegMgr.GetRegisterItem("OSC");
            RegisterItem ML_ISEN_VREF_SET = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_SET");
            RegisterItem ML_VILIM_SET = Parent.RegMgr.GetRegisterItem("ML_VILIM_SET");
            RegisterItem VRECT_SV_0 = Parent.RegMgr.GetRegisterItem("VRECT_SV[7:0]");
            RegisterItem VRECT_SV_1 = Parent.RegMgr.GetRegisterItem("VRECT_SV[15:8]");
            RegisterItem Vrect_Const_0 = Parent.RegMgr.GetRegisterItem("Vrect_Const[7:0]");
            RegisterItem Vrect_Const_1 = Parent.RegMgr.GetRegisterItem("Vrect_Const[15:8]");
            RegisterItem Vout_SET_SV_0 = Parent.RegMgr.GetRegisterItem("Vout_SET_SV[7:0]");
            RegisterItem Vout_SET_SV_1 = Parent.RegMgr.GetRegisterItem("Vout_SET_SV[15:8]");
            RegisterItem Vout_SET_Const_0 = Parent.RegMgr.GetRegisterItem("Vout_SET_Const[7:0]");
            RegisterItem Vout_SET_Const_1 = Parent.RegMgr.GetRegisterItem("Vout_SET_Const[15:0]");
            RegisterItem Vout_SV_0 = Parent.RegMgr.GetRegisterItem("Vout_SV[7:0]");
            RegisterItem Vout_SV_1 = Parent.RegMgr.GetRegisterItem("Vout_SV[15:8]");
            RegisterItem Vout_Const_0 = Parent.RegMgr.GetRegisterItem("Vout_Const[7:0]");
            RegisterItem Vout_Const_1 = Parent.RegMgr.GetRegisterItem("Vout_Const[15:8]");
            RegisterItem ADC_SV_0 = Parent.RegMgr.GetRegisterItem("ADC_SV[7:0]");
            RegisterItem ADC_SV_1 = Parent.RegMgr.GetRegisterItem("ADC_SV[15:8]");
            RegisterItem ADC_Const_0 = Parent.RegMgr.GetRegisterItem("ADC_Const[7:0]");
            RegisterItem ADC_Const_1 = Parent.RegMgr.GetRegisterItem("ADC_Const[15:8]");
            RegisterItem ISen2_SV_0 = Parent.RegMgr.GetRegisterItem("ISen2_SV[7:0]");
            RegisterItem ISen2_SV_1 = Parent.RegMgr.GetRegisterItem("ISen2_SV[15:8]");
            RegisterItem ISen2_Const_0 = Parent.RegMgr.GetRegisterItem("ISen2_Const[7:0]");
            RegisterItem ISen2_Const_1 = Parent.RegMgr.GetRegisterItem("ISen2_Const[15:8]");
            RegisterItem ISenLow_0 = Parent.RegMgr.GetRegisterItem("ISenLow[7:0]");
            RegisterItem ISenLow_1 = Parent.RegMgr.GetRegisterItem("ISenLow[15:8]");
            RegisterItem ISenDiff_0 = Parent.RegMgr.GetRegisterItem("ISenDiff[7:0]");
            RegisterItem ISenDiff_1 = Parent.RegMgr.GetRegisterItem("ISenDiff[15:8]");
            RegisterItem ISen0_SV_0 = Parent.RegMgr.GetRegisterItem("ISen0_SV[7:0]");
            RegisterItem ISen0_SV_1 = Parent.RegMgr.GetRegisterItem("ISen0_SV[15:8]");
            RegisterItem ISen0_Const_0 = Parent.RegMgr.GetRegisterItem("ISen0_Const[7:0]");
            RegisterItem ISen0_Const_1 = Parent.RegMgr.GetRegisterItem("ISen0_Const[15:8]");
            RegisterItem ISen1_SV_0 = Parent.RegMgr.GetRegisterItem("ISen1_SV[7:0]");
            RegisterItem ISen1_SV_1 = Parent.RegMgr.GetRegisterItem("ISen1_SV[15:8]");
            RegisterItem ISen1_Const_0 = Parent.RegMgr.GetRegisterItem("ISen1_Const[7:0]");
            RegisterItem ISen1_Const_1 = Parent.RegMgr.GetRegisterItem("ISen1_Const[15:8]");

            CalibrationData[0x24] = (byte)BGR.Value;
            CalibrationData[0x28] = (byte)LDO5P0.Value;
            CalibrationData[0x2C] = (byte)LDO1P5.Value;
            CalibrationData[0x30] = (byte)ADC_LDO3P3.Value;
            CalibrationData[0x34] = (byte)FSK_LDO3P3.Value;
            CalibrationData[0x38] = (byte)LDO1P8.Value;
            CalibrationData[0x3C] = (byte)OSC.Value;
            CalibrationData[0x40] = (byte)ML_ISEN_VREF_SET.Value;
            CalibrationData[0x44] = (byte)ML_VILIM_SET.Value;
            CalibrationData[0x48] = (byte)VRECT_SV_0.Value;
            CalibrationData[0x49] = (byte)VRECT_SV_1.Value;
            CalibrationData[0x4C] = (byte)Vrect_Const_0.Value;
            CalibrationData[0x4D] = (byte)Vrect_Const_1.Value;
            CalibrationData[0x50] = (byte)Vout_SET_SV_0.Value;
            CalibrationData[0x51] = (byte)Vout_SET_SV_1.Value;
            CalibrationData[0x54] = (byte)Vout_SET_Const_0.Value;
            CalibrationData[0x55] = (byte)Vout_SET_Const_1.Value;
            CalibrationData[0x58] = (byte)Vout_SV_0.Value;
            CalibrationData[0x59] = (byte)Vout_SV_1.Value;
            CalibrationData[0x5C] = (byte)Vout_Const_0.Value;
            CalibrationData[0x5D] = (byte)Vout_Const_1.Value;
            CalibrationData[0x60] = (byte)ADC_SV_0.Value;
            CalibrationData[0x61] = (byte)ADC_SV_1.Value;
            CalibrationData[0x64] = (byte)ADC_Const_0.Value;
            CalibrationData[0x65] = (byte)ADC_Const_1.Value;
            CalibrationData[0x70] = (byte)ISen2_SV_0.Value;
            CalibrationData[0x71] = (byte)ISen2_SV_1.Value;
            CalibrationData[0x74] = (byte)ISen2_Const_0.Value;
            CalibrationData[0x75] = (byte)ISen2_Const_1.Value;
            CalibrationData[0x78] = (byte)ISenLow_0.Value;
            CalibrationData[0x79] = (byte)ISenLow_1.Value;
            CalibrationData[0x7C] = (byte)ISenDiff_0.Value;
            CalibrationData[0x7D] = (byte)ISenDiff_1.Value;
            CalibrationData[0x80] = (byte)ISen0_SV_0.Value;
            CalibrationData[0x81] = (byte)ISen0_SV_1.Value;
            CalibrationData[0x84] = (byte)ISen0_Const_0.Value;
            CalibrationData[0x85] = (byte)ISen0_Const_1.Value;
            CalibrationData[0x88] = (byte)ISen1_SV_0.Value;
            CalibrationData[0x89] = (byte)ISen1_SV_1.Value;
            CalibrationData[0x8C] = (byte)ISen1_Const_0.Value;
            CalibrationData[0x8D] = (byte)ISen1_Const_1.Value;
#endif            
            // Test code start, Initialize calibration data
            //for (int i = 0; i < CalibrationData.Length; i++)
            //    CalibrationData[i] = (byte)i;
            // Test code end
            WriteCalibrationData(EQUUS_WriteCalibrationData);
        }

        private void RemoveCalibrationData_Click(object sender, EventArgs e)
        {
            RemoveCalibrationData(EQUUS_RemoveCalibrationData);
        }

        private void ReadCalibrationData_Click(object sender, EventArgs e)
        {
            ReadCalibrationData(EQUUS_ReadCalibrationData);
        }

        private void RunCalibration_Click(object sender, EventArgs e)
        {
            Parent.StopLogThread();
            if (IsRunCal == false)
            {
                button_RunCal0.Text = "Stop Cal";
                IsRunCal = true;
                RunBackgroudFunc(EQUUS_RunCalibration, EQUUS_StopCalibration);
            }
            else
                EQUUS_StopCalibration();
        }

        private void DownloadFirmwareToRAM_Click(object sender, EventArgs e)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = FW_TARGET.RAM;
                DownloadFirmware(EQUUS_DownloadFW);
            }
        }

        private void DownloadFirmwareToNV_Click(object sender, EventArgs e)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = FW_TARGET.NV_MEM;
                DownloadFirmware(EQUUS_DownloadFW);
            }
        }

        private void EraseFW()
        {
            EraseFirmware(EQUUS_EraseFW);
        }

        private void DumpFW()
        {
            ReadFirmwareSize = GetInt_ChipCtrlTextboxes(9);
            DumpFirmware(EQUUS_DumpFW);
        }

        #region Firmware control methods
        private void EQUUS_DownloadFW()
        {
            int PageSize = 4;
            uint FlashAddress = 0;
            byte[] SendBytes = new byte[PageSize];
            System.IO.FileStream fs = null;
            System.IO.BinaryReader br = null;

            Status = false;
            if (I2C == null)
                return;

            /* Read Firmware file */
            fs = new System.IO.FileStream(FirmwareName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            br = new System.IO.BinaryReader(fs);
            FirmwareData = br.ReadBytes((int)fs.Length);
            FirmwareSize = (int)fs.Length;

            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (FirmwareData.Length + PageSize - 1) / PageSize;
            })));

            /* Mass erase */
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                EQUUS_EraseFW();
                if (Status == false)
                    goto EXIT;
            }
            else
                PageSize = 128;

            /* Program */
            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address & data latch
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                if (WaitFlashState(FLASH_STATE.DATA_LATCHED) == false)
                    goto EXIT;
            }
            for (FlashAddress = 0; FlashAddress < FirmwareData.Length; FlashAddress += (uint)PageSize)
            {
                // Copy FW data
                for (int i = 0; i < PageSize; i++)
                {
                    if (FlashAddress + i < FirmwareData.Length)
                        SendBytes[i] = FirmwareData[FlashAddress + i];
                    else
                        SendBytes[i] = 0xFF;
                }
                if (FirmwareTarget == FW_TARGET.NV_MEM)
                {
                    // a. Write flash address
                    WriteHostMemory(0x0FFE0200, FlashAddress);
                    WriteHostMemory(0x0FFE0204, SendBytes);
                    // b. Wait for command latched signal
                    if (WaitFlashState(FLASH_STATE.CMD_READY) == false)
                        goto EXIT;
                    // c. Write program command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
                    WriteHostMemory(0x0FFE0000, 0x02);
                    // 7, Check erase complete
                    if (WaitFlashState(FLASH_STATE.MAT_BUSY) == false)
                        goto EXIT;
                }
                else // Firmware target is SRAM
                {
                    // 4. Write firmware data
                    WriteMemory(0x20000000 + FlashAddress, SendBytes);
                }
                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            if (FirmwareTarget == FW_TARGET.RAM)
            {
                WriteHostMemory(0x4001F000, 0x01); // REMAP = 1
                ResetSystem();
            }
            Status = true;
            br.Close();
            fs.Close();

        EXIT:
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }

        public void EQUUS_EraseFW()
        {
            Status = false;
            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address latch
            if (WaitFlashState(FLASH_STATE.ADDR_LATCHED, 1000) == false)
                return;
            // 5. Write flash address
            WriteHostMemory(0x0FFE0200, 0x00);
            // 6. Write erase command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
            WriteHostMemory(0x0FFE0000, 0x11);
            // 7, Check erase complete
            if (WaitFlashState(FLASH_STATE.MAT_BUSY, 5000) == false) // timeout 5secI
                return;
            Status = true;
        }

        public void EQUUS_DumpFW()
        {
            const int PageSize = 128;
            byte[] RcvBytes;
            List<byte> FirmwareData = new List<byte>();

            Status = false;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (ReadFirmwareSize + PageSize - 1) / PageSize + 1; // Mass erase
            })));

            // 1. Set I2C host mode
            SetHostMode();
            // 2. Read FW
            for (uint Addr = 0; Addr < ReadFirmwareSize; Addr += PageSize)
            {
                RcvBytes = ReadMemory(Addr, PageSize);
                if (RcvBytes != null && RcvBytes.Length > 0)
                {
#if true
                    // a byte operation
                    for (int i = 0; i < RcvBytes.Length; i++)
                        FirmwareData.Add(RcvBytes[i]);
#else
                    // 32bits double word operation
                    for (int i = 0; i < RcvBytes.Length; i += 4)
                    {
                        FirmwareData.Add(RcvBytes[i + 3]); // [ 7: 0]
                        FirmwareData.Add(RcvBytes[i + 2]); // [15: 8]
                        FirmwareData.Add(RcvBytes[i + 1]); // [23:16]
                        FirmwareData.Add(RcvBytes[i + 0]); // [31:24]
                    }
#endif
                }
                else
                    goto EXIT;

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            Status = true;
            ReadFirmwareData = FirmwareData.ToArray();
            System.IO.FileStream fs = new System.IO.FileStream("ReadFirmwareBinary.bin", System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(ReadFirmwareData);
            bw.Close();
            fs.Close();
        EXIT:
            // 3. Reset system
            ResetSystem();
            Status = true;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }
        #endregion Firmware control methods

        #region Calibration control methods
        private void EQUUS_WriteCalibrationData()
        {
            uint PageSize = 4;
            int Count = 0;
            byte[] SendBytes = new byte[PageSize];

            Status = false;
            // 1~6. Remove calibartion data (Page erase)
            EQUUS_RemoveCalibrationData();

            // Write calibration data
            //for (uint FlashAddress = 0xFFFD18FC; FlashAddress >= 0xFFFD1800; FlashAddress -= PageSize)
            for (uint FlashAddress = 0xFFFD1800; FlashAddress < 0xFFFD1900; FlashAddress += PageSize)
            {
                // Copy FW data
                for (int i = 0; i < PageSize; i++)
                {
                    if (Count < CalibrationData.Length)
                        SendBytes[i] = CalibrationData[Count];
                    else
                        SendBytes[i] = 0xFF;
                    Count++;
                }
                // a. Write flash address
                WriteHostMemory(0x0FFE0200, FlashAddress);
                WriteHostMemory(0x0FFE0204, SendBytes);
                // b. Wait for command latched signal
                if (WaitFlashState(FLASH_STATE.CMD_READY) == false)
                    return;
                // c. Write program command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
                WriteHostMemory(0x0FFE0000, 0x02);
                // 7, Check erase complete
                if (WaitFlashState(FLASH_STATE.MAT_BUSY) == false)
                    return;
                Status = true;
            }
        }

        private void EQUUS_RemoveCalibrationData()
        {
            Status = false;
            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address latch
            //if (WaitFlashState(FLASH_STATE.READY, 1000) == false)
            if (WaitFlashState(FLASH_STATE.ADDR_LATCHED, 1000) == false)
                return;
            // 5,6. Erase
            WriteHostMemory(0x0FFE0200, 0xFFFD1800); // NV memory address for calibration data
            // 6. Write erase command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
            WriteHostMemory(0x0FFE0000, 0x01); // 256 bytes page erase command
            // 7, Check erase complete
            if (WaitFlashState(FLASH_STATE.MAT_BUSY, 5000) == false) // timeout 5secI
                return;
            Status = true;
        }

        private void EQUUS_ReadCalibrationData()
        {
            uint PageSize = 4;
            int Data;

            Status = false;

            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address latch
            //if (WaitFlashState(FLASH_STATE.READY, 1000) == false)
            if (WaitFlashState(FLASH_STATE.ADDR_LATCHED, 1000) == false)
                return;

            // Read calibration data
            for (uint FlashAddress = 0x0FFD1800; FlashAddress < 0x0FFD1900; FlashAddress += PageSize)
            {
                Data = ReadHostMemory(FlashAddress);
                Log.WriteLine(FlashAddress.ToString("X8") + " : " + Data.ToString("X8"));

                Status = true;
            }
        }
        #endregion Calibration control methods

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
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
        }

        public override void SendCommand(string Command)
        {
            // Test only
            Command += "\r\n";
            byte[] Var = Encoding.ASCII.GetBytes(Command);
            Serial.WriteBytes(Var, Var.Length, true);
        }

        private void StartCommandProcessing(RegisterItem CommandReg, int Timeout = 1000)
        {
            CommandReg.Write();
            for (int i = 0; i < (Timeout / 10); i++)
            {
                System.Threading.Thread.Sleep(10);
                CommandReg.Read();
                if (CommandReg.Value == 0)
                    break;
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
                case TEST_ITEMS.TEST_ISEN:
                    TEST_RunIsen();
                    break;
                case TEST_ITEMS.TEST_VOUT:
                    TEST_RunVout();
                    break;
                case TEST_ITEMS.TEST_VOUT_WPC:
                    TEST_RunVout_wpc();
                    break;
                case TEST_ITEMS.TEST_VRECT_WPC:
                    TEST_RunVrect_wpc();
                    break;
                case TEST_ITEMS.TEST_ISEN_REG:
                    TEST_RunIsen_Register();
                    break;
                case TEST_ITEMS.TEST_ACTIVE_LOAD:
                    TEST_RunActiveLoad();
                    break;
                case TEST_ITEMS.TEST_LOAD_SWEEP:
                    TEST_RunLOAD();
                    break;
                case TEST_ITEMS.TEST_ADC_CODE:
                    TEST_RunADCCODE();
                    break;

                // FW test functions
                default:
                    Result = RunTest(TestItem, iVal);
                    break;
            }
            Log.WriteLine(TestItem.ToString() + ":" + iVal.ToString() + ":" + Result.ToString());
        }

        public int RunTest(TEST_ITEMS TestItem, int Arg)
        {
            int iVal = 0;
            RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");   // 0x1C
            RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");     // 0x1D
            RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");         // 0x1E
            RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");         // 0x1F

            switch (TestItem)
            {
                // Test mode control
                case TEST_ITEMS.ENTER_TEST_PT:
                case TEST_ITEMS.EXIT_TEST_PT:
                case TEST_ITEMS.SET_TEST_VRECT:
                    CommandReg.Value = 0x01 + (uint)TestItem - (uint)TEST_ITEMS.ENTER_TEST_PT;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg);
                    StartCommandProcessing(CommandReg);
                    break;
                // Write command only
                case TEST_ITEMS.LDO_TURNON:
                case TEST_ITEMS.LDO_TURNOFF:
                case TEST_ITEMS.LDO_SET:
                    CommandReg.Value = 0x08 + (uint)TestItem - (uint)TEST_ITEMS.LDO_TURNON;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg);
                    StartCommandProcessing(CommandReg);
                    break;
                case TEST_ITEMS.ADC_GET_VRECT:
                case TEST_ITEMS.ADC_GET_VOUT:
                case TEST_ITEMS.ADC_GET_IOUT:
                case TEST_ITEMS.ADC_GET_VBAT_SENS:
                case TEST_ITEMS.ADC_GET_VPTAT:
                case TEST_ITEMS.ADC_GET_NTC:
                    CommandReg.Value = 0x10 + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_VRECT;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
                case TEST_ITEMS.GET_VRECT_mV: // 0x20
                case TEST_ITEMS.GET_VOUT_mV:
                case TEST_ITEMS.GET_IOUT_mA:
                case TEST_ITEMS.GET_VBAT_SENS_mV:
                case TEST_ITEMS.GET_VPTAT_mV:
                case TEST_ITEMS.GET_NTC_mV:
                case TEST_ITEMS.GET_IOUT_MIN_mA:
                case TEST_ITEMS.GET_IOUT_MAX_mA:
                    CommandReg.Value = 0x20 + (uint)TestItem - (uint)TEST_ITEMS.GET_VRECT_mV;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
                case TEST_ITEMS.ADC_GET_VRECT_MIN:
                case TEST_ITEMS.ADC_GET_VOUT_MIN:
                case TEST_ITEMS.ADC_GET_IOUT_MIN:
                case TEST_ITEMS.ADC_GET_VBAT_SENS_MIN:
                case TEST_ITEMS.ADC_GET_VPTAT_MIN:
                case TEST_ITEMS.ADC_GET_NTC_MIN:
                    CommandReg.Value = 0x40 + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_VRECT_MIN;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;

                case TEST_ITEMS.ADC_GET_VRECT_MAX:
                case TEST_ITEMS.ADC_GET_VOUT_MAX:
                case TEST_ITEMS.ADC_GET_IOUT_MAX:
                case TEST_ITEMS.ADC_GET_VBAT_SENS_MAX:
                case TEST_ITEMS.ADC_GET_VPTAT_MAX:
                case TEST_ITEMS.ADC_GET_NTC_MAX:
                    CommandReg.Value = 0x50 + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_VRECT_MAX;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
            }
            return iVal;
        }
        #region Chip test methods
#if false
        private void TEST_RunIsen()
        {
            double Vrect, Vout;
            int IoutCode, VrectCode, VoutCode;
            int xPos = 3, yPos = 5;

            Log.WriteLine("Start TEST_ISEN");
            
            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;
            if (Parent.xlMgr.Sheet.Select("IoutTest") == false)
                Parent.xlMgr.Sheet.Add("IoutTest");

            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : N/C\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 5V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 5"); // offset 5V
#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200);

            for (double Vr = 4.9; Vr <= 5.5; Vr += 0.1)
            {
                PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                ElectronicLoad.Write("CURR 0");
                JLcLib.Delay.Sleep(500);
                yPos = 5;
                for (double Iload = 0; Iload <= 1200; Iload += 50)
                {
                    ElectronicLoad.Write("CURR " + (Iload / 1000).ToString("F2"));
                    JLcLib.Delay.Sleep(100);

                    IoutCode = RunTest(TEST_ITEMS.ADC_GET_IOUT, 256);
                    VrectCode = RunTest(TEST_ITEMS.ADC_GET_VRECT, 256);
                    VoutCode = RunTest(TEST_ITEMS.ADC_GET_VOUT, 256);

                    Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                    Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                    //Vrect = RunTest(TEST_ITEMS.GET_VRECT_mV, 256);
                    //Vout = RunTest(TEST_ITEMS.GET_VOUT_mV, 256);

                    Parent.xlMgr.Cell.Write(xPos + 0, yPos, VrectCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 1, yPos, VoutCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 2, yPos, IoutCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 3, yPos, (Vrect * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 4, yPos, (Vout * 1000.0).ToString());
                    //Parent.xlMgr.Cell.Write(xPos + 3, yPos, Vrect.ToString());
                    //Parent.xlMgr.Cell.Write(xPos + 4, yPos, Vout.ToString());

                    yPos++;
                }
                xPos += 5;
            }
            Log.WriteLine("End TEST_ISEN");
        }
#else
        private void TEST_RunIsen()
        {
            double Vrect, Vout;
            int IoutADC, IoutCode, VrectCode, VoutCode;
            int xPos = 1, yPos = 1;
            string sheet_name, time;

            Log.WriteLine("Start TEST_ISEN");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : N/C\n- CH4 : N/C\n\nElectronicLoad\n- Out : VOUT");
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 5V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 5"); // offset 5V
#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            time = DateTime.Now.ToString("MMddHHmmss_");
            sheet_name = time + "Iout";
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "I_load(mA)");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "Iout ADC");
            Parent.xlMgr.Cell.Write(xPos + 8, yPos, "Vrect Code");
            Parent.xlMgr.Cell.Write(xPos + 15, yPos, "Vout Code");
            Parent.xlMgr.Cell.Write(xPos + 22, yPos, "Vrect Scope");
            Parent.xlMgr.Cell.Write(xPos + 29, yPos, "Vout Scope");
            Parent.xlMgr.Cell.Write(xPos + 36, yPos, "Iout Code");

            for (double Iload = 0; Iload <= 1200; Iload += 50) // 0~1200(mA) Electronic Load
            {
                ElectronicLoad.Write("CURR " + (Iload / 1000).ToString());
                xPos = 1;
                yPos++;
                Parent.xlMgr.Cell.Write(xPos, yPos + 1, Iload.ToString());
                for (double Vr = 4.9; Vr <= 5.5; Vr += 0.1) // 4.9~5.5(V) Power Supply
                {
                    PowerSupply0.Write("VOLT " + Vr.ToString());
                    JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
                    xPos++;
                    JLcLib.Delay.Sleep(200);
                    if (yPos == 2)
                    {
                        Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 7, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 14, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 21, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 28, yPos, (Vr.ToString() + "V"));
                    }

                    IoutADC = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1024);
                    IoutCode = RunTest(TEST_ITEMS.GET_IOUT_mA, 1024);
                    VrectCode = RunTest(TEST_ITEMS.GET_VRECT_mV, 256);
                    VoutCode = RunTest(TEST_ITEMS.GET_VOUT_mV, 256);
                    Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                    Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));

                    Parent.xlMgr.Cell.Write(xPos + 0, yPos + 1, IoutADC.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 7, yPos + 1, VrectCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 14, yPos + 1, VoutCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 21, yPos + 1, (Vrect * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 28, yPos + 1, (Vout * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 35, yPos + 1, IoutCode.ToString());
                }
            }
            Log.WriteLine("End TEST_ISEN");
        }
#endif
        private void CAL_VOUT_Reg()
        {
            double Result;
            double TargetVoutValue = 5.0;
            double Margin = 0.05;

            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]

            for (uint i = 1; i < 2; i++)
            {
                VOUT_SET_RANGE.Value = i;
                TRIM_VOUT_SET.Value = TRIM_VOUT_SET.Read();

                for (int j = 0; j < 100; j++)
                {
                    JLcLib.Delay.Sleep(50);
                    Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                    if (Result > TargetVoutValue + Margin)
                    {
                        TRIM_VOUT_SET.Value -= 1;
                    }
                    else if (Result < TargetVoutValue - Margin)
                    {
                        TRIM_VOUT_SET.Value += 1;
                        if (TRIM_VOUT_SET.Value > 100)
                            TRIM_VOUT_SET.Value = 100;
                    }
                    else
                        break;
                    TRIM_VOUT_SET.Write();
                }
            }
        }

        private void TEST_RunIsen_Register()
        {
            double Vrect_vpp, Vrect, Vout, Iout;
            int IoutCode_Min, IoutCode_Max, IoutCode, IoutADC_Min, IoutADC_Max, IoutADC, VrectCode, VoutCode;
            int xPos = 0, yPos = 3;
            string sheet_name, time;

            RegisterItem FB = Parent.RegMgr.GetRegisterItem("ML_FB_TRIM<3:0>");
            RegisterItem ISEN_VREF = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>");
            RegisterItem ISEN_RVAR = Parent.RegMgr.GetRegisterItem("ML_ISEN_RVAR<1:0>");
            RegisterItem VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>");
            RegisterItem VREF = Parent.RegMgr.GetRegisterItem("ML_VREF_TRIM<3:0>");
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN");
            RegisterItem FC_R = Parent.RegMgr.GetRegisterItem("ML_FC_RTRIM<3:0>");
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>");
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");
            RegisterItem ANA_MUX = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");

            Log.WriteLine("Start TEST_ISEN Sweep Register");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : VBAT_SENS\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5.3");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 5V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 5"); // offset 5V
            // Cannel.3 use for Iout(mV)
            Oscilloscope.Write(":CHAN3:DISP 1");
            Oscilloscope.Write(":CHAN3:SCAL 5E-1"); // 500mV/div
            Oscilloscope.Write(":CHAN3:OFFS 1.5"); // offset 5V
#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            //ElectronicLoad.Write("CURR 0.6");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            OCL_EN.Value = 0;
            OCL_EN.Write();
            FC_R.Value = 15;
            FC_R.Write();
            ANA_SEL.Value = 3;
            ANA_SEL.Write();
            ANA_PEN.Value = 1;
            ANA_PEN.Write();
            ANA_MUX.Value = 1;
            ANA_MUX.Write();

            time = DateTime.Now.ToString("MMddHHmmss_");

            for (double Iload = 600; Iload <= 600; Iload += 50) // 0~1200(mA) Electronic Load
            {
                sheet_name = time + Iload.ToString() + "mA";
                Parent.xlMgr.Sheet.Add(sheet_name);

                xPos = 3;
                yPos = 4;
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, ("Load = " + Iload.ToString() + "mA"));
                yPos++;
                Parent.xlMgr.Cell.Write(xPos + 5, yPos, "IOUT_ADC_AVG(Max 4095)");
                Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, "IOUT_PP(mV)");
                Parent.xlMgr.Cell.Write(xPos + 5 + 14, yPos, "IOUT_Code_AVG(Max 1200)");
                yPos++;
                Parent.xlMgr.Cell.Write(xPos + 0, yPos, "ML_FB_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 1, yPos, "ML_VREF_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 2, yPos, "ML_ISEN_RVAR");
                Parent.xlMgr.Cell.Write(xPos + 3, yPos, "ML_VILIM_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, "ML_ISEN_VREF");

                for (double Vr = 4.9; Vr <= 5.5; Vr += 0.1) // 4.9~5.5(V) Power Supply
                {
                    yPos = 6;
                    Parent.xlMgr.Cell.Write(xPos + 5, yPos, (Vr.ToString() + "V"));
                    Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, (Vr.ToString() + "V"));
                    Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, (Vr.ToString() + "V"));
                    yPos++;
                    for (uint i = 0; i <= 15; i++) // 0~15 ML_FB_TRIM<3:0>
                    {
                        FB.Value = i;
                        FB.Write();
                        for (uint j = 0; j <= 15; j++) // 0~15 ML_VREF_TRIM<3:0>
                        {
                            VREF.Value = j;
                            VREF.Write();
                            ElectronicLoad.Write("CURR 0");
                            PowerSupply0.Write("VOLT 6V");
                            JLcLib.Delay.Sleep(100);
                            CAL_VOUT_Reg();
                            PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                            ElectronicLoad.Write("CURR " + (Iload / 1000).ToString("F2"));
                            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
#if true
                            for (uint k = 0; k <= 3; k++) // 0~3 ML_ISEN_RVAR<1:0>
                            {
                                ISEN_RVAR.Value = k;
                                ISEN_RVAR.Write();
                                for (uint l = 0; l <= 15; l++) // 0~15 ML_VILIM_TRIM<3:0>
                                {
                                    VILIM.Value = l;
                                    VILIM.Write();
                                    for (uint m = 0; m <= 15; m++) // 0~15 ML_ISEN_VREF_TRIM<3:0>
                                    {
                                        ISEN_VREF.Value = m;
                                        ISEN_VREF.Write();
                                        JLcLib.Delay.Sleep(50);
#if true
                                        IoutCode = RunTest(TEST_ITEMS.GET_IOUT_mA, 256);
                                        IoutADC_Min = RunTest(TEST_ITEMS.ADC_GET_IOUT_MIN, 256);
                                        IoutADC_Max = RunTest(TEST_ITEMS.ADC_GET_IOUT_MAX, 256);
                                        IoutADC = RunTest(TEST_ITEMS.ADC_GET_IOUT, 256);
                                        //VrectCode = RunTest(TEST_ITEMS.GET_VRECT_mV, 256);
                                        //VoutCode = RunTest(TEST_ITEMS.GET_VOUT_mV, 256);

                                        //Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                                        //Vrect_vpp = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VPP? CHAN1"));
                                        //Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                                        Iout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VPP? CHAN3"));

                                        if (xPos == 3)
                                        {
                                            Parent.xlMgr.Cell.Write(xPos + 0, yPos, i.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 1, yPos, j.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 2, yPos, k.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 3, yPos, l.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 4, yPos, m.ToString());
                                        }
                                        Parent.xlMgr.Cell.Write(xPos + 5, yPos, IoutADC.ToString());
                                        Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, (Iout * 1000.0).ToString());
                                        Parent.xlMgr.Cell.Write(xPos + 5 + 14, yPos, IoutCode.ToString());

                                        //Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, IoutADC_Min.ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 5 + 14, yPos, IoutADC_Max.ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vrect_vpp * 1000.0).ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 2, yPos, (Vrect * 1000.0).ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 3, yPos, VrectCode.ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 4, yPos, (Vout * 1000.0).ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 5, yPos, VoutCode.ToString());                                        
#endif
                                        yPos++;
                                    }
                                }
                            }
#endif
                        }
                    }
                    xPos++;


                }
                //xPos += 5;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write();
            ANA_PEN.Value = 0;
            ANA_PEN.Write();
            ANA_MUX.Value = 0;
            ANA_MUX.Write();
            Log.WriteLine("End TEST_ISEN Sweep Register");
        }

        private void TEST_RunActiveLoad()
        {
            double Irect;
            int xPos = 1, yPos = 1;
            string sheet_name;

            RegisterItem SEL = Parent.RegMgr.GetRegisterItem("AL_SEL");
            RegisterItem PEN = Parent.RegMgr.GetRegisterItem("AL_PEN");
            RegisterItem ILOAD = Parent.RegMgr.GetRegisterItem("AL_ILOAD_SET<3:0>");
            RegisterItem ISLOPE = Parent.RegMgr.GetRegisterItem("AL_ISLOPE_TRIM<4:0>");
            RegisterItem IOFFSET = Parent.RegMgr.GetRegisterItem("AL_IOFFSET_TRIM<4:0>");


            Log.WriteLine("Start Active Load Sweep Register");

            CAL_CheckInstrument();

            if (PowerSupply0 == null || PowerSupply0.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C");

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 6.0");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif
            JLcLib.Delay.Sleep(200);

            SEL.Value = 1;
            SEL.Write();
            PEN.Value = 1;
            PEN.Write();
            sheet_name = DateTime.Now.ToString("MMddHHmmss") + "_AL";
            Parent.xlMgr.Sheet.Add(sheet_name);

            Parent.xlMgr.Cell.Write(xPos, yPos, "Active Load = on");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "VRECT = 6V");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "IRECT = 0mA");
            xPos = 1;
            for (uint i = 0; i < 32; i++) // AL_ILOAD_SET
            {
                ILOAD.Value = i;
                ILOAD.Write();
                Parent.xlMgr.Cell.Write(xPos, yPos + 1, "AL_ILOAD_SET = " + i.ToString());
                Parent.xlMgr.Cell.Write(xPos, yPos + 2, "AL_ISLOPE_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 1, yPos + 1, "AL_IOFFSET_TRIM");
                for (uint j = 0; j < 32; j++) // AL_ISLOPE_TRIM
                {
                    ISLOPE.Value = j;
                    ISLOPE.Write();
                    Parent.xlMgr.Cell.Write(xPos, (int)(yPos + 3 + j), j.ToString());
                    for (uint k = 0; k < 32; k++) // AL_IOFFSET_TRIM
                    {
                        IOFFSET.Value = k;
                        IOFFSET.Write();

                        if (j == 0) Parent.xlMgr.Cell.Write((int)(xPos + 1 + k), (int)(yPos + 2 + j), k.ToString());

#if (POWER_SUPPLY_E36313A)
                        Irect = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR?"));
#else
                        Irect = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR?"));
#endif
                        Parent.xlMgr.Cell.Write((int)(xPos + 1 + k), (int)(yPos + 3 + j), (Irect * 1000).ToString());
                    }
                }
                yPos += 36;
            }
            Log.WriteLine("End Active Load Sweep Register");
        }

        private void TEST_RunLOAD()
        {
            Log.WriteLine("Start Loader Sweep");

            CAL_CheckInstrument();

            if (ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("INP 1");
            ElectronicLoad.Write("CURR 0.3");
            System.Threading.Thread.Sleep(1);
            ElectronicLoad.Write("CURR 0.5");
            System.Threading.Thread.Sleep(1);
            ElectronicLoad.Write("CURR 0.8");
            System.Threading.Thread.Sleep(1);
            ElectronicLoad.Write("CURR 1");
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);

            for (double i = 1000; i >= 0; i--)
            {
                ElectronicLoad.Write("CURR " + (i / 1000).ToString());
                //System.Threading.Thread.Sleep(1);

            }
            ElectronicLoad.Write("INP 0");
            ElectronicLoad.Write("CURR 0");
            Log.WriteLine("End Loader Sweep");
        }

        private void TEST_RunADCCODE()
        {
            int xPos = 1, yPos = 1;
            string sheet_name;
            double Vrect, Vntc;

            Log.WriteLine("Start TEST_ADC_CODE");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || PowerSupply0 == null || PowerSupply0.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : GPIO\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : GPIO\n- CH3 : N/C\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("CURR 1");
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT 0");
            PowerSupply0.Write("CURR 0.5");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:3)");
#else
            PowerSupply0.Write("OUTP ON");
#endif
            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 0"); // offset 0V
            // Cannel.2 use for GPIO
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 0"); // offset 0V

            JLcLib.Delay.Sleep(200);

            sheet_name = DateTime.Now.ToString("MMddHHmmss_") + "ADC_CODE";
            Parent.xlMgr.Sheet.Add(sheet_name);

            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "Vrect");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "Vrect_Scope");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "Vrect_Code_AVG");
            Parent.xlMgr.Cell.Write(xPos + 3, yPos, "Vrect_Code_Min");
            Parent.xlMgr.Cell.Write(xPos + 4, yPos, "Vrect_Code_Max");

            PowerSupply0.Write("INST:NSEL 2");
            for (double Vr = 3.7; Vr <= 20.1; Vr += 0.1) // Vrect Power Supply (V) 3.7~20
            {
                yPos++;
                Oscilloscope.Write(":CHAN1:OFFS " + Vr.ToString("F2"));
                PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

                Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vr.ToString("F2") + "V"));
                Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vrect * 1000).ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_VRECT, 255);
                Parent.xlMgr.Cell.Write(xPos + 2, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_VRECT_MIN, 255);
                Parent.xlMgr.Cell.Write(xPos + 3, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_VRECT_MAX, 255);
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, Vrect.ToString("F3"));
            }

            xPos = 7;
            yPos = 1;

            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "Vntc");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "Vntc_Scope");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "Vntc_Code_AVG");
            Parent.xlMgr.Cell.Write(xPos + 3, yPos, "Vntc_Code_Min");
            Parent.xlMgr.Cell.Write(xPos + 4, yPos, "Vntc_Code_Max");

            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("INST:NSEL 3");
            for (double Vr = 0; Vr <= 3.3; Vr += 0.05) // Vrect Power Supply (V) 0~3.3
            {
                yPos++;
                Oscilloscope.Write(":CHAN2:OFFS " + Vr.ToString("F2"));
                PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

                Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vr.ToString("F2") + "V"));
                Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vrect * 1000).ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_NTC, 255);
                Parent.xlMgr.Cell.Write(xPos + 2, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_NTC_MIN, 255);
                Parent.xlMgr.Cell.Write(xPos + 3, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_NTC_MAX, 255);
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, Vrect.ToString("F3"));
            }

            Log.WriteLine("End TEST_ADC_CODE");
        }

        private void TEST_RunVout()
        {
            double Vrect, Vout;
            int xPos = 3, yPos = 5;

            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]

            Log.WriteLine("Start TEST_Vout");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;
            if (Parent.xlMgr.Sheet.Select("VoutTest") == false)
                Parent.xlMgr.Sheet.Add("VoutTest");

            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : N/C\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 8");
            PowerSupply0.Write("CURR 1");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 8"); // offset 8V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 1"); // 1V/div
            Oscilloscope.Write(":CHAN2:OFFS 4"); // offset 4V

            JLcLib.Delay.Sleep(200);

            for (uint ct = 0; ct <= 1; ct++)
            {
                VOUT_SET_RANGE.Value = ct;
                VOUT_SET_RANGE.Write();
                yPos = 5;
                for (uint set = 0; set < 101; set++)
                {
                    TRIM_VOUT_SET.Value = set;
                    TRIM_VOUT_SET.Write();
                    JLcLib.Delay.Sleep(100);

                    Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                    Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));

                    Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vrect * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vout * 1000.0).ToString());

                    yPos++;
                }
                xPos += 2;
            }
            Log.WriteLine("End TEST_Vout");
        }

        private void TEST_RunVout_wpc()
        {
            RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");
            RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");
            RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");
            RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");

            Log.WriteLine("Start TEST_Vout_WPC");

            // LDO_SET 3700mV
            CommandReg.Value = 0x08 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 3700);
            StartCommandProcessing(CommandReg);
            JLcLib.Delay.Sleep(500);

            for (int i = 0; i < 1; i++)
            {
                for (double vr = 3.7; vr < 6.5; vr += 0.1)
                {
                    // SET_TEST_VOUT
                    CommandReg.Value = 0x08 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                    //JLcLib.Delay.Sleep(10);
                }
                for (double vr = 6.5; vr >= 3.7; vr -= 0.1)
                {
                    // SET_TEST_VOUT
                    CommandReg.Value = 0x08 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                    //JLcLib.Delay.Sleep(10);
                }
            }
            JLcLib.Delay.Sleep(500);
            // LDO_SET 5000mV
            CommandReg.Value = 0x08 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 5000);
            StartCommandProcessing(CommandReg);

            Log.WriteLine("End TEST_Vout_WPC");
        }

        private void TEST_RunVrect_wpc()
        {
            int iVal = 0;
            RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");
            RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");
            RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");
            RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");

            Log.WriteLine("Start TEST_Vrect_WPC");

            // ENTER_TEST_PT
            CommandReg.Value = 0x01 + 0x00;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 0);
            StartCommandProcessing(CommandReg);

            // LDO_SET 6500mV
            CommandReg.Value = 0x08 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 6500);
            StartCommandProcessing(CommandReg);

            // SET_TEST_VRECT 3.7V
            CommandReg.Value = 0x01 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, (int)(3.7 * 1000));
            StartCommandProcessing(CommandReg);
            JLcLib.Delay.Sleep(500);

            for (int i = 0; i < 3; i++)
            {
                for (double vr = 3.7; vr <= 11.5; vr += 0.5)
                {
                    // SET_TEST_VRECT
                    CommandReg.Value = 0x01 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                }
                for (double vr = 12; vr >= 3.7; vr -= 0.5)
                {
                    // SET_TEST_VRECT
                    CommandReg.Value = 0x01 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                }
            }
            JLcLib.Delay.Sleep(500);

            // EXIT_TEST_PT
            CommandReg.Value = 0x01 + 0x01;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 0);
            StartCommandProcessing(CommandReg);

            Log.WriteLine("End TEST_Vrect_WPC");
        }
        #endregion Chip test methods

        #region Calibration methods
        Button button_RunCal0 = null;
        private void EQUUS_RunCalibration()
        {
            Log.WriteLine("Run Calibratoin!!");
            CAL_CheckInstrument();
            CAL_InitInstrument();
            JLcLib.Delay.Sleep(2000);
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
            PowerSupply1.Write("OUTP OFF");
#endif
            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : GPIO\n\nOscilloscope\n- CH1 : LDO5P0\n- CH2 : LDO1P8\n- CH3 : VBAT_SENS\n- CH4 : INT_A");
            if (!IsRunCal) return;
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:3)");
#else
            PowerSupply0.Write("OUTP ON");
            PowerSupply1.Write("OUTP ON");
#endif
            System.Threading.Thread.Sleep(100);
            // BGR calibration
            CAL_RunBGR();
            if (!IsRunCal) return;
            CAL_RunLDO5P0();
            if (!IsRunCal) return;
            CAL_RunLDO1P5();
            if (!IsRunCal) return;
            CAL_RunADC_LDO3P3();
            if (!IsRunCal) return;
            CAL_RunFSK_LDO3P3();
            if (!IsRunCal) return;
            CAL_RunLDO1P8();
            if (!IsRunCal) return;
            CAL_RunOSC();
            if (!IsRunCal) return;
            MessageBox.Show("Change Oscilloscope CH2 from LDO1P8 to VOUT");
            if (!IsRunCal) return;
            CAL_RunVRECT();
            if (!IsRunCal) return;
            CAL_RunVOUT();
            if (!IsRunCal) return;
            CAL_RunADC();
            if (!IsRunCal) return;
            CAL_RunISENVREF();
            if (!IsRunCal) return;
            CAL_RunVILIM();
            if (!IsRunCal) return;
#if false
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
            PowerSupply1.Write("OUTP OFF");
#endif
            MessageBox.Show("Change Power Supply CH1 and CH2");
#endif
            if (!IsRunCal) return;
            CAL_RunIOUT();
            Log.WriteLine("Done Calibratoin!!");
        }

        private void EQUUS_StopCalibration()
        {
            IsRunCal = false;
            button_RunCal0.Text = "RUN CAL";
        }

        private void CAL_CheckInstrument()
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
                    case JLcLib.Instrument.InstrumentTypes.OscilloScope1:
                        if (Oscilloscope == null)
                            Oscilloscope = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (Oscilloscope.IsOpen == false)
                            Oscilloscope.Open();
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

        private void CAL_InitInstrument()
        {
            if (PowerSupply0 != null && PowerSupply0.IsOpen)
            {
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
                PowerSupply0.Write("OUTP OFF,(@1:3)");
#else // 2port Power supply
                PowerSupply0.Write("OUTP 0");
                PowerSupply1.Write("OUTP 0");
#endif
                PowerSupply0.Write("INST:NSEL 1");// AP1P8 power supply
                PowerSupply0.Write("VOLT 1.8");
                PowerSupply0.Write("CURR 0.3");
                PowerSupply0.Write("INST:NSEL 2"); // VRECT power supply
                PowerSupply0.Write("VOLT 6.0");
                PowerSupply0.Write("CURR 2.0");
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
                PowerSupply0.Write("INST:NSEL 3");
                PowerSupply0.Write("VOLT 0.3");
                PowerSupply0.Write("CURR 0.3");
#else // 2port Power supply
                PowerSupply1.Write("INST:NSEL 1");
                PowerSupply1.Write("VOLT 0.3");
                PowerSupply1.Write("CURR 0.2");
#endif
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
                PowerSupply0.Write("OUTP ON,(@1:3)");
#else // 2port Power supply
                PowerSupply0.Write("OUTP 1");
                PowerSupply1.Write("OUTP 1");
#endif
            }
            if (Oscilloscope != null && Oscilloscope.IsOpen)
            {
                //Oscilloscope.Write(":TIM:SCAL 2E-4"); // 200usec/div, Scale / 1div
                Oscilloscope.Write(":TIM:SCAL 1E-7"); // 100nsec/div, Scale / 1div
                Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
                Oscilloscope.Write(":TIM:DEL 0"); // delay
                // Cannel.1 use for LDO 5p0
                Oscilloscope.Write(":CHAN1:DISP 1");
                Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
                Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 6V
                // Cannel.2 use for LDO 1p8, VOUT
                Oscilloscope.Write(":CHAN2:DISP 1");
                Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 2V/div
                Oscilloscope.Write(":CHAN2:OFFS 1.8"); // offset 1V
                // Cannel.3 use for ANA_TEST
                Oscilloscope.Write(":CHAN3:DISP 1");
                Oscilloscope.Write(":CHAN3:SCAL 2E-1"); // 200mV/div
                Oscilloscope.Write(":CHAN3:OFFS 1.2"); // offset 1V
                // Cannel.4 use for OSC test
                Oscilloscope.Write(":CHAN4:DISP 1");
                Oscilloscope.Write(":CHAN4:SCAL 1"); // 1V/div
                Oscilloscope.Write(":CHAN4:OFFS 3"); // offset 1V
                Oscilloscope.Write(":TRIG:MODE EDGE"); // trigger mode edge
                Oscilloscope.Write(":TRIG:EDGE:SOUR CHAN4"); // triger source 4
                Oscilloscope.Write(":TRIG:EDGE:LEV 9E-1"); // trigger level 900mV
            }
            if (ElectronicLoad != null && ElectronicLoad.IsOpen)
            {
#if (DC_LOADER_KIKUSUI) // KIKUSUI
                ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
                ElectronicLoad.Write("SYST:REM");
                ElectronicLoad.Write("FUNC CURR");
#endif
                ElectronicLoad.Write("INP 0");
                ElectronicLoad.Write("CURR 0");
                System.Threading.Thread.Sleep(10);
                ElectronicLoad.Write("INP 1");
            }
        }

        private void CAL_RunBGR()
        {
            double Result = 0.0;
            double Target = 1.22, Margin = 0.005;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "BGR";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR<4:0>"); // 0x91[4:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 0; // BGR
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_BGR.Read();
            TRIM_BGR.Value = 0;
            TRIM_BGR.Write();
            for (uint BitPos = 10; BitPos < 26; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_BGR.Value = BitPos;
                TRIM_BGR.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_BGR.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_BGR.Value = (byte)reg_value;
            TRIM_BGR.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x24] = (byte)TRIM_BGR.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x24] = (byte)0xFF;
            }
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunLDO5P0()
        {
            double Result = 0.0;
            double Target = 5, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "LDO5P0";
            RegisterItem TRIM_LDO5P0 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO5P0<3:0>"); // 0x94[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            TRIM_LDO5P0.Read();
            TRIM_LDO5P0.Value = 0;
            TRIM_LDO5P0.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO5P0.Value = BitPos;
                TRIM_LDO5P0.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO5P0.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO5P0.Value = reg_value;
            TRIM_LDO5P0.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x28] = (byte)TRIM_LDO5P0.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x28] = (byte)0xFF;
            }
        }

        private void CAL_RunLDO1P5()
        {
            double Result = 0.0;
            double Target = 1.5, Margin = 0.01;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.02; // pass/fail limit
            string Name = "LDO1P5";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_LDO1P5 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO1P5<3:0>"); // 0x9C[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 1.5"); // offset 1.5V
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 4; // LDO1P5
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_LDO1P5.Read();
            TRIM_LDO1P5.Value = 0x08;
            TRIM_LDO1P5.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO1P5.Value = BitPos;
                TRIM_LDO1P5.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO1P5.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO1P5.Value = reg_value;
            TRIM_LDO1P5.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x2C] = (byte)TRIM_LDO1P5.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x2C] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write(); // Write 0xBC
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunADC_LDO3P3()
        {
            double Result = 0.0;
            double Target = 3.3, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "ADC_LDO3P3";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_ADC_LDO3P3<3:0>"); // 0x96[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 3.3"); // offset 3.3V
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 6; // ADC LDO 3.3
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_LDO3P3.Read();
            TRIM_LDO3P3.Value = 0x08;
            TRIM_LDO3P3.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO3P3.Value = BitPos;
                TRIM_LDO3P3.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO3P3.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO3P3.Value = reg_value;
            TRIM_LDO3P3.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x30] = (byte)TRIM_LDO3P3.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x30] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write(); // Write 0xBC
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunFSK_LDO3P3()
        {
            double Result = 0.0;
            double Target = 3.3, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "FSK_LDO3P3";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_FSK_LDO3P3<3:0>"); // 0x98[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 3.3"); // offset 3.3V
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 5; // FSK LDO 3.3
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_LDO3P3.Read();
            TRIM_LDO3P3.Value = 0x08;
            TRIM_LDO3P3.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO3P3.Value = BitPos;
                TRIM_LDO3P3.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO3P3.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO3P3.Value = reg_value;
            TRIM_LDO3P3.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x34] = (byte)TRIM_LDO3P3.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x34] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write(); // Write 0xBC
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunLDO1P8()
        {
            double Result = 0.0;
            double Target = 1.8, Margin = 0.04;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "LDO1P8";
            RegisterItem TRIM_LDO1P8 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO18<3:0>"); // 0x9A[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            TRIM_LDO1P8.Read();
            TRIM_LDO1P8.Value = 0x08;
            TRIM_LDO1P8.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO1P8.Value = BitPos;
                TRIM_LDO1P8.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO1P8.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO1P8.Value = reg_value;
            TRIM_LDO1P8.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x38] = (byte)TRIM_LDO1P8.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x38] = (byte)0xFF;
            }
        }

        private void CAL_RunOSC()
        {
            double Result = 0.0;
            double Target = 16000000, Margin = 100000;
            double min = 20000000.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "OSC";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_OSC = Parent.RegMgr.GetRegisterItem("PU_TRIM_FOSC_16M<7:0>"); // 0x92[7:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
            DIG_SEL.Read();
            DIG_SEL.Value = 0; // OSC
            ANA_SEL.Value = 0; // BGR
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 1;
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_OSC.Read();
            TRIM_OSC.Value = 0;
            TRIM_OSC.Write();
            for (uint BitPos = 10; BitPos < 21; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_OSC.Value = BitPos;
                TRIM_OSC.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
                Result += double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
                Result /= 2;
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_OSC.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_OSC.Value = reg_value;
            TRIM_OSC.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
            Result += double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
            Result /= 2;
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x3C] = (byte)TRIM_OSC.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x3C] = 0xFF;
            }
            DIG_MUX_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunISENVREF()
        {
            double Vrect = 4.9;
            double Result = 0.0;
            double Target = 300, Margin = 10;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.05; // pass/fail limit
            string Name = "ML_ISEN_VREF";
            RegisterItem TRIM_ISEN = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>"); // 0xA2[7:4]
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>"); // 0xA1[7:4]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));

            ElectronicLoad.Write("CURR 0.1");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            TRIM_VILIM.Read();
            TRIM_VILIM.Value = 0;
            TRIM_VILIM.Write();

            TRIM_ISEN.Read();
            TRIM_ISEN.Value = 0x07;
            TRIM_ISEN.Write();
            for (uint BitPos = 0; BitPos < 5; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_ISEN.Value = BitPos;
                TRIM_ISEN.Write();
                JLcLib.Delay.Sleep(50);
                Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_ISEN.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_ISEN.Value = reg_value;
            TRIM_ISEN.Write();
            JLcLib.Delay.Sleep(50);
            Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
#if false
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)TRIM_ISEN.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)0xFF;
            }
#else
            if (Result >= 0 && Result <= 800)
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)TRIM_ISEN.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)0xFF;
            }
#endif
            ElectronicLoad.Write("CURR 0");
        }

        private void CAL_RunVILIM()
        {
            double Vrect = 5.3;
            double Result = 0.0;
            double Target = 855, Margin = 10;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "ML_VILIM";
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>"); // 0xA1[7:4]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));

            ElectronicLoad.Write("CURR 0.6");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            TRIM_VILIM.Read();
            TRIM_VILIM.Value = 0x07;
            TRIM_VILIM.Write();
            for (uint BitPos = 0; BitPos < 10; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_VILIM.Value = BitPos;
                TRIM_VILIM.Write();
                JLcLib.Delay.Sleep(50);
                Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_VILIM.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_VILIM.Value = reg_value;
            TRIM_VILIM.Write();
            JLcLib.Delay.Sleep(50);
            Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
#if false
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)TRIM_VILIM.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)0xFF;
            }
#else
            if (Result >= 700 && Result <= 1300)
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)TRIM_VILIM.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)0xFF;
            }
#endif
            ElectronicLoad.Write("CURR 0");
        }

        private void CAL_RunVRECT()
        {
            double LowVrect = 4, HighVrect = 8;
            int LowVrectCode, HighVrectCode;
            short Slope, Const;
            string Name = "VRECT";

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            // 4V low VRECT calibration
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + LowVrect.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            LowVrectCode = RunTest(TEST_ITEMS.ADC_GET_VRECT, 1023); // 1024 average
            Log.WriteLine(Name + ":LOW:" + LowVrect.ToString("F1") + ":" + LowVrectCode.ToString("F3"));
            if (!IsRunCal)
                return;
            // 8V high VRECT calibration
            PowerSupply0.Write("VOLT " + HighVrect.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            HighVrectCode = RunTest(TEST_ITEMS.ADC_GET_VRECT, 1023); // 1024 average
            Log.WriteLine(Name + ":HIGH:" + HighVrect.ToString("F1") + ":" + HighVrectCode.ToString("F3"));

            Slope = (short)((HighVrect - LowVrect) * 1000 * 1024 / (HighVrectCode - LowVrectCode));
            Const = (short)(LowVrect * 1000 - (Slope * LowVrectCode) / 1024);

            if ((Slope >= 4250 && Slope <= 5750) && (Const >= -512 && Const <= 512))
            {
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 5000;
                Const = 0;
            }
            CalibrationData[0x48] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x49] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x4C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x4D] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunVOUT()
        {
            double Result, Vrect = 8.0;
            double[] TargetVoutValues = new double[2] { 3.5, 5.0 };
            double[] VoutValues = new double[2];
            double[] VoutSetValues = new double[2];
            double[] Limits = new double[2] { TargetVoutValues[0] * 0.03, TargetVoutValues[1] * 0.03 };
            int[] VoutCodes = new int[2];
            double Margin = 0.015;
            double min = 4095.0;
            uint reg_value = 0;
            short Slope, Const;
            string Name = "VOUT";
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN"); // 0x9E[3:3]
            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]

            // Sets 8V VRECT
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            OCL_EN.Read();
            OCL_EN.Value = 0;
            OCL_EN.Write();

            RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //ElectronicLoad.Write("CURR 0.1");
            //ElectronicLoad.Write("INP 1");

            VOUT_SET_RANGE.Read();
            TRIM_VOUT_SET.Read();
            for (uint i = 0; i < 2; i++)
            {
                min = 4095.0;

                VOUT_SET_RANGE.Value = i;
                VOUT_SET_RANGE.Write();

                Oscilloscope.Write(":CHAN2:SCAL 2"); // 2V/div
                Oscilloscope.Write(":CHAN2:OFFS 8"); // offset 8V
                TRIM_VOUT_SET.Value = 0x40;
                TRIM_VOUT_SET.Write();
                JLcLib.Delay.Sleep(200);

                for (uint BitPos = 85; BitPos < 100; BitPos++)
                {
                    if (!IsRunCal)
                        return;
                    if (i == 0)
                    {
                        TRIM_VOUT_SET.Value = BitPos;
                    }
                    else
                    {
                        TRIM_VOUT_SET.Value = BitPos - 40;
                    }
                    TRIM_VOUT_SET.Write();
                    JLcLib.Delay.Sleep(50);
                    Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                    if (Math.Abs(TargetVoutValues[i] - Result) < min)
                    {
                        min = Math.Abs(TargetVoutValues[i] - Result);
                        reg_value = TRIM_VOUT_SET.Value;
                    }
                    Log.WriteLine(Name + i.ToString() + ":" + BitPos.ToString("D2") + ":" + TRIM_VOUT_SET.Value.ToString("D3") + ":" + Result.ToString("F3"));
                }
                TRIM_VOUT_SET.Value = reg_value;
                TRIM_VOUT_SET.Write();
                JLcLib.Delay.Sleep(50);
                VoutValues[i] = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                VoutSetValues[i] = TRIM_VOUT_SET.Value;
                VoutCodes[i] = RunTest(TEST_ITEMS.ADC_GET_VOUT, 1023); // 1024 average

                if (VoutValues[i] >= TargetVoutValues[i] - Limits[i] && VoutValues[i] <= TargetVoutValues[i] + Limits[i])
                    Log.WriteLine(Name + "_LVL:PASS:" + reg_value.ToString() + ":" + VoutValues[i].ToString() + ":" + VoutCodes[i].ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
                else
                    Log.WriteLine(Name + "_LVL:FAIL:" + reg_value.ToString() + ":" + VoutValues[i].ToString() + ":" + VoutCodes[i].ToString(), Color.Coral, Log.RichTextBox.BackColor);
            }
            //RunTest(TEST_ITEMS.LDO_TURNOFF, 0);
            CalibrationData[0xB0] = (byte)TRIM_VOUT_SET.Value;
            //ElectronicLoad.Write("CURR 0");

            // Vout voltage calibration
            Slope = (short)((VoutValues[1] - VoutValues[0]) * 1000 * 1024 / (VoutCodes[1] - VoutCodes[0]));
            Const = (short)(VoutValues[0] * 1000 - (Slope * VoutCodes[0]) / 1024);
            if ((Slope >= 2000 && Slope <= 3000) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 2500;
                Const = 0;
            }
            CalibrationData[0x58] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x59] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x5C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x5D] = (byte)((Const >> 8) & 0xFF);

            // Vout setting register calibration
            Slope = (short)((VoutSetValues[1] + 100 - VoutSetValues[0]) * 8192 / ((TargetVoutValues[1] - TargetVoutValues[0]) * 1000));
            Const = (short)(VoutSetValues[0] - (Slope * TargetVoutValues[0] * 1000) / 8192);
            if ((Slope >= 300 && Slope <= 400) && (Const >= -200 && Const <= 200))
                Log.WriteLine(Name + "_SET:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_SET:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 341;
                Const = -82;
            }
            CalibrationData[0x50] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x51] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x54] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x55] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunADC()
        {
            double LowVolt = 0.3, HighVolt = 1.5;
            int LowVoltCode, HighVoltCode;
            short Slope, Const;
            string Name = "ADC";

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT " + LowVolt.ToString("F1"));
            PowerSupply0.Write("OUTP 1");
#else // 2 port power supply
            PowerSupply1.Write("INST:NSEL 1");
            PowerSupply1.Write("VOLT " + LowVolt.ToString("F1"));
            PowerSupply1.Write("OUTP 1");
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            LowVoltCode = RunTest(TEST_ITEMS.ADC_GET_NTC, 1023); // 1024 average
            Log.WriteLine(Name + ":LOW:" + LowVolt.ToString("F1") + ":" + LowVoltCode.ToString("F3"));
            if (!IsRunCal)
                return;
            // 8V high VRECT calibration
#if (POWER_SUPPLY_E36313A) // 3 port power supply
            PowerSupply0.Write("VOLT " + HighVolt.ToString("F1"));
#else // 2 power power supply
            PowerSupply1.Write("VOLT " + HighVolt.ToString("F1"));
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            HighVoltCode = RunTest(TEST_ITEMS.ADC_GET_NTC, 1023); // 1024 average
            Log.WriteLine(Name + ":HIGH:" + HighVolt.ToString("F1") + ":" + HighVoltCode.ToString("F3"));

            Slope = (short)((HighVolt - LowVolt) * 1000 * 1024 / (HighVoltCode - LowVoltCode));
            Const = (short)(LowVolt * 1000 - (Slope * LowVoltCode) / 1024);

            if ((Slope >= 750 && Slope <= 900) && (Const >= -64 && Const <= 64))
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 825;
                Const = 0;
            }
            CalibrationData[0x60] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x61] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x64] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x65] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunIOUT()
        {
            double Result;
            double[] VrectValues = new double[2] { 4.9, 5.5 }; // unit V
            double[] IoutValues = new double[3] { 0, 0.15, 1 }; // unit A
            int[,] ISenCodes = new int[IoutValues.Length, VrectValues.Length];
            double Margin = 0.03;
            short Slope, Const;
            string Name = "IOUT";
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN"); // 0x9E[3:3]
            RegisterItem TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR<4:0>"); // 0x91[4:0]
            RegisterItem TRIM_LDO5P0 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO5P0<3:0>"); // 0x94[3:0]
            RegisterItem TRIM_LDO1P5 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO1P5<3:0>"); // 0x9C[3:0]
            RegisterItem TRIM_ADC_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_ADC_LDO3P3<3:0>"); // 0x96[3:0]
            RegisterItem TRIM_FSK_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_FSK_LDO3P3<3:0>"); // 0x98[3:0]
            RegisterItem TRIM_LDO1P8 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO18<3:0>"); // 0x9A[3:0]
            RegisterItem TRIM_OSC = Parent.RegMgr.GetRegisterItem("PU_TRIM_FOSC_16M<7:0>"); // 0x92[7:0]
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>"); // 0xA1[7:4]
            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]
            RegisterItem TRIM_ISEN = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>"); // 0xA21[7:4]

            // Sets 8V VRECT
            /*
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5.3");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif
            */
            PowerSupply0.Write("INST:NSEL 2");

            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            //OCL_EN.Value = 0;
            //OCL_EN.Write();

            RunTest(TEST_ITEMS.LDO_TURNON, 0);
            // Set 5V VOUT
            /*TRIM_BGR.Value = CalibrationData[0x24];
            TRIM_BGR.Write();
            TRIM_LDO5P0.Value = CalibrationData[0x28];
            TRIM_LDO5P0.Write();
            TRIM_LDO1P5.Value = CalibrationData[0x2C];
            TRIM_LDO1P5.Write();
            TRIM_ADC_LDO3P3.Value = CalibrationData[0x30];
            TRIM_ADC_LDO3P3.Write();
            TRIM_FSK_LDO3P3.Value = CalibrationData[0x34];
            TRIM_FSK_LDO3P3.Write();
            TRIM_LDO1P8.Value = CalibrationData[0x38];
            TRIM_LDO1P8.Write();
            TRIM_OSC.Value = CalibrationData[0x3C];
            TRIM_OSC.Write();
            TRIM_VILIM.Value = CalibrationData[0x44];
            TRIM_VILIM.Write();
            VOUT_SET_RANGE.Value = 1;
            VOUT_SET_RANGE.Write();
            TRIM_VOUT_SET.Value = CalibrationData[0xB0];
            TRIM_VOUT_SET.Write();
            TRIM_ISEN.Read();
            TRIM_ISEN.Value = CalibrationData[0x40];
            TRIM_ISEN.Write();*/

            for (int i = 0; i < IoutValues.Length; i++) // Zero(0)/Low(1)/High(2) current
            {
                ElectronicLoad.Write("CURR " + IoutValues[i].ToString("F2"));
                for (int j = 0; j < VrectValues.Length; j++) // Low(0)/High(1) VRECT
                {
                    PowerSupply0.Write("VOLT " + VrectValues[j].ToString("F1")); // Set VRECT voltage
                    JLcLib.Delay.Sleep(500);
                    ISenCodes[i, j] = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
                    Log.WriteLine(Name + ":" + IoutValues[i].ToString("F2") + ":" + VrectValues[j].ToString("F1") + ":" + ISenCodes[i, j].ToString());
                }
            }
            //RunTest(TEST_ITEMS.LDO_TURNOFF, 0);
            ElectronicLoad.Write("CURR 0");

            // Store ADC Iout codes zero load current and difference between high and low Vrect at high load current
            //Const = (short)ISenCodes[0, 1]; // ADC code of zero load current, high Vrect
            //CalibrationData[0x70] = (byte)((Const >> 0) & 0xFF);
            //CalibrationData[0x71] = (byte)((Const >> 8) & 0xFF);
            //Const = (short)ISenCodes[1, 1]; // ADC code of low load current, high Vrect
            //CalibrationData[0x74] = (byte)((Const >> 0) & 0xFF);
            //CalibrationData[0x75] = (byte)((Const >> 8) & 0xFF);
            Const = (short)IoutValues[0]; // Low load current
            CalibrationData[0x78] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x79] = (byte)((Const >> 8) & 0xFF);
            Log.WriteLine(Name + "_Low:PASS:" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            Const = (short)(ISenCodes[2, 0] - ISenCodes[2, 1]);
            CalibrationData[0x7C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x7D] = (byte)((Const >> 8) & 0xFF);
            Log.WriteLine(Name + "_Diff:PASS:" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);

            // Calculate Isen0
            Slope = (short)((IoutValues[2] - IoutValues[0]) * 1000 * 1024 / (ISenCodes[2, 0] - ISenCodes[0, 0]));
            Const = (short)(IoutValues[0] * 1000 - (Slope * ISenCodes[0, 0]) / 1024);
            Log.WriteLine(Name + "_0:" + Slope.ToString() + ":" + Const.ToString(), Color.DarkGray, Log.RichTextBox.BackColor);

            if ((Slope >= 500 && Slope <= 1500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_0:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_0:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x80] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x81] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x84] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x85] = (byte)((Const >> 8) & 0xFF);

            // Calculate Isen1
            Slope = (short)((IoutValues[2] - IoutValues[1]) * 1000 * 1024 / (ISenCodes[2, 1] - ISenCodes[1, 1]));
            Const = (short)(IoutValues[1] * 1000 - (Slope * ISenCodes[1, 1]) / 1024);
            Log.WriteLine(Name + "_1:" + Slope.ToString() + ":" + Const.ToString(), Color.DarkGray, Log.RichTextBox.BackColor);
            if ((Slope >= 500 && Slope <= 1500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_1:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_1:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x88] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x89] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x8C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x8D] = (byte)((Const >> 8) & 0xFF);

            // Calculate Isen2
            Slope = (short)((IoutValues[1] - IoutValues[0]) * 1000 * 1024 / (ISenCodes[1, 1] - ISenCodes[0, 1]));
            Const = (short)(IoutValues[0] * 1000 - (Slope * ISenCodes[0, 1]) / 1024);
            if ((Slope >= 500 && Slope <= 1500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_2:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_2:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x70] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x71] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x74] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x75] = (byte)((Const >> 8) & 0xFF);
        }
        #endregion Calibration methods
    }

    public class TWS : ChipControl
    {
        #region Variable and declaration
        public enum FLASH_STATE
        {
            CMD_READY = (1 << 4),
            ADDR_DATA_READY = (1 << 5),
            READY = (1 << 4) | (1 << 5),
            ADDR_LATCHED = (1 << 6),
            DATA_LATCHED = (1 << 7),
            MAT_BUSY = (1 << 8),
        }

        public enum FW_TARGET
        {
            NV_MEM = 0,
            RAM = 1,
        }

        public enum COMBOBOX_ITEMS
        {
            TEST,
            AUTO,
            CAL,
            FW,
        }

        public enum TEST_ITEMS
        {
            ENTER_TEST_PT, // 0x01 (FW command)
            EXIT_TEST_PT, // 0x02

            SET_TEST_VRECT, // 0x04

            LDO_TURNON, // 0x08 (FW command)
            LDO_TURNOFF, // 0x09
            LDO_SET, // 0x0A

            ADC_GET_VRECT, // 0x10,
            ADC_GET_VOUT, // 0x11,
            ADC_GET_IOUT, // 0x12,
            ADC_GET_VBAT_SENS, // 0x13,
            ADC_GET_VPTAT, // 0x14,
            ADC_GET_NTC, // 0x15,
            //ADC_GET_BGR, // 0x16
            //ADC_GET_ILIM, // 0x17
            //ADC_GET_REF, // 0x18
            //ADC_GET_V_DL, // 0x19
            //ADC_GET_I_DL, // 0x1A
            ADC_GET_EB1, // 0x1B
            ADC_GET_EB2, // 0x1C

            GET_VRECT_mV, // 0x20,
            GET_VOUT_mV, // 0x21,
            GET_IOUT_mA, // 0x22,
            GET_VBAT_SENS_mV, // 0x23,
            GET_VPTAT_mV, // 0x24,
            GET_NTC_mV, // 0x25,
            GET_IOUT_MIN_mA, // 0x26, // test only
            GET_IOUT_MAX_mA, // 0x27, // test only

            /* 0x40 ~ 0x4F : ADC test MIN */
            ADC_GET_VRECT_MIN,
            ADC_GET_VOUT_MIN,
            ADC_GET_IOUT_MIN,
            ADC_GET_VBAT_SENS_MIN,
            ADC_GET_VPTAT_MIN,
            ADC_GET_NTC_MIN,
            //ADC_GET_BGR_MIN,
            //ADC_GET_ILIM_MIN,
            //ADC_GET_REF_MIN,
            //ADC_GET_V_DL_MIN,
            //ADC_GET_I_DL_MIN,

            /* 0x50 ~ 0x5F : ADC test MIN */
            ADC_GET_VRECT_MAX,
            ADC_GET_VOUT_MAX,
            ADC_GET_IOUT_MAX,
            ADC_GET_VBAT_SENS_MAX,
            ADC_GET_VPTAT_MAX,
            ADC_GET_NTC_MAX,
            //ADC_GET_BGR_MAX,
            //ADC_GET_ILIM_MAX,
            //ADC_GET_REF_MAX,
            //ADC_GET_V_DL_MAX,
            //ADC_GET_I_DL_MAX,

            NUM_TEST_ITEMS,
        }

        public enum AUTO_TEST_ITEMS
        {
            TEST_ISEN, // GUI test function
            TEST_VOUT, // GUI test function
            TEST_VOUT_WPC, // GUI test function
            TEST_VRECT_WPC, // GUI test function
            TEST_ISEN_REG, // GUI test function
            TEST_ACTIVE_LOAD, // GUI test function
            TEST_LOAD_SWEEP, // GUI test function
            TEST_ADC_CODE, // GUI test function
            TEST_JISU,

            NUM_TEST_ITEMS,
        }

        public enum CAL_ITEMS
        {
            CAL_RUN,
            CAL_ERASE,
            CAL_WRITE,
            CAL_READ,

            NUM_TEST_ITEMS,
        }

        public enum FW_DN_ITEMS
        {
            ERASE,
            WRITE,
            READ,

            NUM_TEST_ITEMS,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private FW_TARGET FirmwareTarget { get; set; } = FW_TARGET.NV_MEM;

        public int F1_SlaveAddress { get; private set; } = 0x24;
        public int F2_SlaveAddress { get; } = 0x04;
        public int F3_SlaveAddress { get; } = 0x05;

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        private bool IsRunCal = false;
        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.TEST;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI PowerSupply1 = null;
        JLcLib.Instrument.SCPI Oscilloscope = null;
        JLcLib.Instrument.SCPI ElectronicLoad = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        #endregion Variable and declaration

        public TWS(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "TWS");
            Serial.DataReceived += Serial_DataReceived;
            CalibrationData = new byte[256];

            /* Init test items combo box */
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Serial_DataReceived(object sender, JLcLib.Comn.RcvEventArgs e)
        {
            IsSerialReceivedData = true;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[2];

            SendData[0] = (byte)(Address & 0xFF);
            SendData[1] = (byte)(Data & 0xFF);

            iComn.WriteBytes(SendData, SendData.Length, true);
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[1];
            byte[] RcvData = new byte[1];

            SendData[0] = (byte)(Address & 0xFF);

            iComn.WriteBytes(SendData, SendData.Length, true);
            RcvData = iComn.ReadBytes(RcvData.Length);

            return RcvData[0];
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

        #region SKAI TWS host control methods
        public void SetHostMode()
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56); // MGC
            SendBytes.Add(0x81); // Command (CH_HOST)
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        public void ResetSystem()
        {
            byte[] SendBytes = new byte[2] { 0x56, 0xF0 };

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        int ReadHostMemory(uint Address)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56); // MGC
            SendBytes.Add(0xB0); // Command (Set read address)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            byte[] RcvBytes = I2C.ReadBytes(6);
            int Data = -1;
            if (RcvBytes != null && RcvBytes.Length >= 6 && RcvBytes[0] == 0x56 && RcvBytes[1] == 0x01) // Check MGC and STA
                Data = (RcvBytes[2] << 24) | (RcvBytes[3] << 16) | (RcvBytes[4] << 8) | RcvBytes[5];

            I2C.Config.SlaveAddress = F1_SlaveAddress;
            return Data;
        }

        byte[] ReadMemory(uint Address, int Length)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56); // MGC
            SendBytes.Add(0xB0); // Command (Set write data)
            //SendBytes.Add(0xB0); // Command (Set read address)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            byte[] RcvBytes = I2C.ReadBytes(Length + 2);
#if true
            // if 32bit operation
            List<byte> Data = new List<byte>();
            for (int i = 2; i < RcvBytes.Length; i += 4)
            {
                Data.Add(RcvBytes[3 + i]);
                Data.Add(RcvBytes[2 + i]);
                Data.Add(RcvBytes[1 + i]);
                Data.Add(RcvBytes[0 + i]);
            }
            RcvBytes = Data.ToArray();
            // endif 32bit operation
#endif
            I2C.Config.SlaveAddress = F1_SlaveAddress;

            return RcvBytes;
        }

        void WriteHostMemory(uint Address, uint Data)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56);
            SendBytes.Add(0xA0); // Command (Set write data)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        void WriteHostMemory(int Address, byte[] Data)
        {
            List<byte> SendBytes = new List<byte>();

            if (Data.Length < 4)
                return;

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56);
            SendBytes.Add(0xA0); // Command (Set write data)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            SendBytes.Add(Data[3]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[0]);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        void WriteMemory(uint Address, byte[] Data)
        {
            List<byte> SendBytes = new List<byte>();

            F1_SlaveAddress = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = F2_SlaveAddress;

            SendBytes.Add(0x56);
            SendBytes.Add(0xA0); // Command (Set write data)
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));

#if false
            // Normal byte operation
            for (int i = 0; i < Data.Length; i++)
                SendBytes.Add(Data[i]);
#else
            // 32bits double word operation
            for (int i = 0; i < Data.Length; i += 4)
            {
                SendBytes.Add(Data[3 + i]);
                SendBytes.Add(Data[2 + i]);
                SendBytes.Add(Data[1 + i]);
                SendBytes.Add(Data[0 + i]);
            }
#endif
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = F1_SlaveAddress;
        }

        bool WaitFlashState(FLASH_STATE State, int Timeout = 300)
        {
            bool Completed = false;
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();

            sw.Start();
            do
            {
                int Data = ReadHostMemory(0x0FFE0004);
                switch (State)
                {
                    case FLASH_STATE.CMD_READY:
                    case FLASH_STATE.ADDR_DATA_READY:
                        if ((Data & (int)State) == (int)State)
                            Completed = true;
                        break;
                    case FLASH_STATE.READY:
                        if (Data == (int)State)
                            Completed = true;
                        break;
                    case FLASH_STATE.ADDR_LATCHED:
                    case FLASH_STATE.DATA_LATCHED:
                    case FLASH_STATE.MAT_BUSY:
                        if ((Data & (int)State) == 0)
                        {
                            if (State == FLASH_STATE.MAT_BUSY)
                            {
                                if ((Data & 0xFF) == 0x30)
                                    Completed = true;
                            }
                            else
                                Completed = true;
                        }
                        break;
                }
            } while (sw.Elapsed.TotalMilliseconds < Timeout && Completed == false);

            return Completed;
        }
        #endregion SKAI TWS host control methods

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
#if false
            Parent.ChipCtrlButtons[4].Text = "Load";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Load_On_Off;
#else
            Parent.ChipCtrlButtons[4].Text = "ISNS";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += ISNS_Load_sweep;
#endif
            Parent.ChipCtrlButtons[8].Text = "TEST";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Go_To_Menual_Test_Control;

            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Go_To_Auto_Test_Control;

            Parent.ChipCtrlButtons[10].Text = "CAL";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Go_To_CAL_Control;

            Parent.ChipCtrlButtons[11].Text = "FW";
            Parent.ChipCtrlButtons[11].Visible = true;
            Parent.ChipCtrlButtons[11].Click += Go_To_FW_Control;
        }

        private void Load_On_Off(object sender, EventArgs e)
        {
            ElectronicLoad = new JLcLib.Instrument.SCPI(JLcLib.Instrument.InstrumentTypes.ElectronicLoad);
            ElectronicLoad.Open();

            if (ElectronicLoad.IsOpen == false)
            {
                MessageBox.Show("Check Loader");
                return;
            }
            ElectronicLoad.Write("INP 1");
            System.Threading.Thread.Sleep(10);
            ElectronicLoad.Write("INP 0");
        }

        private void ISNS_Load_sweep(object sender, EventArgs e)
        {
            double volt_v;
            int xPos = 1, yPos = 1;
            string sheet_name, time;


            RegisterItem CS_IBIAS = Parent.RegMgr.GetRegisterItem("CS_IBIAS");    //0x31
            RegisterItem CS_GAIN = Parent.RegMgr.GetRegisterItem("CS_GAIN");    //0x32
            RegisterItem CS_IV = Parent.RegMgr.GetRegisterItem("CS_IV");    //0x32

            DigitalMultimeter0 = new JLcLib.Instrument.SCPI(JLcLib.Instrument.InstrumentTypes.DigitalMultimeter0);
            DigitalMultimeter0.Open();

            if (DigitalMultimeter0.IsOpen == false)
            {
                MessageBox.Show("Check DMM");
                return;
            }

            ElectronicLoad = new JLcLib.Instrument.SCPI(JLcLib.Instrument.InstrumentTypes.ElectronicLoad);
            ElectronicLoad.Open();

            if (ElectronicLoad.IsOpen == false)
            {
                MessageBox.Show("Check load");
                return;
            }

            PowerSupply0 = new JLcLib.Instrument.SCPI(JLcLib.Instrument.InstrumentTypes.PowerSupply0);
            PowerSupply0.Open();

            if (PowerSupply0.IsOpen == false)
            {
                MessageBox.Show("Check supply");
                return;
            }

            Log.WriteLine("Start TEST_ISEN");
            //ch1 PVIN
            PowerSupply0.Write("OUTP OFF,(@1)");
            PowerSupply0.Write("VOLT 5,(@1)");
            PowerSupply0.Write("CURR 3,(@1)");
            PowerSupply0.Write("OUTP ON,(@1)");
            //ch2 VIN
            PowerSupply0.Write("OUTP OFF,(@2)");
            PowerSupply0.Write("VOLT 5,(@2)");
            PowerSupply0.Write("CURR 1,(@2)");
            PowerSupply0.Write("OUTP ON,(@2)");

            //ch 3 BST
            PowerSupply0.Write("OUTP OFF,(@3)");
            PowerSupply0.Write("VOLT 10,(@3)");
            PowerSupply0.Write("CURR 0.1,(@3)");
            PowerSupply0.Write("OUTP ON,(@3)");

            DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?");



            CS_IBIAS.Value = 0; CS_IBIAS.Write();
            CS_IV.Value = 0; CS_IV.Write();
            CS_GAIN.Value = 0; CS_GAIN.Write();



#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            time = DateTime.Now.ToString("MMddHHmmss_");
            sheet_name = time + "Iout";
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "VIN(V)");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "LOAD(mA)");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "CS_GAIN Code");
            Parent.xlMgr.Cell.Write(xPos + 3, yPos, "CS_IBIAS Code");
            Parent.xlMgr.Cell.Write(xPos + 4, yPos, "CS_IV Code");
            Parent.xlMgr.Cell.Write(xPos + 5, yPos, "ISEN(V)");

            int yCount = 2;
            int Count = 0;
            for (int GAIN = 0; GAIN < 2; GAIN++)
            {
                CS_GAIN.Value = (uint)GAIN;
                CS_GAIN.Write();

                yPos = 2;
                if (GAIN == 1)
                {
                    xPos = 72;
                    yPos = 2;
                }
                JLcLib.Delay.Sleep(10);
                for (int IBIAS = 0; IBIAS < 32; IBIAS++)
                {
                    CS_IBIAS.Value = (uint)IBIAS;
                    CS_IBIAS.Write();
                    JLcLib.Delay.Sleep(10);

                    for (int loadcnt = 0; loadcnt <= 20; loadcnt++) // 0~2000(mA) Electronic Load
                    //for (int IV = 0; IV < 63; IV++)
                    {

                        double Iload = loadcnt * 100;
                        ElectronicLoad.Write("CURR " + (Iload / 1000).ToString());
                        JLcLib.Delay.Sleep(150);

                        for (int IV = 0; IV < 64; IV++)
                        {
                            CS_IV.Value = (uint)IV;
                            CS_IV.Write();
                            JLcLib.Delay.Sleep(20);
                            volt_v = double.Parse(DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?"));

                            //xPos = 1;
                            //yPos++;

                            Parent.xlMgr.Cell.Write(xPos + 0, yPos + 1 + loadcnt, "5");
                            Parent.xlMgr.Cell.Write(xPos + 1, yPos + 1 + loadcnt, Iload.ToString());
                            Parent.xlMgr.Cell.Write(xPos + 2, yPos + 1 + loadcnt, GAIN.ToString());
                            Parent.xlMgr.Cell.Write(xPos + 3, yPos + 1 + loadcnt, IBIAS.ToString());
                            Parent.xlMgr.Cell.Write(xPos + 4 + IV, yPos, IV.ToString());
                            Parent.xlMgr.Cell.Write(xPos + 4 + IV, yPos + 1 + loadcnt, volt_v.ToString());
                            //yPos++;
                        }
                        ElectronicLoad.Write("CURR 0");
                    }

                    yPos += 25;
                }
            }

            ElectronicLoad.Write("INP 0");
            PowerSupply0.Write("OUTP OFF,(@3)"); //BST 1st off
            PowerSupply0.Write("OUTP OFF,(@1)");
            PowerSupply0.Write("OUTP OFF,(@2)");
            Log.WriteLine("End TEST_ISEN");
        }

        private void Go_To_Menual_Test_Control(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.TEST;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Go_To_Auto_Test_Control(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.AUTO;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)AUTO_TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((AUTO_TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Go_To_CAL_Control(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.CAL;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)CAL_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((CAL_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Go_To_FW_Control(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.FW;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)FW_DN_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((FW_DN_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void SerialSetting_Click(object sender, EventArgs e)
        {
            JLcLib.Comn.WireComnForm.Show(Serial);
            Serial.WriteSettingFile(Parent.IniFile, "TWS");
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

        private void WriteCalibrationData_Click()
        {
#if true // For LSI
            RegisterItem BGR = Parent.RegMgr.GetRegisterItem("BGR");
            RegisterItem LDO5P0 = Parent.RegMgr.GetRegisterItem("LDO5P0");
            RegisterItem LDO1P5 = Parent.RegMgr.GetRegisterItem("LDO1P5");
            RegisterItem ADC_LDO3P3 = Parent.RegMgr.GetRegisterItem("ADC_LDO3P3");
            RegisterItem FSK_LDO3P3 = Parent.RegMgr.GetRegisterItem("FSK_LDO3P3");
            RegisterItem LDO1P8 = Parent.RegMgr.GetRegisterItem("LDO1P8");
            RegisterItem OSC = Parent.RegMgr.GetRegisterItem("OSC");
            RegisterItem ML_ISEN_VREF_SET = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_SET");
            RegisterItem ML_VILIM_SET = Parent.RegMgr.GetRegisterItem("ML_VILIM_SET");
            RegisterItem VRECT_SV_0 = Parent.RegMgr.GetRegisterItem("VRECT_SV[7:0]");
            RegisterItem VRECT_SV_1 = Parent.RegMgr.GetRegisterItem("VRECT_SV[15:8]");
            RegisterItem Vrect_Const_0 = Parent.RegMgr.GetRegisterItem("Vrect_Const[7:0]");
            RegisterItem Vrect_Const_1 = Parent.RegMgr.GetRegisterItem("Vrect_Const[15:8]");
            RegisterItem Vout_SET_SV_0 = Parent.RegMgr.GetRegisterItem("Vout_SET_SV[7:0]");
            RegisterItem Vout_SET_SV_1 = Parent.RegMgr.GetRegisterItem("Vout_SET_SV[15:8]");
            RegisterItem Vout_SET_Const_0 = Parent.RegMgr.GetRegisterItem("Vout_SET_Const[7:0]");
            RegisterItem Vout_SET_Const_1 = Parent.RegMgr.GetRegisterItem("Vout_SET_Const[15:0]");
            RegisterItem Vout_SV_0 = Parent.RegMgr.GetRegisterItem("Vout_SV[7:0]");
            RegisterItem Vout_SV_1 = Parent.RegMgr.GetRegisterItem("Vout_SV[15:8]");
            RegisterItem Vout_Const_0 = Parent.RegMgr.GetRegisterItem("Vout_Const[7:0]");
            RegisterItem Vout_Const_1 = Parent.RegMgr.GetRegisterItem("Vout_Const[15:8]");
            RegisterItem ADC_SV_0 = Parent.RegMgr.GetRegisterItem("ADC_SV[7:0]");
            RegisterItem ADC_SV_1 = Parent.RegMgr.GetRegisterItem("ADC_SV[15:8]");
            RegisterItem ADC_Const_0 = Parent.RegMgr.GetRegisterItem("ADC_Const[7:0]");
            RegisterItem ADC_Const_1 = Parent.RegMgr.GetRegisterItem("ADC_Const[15:8]");
            RegisterItem ISen2_SV_0 = Parent.RegMgr.GetRegisterItem("ISen2_SV[7:0]");
            RegisterItem ISen2_SV_1 = Parent.RegMgr.GetRegisterItem("ISen2_SV[15:8]");
            RegisterItem ISen2_Const_0 = Parent.RegMgr.GetRegisterItem("ISen2_Const[7:0]");
            RegisterItem ISen2_Const_1 = Parent.RegMgr.GetRegisterItem("ISen2_Const[15:8]");
            RegisterItem ISenLow_0 = Parent.RegMgr.GetRegisterItem("ISenLow[7:0]");
            RegisterItem ISenLow_1 = Parent.RegMgr.GetRegisterItem("ISenLow[15:8]");
            RegisterItem ISenDiff_0 = Parent.RegMgr.GetRegisterItem("ISenDiff[7:0]");
            RegisterItem ISenDiff_1 = Parent.RegMgr.GetRegisterItem("ISenDiff[15:8]");
            RegisterItem ISen0_SV_0 = Parent.RegMgr.GetRegisterItem("ISen0_SV[7:0]");
            RegisterItem ISen0_SV_1 = Parent.RegMgr.GetRegisterItem("ISen0_SV[15:8]");
            RegisterItem ISen0_Const_0 = Parent.RegMgr.GetRegisterItem("ISen0_Const[7:0]");
            RegisterItem ISen0_Const_1 = Parent.RegMgr.GetRegisterItem("ISen0_Const[15:8]");
            RegisterItem ISen1_SV_0 = Parent.RegMgr.GetRegisterItem("ISen1_SV[7:0]");
            RegisterItem ISen1_SV_1 = Parent.RegMgr.GetRegisterItem("ISen1_SV[15:8]");
            RegisterItem ISen1_Const_0 = Parent.RegMgr.GetRegisterItem("ISen1_Const[7:0]");
            RegisterItem ISen1_Const_1 = Parent.RegMgr.GetRegisterItem("ISen1_Const[15:8]");

            CalibrationData[0x24] = (byte)BGR.Value;
            CalibrationData[0x28] = (byte)LDO5P0.Value;
            CalibrationData[0x2C] = (byte)LDO1P5.Value;
            CalibrationData[0x30] = (byte)ADC_LDO3P3.Value;
            CalibrationData[0x34] = (byte)FSK_LDO3P3.Value;
            CalibrationData[0x38] = (byte)LDO1P8.Value;
            CalibrationData[0x3C] = (byte)OSC.Value;
            CalibrationData[0x40] = (byte)ML_ISEN_VREF_SET.Value;
            CalibrationData[0x44] = (byte)ML_VILIM_SET.Value;
            CalibrationData[0x48] = (byte)VRECT_SV_0.Value;
            CalibrationData[0x49] = (byte)VRECT_SV_1.Value;
            CalibrationData[0x4C] = (byte)Vrect_Const_0.Value;
            CalibrationData[0x4D] = (byte)Vrect_Const_1.Value;
            CalibrationData[0x50] = (byte)Vout_SET_SV_0.Value;
            CalibrationData[0x51] = (byte)Vout_SET_SV_1.Value;
            CalibrationData[0x54] = (byte)Vout_SET_Const_0.Value;
            CalibrationData[0x55] = (byte)Vout_SET_Const_1.Value;
            CalibrationData[0x58] = (byte)Vout_SV_0.Value;
            CalibrationData[0x59] = (byte)Vout_SV_1.Value;
            CalibrationData[0x5C] = (byte)Vout_Const_0.Value;
            CalibrationData[0x5D] = (byte)Vout_Const_1.Value;
            CalibrationData[0x60] = (byte)ADC_SV_0.Value;
            CalibrationData[0x61] = (byte)ADC_SV_1.Value;
            CalibrationData[0x64] = (byte)ADC_Const_0.Value;
            CalibrationData[0x65] = (byte)ADC_Const_1.Value;
            CalibrationData[0x70] = (byte)ISen2_SV_0.Value;
            CalibrationData[0x71] = (byte)ISen2_SV_1.Value;
            CalibrationData[0x74] = (byte)ISen2_Const_0.Value;
            CalibrationData[0x75] = (byte)ISen2_Const_1.Value;
            CalibrationData[0x78] = (byte)ISenLow_0.Value;
            CalibrationData[0x79] = (byte)ISenLow_1.Value;
            CalibrationData[0x7C] = (byte)ISenDiff_0.Value;
            CalibrationData[0x7D] = (byte)ISenDiff_1.Value;
            CalibrationData[0x80] = (byte)ISen0_SV_0.Value;
            CalibrationData[0x81] = (byte)ISen0_SV_1.Value;
            CalibrationData[0x84] = (byte)ISen0_Const_0.Value;
            CalibrationData[0x85] = (byte)ISen0_Const_1.Value;
            CalibrationData[0x88] = (byte)ISen1_SV_0.Value;
            CalibrationData[0x89] = (byte)ISen1_SV_1.Value;
            CalibrationData[0x8C] = (byte)ISen1_Const_0.Value;
            CalibrationData[0x8D] = (byte)ISen1_Const_1.Value;
#endif
            // Test code start, Initialize calibration data
            //for (int i = 0; i < CalibrationData.Length; i++)
            //    CalibrationData[i] = (byte)i;
            // Test code end
            WriteCalibrationData(TWS_WriteCalibrationData);
        }

        private void RemoveCalibrationData_Click()
        {
            RemoveCalibrationData(TWS_RemoveCalibrationData);
        }

        private void ReadCalibrationData_Click()
        {
            ReadCalibrationData(TWS_ReadCalibrationData);
        }

        private void RunCalibration_Click()
        {
            Parent.StopLogThread();
            if (IsRunCal == false)
            {
                IsRunCal = true;
                RunBackgroudFunc(TWS_RunCalibration, TWS_StopCalibration);
            }
            else
                TWS_StopCalibration();
        }

        private void DownloadFirmwareToRAM_Click(object sender, EventArgs e)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = FW_TARGET.RAM;
                DownloadFirmware(TWS_DownloadFW);
            }
        }

        private void DownloadFirmwareToNV_Click()
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = FW_TARGET.NV_MEM;
                DownloadFirmware(TWS_DownloadFW);
            }
        }

        private void EraseFW()
        {
            EraseFirmware(TWS_EraseFW);
        }

        private void DumpFW()
        {
            ReadFirmwareSize = GetInt_ChipCtrlTextboxes(9);
            DumpFirmware(TWS_DumpFW);
        }

        #region Firmware control methods
        private void TWS_DownloadFW()
        {
            int PageSize = 4;
            uint FlashAddress = 0;
            byte[] SendBytes = new byte[PageSize];
            System.IO.FileStream fs = null;
            System.IO.BinaryReader br = null;

            Status = false;
            if (I2C == null)
                return;

            /* Read Firmware file */
            fs = new System.IO.FileStream(FirmwareName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            br = new System.IO.BinaryReader(fs);
            FirmwareData = br.ReadBytes((int)fs.Length);
            FirmwareSize = (int)fs.Length;

            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (FirmwareData.Length + PageSize - 1) / PageSize;
            })));

            /* Mass erase */
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                TWS_EraseFW();
                if (Status == false)
                    goto EXIT;
            }
            else
                PageSize = 128;

            /* Program */
            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address & data latch
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                if (WaitFlashState(FLASH_STATE.DATA_LATCHED) == false)
                    goto EXIT;
            }
            for (FlashAddress = 0; FlashAddress < FirmwareData.Length; FlashAddress += (uint)PageSize)
            {
                // Copy FW data
                for (int i = 0; i < PageSize; i++)
                {
                    if (FlashAddress + i < FirmwareData.Length)
                        SendBytes[i] = FirmwareData[FlashAddress + i];
                    else
                        SendBytes[i] = 0xFF;
                }
                if (FirmwareTarget == FW_TARGET.NV_MEM)
                {
                    // a. Write flash address
                    WriteHostMemory(0x0FFE0200, FlashAddress);
                    WriteHostMemory(0x0FFE0204, SendBytes);
                    // b. Wait for command latched signal
                    if (WaitFlashState(FLASH_STATE.CMD_READY) == false)
                        goto EXIT;
                    // c. Write program command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
                    WriteHostMemory(0x0FFE0000, 0x02);
                    // 7, Check erase complete
                    if (WaitFlashState(FLASH_STATE.MAT_BUSY) == false)
                        goto EXIT;
                }
                else // Firmware target is SRAM
                {
                    // 4. Write firmware data
                    WriteMemory(0x20000000 + FlashAddress, SendBytes);
                }
                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            if (FirmwareTarget == FW_TARGET.RAM)
            {
                WriteHostMemory(0x4001F000, 0x01); // REMAP = 1
                ResetSystem();
            }
            Status = true;
            br.Close();
            fs.Close();

        EXIT:
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }

        public void TWS_EraseFW()
        {
            Status = false;
            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address latch
            if (WaitFlashState(FLASH_STATE.ADDR_LATCHED, 1000) == false)
                return;
            // 5. Write flash address
            WriteHostMemory(0x0FFE0200, 0x00);
            // 6. Write erase command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
            WriteHostMemory(0x0FFE0000, 0x11);
            // 7, Check erase complete
            if (WaitFlashState(FLASH_STATE.MAT_BUSY, 5000) == false) // timeout 5secI
                return;
            Status = true;
        }

        public void TWS_DumpFW()
        {
            const int PageSize = 128;
            byte[] RcvBytes;
            List<byte> FirmwareData = new List<byte>();

            Status = false;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (ReadFirmwareSize + PageSize - 1) / PageSize + 1; // Mass erase
            })));

            // 1. Set I2C host mode
            SetHostMode();
            // 2. Read FW
            for (uint Addr = 0; Addr < ReadFirmwareSize; Addr += PageSize)
            {
                RcvBytes = ReadMemory(Addr, PageSize);
                if (RcvBytes != null && RcvBytes.Length > 0)
                {
                    // a byte operation
                    for (int i = 0; i < RcvBytes.Length; i++)
                        FirmwareData.Add(RcvBytes[i]);
                }
                else
                    goto EXIT;

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            Status = true;
            ReadFirmwareData = FirmwareData.ToArray();
            System.IO.FileStream fs = new System.IO.FileStream("ReadFirmwareBinary.bin", System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(ReadFirmwareData);
            bw.Close();
            fs.Close();
        EXIT:
            // 3. Reset system
            ResetSystem();
            Status = true;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }
        #endregion Firmware control methods

        #region Calibration control methods
        private void TWS_WriteCalibrationData()
        {
            uint PageSize = 4;
            int Count = 0;
            byte[] SendBytes = new byte[PageSize];

            Status = false;
            // 1~6. Remove calibartion data (Page erase)
            TWS_RemoveCalibrationData();

            // Write calibration data
            //for (uint FlashAddress = 0xFFFD18FC; FlashAddress >= 0xFFFD1800; FlashAddress -= PageSize)
            for (uint FlashAddress = 0xFFFD1800; FlashAddress < 0xFFFD1900; FlashAddress += PageSize)
            {
                // Copy FW data
                for (int i = 0; i < PageSize; i++)
                {
                    if (Count < CalibrationData.Length)
                        SendBytes[i] = CalibrationData[Count];
                    else
                        SendBytes[i] = 0xFF;
                    Count++;
                }
                // a. Write flash address
                WriteHostMemory(0x0FFE0200, FlashAddress);
                WriteHostMemory(0x0FFE0204, SendBytes);
                // b. Wait for command latched signal
                if (WaitFlashState(FLASH_STATE.CMD_READY) == false)
                    return;
                // c. Write program command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
                WriteHostMemory(0x0FFE0000, 0x02);
                // 7, Check erase complete
                if (WaitFlashState(FLASH_STATE.MAT_BUSY) == false)
                    return;
                Status = true;
            }
        }

        private void TWS_RemoveCalibrationData()
        {
            Status = false;
            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address latch
            //if (WaitFlashState(FLASH_STATE.READY, 1000) == false)
            if (WaitFlashState(FLASH_STATE.ADDR_LATCHED, 1000) == false)
                return;
            // 5,6. Erase
            WriteHostMemory(0x0FFE0200, 0xFFFD1800); // NV memory address for calibration data
            // 6. Write erase command, [4:3]ERASE_MODE(0=page(256Bytes), 1=sector(1kBytes), 2=Mass), [1]PROGRAM_EN, [0]ERASE_EN
            WriteHostMemory(0x0FFE0000, 0x01); // 256 bytes page erase command
            // 7, Check erase complete
            if (WaitFlashState(FLASH_STATE.MAT_BUSY, 5000) == false) // timeout 5secI
                return;
            Status = true;
        }

        private void TWS_ReadCalibrationData()
        {
            uint PageSize = 4;
            int Data;

            Status = false;

            // 1. Set I2C host mode
            SetHostMode();
            // 2,3. Check Remap
            //CheckRemap();
            // 4. Check ready signal for address latch
            //if (WaitFlashState(FLASH_STATE.READY, 1000) == false)
            if (WaitFlashState(FLASH_STATE.ADDR_LATCHED, 1000) == false)
                return;

            // Read calibration data
            for (uint FlashAddress = 0x0FFD1800; FlashAddress < 0x0FFD1900; FlashAddress += PageSize)
            {
                Data = ReadHostMemory(FlashAddress);
                Log.WriteLine(FlashAddress.ToString("X8") + " : " + Data.ToString("X8"));

                Status = true;
            }
        }
        #endregion Calibration control methods

        public override bool CheckConnectionForLog()
        {
            return ((Serial != null) && Serial.IsOpen);
        }

        public override void RunLog()
        {
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
        }

        public override void SendCommand(string Command)
        {
            // Test only
            Command += "\r\n";
            byte[] Var = Encoding.ASCII.GetBytes(Command);
            Serial.WriteBytes(Var, Var.Length, true);
        }

        private void StartCommandProcessing(RegisterItem CommandReg, int Timeout = 1000)
        {
            CommandReg.Write();
            for (int i = 0; i < (Timeout / 10); i++)
            {
                System.Threading.Thread.Sleep(10);
                CommandReg.Read();
                if (CommandReg.Value == 0)
                    break;
            }
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal, Result = 0;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.TEST:
                    // FW test functions
                    Result = RunTest((TEST_ITEMS)TestItemIndex, iVal);
                    Log.WriteLine(TestItemIndex.ToString() + ":" + iVal.ToString() + ":" + Result.ToString());
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((AUTO_TEST_ITEMS)TestItemIndex)
                    {
                        // GUI test functions
                        case AUTO_TEST_ITEMS.TEST_ISEN:
                            TEST_RunIsen();
                            break;
                        case AUTO_TEST_ITEMS.TEST_VOUT:
                            TEST_RunVout();
                            break;
                        case AUTO_TEST_ITEMS.TEST_VOUT_WPC:
                            TEST_RunVout_wpc();
                            break;
                        case AUTO_TEST_ITEMS.TEST_VRECT_WPC:
                            TEST_RunVrect_wpc();
                            break;
                        case AUTO_TEST_ITEMS.TEST_ISEN_REG:
                            TEST_RunIsen_Register();
                            break;
                        case AUTO_TEST_ITEMS.TEST_ACTIVE_LOAD:
                            TEST_RunActiveLoad();
                            break;
                        case AUTO_TEST_ITEMS.TEST_LOAD_SWEEP:
                            TEST_RunLOAD();
                            break;
                        case AUTO_TEST_ITEMS.TEST_ADC_CODE:
                            TEST_RunADCCODE();
                            break;
                        case AUTO_TEST_ITEMS.TEST_JISU:
                            TEST_RunJISUCODE();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.CAL:
                    switch ((CAL_ITEMS)TestItemIndex)
                    {
                        case CAL_ITEMS.CAL_RUN:
                            RunCalibration_Click();
                            break;
                        case CAL_ITEMS.CAL_ERASE:
                            RemoveCalibrationData_Click();
                            break;
                        case CAL_ITEMS.CAL_WRITE:
                            WriteCalibrationData_Click();
                            break;
                        case CAL_ITEMS.CAL_READ:
                            ReadCalibrationData_Click();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.FW:
                    switch ((FW_DN_ITEMS)TestItemIndex)
                    {
                        case FW_DN_ITEMS.ERASE:
                            EraseFW();
                            break;
                        case FW_DN_ITEMS.WRITE:
                            DownloadFirmwareToNV_Click();
                            break;
                        case FW_DN_ITEMS.READ:
                            DumpFW();
                            break;
                        default:
                            break;
                    }
                    break;
                default:
                    break;

            }
        }

        public int RunTest(TEST_ITEMS TestItem, int Arg)
        {
            int iVal = 0;
            RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");   // 0x1C
            RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");         // 0x1E
            RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");         // 0x1F

            switch (TestItem)
            {
                // Test mode control
                case TEST_ITEMS.ENTER_TEST_PT:
                case TEST_ITEMS.EXIT_TEST_PT:
                case TEST_ITEMS.SET_TEST_VRECT:
                    CommandReg.Value = 0x01 + (uint)TestItem - (uint)TEST_ITEMS.ENTER_TEST_PT;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg);
                    StartCommandProcessing(CommandReg);
                    break;
                // Write command only
                case TEST_ITEMS.LDO_TURNON:
                case TEST_ITEMS.LDO_TURNOFF:
                case TEST_ITEMS.LDO_SET:
                    CommandReg.Value = 0x08 + (uint)TestItem - (uint)TEST_ITEMS.LDO_TURNON;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg);
                    StartCommandProcessing(CommandReg);
                    break;
                case TEST_ITEMS.ADC_GET_VRECT:
                case TEST_ITEMS.ADC_GET_VOUT:
                case TEST_ITEMS.ADC_GET_IOUT:
                case TEST_ITEMS.ADC_GET_VBAT_SENS:
                case TEST_ITEMS.ADC_GET_VPTAT:
                case TEST_ITEMS.ADC_GET_NTC:
                    CommandReg.Value = 0x10 + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_VRECT;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
                case TEST_ITEMS.ADC_GET_EB1:
                case TEST_ITEMS.ADC_GET_EB2:
                    CommandReg.Value = 0x1B + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_EB1;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
                case TEST_ITEMS.GET_VRECT_mV: // 0x20
                case TEST_ITEMS.GET_VOUT_mV:
                case TEST_ITEMS.GET_IOUT_mA:
                case TEST_ITEMS.GET_VBAT_SENS_mV:
                case TEST_ITEMS.GET_VPTAT_mV:
                case TEST_ITEMS.GET_NTC_mV:
                case TEST_ITEMS.GET_IOUT_MIN_mA:
                case TEST_ITEMS.GET_IOUT_MAX_mA:
                    CommandReg.Value = 0x20 + (uint)TestItem - (uint)TEST_ITEMS.GET_VRECT_mV;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
                case TEST_ITEMS.ADC_GET_VRECT_MIN:
                case TEST_ITEMS.ADC_GET_VOUT_MIN:
                case TEST_ITEMS.ADC_GET_IOUT_MIN:
                case TEST_ITEMS.ADC_GET_VBAT_SENS_MIN:
                case TEST_ITEMS.ADC_GET_VPTAT_MIN:
                case TEST_ITEMS.ADC_GET_NTC_MIN:
                    CommandReg.Value = 0x40 + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_VRECT_MIN;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;

                case TEST_ITEMS.ADC_GET_VRECT_MAX:
                case TEST_ITEMS.ADC_GET_VOUT_MAX:
                case TEST_ITEMS.ADC_GET_IOUT_MAX:
                case TEST_ITEMS.ADC_GET_VBAT_SENS_MAX:
                case TEST_ITEMS.ADC_GET_VPTAT_MAX:
                case TEST_ITEMS.ADC_GET_NTC_MAX:
                    CommandReg.Value = 0x50 + (uint)TestItem - (uint)TEST_ITEMS.ADC_GET_VRECT_MAX;
                    if (Arg < 0) Arg = 0;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)Arg + 1);
                    StartCommandProcessing(CommandReg);
                    iVal = (int)GetInt16fromRegItems(Arg1Reg, Arg0Reg);
                    break;
            }
            return iVal;
        }

        #region Chip test methods
        private void TEST_RunIsen()
        {
            double Vrect, Vout;
            int IoutADC, IoutCode, VrectCode, VoutCode;
            int xPos = 1, yPos = 1;
            string sheet_name, time;

            Log.WriteLine("Start TEST_ISEN");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : N/C\n- CH4 : N/C\n\nElectronicLoad\n- Out : VOUT");
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 5V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 5"); // offset 5V
#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            time = DateTime.Now.ToString("MMddHHmmss_");
            sheet_name = time + "Iout";
            Parent.xlMgr.Sheet.Add(sheet_name);
            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "I_load(mA)");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "Iout ADC");
            Parent.xlMgr.Cell.Write(xPos + 8, yPos, "Vrect Code");
            Parent.xlMgr.Cell.Write(xPos + 15, yPos, "Vout Code");
            Parent.xlMgr.Cell.Write(xPos + 22, yPos, "Vrect Scope");
            Parent.xlMgr.Cell.Write(xPos + 29, yPos, "Vout Scope");
            Parent.xlMgr.Cell.Write(xPos + 36, yPos, "Iout Code");

            for (double Iload = 0; Iload <= 1200; Iload += 50) // 0~1200(mA) Electronic Load
            {
                ElectronicLoad.Write("CURR " + (Iload / 1000).ToString());
                xPos = 1;
                yPos++;
                Parent.xlMgr.Cell.Write(xPos, yPos + 1, Iload.ToString());
                for (double Vr = 4.9; Vr <= 5.5; Vr += 0.1) // 4.9~5.5(V) Power Supply
                {
                    PowerSupply0.Write("VOLT " + Vr.ToString());
                    JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
                    xPos++;
                    JLcLib.Delay.Sleep(200);
                    if (yPos == 2)
                    {
                        Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 7, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 14, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 21, yPos, (Vr.ToString() + "V"));
                        Parent.xlMgr.Cell.Write(xPos + 28, yPos, (Vr.ToString() + "V"));
                    }

                    IoutADC = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1024);
                    IoutCode = RunTest(TEST_ITEMS.GET_IOUT_mA, 1024);
                    VrectCode = RunTest(TEST_ITEMS.GET_VRECT_mV, 256);
                    VoutCode = RunTest(TEST_ITEMS.GET_VOUT_mV, 256);
                    Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                    Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));

                    Parent.xlMgr.Cell.Write(xPos + 0, yPos + 1, IoutADC.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 7, yPos + 1, VrectCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 14, yPos + 1, VoutCode.ToString());
                    Parent.xlMgr.Cell.Write(xPos + 21, yPos + 1, (Vrect * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 28, yPos + 1, (Vout * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 35, yPos + 1, IoutCode.ToString());
                }
            }
            Log.WriteLine("End TEST_ISEN");
        }

        private void CAL_VOUT_Reg()
        {
            double Result;
            double TargetVoutValue = 5.0;
            double Margin = 0.05;

            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]

            for (uint i = 1; i < 2; i++)
            {
                VOUT_SET_RANGE.Value = i;
                TRIM_VOUT_SET.Value = TRIM_VOUT_SET.Read();

                for (int j = 0; j < 100; j++)
                {
                    JLcLib.Delay.Sleep(50);
                    Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                    if (Result > TargetVoutValue + Margin)
                    {
                        TRIM_VOUT_SET.Value -= 1;
                    }
                    else if (Result < TargetVoutValue - Margin)
                    {
                        TRIM_VOUT_SET.Value += 1;
                        if (TRIM_VOUT_SET.Value > 100)
                            TRIM_VOUT_SET.Value = 100;
                    }
                    else
                        break;
                    TRIM_VOUT_SET.Write();
                }
            }
        }

        private void TEST_RunIsen_Register()
        {
            double Vrect_vpp, Vrect, Vout, Iout;
            int IoutCode_Min, IoutCode_Max, IoutCode, IoutADC_Min, IoutADC_Max, IoutADC, VrectCode, VoutCode;
            int xPos = 0, yPos = 3;
            string sheet_name, time;

            RegisterItem FB = Parent.RegMgr.GetRegisterItem("ML_FB_TRIM<3:0>");
            RegisterItem ISEN_VREF = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>");
            RegisterItem ISEN_RVAR = Parent.RegMgr.GetRegisterItem("ML_ISEN_RVAR<1:0>");
            RegisterItem VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>");
            RegisterItem VREF = Parent.RegMgr.GetRegisterItem("ML_VREF_TRIM<3:0>");
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN");
            RegisterItem FC_R = Parent.RegMgr.GetRegisterItem("ML_FC_RTRIM<3:0>");
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>");
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");
            RegisterItem ANA_MUX = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN");

            Log.WriteLine("Start TEST_ISEN Sweep Register");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : VBAT_SENS\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5.3");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 5V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 5"); // offset 5V
            // Cannel.3 use for Iout(mV)
            Oscilloscope.Write(":CHAN3:DISP 1");
            Oscilloscope.Write(":CHAN3:SCAL 5E-1"); // 500mV/div
            Oscilloscope.Write(":CHAN3:OFFS 1.5"); // offset 5V
#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            //ElectronicLoad.Write("CURR 0.6");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            OCL_EN.Value = 0;
            OCL_EN.Write();
            FC_R.Value = 15;
            FC_R.Write();
            ANA_SEL.Value = 3;
            ANA_SEL.Write();
            ANA_PEN.Value = 1;
            ANA_PEN.Write();
            ANA_MUX.Value = 1;
            ANA_MUX.Write();

            time = DateTime.Now.ToString("MMddHHmmss_");

            for (double Iload = 600; Iload <= 600; Iload += 50) // 0~1200(mA) Electronic Load
            {
                sheet_name = time + Iload.ToString() + "mA";
                Parent.xlMgr.Sheet.Add(sheet_name);

                xPos = 3;
                yPos = 4;
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, ("Load = " + Iload.ToString() + "mA"));
                yPos++;
                Parent.xlMgr.Cell.Write(xPos + 5, yPos, "IOUT_ADC_AVG(Max 4095)");
                Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, "IOUT_PP(mV)");
                Parent.xlMgr.Cell.Write(xPos + 5 + 14, yPos, "IOUT_Code_AVG(Max 1200)");
                yPos++;
                Parent.xlMgr.Cell.Write(xPos + 0, yPos, "ML_FB_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 1, yPos, "ML_VREF_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 2, yPos, "ML_ISEN_RVAR");
                Parent.xlMgr.Cell.Write(xPos + 3, yPos, "ML_VILIM_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, "ML_ISEN_VREF");

                for (double Vr = 4.9; Vr <= 5.5; Vr += 0.1) // 4.9~5.5(V) Power Supply
                {
                    yPos = 6;
                    Parent.xlMgr.Cell.Write(xPos + 5, yPos, (Vr.ToString() + "V"));
                    Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, (Vr.ToString() + "V"));
                    Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, (Vr.ToString() + "V"));
                    yPos++;
                    for (uint i = 0; i <= 15; i++) // 0~15 ML_FB_TRIM<3:0>
                    {
                        FB.Value = i;
                        FB.Write();
                        for (uint j = 0; j <= 15; j++) // 0~15 ML_VREF_TRIM<3:0>
                        {
                            VREF.Value = j;
                            VREF.Write();
                            ElectronicLoad.Write("CURR 0");
                            PowerSupply0.Write("VOLT 6V");
                            JLcLib.Delay.Sleep(100);
                            CAL_VOUT_Reg();
                            PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                            ElectronicLoad.Write("CURR " + (Iload / 1000).ToString("F2"));
                            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
#if true
                            for (uint k = 0; k <= 3; k++) // 0~3 ML_ISEN_RVAR<1:0>
                            {
                                ISEN_RVAR.Value = k;
                                ISEN_RVAR.Write();
                                for (uint l = 0; l <= 15; l++) // 0~15 ML_VILIM_TRIM<3:0>
                                {
                                    VILIM.Value = l;
                                    VILIM.Write();
                                    for (uint m = 0; m <= 15; m++) // 0~15 ML_ISEN_VREF_TRIM<3:0>
                                    {
                                        ISEN_VREF.Value = m;
                                        ISEN_VREF.Write();
                                        JLcLib.Delay.Sleep(50);
#if true
                                        IoutCode = RunTest(TEST_ITEMS.GET_IOUT_mA, 256);
                                        IoutADC_Min = RunTest(TEST_ITEMS.ADC_GET_IOUT_MIN, 256);
                                        IoutADC_Max = RunTest(TEST_ITEMS.ADC_GET_IOUT_MAX, 256);
                                        IoutADC = RunTest(TEST_ITEMS.ADC_GET_IOUT, 256);
                                        //VrectCode = RunTest(TEST_ITEMS.GET_VRECT_mV, 256);
                                        //VoutCode = RunTest(TEST_ITEMS.GET_VOUT_mV, 256);

                                        //Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                                        //Vrect_vpp = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VPP? CHAN1"));
                                        //Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                                        Iout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VPP? CHAN3"));

                                        if (xPos == 3)
                                        {
                                            Parent.xlMgr.Cell.Write(xPos + 0, yPos, i.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 1, yPos, j.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 2, yPos, k.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 3, yPos, l.ToString());
                                            Parent.xlMgr.Cell.Write(xPos + 4, yPos, m.ToString());
                                        }
                                        Parent.xlMgr.Cell.Write(xPos + 5, yPos, IoutADC.ToString());
                                        Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, (Iout * 1000.0).ToString());
                                        Parent.xlMgr.Cell.Write(xPos + 5 + 14, yPos, IoutCode.ToString());

                                        //Parent.xlMgr.Cell.Write(xPos + 5 + 7, yPos, IoutADC_Min.ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 5 + 14, yPos, IoutADC_Max.ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vrect_vpp * 1000.0).ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 2, yPos, (Vrect * 1000.0).ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 3, yPos, VrectCode.ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 4, yPos, (Vout * 1000.0).ToString());
                                        //Parent.xlMgr.Cell.Write(xPos + 5, yPos, VoutCode.ToString());                                        
#endif
                                        yPos++;
                                    }
                                }
                            }
#endif
                        }
                    }
                    xPos++;


                }
                //xPos += 5;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write();
            ANA_PEN.Value = 0;
            ANA_PEN.Write();
            ANA_MUX.Value = 0;
            ANA_MUX.Write();
            Log.WriteLine("End TEST_ISEN Sweep Register");
        }

        private void TEST_RunActiveLoad()
        {
            double Irect;
            int xPos = 1, yPos = 1;
            string sheet_name;

            RegisterItem SEL = Parent.RegMgr.GetRegisterItem("AL_SEL");
            RegisterItem PEN = Parent.RegMgr.GetRegisterItem("AL_PEN");
            RegisterItem ILOAD = Parent.RegMgr.GetRegisterItem("AL_ILOAD_SET<3:0>");
            RegisterItem ISLOPE = Parent.RegMgr.GetRegisterItem("AL_ISLOPE_TRIM<4:0>");
            RegisterItem IOFFSET = Parent.RegMgr.GetRegisterItem("AL_IOFFSET_TRIM<4:0>");


            Log.WriteLine("Start Active Load Sweep Register");

            CAL_CheckInstrument();

            if (PowerSupply0 == null || PowerSupply0.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            MessageBox.Show("Power Supply\n- CH1 : VRECT\n- CH2 : AP1P8\n- CH3 : N/C");

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 6.0");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif
            JLcLib.Delay.Sleep(200);

            SEL.Value = 1;
            SEL.Write();
            PEN.Value = 1;
            PEN.Write();
            sheet_name = DateTime.Now.ToString("MMddHHmmss") + "_AL";
            Parent.xlMgr.Sheet.Add(sheet_name);

            Parent.xlMgr.Cell.Write(xPos, yPos, "Active Load = on");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "VRECT = 6V");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "IRECT = 0mA");
            xPos = 1;
            for (uint i = 0; i < 32; i++) // AL_ILOAD_SET
            {
                ILOAD.Value = i;
                ILOAD.Write();
                Parent.xlMgr.Cell.Write(xPos, yPos + 1, "AL_ILOAD_SET = " + i.ToString());
                Parent.xlMgr.Cell.Write(xPos, yPos + 2, "AL_ISLOPE_TRIM");
                Parent.xlMgr.Cell.Write(xPos + 1, yPos + 1, "AL_IOFFSET_TRIM");
                for (uint j = 0; j < 32; j++) // AL_ISLOPE_TRIM
                {
                    ISLOPE.Value = j;
                    ISLOPE.Write();
                    Parent.xlMgr.Cell.Write(xPos, (int)(yPos + 3 + j), j.ToString());
                    for (uint k = 0; k < 32; k++) // AL_IOFFSET_TRIM
                    {
                        IOFFSET.Value = k;
                        IOFFSET.Write();

                        if (j == 0) Parent.xlMgr.Cell.Write((int)(xPos + 1 + k), (int)(yPos + 2 + j), k.ToString());

#if (POWER_SUPPLY_E36313A)
                        Irect = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR?"));
#else
                        Irect = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR?"));
#endif
                        Parent.xlMgr.Cell.Write((int)(xPos + 1 + k), (int)(yPos + 3 + j), (Irect * 1000).ToString());
                    }
                }
                yPos += 36;
            }
            Log.WriteLine("End Active Load Sweep Register");
        }

        private void TEST_RunLOAD()
        {
            Log.WriteLine("Start Loader Sweep");

            CAL_CheckInstrument();

            if (ElectronicLoad == null || ElectronicLoad.IsOpen == false)
                return;

#if (DC_LOADER_KIKUSUI) // KIKUSUI
            ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
#endif
            ElectronicLoad.Write("INP 1");
            ElectronicLoad.Write("CURR 0.3");
            System.Threading.Thread.Sleep(1);
            ElectronicLoad.Write("CURR 0.5");
            System.Threading.Thread.Sleep(1);
            ElectronicLoad.Write("CURR 0.8");
            System.Threading.Thread.Sleep(1);
            ElectronicLoad.Write("CURR 1");
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);
            System.Threading.Thread.Sleep(1000);

            for (double i = 1000; i >= 0; i--)
            {
                ElectronicLoad.Write("CURR " + (i / 1000).ToString());
                //System.Threading.Thread.Sleep(1);

            }
            ElectronicLoad.Write("INP 0");
            ElectronicLoad.Write("CURR 0");
            Log.WriteLine("End Loader Sweep");
        }

        private void TEST_RunADCCODE()
        {
            int xPos = 1, yPos = 1;
            string sheet_name;
            double Vrect, Vntc;

            Log.WriteLine("Start TEST_ADC_CODE");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || PowerSupply0 == null || PowerSupply0.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;

            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : GPIO\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : GPIO\n- CH3 : N/C\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("CURR 1");
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT 0");
            PowerSupply0.Write("CURR 0.5");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:3)");
#else
            PowerSupply0.Write("OUTP ON");
#endif
            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 0"); // offset 0V
            // Cannel.2 use for GPIO
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 0"); // offset 0V

            JLcLib.Delay.Sleep(200);

            sheet_name = DateTime.Now.ToString("MMddHHmmss_") + "ADC_CODE";
            Parent.xlMgr.Sheet.Add(sheet_name);

            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "Vrect");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "Vrect_Scope");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "Vrect_Code_AVG");
            Parent.xlMgr.Cell.Write(xPos + 3, yPos, "Vrect_Code_Min");
            Parent.xlMgr.Cell.Write(xPos + 4, yPos, "Vrect_Code_Max");

            PowerSupply0.Write("INST:NSEL 2");
            for (double Vr = 3.7; Vr <= 20.1; Vr += 0.1) // Vrect Power Supply (V) 3.7~20
            {
                yPos++;
                Oscilloscope.Write(":CHAN1:OFFS " + Vr.ToString("F2"));
                PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

                Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vr.ToString("F2") + "V"));
                Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vrect * 1000).ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_VRECT, 255);
                Parent.xlMgr.Cell.Write(xPos + 2, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_VRECT_MIN, 255);
                Parent.xlMgr.Cell.Write(xPos + 3, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_VRECT_MAX, 255);
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, Vrect.ToString("F3"));
            }

            xPos = 7;
            yPos = 1;

            Parent.xlMgr.Cell.Write(xPos + 0, yPos, "Vntc");
            Parent.xlMgr.Cell.Write(xPos + 1, yPos, "Vntc_Scope");
            Parent.xlMgr.Cell.Write(xPos + 2, yPos, "Vntc_Code_AVG");
            Parent.xlMgr.Cell.Write(xPos + 3, yPos, "Vntc_Code_Min");
            Parent.xlMgr.Cell.Write(xPos + 4, yPos, "Vntc_Code_Max");

            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("INST:NSEL 3");
            for (double Vr = 0; Vr <= 3.3; Vr += 0.05) // Vrect Power Supply (V) 0~3.3
            {
                yPos++;
                Oscilloscope.Write(":CHAN2:OFFS " + Vr.ToString("F2"));
                PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

                Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vr.ToString("F2") + "V"));
                Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vrect * 1000).ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_NTC, 255);
                Parent.xlMgr.Cell.Write(xPos + 2, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_NTC_MIN, 255);
                Parent.xlMgr.Cell.Write(xPos + 3, yPos, Vrect.ToString("F3"));
                Vrect = RunTest(TEST_ITEMS.ADC_GET_NTC_MAX, 255);
                Parent.xlMgr.Cell.Write(xPos + 4, yPos, Vrect.ToString("F3"));
            }

            Log.WriteLine("End TEST_ADC_CODE");
        }

        private void TEST_RunJISUCODE()
        {
            int Result;
            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]

            Result = RunTest(TEST_ITEMS.ADC_GET_EB1, 1023); // 1024 average
        }

        private void TEST_RunVout()
        {
            double Vrect, Vout;
            int xPos = 3, yPos = 5;

            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]

            Log.WriteLine("Start TEST_Vout");

            CAL_CheckInstrument();

            if (Oscilloscope == null || Oscilloscope.IsOpen == false)
                return;

            if (Parent.xlMgr == null)
                return;
            if (Parent.xlMgr.Sheet.Select("VoutTest") == false)
                Parent.xlMgr.Sheet.Add("VoutTest");

            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : N/C\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : N/C\n- CH4 : N/C");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
#endif
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 8");
            PowerSupply0.Write("CURR 1");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif

            //RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //RunTest(TEST_ITEMS.LDO_SET, 5000);

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 8"); // offset 8V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 1"); // 1V/div
            Oscilloscope.Write(":CHAN2:OFFS 4"); // offset 4V

            JLcLib.Delay.Sleep(200);

            for (uint ct = 0; ct <= 1; ct++)
            {
                VOUT_SET_RANGE.Value = ct;
                VOUT_SET_RANGE.Write();
                yPos = 5;
                for (uint set = 0; set < 101; set++)
                {
                    TRIM_VOUT_SET.Value = set;
                    TRIM_VOUT_SET.Write();
                    JLcLib.Delay.Sleep(100);

                    Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                    Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));

                    Parent.xlMgr.Cell.Write(xPos + 0, yPos, (Vrect * 1000.0).ToString());
                    Parent.xlMgr.Cell.Write(xPos + 1, yPos, (Vout * 1000.0).ToString());

                    yPos++;
                }
                xPos += 2;
            }
            Log.WriteLine("End TEST_Vout");
        }

        private void TEST_RunVout_wpc()
        {
            RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");
            RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");
            RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");
            RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");

            Log.WriteLine("Start TEST_Vout_WPC");

            // LDO_SET 3700mV
            CommandReg.Value = 0x08 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 3700);
            StartCommandProcessing(CommandReg);
            JLcLib.Delay.Sleep(500);

            for (int i = 0; i < 1; i++)
            {
                for (double vr = 3.7; vr < 6.5; vr += 0.1)
                {
                    // SET_TEST_VOUT
                    CommandReg.Value = 0x08 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                    //JLcLib.Delay.Sleep(10);
                }
                for (double vr = 6.5; vr >= 3.7; vr -= 0.1)
                {
                    // SET_TEST_VOUT
                    CommandReg.Value = 0x08 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                    //JLcLib.Delay.Sleep(10);
                }
            }
            JLcLib.Delay.Sleep(500);
            // LDO_SET 5000mV
            CommandReg.Value = 0x08 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 5000);
            StartCommandProcessing(CommandReg);

            Log.WriteLine("End TEST_Vout_WPC");
        }

        private void TEST_RunVrect_wpc()
        {
            int iVal = 0;
            RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");
            RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");
            RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");
            RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");

            Log.WriteLine("Start TEST_Vrect_WPC");

            // ENTER_TEST_PT
            CommandReg.Value = 0x01 + 0x00;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 0);
            StartCommandProcessing(CommandReg);

            // LDO_SET 6500mV
            CommandReg.Value = 0x08 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 6500);
            StartCommandProcessing(CommandReg);

            // SET_TEST_VRECT 3.7V
            CommandReg.Value = 0x01 + 0x02;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, (int)(3.7 * 1000));
            StartCommandProcessing(CommandReg);
            JLcLib.Delay.Sleep(500);

            for (int i = 0; i < 3; i++)
            {
                for (double vr = 3.7; vr <= 11.5; vr += 0.5)
                {
                    // SET_TEST_VRECT
                    CommandReg.Value = 0x01 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                }
                for (double vr = 12; vr >= 3.7; vr -= 0.5)
                {
                    // SET_TEST_VRECT
                    CommandReg.Value = 0x01 + 0x02;
                    SetInt16byRegItems(Arg1Reg, Arg0Reg, (uint)(vr * 1000));
                    StartCommandProcessing(CommandReg);
                }
            }
            JLcLib.Delay.Sleep(500);

            // EXIT_TEST_PT
            CommandReg.Value = 0x01 + 0x01;
            SetInt16byRegItems(Arg1Reg, Arg0Reg, 0);
            StartCommandProcessing(CommandReg);

            Log.WriteLine("End TEST_Vrect_WPC");
        }
        #endregion Chip test methods

        #region Calibration methods
        private void TWS_RunCalibration()
        {
            Log.WriteLine("Run Calibratoin!!");
            CAL_CheckInstrument();
            CAL_InitInstrument();
            JLcLib.Delay.Sleep(2000);
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
            PowerSupply1.Write("OUTP OFF");
#endif
            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : GPIO\n\nOscilloscope\n- CH1 : LDO5P0\n- CH2 : LDO1P8\n- CH3 : VBAT_SENS\n- CH4 : INT_A");
            if (!IsRunCal) return;
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:3)");
#else
            PowerSupply0.Write("OUTP ON");
            PowerSupply1.Write("OUTP ON");
#endif
            System.Threading.Thread.Sleep(100);
            // BGR calibration
            CAL_RunBGR();
            if (!IsRunCal) return;
            CAL_RunLDO5P0();
            if (!IsRunCal) return;
            CAL_RunLDO1P5();
            if (!IsRunCal) return;
            CAL_RunADC_LDO3P3();
            if (!IsRunCal) return;
            CAL_RunFSK_LDO3P3();
            if (!IsRunCal) return;
            CAL_RunLDO1P8();
            if (!IsRunCal) return;
            CAL_RunOSC();
            if (!IsRunCal) return;
            MessageBox.Show("Change Oscilloscope CH2 from LDO1P8 to VOUT");
            if (!IsRunCal) return;
            CAL_RunVRECT();
            if (!IsRunCal) return;
            CAL_RunVOUT();
            if (!IsRunCal) return;
            CAL_RunADC();
            if (!IsRunCal) return;
            CAL_RunISENVREF();
            if (!IsRunCal) return;
            CAL_RunVILIM();
            if (!IsRunCal) return;
#if false
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP OFF,(@1:3)");
#else
            PowerSupply0.Write("OUTP OFF");
            PowerSupply1.Write("OUTP OFF");
#endif
            MessageBox.Show("Change Power Supply CH1 and CH2");
#endif
            if (!IsRunCal) return;
            CAL_RunIOUT();
            Log.WriteLine("Done Calibratoin!!");
        }

        private void TWS_StopCalibration()
        {
            IsRunCal = false;
        }

        private void CAL_CheckInstrument()
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
                    case JLcLib.Instrument.InstrumentTypes.OscilloScope1:
                        if (Oscilloscope == null)
                            Oscilloscope = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (Oscilloscope.IsOpen == false)
                            Oscilloscope.Open();
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

        private void CAL_InitInstrument()
        {
            if (PowerSupply0 != null && PowerSupply0.IsOpen)
            {
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
                PowerSupply0.Write("OUTP OFF,(@1:3)");
#else // 2port Power supply
                PowerSupply0.Write("OUTP 0");
                PowerSupply1.Write("OUTP 0");
#endif
                PowerSupply0.Write("INST:NSEL 1");// AP1P8 power supply
                PowerSupply0.Write("VOLT 1.8");
                PowerSupply0.Write("CURR 0.3");
                PowerSupply0.Write("INST:NSEL 2"); // VRECT power supply
                PowerSupply0.Write("VOLT 6.0");
                PowerSupply0.Write("CURR 2.0");
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
                PowerSupply0.Write("INST:NSEL 3");
                PowerSupply0.Write("VOLT 0.3");
                PowerSupply0.Write("CURR 0.3");
#else // 2port Power supply
                PowerSupply1.Write("INST:NSEL 1");
                PowerSupply1.Write("VOLT 0.3");
                PowerSupply1.Write("CURR 0.2");
#endif
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
                PowerSupply0.Write("OUTP ON,(@1:3)");
#else // 2port Power supply
                PowerSupply0.Write("OUTP 1");
                PowerSupply1.Write("OUTP 1");
#endif
            }
            if (Oscilloscope != null && Oscilloscope.IsOpen)
            {
                //Oscilloscope.Write(":TIM:SCAL 2E-4"); // 200usec/div, Scale / 1div
                Oscilloscope.Write(":TIM:SCAL 1E-7"); // 100nsec/div, Scale / 1div
                Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
                Oscilloscope.Write(":TIM:DEL 0"); // delay
                // Cannel.1 use for LDO 5p0
                Oscilloscope.Write(":CHAN1:DISP 1");
                Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
                Oscilloscope.Write(":CHAN1:OFFS 5"); // offset 6V
                // Cannel.2 use for LDO 1p8, VOUT
                Oscilloscope.Write(":CHAN2:DISP 1");
                Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 2V/div
                Oscilloscope.Write(":CHAN2:OFFS 1.8"); // offset 1V
                // Cannel.3 use for ANA_TEST
                Oscilloscope.Write(":CHAN3:DISP 1");
                Oscilloscope.Write(":CHAN3:SCAL 2E-1"); // 200mV/div
                Oscilloscope.Write(":CHAN3:OFFS 1.2"); // offset 1V
                // Cannel.4 use for OSC test
                Oscilloscope.Write(":CHAN4:DISP 1");
                Oscilloscope.Write(":CHAN4:SCAL 1"); // 1V/div
                Oscilloscope.Write(":CHAN4:OFFS 3"); // offset 1V
                Oscilloscope.Write(":TRIG:MODE EDGE"); // trigger mode edge
                Oscilloscope.Write(":TRIG:EDGE:SOUR CHAN4"); // triger source 4
                Oscilloscope.Write(":TRIG:EDGE:LEV 9E-1"); // trigger level 900mV
            }
            if (ElectronicLoad != null && ElectronicLoad.IsOpen)
            {
#if (DC_LOADER_KIKUSUI) // KIKUSUI
                ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
#else // ITECH
                ElectronicLoad.Write("SYST:REM");
                ElectronicLoad.Write("FUNC CURR");
#endif
                ElectronicLoad.Write("INP 0");
                ElectronicLoad.Write("CURR 0");
                System.Threading.Thread.Sleep(10);
                ElectronicLoad.Write("INP 1");
            }
        }

        private void CAL_RunBGR()
        {
            double Result = 0.0;
            double Target = 1.22, Margin = 0.005;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "BGR";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR<4:0>"); // 0x91[4:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 0; // BGR
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_BGR.Read();
            TRIM_BGR.Value = 0;
            TRIM_BGR.Write();
            for (uint BitPos = 10; BitPos < 26; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_BGR.Value = BitPos;
                TRIM_BGR.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_BGR.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_BGR.Value = (byte)reg_value;
            TRIM_BGR.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x24] = (byte)TRIM_BGR.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x24] = (byte)0xFF;
            }
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunLDO5P0()
        {
            double Result = 0.0;
            double Target = 5, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "LDO5P0";
            RegisterItem TRIM_LDO5P0 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO5P0<3:0>"); // 0x94[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            TRIM_LDO5P0.Read();
            TRIM_LDO5P0.Value = 0;
            TRIM_LDO5P0.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO5P0.Value = BitPos;
                TRIM_LDO5P0.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO5P0.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO5P0.Value = reg_value;
            TRIM_LDO5P0.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x28] = (byte)TRIM_LDO5P0.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x28] = (byte)0xFF;
            }
        }

        private void CAL_RunLDO1P5()
        {
            double Result = 0.0;
            double Target = 1.5, Margin = 0.01;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.02; // pass/fail limit
            string Name = "LDO1P5";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_LDO1P5 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO1P5<3:0>"); // 0x9C[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 1.5"); // offset 1.5V
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 4; // LDO1P5
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_LDO1P5.Read();
            TRIM_LDO1P5.Value = 0x08;
            TRIM_LDO1P5.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO1P5.Value = BitPos;
                TRIM_LDO1P5.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO1P5.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO1P5.Value = reg_value;
            TRIM_LDO1P5.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x2C] = (byte)TRIM_LDO1P5.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x2C] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write(); // Write 0xBC
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunADC_LDO3P3()
        {
            double Result = 0.0;
            double Target = 3.3, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "ADC_LDO3P3";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_ADC_LDO3P3<3:0>"); // 0x96[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 3.3"); // offset 3.3V
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 6; // ADC LDO 3.3
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_LDO3P3.Read();
            TRIM_LDO3P3.Value = 0x08;
            TRIM_LDO3P3.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO3P3.Value = BitPos;
                TRIM_LDO3P3.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO3P3.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO3P3.Value = reg_value;
            TRIM_LDO3P3.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x30] = (byte)TRIM_LDO3P3.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x30] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write(); // Write 0xBC
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunFSK_LDO3P3()
        {
            double Result = 0.0;
            double Target = 3.3, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "FSK_LDO3P3";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_FSK_LDO3P3<3:0>"); // 0x98[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 3.3"); // offset 3.3V
            DIG_SEL.Read();
            DIG_SEL.Value = 0;
            ANA_SEL.Value = 5; // FSK LDO 3.3
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 0;
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_LDO3P3.Read();
            TRIM_LDO3P3.Value = 0x08;
            TRIM_LDO3P3.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO3P3.Value = BitPos;
                TRIM_LDO3P3.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO3P3.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO3P3.Value = reg_value;
            TRIM_LDO3P3.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN3"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x34] = (byte)TRIM_LDO3P3.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x34] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_SEL.Write(); // Write 0xBC
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunLDO1P8()
        {
            double Result = 0.0;
            double Target = 1.8, Margin = 0.04;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "LDO1P8";
            RegisterItem TRIM_LDO1P8 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO18<3:0>"); // 0x9A[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            TRIM_LDO1P8.Read();
            TRIM_LDO1P8.Value = 0x08;
            TRIM_LDO1P8.Write();
            for (uint BitPos = 5; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO1P8.Value = BitPos;
                TRIM_LDO1P8.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_LDO1P8.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_LDO1P8.Value = reg_value;
            TRIM_LDO1P8.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x38] = (byte)TRIM_LDO1P8.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x38] = (byte)0xFF;
            }
        }

        private void CAL_RunOSC()
        {
            double Result = 0.0;
            double Target = 16000000, Margin = 100000;
            double min = 20000000.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "OSC";
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>"); // 0xBC[7:4]
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>"); // 0xBC[3:0]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MUX_EN"); // 0xBB[7]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MUX_EN"); // 0xBB[6]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN"); // 0xBB[5]
            RegisterItem TRIM_OSC = Parent.RegMgr.GetRegisterItem("PU_TRIM_FOSC_16M<7:0>"); // 0x92[7:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
            DIG_SEL.Read();
            DIG_SEL.Value = 0; // OSC
            ANA_SEL.Value = 0; // BGR
            ANA_SEL.Write(); // Write 0xBC
            DIG_MUX_PEN.Read();
            DIG_MUX_PEN.Value = 1;
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB

            TRIM_OSC.Read();
            TRIM_OSC.Value = 0;
            TRIM_OSC.Write();
            for (uint BitPos = 10; BitPos < 21; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_OSC.Value = BitPos;
                TRIM_OSC.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
                Result += double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
                Result /= 2;
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_OSC.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_OSC.Value = reg_value;
            TRIM_OSC.Write();
            JLcLib.Delay.Sleep(50);
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
            Result += double.Parse(Oscilloscope.WriteAndReadString("MEAS:FREQ? CHAN4"));
            Result /= 2;
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x3C] = (byte)TRIM_OSC.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x3C] = 0xFF;
            }
            DIG_MUX_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0xBB
        }

        private void CAL_RunISENVREF()
        {
            double Vrect = 4.9;
            double Result = 0.0;
            double Target = 300, Margin = 10;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.05; // pass/fail limit
            string Name = "ML_ISEN_VREF";
            RegisterItem TRIM_ISEN = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>"); // 0xA2[7:4]
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>"); // 0xA1[7:4]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));

            ElectronicLoad.Write("CURR 0.1");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            TRIM_VILIM.Read();
            TRIM_VILIM.Value = 0;
            TRIM_VILIM.Write();

            TRIM_ISEN.Read();
            TRIM_ISEN.Value = 0x07;
            TRIM_ISEN.Write();
            for (uint BitPos = 0; BitPos < 5; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_ISEN.Value = BitPos;
                TRIM_ISEN.Write();
                JLcLib.Delay.Sleep(50);
                Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_ISEN.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_ISEN.Value = reg_value;
            TRIM_ISEN.Write();
            JLcLib.Delay.Sleep(50);
            Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
#if false
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)TRIM_ISEN.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)0xFF;
            }
#else
            if (Result >= 0 && Result <= 800)
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)TRIM_ISEN.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_ISEN.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x40] = (byte)0xFF;
            }
#endif
            ElectronicLoad.Write("CURR 0");
        }

        private void CAL_RunVILIM()
        {
            double Vrect = 5.3;
            double Result = 0.0;
            double Target = 855, Margin = 10;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "ML_VILIM";
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>"); // 0xA1[7:4]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));

            ElectronicLoad.Write("CURR 0.6");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            TRIM_VILIM.Read();
            TRIM_VILIM.Value = 0x07;
            TRIM_VILIM.Write();
            for (uint BitPos = 0; BitPos < 10; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_VILIM.Value = BitPos;
                TRIM_VILIM.Write();
                JLcLib.Delay.Sleep(50);
                Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_VILIM.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_VILIM.Value = reg_value;
            TRIM_VILIM.Write();
            JLcLib.Delay.Sleep(50);
            Result = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
#if false
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)TRIM_VILIM.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)0xFF;
            }
#else
            if (Result >= 700 && Result <= 1300)
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)TRIM_VILIM.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x44] = (byte)0xFF;
            }
#endif
            ElectronicLoad.Write("CURR 0");
        }

        private void CAL_RunVRECT()
        {
            double LowVrect = 4, HighVrect = 8;
            int LowVrectCode, HighVrectCode;
            short Slope, Const;
            string Name = "VRECT";

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            // 4V low VRECT calibration
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + LowVrect.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            LowVrectCode = RunTest(TEST_ITEMS.ADC_GET_VRECT, 1023); // 1024 average
            Log.WriteLine(Name + ":LOW:" + LowVrect.ToString("F1") + ":" + LowVrectCode.ToString("F3"));
            if (!IsRunCal)
                return;
            // 8V high VRECT calibration
            PowerSupply0.Write("VOLT " + HighVrect.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            HighVrectCode = RunTest(TEST_ITEMS.ADC_GET_VRECT, 1023); // 1024 average
            Log.WriteLine(Name + ":HIGH:" + HighVrect.ToString("F1") + ":" + HighVrectCode.ToString("F3"));

            Slope = (short)((HighVrect - LowVrect) * 1000 * 1024 / (HighVrectCode - LowVrectCode));
            Const = (short)(LowVrect * 1000 - (Slope * LowVrectCode) / 1024);

            if ((Slope >= 4250 && Slope <= 5750) && (Const >= -512 && Const <= 512))
            {
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 5000;
                Const = 0;
            }
            CalibrationData[0x48] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x49] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x4C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x4D] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunVOUT()
        {
            double Result, Vrect = 8.0;
            double[] TargetVoutValues = new double[2] { 3.5, 5.0 };
            double[] VoutValues = new double[2];
            double[] VoutSetValues = new double[2];
            double[] Limits = new double[2] { TargetVoutValues[0] * 0.03, TargetVoutValues[1] * 0.03 };
            int[] VoutCodes = new int[2];
            double Margin = 0.015;
            double min = 4095.0;
            uint reg_value = 0;
            short Slope, Const;
            string Name = "VOUT";
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN"); // 0x9E[3:3]
            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]

            // Sets 8V VRECT
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            OCL_EN.Read();
            OCL_EN.Value = 0;
            OCL_EN.Write();

            RunTest(TEST_ITEMS.LDO_TURNON, 0);
            //ElectronicLoad.Write("CURR 0.1");
            //ElectronicLoad.Write("INP 1");

            VOUT_SET_RANGE.Read();
            TRIM_VOUT_SET.Read();
            for (uint i = 0; i < 2; i++)
            {
                min = 4095.0;

                VOUT_SET_RANGE.Value = i;
                VOUT_SET_RANGE.Write();

                Oscilloscope.Write(":CHAN2:SCAL 2"); // 2V/div
                Oscilloscope.Write(":CHAN2:OFFS 8"); // offset 8V
                TRIM_VOUT_SET.Value = 0x40;
                TRIM_VOUT_SET.Write();
                JLcLib.Delay.Sleep(200);

                for (uint BitPos = 85; BitPos < 100; BitPos++)
                {
                    if (!IsRunCal)
                        return;
                    if (i == 0)
                    {
                        TRIM_VOUT_SET.Value = BitPos;
                    }
                    else
                    {
                        TRIM_VOUT_SET.Value = BitPos - 40;
                    }
                    TRIM_VOUT_SET.Write();
                    JLcLib.Delay.Sleep(50);
                    Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                    if (Math.Abs(TargetVoutValues[i] - Result) < min)
                    {
                        min = Math.Abs(TargetVoutValues[i] - Result);
                        reg_value = TRIM_VOUT_SET.Value;
                    }
                    Log.WriteLine(Name + i.ToString() + ":" + BitPos.ToString("D2") + ":" + TRIM_VOUT_SET.Value.ToString("D3") + ":" + Result.ToString("F3"));
                }
                TRIM_VOUT_SET.Value = reg_value;
                TRIM_VOUT_SET.Write();
                JLcLib.Delay.Sleep(50);
                VoutValues[i] = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                VoutSetValues[i] = TRIM_VOUT_SET.Value;
                VoutCodes[i] = RunTest(TEST_ITEMS.ADC_GET_VOUT, 1023); // 1024 average

                if (VoutValues[i] >= TargetVoutValues[i] - Limits[i] && VoutValues[i] <= TargetVoutValues[i] + Limits[i])
                    Log.WriteLine(Name + "_LVL:PASS:" + reg_value.ToString() + ":" + VoutValues[i].ToString() + ":" + VoutCodes[i].ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
                else
                    Log.WriteLine(Name + "_LVL:FAIL:" + reg_value.ToString() + ":" + VoutValues[i].ToString() + ":" + VoutCodes[i].ToString(), Color.Coral, Log.RichTextBox.BackColor);
            }
            //RunTest(TEST_ITEMS.LDO_TURNOFF, 0);
            CalibrationData[0xB0] = (byte)TRIM_VOUT_SET.Value;
            //ElectronicLoad.Write("CURR 0");

            // Vout voltage calibration
            Slope = (short)((VoutValues[1] - VoutValues[0]) * 1000 * 1024 / (VoutCodes[1] - VoutCodes[0]));
            Const = (short)(VoutValues[0] * 1000 - (Slope * VoutCodes[0]) / 1024);
            if ((Slope >= 2000 && Slope <= 3000) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 2500;
                Const = 0;
            }
            CalibrationData[0x58] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x59] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x5C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x5D] = (byte)((Const >> 8) & 0xFF);

            // Vout setting register calibration
            Slope = (short)((VoutSetValues[1] + 100 - VoutSetValues[0]) * 8192 / ((TargetVoutValues[1] - TargetVoutValues[0]) * 1000));
            Const = (short)(VoutSetValues[0] - (Slope * TargetVoutValues[0] * 1000) / 8192);
            if ((Slope >= 300 && Slope <= 400) && (Const >= -200 && Const <= 200))
                Log.WriteLine(Name + "_SET:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_SET:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 341;
                Const = -82;
            }
            CalibrationData[0x50] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x51] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x54] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x55] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunADC()
        {
            double LowVolt = 0.3, HighVolt = 1.5;
            int LowVoltCode, HighVoltCode;
            short Slope, Const;
            string Name = "ADC";

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
#if (POWER_SUPPLY_E36313A) // 3 port Power supply
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT " + LowVolt.ToString("F1"));
            PowerSupply0.Write("OUTP 1");
#else // 2 port power supply
            PowerSupply1.Write("INST:NSEL 1");
            PowerSupply1.Write("VOLT " + LowVolt.ToString("F1"));
            PowerSupply1.Write("OUTP 1");
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            LowVoltCode = RunTest(TEST_ITEMS.ADC_GET_NTC, 1023); // 1024 average
            Log.WriteLine(Name + ":LOW:" + LowVolt.ToString("F1") + ":" + LowVoltCode.ToString("F3"));
            if (!IsRunCal)
                return;
            // 8V high VRECT calibration
#if (POWER_SUPPLY_E36313A) // 3 port power supply
            PowerSupply0.Write("VOLT " + HighVolt.ToString("F1"));
#else // 2 power power supply
            PowerSupply1.Write("VOLT " + HighVolt.ToString("F1"));
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            HighVoltCode = RunTest(TEST_ITEMS.ADC_GET_NTC, 1023); // 1024 average
            Log.WriteLine(Name + ":HIGH:" + HighVolt.ToString("F1") + ":" + HighVoltCode.ToString("F3"));

            Slope = (short)((HighVolt - LowVolt) * 1000 * 1024 / (HighVoltCode - LowVoltCode));
            Const = (short)(LowVolt * 1000 - (Slope * LowVoltCode) / 1024);

            if ((Slope >= 750 && Slope <= 900) && (Const >= -64 && Const <= 64))
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 825;
                Const = 0;
            }
            CalibrationData[0x60] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x61] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x64] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x65] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunIOUT()
        {
            double Result;
            double[] VrectValues = new double[2] { 4.9, 5.5 }; // unit V
            double[] IoutValues = new double[3] { 0, 0.15, 1 }; // unit A
            int[,] ISenCodes = new int[IoutValues.Length, VrectValues.Length];
            double Margin = 0.03;
            short Slope, Const;
            string Name = "IOUT";
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN"); // 0x9E[3:3]
            RegisterItem TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR<4:0>"); // 0x91[4:0]
            RegisterItem TRIM_LDO5P0 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO5P0<3:0>"); // 0x94[3:0]
            RegisterItem TRIM_LDO1P5 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO1P5<3:0>"); // 0x9C[3:0]
            RegisterItem TRIM_ADC_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_ADC_LDO3P3<3:0>"); // 0x96[3:0]
            RegisterItem TRIM_FSK_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_FSK_LDO3P3<3:0>"); // 0x98[3:0]
            RegisterItem TRIM_LDO1P8 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO18<3:0>"); // 0x9A[3:0]
            RegisterItem TRIM_OSC = Parent.RegMgr.GetRegisterItem("PU_TRIM_FOSC_16M<7:0>"); // 0x92[7:0]
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>"); // 0xA1[7:4]
            RegisterItem VOUT_SET_RANGE = Parent.RegMgr.GetRegisterItem("ML_REF_VOUT_CT"); // 0x8B[2:2]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<6:0>"); // 0xA0[6:0]
            RegisterItem TRIM_ISEN = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>"); // 0xA21[7:4]

            // Sets 8V VRECT
            /*
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 1.8");
            PowerSupply0.Write("CURR 0.3");
            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 5.3");
            PowerSupply0.Write("CURR 2");
#if (POWER_SUPPLY_E36313A)
            PowerSupply0.Write("OUTP ON,(@1:2)");
#else
            PowerSupply0.Write("OUTP ON");
#endif
            */
            PowerSupply0.Write("INST:NSEL 2");

            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            //OCL_EN.Value = 0;
            //OCL_EN.Write();

            RunTest(TEST_ITEMS.LDO_TURNON, 0);
            // Set 5V VOUT
            /*TRIM_BGR.Value = CalibrationData[0x24];
            TRIM_BGR.Write();
            TRIM_LDO5P0.Value = CalibrationData[0x28];
            TRIM_LDO5P0.Write();
            TRIM_LDO1P5.Value = CalibrationData[0x2C];
            TRIM_LDO1P5.Write();
            TRIM_ADC_LDO3P3.Value = CalibrationData[0x30];
            TRIM_ADC_LDO3P3.Write();
            TRIM_FSK_LDO3P3.Value = CalibrationData[0x34];
            TRIM_FSK_LDO3P3.Write();
            TRIM_LDO1P8.Value = CalibrationData[0x38];
            TRIM_LDO1P8.Write();
            TRIM_OSC.Value = CalibrationData[0x3C];
            TRIM_OSC.Write();
            TRIM_VILIM.Value = CalibrationData[0x44];
            TRIM_VILIM.Write();
            VOUT_SET_RANGE.Value = 1;
            VOUT_SET_RANGE.Write();
            TRIM_VOUT_SET.Value = CalibrationData[0xB0];
            TRIM_VOUT_SET.Write();
            TRIM_ISEN.Read();
            TRIM_ISEN.Value = CalibrationData[0x40];
            TRIM_ISEN.Write();*/

            for (int i = 0; i < IoutValues.Length; i++) // Zero(0)/Low(1)/High(2) current
            {
                ElectronicLoad.Write("CURR " + IoutValues[i].ToString("F2"));
                for (int j = 0; j < VrectValues.Length; j++) // Low(0)/High(1) VRECT
                {
                    PowerSupply0.Write("VOLT " + VrectValues[j].ToString("F1")); // Set VRECT voltage
                    JLcLib.Delay.Sleep(500);
                    ISenCodes[i, j] = RunTest(TEST_ITEMS.ADC_GET_IOUT, 1023); // 1024 average
                    Log.WriteLine(Name + ":" + IoutValues[i].ToString("F2") + ":" + VrectValues[j].ToString("F1") + ":" + ISenCodes[i, j].ToString());
                }
            }
            //RunTest(TEST_ITEMS.LDO_TURNOFF, 0);
            ElectronicLoad.Write("CURR 0");

            // Store ADC Iout codes zero load current and difference between high and low Vrect at high load current
            //Const = (short)ISenCodes[0, 1]; // ADC code of zero load current, high Vrect
            //CalibrationData[0x70] = (byte)((Const >> 0) & 0xFF);
            //CalibrationData[0x71] = (byte)((Const >> 8) & 0xFF);
            //Const = (short)ISenCodes[1, 1]; // ADC code of low load current, high Vrect
            //CalibrationData[0x74] = (byte)((Const >> 0) & 0xFF);
            //CalibrationData[0x75] = (byte)((Const >> 8) & 0xFF);
            Const = (short)IoutValues[0]; // Low load current
            CalibrationData[0x78] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x79] = (byte)((Const >> 8) & 0xFF);
            Log.WriteLine(Name + "_Low:PASS:" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            Const = (short)(ISenCodes[2, 0] - ISenCodes[2, 1]);
            CalibrationData[0x7C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x7D] = (byte)((Const >> 8) & 0xFF);
            Log.WriteLine(Name + "_Diff:PASS:" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);

            // Calculate Isen0
            Slope = (short)((IoutValues[2] - IoutValues[0]) * 1000 * 1024 / (ISenCodes[2, 0] - ISenCodes[0, 0]));
            Const = (short)(IoutValues[0] * 1000 - (Slope * ISenCodes[0, 0]) / 1024);
            Log.WriteLine(Name + "_0:" + Slope.ToString() + ":" + Const.ToString(), Color.DarkGray, Log.RichTextBox.BackColor);

            if ((Slope >= 500 && Slope <= 1500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_0:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_0:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x80] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x81] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x84] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x85] = (byte)((Const >> 8) & 0xFF);

            // Calculate Isen1
            Slope = (short)((IoutValues[2] - IoutValues[1]) * 1000 * 1024 / (ISenCodes[2, 1] - ISenCodes[1, 1]));
            Const = (short)(IoutValues[1] * 1000 - (Slope * ISenCodes[1, 1]) / 1024);
            Log.WriteLine(Name + "_1:" + Slope.ToString() + ":" + Const.ToString(), Color.DarkGray, Log.RichTextBox.BackColor);
            if ((Slope >= 500 && Slope <= 1500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_1:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_1:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x88] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x89] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x8C] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x8D] = (byte)((Const >> 8) & 0xFF);

            // Calculate Isen2
            Slope = (short)((IoutValues[1] - IoutValues[0]) * 1000 * 1024 / (ISenCodes[1, 1] - ISenCodes[0, 1]));
            Const = (short)(IoutValues[0] * 1000 - (Slope * ISenCodes[0, 1]) / 1024);
            if ((Slope >= 500 && Slope <= 1500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_2:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_2:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x70] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x71] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x74] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x75] = (byte)((Const >> 8) & 0xFF);
        }
        #endregion Calibration methods
    }
}
