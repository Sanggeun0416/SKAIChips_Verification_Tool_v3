using JLcLib.Chip;
using JLcLib.Comn;
using SKAIChips_Verification;
using System;
using System.Text;
using System.Windows.Forms;

namespace SKAI_IRIS
{
    public class SCP1501_N1C : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS_MANUAL
        {
            Write_WURF,
            TX_CH_SEL,
            TX_ON,
            TX_OFF,
            SET_BLE,
            Read_FSM,
            Read_ADC,
            Read_Volt_Temp,
            TESTINOUT_Temp,
            TESTINOUT_Volt,
            TESTINOUT_BGR,
            TESTINOUT_EDOUT,
            TESTINOUT_RCOSC,
            TESTINOUT_Disable,
            NVM_POWER_ON,
            NVM_POWER_OFF,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            LAB_TEST,
            DUT_TempSensor,
            THB_SMT,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_DTM
        {
            SET_CH,
            SET_Length,
            DTM_START_PRNBS9,
            DTM_START_11110000,
            DTM_START_10101010,
            DTM_START_PRNBS15,
            DTM_START_ALL_1,
            DTM_START_ALL_0,
            DTM_START_00001111,
            DTM_START_0101,
            DTM_STOP,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
            DTM,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        public int SlaveAddress { get; private set; } = 0x3A;
        public int DTM_PayLoadLength = 37;
        public int DTM_Channel = 0;

        /* Intrument */
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

        public SCP1501_N1C(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "SCP1501_N1C");
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

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[8];
            I2C.Config.SlaveAddress = SlaveAddress;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "NVM":
                    // bist_sel=1, bist_type=3, bist_cmd=0, bist_addr=address
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0d);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=3, bist_cmd=1, bist_addr=address
                    SendData[4] = (byte)(0x0f);
                    SendData[5] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_wdata=data
                    SendData[0] = (byte)(0x04);            // ADDR = 0x00060004
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=3, bist_cmd=0, bist_addr=address
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0d);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Dummy
                    System.Threading.Thread.Sleep(5);

                    // bist_sel=0, bist_type=3, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0c);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                case "NVM Controller":
                case "NVM BIST":
                case "I2C Slave":
                case "LPCAL":
                case "ADC Controller":
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                default:
                    // Write Address, Data
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x07);            // AREG_SEL=1, AREG_CE=1, AREG_WE=1
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);            // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);            // AREG_WDATA[7:0]
                    SendData[5] = (byte)(0x00);            // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);            // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
            }
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            uint result = 0xffffffff;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "NVM":
                    // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x11);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
                    SendData[4] = (byte)(0x13);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
                    SendData[4] = (byte)(0x11);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Read bist_rdata
                    SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);

                    // bist_sel=0, bist_type=4, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x10);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                case "NVM Controller":
                case "NVM BIST":
                case "I2C Slave":
                case "LPCAL":
                case "ADC Controller":
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);
                    break;

                default:
                    // Write Address
                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);             // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);   // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    SendData[6] = (byte)(0x05);             // AREG_SEL=1, AREG_CE=1, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Read Data
                    SendData[0] = (byte)(0x10);             // ADDR = 0x00010010
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = RcvData[3];

                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[5] = (byte)(0x00);             // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);             // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
            }

            return result;
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
            Parent.ChipCtrlButtons[4].Text = "GH0_H";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;
            Parent.ChipCtrlButtons[5].Text = "WakeUp";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += WakeUp_I2C;
            Parent.ChipCtrlButtons[6].Text = "MD_OOK";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Set_Gecko_OOK_Mode;
            Parent.ChipCtrlButtons[8].Text = "Manual";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;
            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_Auto_Test_Items;
            Parent.ChipCtrlButtons[10].Text = "DTM";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Change_To_DTM_Test_Items;
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
            Serial.WriteSettingFile(Parent.IniFile, "SCP1501_N1C");
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

        private void Change_To_DTM_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.DTM;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_DTM.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_DTM)i).ToString());
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

        private void WakeUp_I2C(object sender, EventArgs e)
        {
            WakeUp_I2C();
        }

        private void Set_Gecko_OOK_Mode(object sender, EventArgs e)
        {
            SendCommand("mode ook");
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
                        case TEST_ITEMS_MANUAL.Write_WURF:
                            if (Arg == "")
                            {
                                Log.WriteLine("WURF에 write할 값을 16진수로 적어주세요.");
                            }
                            else
                            {
                                Log.WriteLine("OOK " + Arg);
                                Write_WURF_AON(Arg);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_CH_SEL:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                Log.WriteLine("Set CH" + Arg);
                                Write_Register_Fractional_Calc_Ch((uint)iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_ON:
                            Log.WriteLine("TX ON");
                            Write_Register_Tx_Tone_Send(true);
                            break;
                        case TEST_ITEMS_MANUAL.TX_OFF:
                            Log.WriteLine("TX OFF");
                            Write_Register_Tx_Tone_Send(false);
                            break;
                        case TEST_ITEMS_MANUAL.SET_BLE:
                            if ((Arg == "") || (iVal < 0) || (iVal > 65535))
                            {
                                Log.WriteLine("WP값을 10진수로 적어주세요. (Range 0~65535)");
                            }
                            else
                            {
                                Log.WriteLine("Write AON for Advertising (WP = )" + Arg);
                                Write_Register_Send_Advertising(iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.Read_FSM:
                            Read_Register_FSM();
                            break;
                        case TEST_ITEMS_MANUAL.Read_ADC:
                            Result = Read_ADC_Result(false, 1, 63, 31);
                            Disable_ADC();
                            Log.WriteLine("Volt : " + (Result >> 8).ToString() + "\tTemp : " + (Result & 0xff).ToString());
                            break;
                        case TEST_ITEMS_MANUAL.Read_Volt_Temp:
                            Calculation_VBAT_Voltage_And_Temperature();
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Temp:
                            Set_TestInOut_For_VTEMP(true);
                            Read_ADC_Result(false, 1, 1, 1);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Volt:
                            Set_TestInOut_For_VS(true);
                            Read_ADC_Result(true, 1, 1, 1);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_BGR:
                            Set_TestInOut_For_BGR(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_EDOUT:
                            Set_TestInOut_For_EDOUT(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_RCOSC:
                            Set_TestInOut_For_RCOSC(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Disable:
                            Set_TestInOut_For_VTEMP(false);
                            Set_TestInOut_For_VS(false);
                            Set_TestInOut_For_BGR(false);
                            Set_TestInOut_For_EDOUT(false);
                            Set_TestInOut_For_RCOSC(false);
                            Disable_ADC();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_ON:
                            Enable_NVM_BIST();
                            Power_On_NVM();
                            Disable_NVM_BIST();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_OFF:
                            Enable_NVM_BIST();
                            Power_Off_NVM();
                            Disable_NVM_BIST();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((TEST_ITEMS_AUTO)TestItemIndex)
                    {
                        case TEST_ITEMS_AUTO.LAB_TEST:
                            Test_Good_Chip_Sorting_Rev3(iVal);
                            break;
                        case TEST_ITEMS_AUTO.DUT_TempSensor:
                            DUT_TempSensor_Test();
                            break;
                        case TEST_ITEMS_AUTO.THB_SMT:
                            Test_THB_SMT_Board(iVal);
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.DTM:
                    switch ((TEST_ITEMS_DTM)TestItemIndex)
                    {
                        case TEST_ITEMS_DTM.SET_CH:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                DTM_Channel = iVal;
                                Log.WriteLine("Channel : " + DTM_Channel.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.SET_Length:
                            if ((Arg == "") || ((iVal != 37) && (iVal != 255)))
                            {
                                Log.WriteLine("Length를 10진수로 적어주세요. (Range 37 또는 255)");
                            }
                            else
                            {
                                DTM_PayLoadLength = iVal;
                                Log.WriteLine("Length : " + DTM_PayLoadLength.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.DTM_START_PRNBS9:
                        case TEST_ITEMS_DTM.DTM_START_11110000:
                        case TEST_ITEMS_DTM.DTM_START_10101010:
                        case TEST_ITEMS_DTM.DTM_START_PRNBS15:
                        case TEST_ITEMS_DTM.DTM_START_ALL_1:
                        case TEST_ITEMS_DTM.DTM_START_ALL_0:
                        case TEST_ITEMS_DTM.DTM_START_00001111:
                        case TEST_ITEMS_DTM.DTM_START_0101:
                            Log.WriteLine("Start DTM\nChannel : " + DTM_Channel.ToString() + "\tLength : " + DTM_PayLoadLength.ToString());
                            Run_BLE_DTM_MODE((uint)(TestItemIndex - TEST_ITEMS_DTM.DTM_START_PRNBS9));
                            break;
                        case TEST_ITEMS_DTM.DTM_STOP:
                            Log.WriteLine("Stop DTM");
                            Stop_BLE_DTM_MODE();
                            break;
                    }
                    break;
                default:
                    break;
            }
        }

        #region Function for Chip Test
        private void WakeUp_I2C()
        {
            byte[] SendData = new byte[2];

            SendData[0] = 0xAA;
            SendData[1] = 0xBB;

            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_CMD_I2C(byte data)
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=1, WURF_WRITE=1
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = data;
            SendData[5] = 0x00;
            SendData[6] = 0x06;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // WURF_WRITE=0
            SendData[6] = 0x02;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_WURF_Sel()
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=0
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_AON(string Arg)
        {
            RegisterItem w_WURF_END = Parent.RegMgr.GetRegisterItem("w_WURF_END");         // 0x38
            byte data;

            switch (Arg)
            {
                case "30":
                    data = 0x30;
                    break;
                case "60":
                    data = 0x60;
                    break;
                case "61":
                    data = 0x61;
                    break;
                case "70":
                    data = 0x70;
                    break;
                case "80":
                    data = 0x80;
                    break;
                case "b0":
                case "B0":
                    data = 0xb0;
                    break;
                case "c0":
                case "C0":
                    data = 0xb0;
                    break;
                default:
                    Log.WriteLine("지원하지 않는 CMD 입니다.");
                    return;
                    break;
            }

            w_WURF_END.Read();
            w_WURF_END.Value = 0;
            w_WURF_END.Write();

            Write_WURF_CMD_I2C(data);

            w_WURF_END.Value = 1;
            w_WURF_END.Write();

            Disable_WURF_Sel();

            w_WURF_END.Value = 0;
            w_WURF_END.Write();
        }

        private void Write_Register_Fractional_Calc_Ch(uint ch)
        {
            uint vco;

            RegisterItem SPI_DSM_F_H = Parent.RegMgr.GetRegisterItem("O_SPI_DSM_F[22:16]"); // 0x3E
            RegisterItem SPI_PS_P = Parent.RegMgr.GetRegisterItem("O_SPI_PS_P[4:0]");       // 0x3F
            RegisterItem SPI_PS_S = Parent.RegMgr.GetRegisterItem("O_SPI_PS_S[2:0]");       // 0x3F                

            vco = 2402 + 2 * ch;
            SPI_PS_P.Value = vco >> 7;
            SPI_PS_S.Value = (vco >> 5) - (SPI_PS_P.Value << 2);
            SPI_DSM_F_H.Value = (vco << 1) - (SPI_PS_P.Value << 8) - (SPI_PS_S.Value << 6);

            SPI_PS_P.Write();
            SPI_DSM_F_H.Write();
        }

        private void Write_Register_Tx_Tone_Send(bool on)
        {
            RegisterItem EXT_CH_MODE = Parent.RegMgr.GetRegisterItem("O_EXT_CH_MODE");             // 0x3E
            RegisterItem TX_SEL = Parent.RegMgr.GetRegisterItem("O_TX_SEL");                       // 0x56
            RegisterItem TX_BUF_PEN = Parent.RegMgr.GetRegisterItem("O_TX_BUF_PEN");               // 0x57
            RegisterItem PRE_DA_PEN = Parent.RegMgr.GetRegisterItem("O_PRE_DA_PEN");               // 0x57
            RegisterItem DA_PEN = Parent.RegMgr.GetRegisterItem("O_DA_PEN");                       // 0x57
            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("O_PLL_PEN");                     // 0x59

            RegisterItem EXT_DA_GAIN_SEL = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_SEL");     // 0x5B
            RegisterItem w_TX_BUF_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_TX_BUF_PEN_MODE");   // 0x5C
            RegisterItem w_PRE_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PRE_DA_PEN_MODE");   // 0x5C
            RegisterItem w_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_DA_PEN_MODE");           // 0x5C
            RegisterItem w_TX_SEL_MODE = Parent.RegMgr.GetRegisterItem("w_TX_SEL_MODE");           // 0x5C
            RegisterItem w_PLL_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PLL_PEN_MODE");         // 0x5D
            RegisterItem XTAL_PLL_CLK_EN_MD = Parent.RegMgr.GetRegisterItem("w_XTAL_PLL_CLK_EN_MD");   // 0x5F

            if (on == true)
            {
                // Tx Active
                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 1;
                EXT_CH_MODE.Write();

                TX_SEL.Read();
                TX_SEL.Value = 1;
                TX_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 1;
                PRE_DA_PEN.Value = 1;
                DA_PEN.Value = 1;
                TX_BUF_PEN.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 0;
                EXT_DA_GAIN_SEL.Write();

                // Dynamic -> Static
                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 1;
                w_PLL_PEN_MODE.Write();

                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 1;
                w_PRE_DA_PEN_MODE.Value = 1;
                w_DA_PEN_MODE.Value = 1;
                w_TX_SEL_MODE.Value = 1;
                w_TX_BUF_PEN_MODE.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 1;
                XTAL_PLL_CLK_EN_MD.Write();

                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                PLL_PEN.Value = 1;
                PLL_PEN.Write();
            }
            else
            {
                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 0;
                XTAL_PLL_CLK_EN_MD.Write();

                // Static -> Dynamic
                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 0;
                w_PRE_DA_PEN_MODE.Value = 0;
                w_DA_PEN_MODE.Value = 0;
                w_TX_SEL_MODE.Value = 0;
                w_TX_BUF_PEN_MODE.Write();

                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 0;
                w_PLL_PEN_MODE.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 1;
                EXT_DA_GAIN_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 0;
                PRE_DA_PEN.Value = 0;
                DA_PEN.Value = 0;
                TX_BUF_PEN.Write();

                TX_SEL.Read();
                TX_SEL.Value = 0;
                TX_SEL.Write();

                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 0;
                EXT_CH_MODE.Write();
            }
        }

        private void Write_Register_Send_Advertising(int iVal)
        {
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");     // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");     // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");     // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");     // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");     // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");     // 0x1B
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");     // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");     // 0x1F

            B22_WP_0.Value = (uint)(iVal & 0xff);
            B22_WP_0.Write();

            B23_WP_1.Value = (uint)((iVal >> 8) & 0xff);
            B23_WP_1.Write();

            B24_AI_0.Value = 210;
            B24_AI_0.Write();

            B25_AI_1.Value = 29;
            B25_AI_1.Write();

            B26_AI_2.Value = 0;
            B26_AI_2.Write();

            B27_AI_3.Value = 0;
            B27_AI_3.Write();

            B30_AD_0.Value = 3;
            B30_AD_0.Write();

            B31_AD_1.Value = 0;
            B31_AD_1.Write();
        }

        private uint Run_Read_FSM_Status()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[1] & 0xf) << 8) | RcvData[0]);
        }

        private uint CRC_FLAG_READ()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[2] & 0x3) << 16) | (RcvData[1] << 8) | RcvData[0]);
        }

        private void Read_Register_FSM()
        {
            uint u_fsm;
            uint crc_flag;
            u_fsm = Run_Read_FSM_Status();
            crc_flag = CRC_FLAG_READ();
            Log.WriteLine("FSM_Status : " + (u_fsm & 0x000f).ToString());
            Log.WriteLine("FSM_State : " + ((u_fsm & 0x0ff0) >> 4).ToString());
            Log.WriteLine("WUR_CRC_ERROR : " + ((crc_flag >> 17) & 0xff).ToString());
            Log.WriteLine("WUR_CRC_DEC : " + (crc_flag & 0xffff).ToString());
        }

        private uint Read_ADC_Result(bool temp_first, uint eoc_skip_temp, uint eoc_skip_volt, uint avg_cnt)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            // Set w_TEST_SELECT
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
#if true // w_TEST_SELECT = 1
            if (temp_first)
            {
                SendData[7] = 0x80; // temp -> volt
            }
            else
            {
                SendData[7] = 0xC0; // volt -> temp
            }
#else // w_TEST_SELECT = 0
            if (temp_first)
            {
                SendData[7] = 0x00; // temp -> volt
            }
            else
            {
                SendData[7] = 0x40; // volt -> temp
            }
#endif
            I2C.WriteBytes(SendData, SendData.Length, true);
            // Enable ADC (TEST_START = 0, I_PEN = 1)
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x01);
            SendData[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            SendData[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            SendData[7] = (byte)((avg_cnt & 0x1f) << 1);
            I2C.WriteBytes(SendData, SendData.Length, true);
            // Enable ADC (TEST_START = 1, I_PEN = 1)
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x01);
            SendData[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            SendData[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            SendData[7] = (byte)(((avg_cnt & 0x1f) << 1) | 0x80);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(10);
            // Read ADC_D_B[7:0], ADC_D_T[7:0]
            SendData[0] = 0x08;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            return (uint)(((RcvData[1] << 8) | RcvData[0]) & 0xffff);
        }

        private void Disable_ADC()
        {
            byte[] SendData = new byte[8];

            // Disable ADC
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private double Calculation_Temperature(uint otp, uint adc)
        {
            double[] lowtempcomp_ADC = { 57.11, 57.23, 57.34, 57.45, 57.56, 57.66, 57.75, 57.84 };
            double temperature;
            uint OTP23p5, OTP85TO23p5;

            OTP23p5 = (otp & 0x1f) + 104;
            OTP85TO23p5 = (otp >> 5) + 35;

            if (adc >= OTP23p5)
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) + 23.5;
            }
            else
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) * ((23.5 - (-40)) / lowtempcomp_ADC[otp >> 5]) + 23.5;
            }

            return temperature;
        }

        private double Calculation_VBAT_Voltage(double temperature, uint adc)
        {
            double voltage;

            if (temperature >= 30)
            {
                voltage = (adc - (11.475 - (-3 / (85 - 30)) * (temperature - 30))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }
            else
            {
                voltage = (adc - (11.475 - (2 / (30 - (-40))) * (30 - temperature))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }

            return voltage;
        }

        private void Calculation_VBAT_Voltage_And_Temperature()
        {
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");       // 0x14
            uint adc;
            double temp, volt;

            B20_DC_0.Read();
            adc = Read_ADC_Result(false, 1, 63, 31);
            Disable_ADC();
            temp = Calculation_Temperature(B20_DC_0.Value, adc & 0xff);
            volt = Calculation_VBAT_Voltage(temp, adc >> 8);
            Log.WriteLine("Volt : " + volt.ToString("F2") + "(" + (adc >> 8).ToString() + ")\tTemp : " + temp.ToString("F3") + "(" + (adc & 0xff).ToString() + ")");
        }

        private void Set_TestInOut_For_VTEMP(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");       // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL");     // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");            // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");            // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BUF_MUX_SEL.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_VS(bool on)
        {
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");        // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");        // 0x4C
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");   // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL"); // 0x5B

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0x20;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();
            }
        }

        private void Set_TestInOut_For_BGR(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");    // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");         // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");         // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_EDOUT(bool on)
        {
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[8]");   // 0x5A
            RegisterItem O_RX_DATAT = Parent.RegMgr.GetRegisterItem("O_RX_DATAT");      // 0x5A

            if (on == true)
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 1;
                O_RX_DATAT.Value = 1;
                ITEST_CONT.Write();
            }
            else
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 0;
                O_RX_DATAT.Value = 0;
                ITEST_CONT.Write();
            }
        }

        private void Set_TestInOut_For_RCOSC(bool on)
        {
            RegisterItem TEST_EN_32K = Parent.RegMgr.GetRegisterItem("O_TEST_EN_32K");     // 0x4C
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");    // 0x4C

            if (on == true)
            {
                TEST_EN_32K.Read();
                TEST_EN_32K.Value = 1;
                TEST_CON_H.Value = 8;
                TEST_EN_32K.Write();
            }
            else
            {
                TEST_EN_32K.Value = 0;
                TEST_CON_H.Value = 0;
                TEST_EN_32K.Write();
            }
        }

        private void Enable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x01;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=0, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_On_NVM()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=1, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x07;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=1, bist_cmd=0
            SendData[4] = 0x05;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_Off_NVM()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=2, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x0B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=2, bist_cmd=0
            SendData[4] = 0x09;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }
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

        private bool Check_Revision_Information(uint value)
        {
            RegisterItem REV_ID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");     // 0x68
            bool result = false;

            REV_ID.Read();

            if (value == REV_ID.Value)
            {
                result = true;
            }

            return result;
        }

        private uint Read_Mac_Address()
        {
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10

            B14_BA_0.Read();
            B15_BA_1.Read();
            B16_BA_2.Read();

            return (B16_BA_2.Value << 16) | (B15_BA_1.Value << 8) | B14_BA_0.Value;
        }

        private bool Check_OTP_VALID(bool start, uint page)
        {
            uint addr_nvm;
            bool result = true;

            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            byte[] Data = new byte[4];

            if (start == false)
            {
                return false;
            }

            Enable_NVM_BIST();
            Power_On_NVM();

            if (page < 19)
            {
                addr_nvm = page * 44; // ble, control
            }
            else
            {
                addr_nvm = page * 44 + (page - 19) * 4; // Analog
            }

            // NVM Verify // VALID
            Data[0] = 0x1D;
            Data[1] = 0xCA;
            Data[2] = 0x1D;
            Data[3] = 0xCA;
            // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
            SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            SendData[4] = (byte)(0x13);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
            SendData[4] = (byte)(0x11);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Read bist_rdata
            SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            for (int j = 0; j < 4; j++)
            {
                if (RcvData[j] != Data[j])
                {
                    result = false;
                }
            }

            Power_Off_NVM();
            Disable_NVM_BIST();

            return result;
        }

        private bool Read_FF_NVM(bool start)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return true;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // bist_sel=1, bist_type=6, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=6, bist_cmd=0
            SendData[4] = 0x19;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] != 0x0b)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void Write_Register_AON_Fix_Value()
        {
            RegisterItem w_ADC_EOC_SKIP_TEMP = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_TEMP[5:0]");  // 0x34
            RegisterItem w_ADC_EOC_SKIP_VOLT = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_VOLT[5:0]");  // 0x35
            RegisterItem w_LPCAL_FINE_EN = Parent.RegMgr.GetRegisterItem("w_LPCAL_FINE_EN");               // 0x37
            RegisterItem w_TAMP_DET_EN = Parent.RegMgr.GetRegisterItem("w_TAMP_DET_EN");                   // 0x39
            RegisterItem w_GPIO_WKUP_MD = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_MD");                 // 0x41
            RegisterItem w_GPIO_WKUP_INTV = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_INTV[1:0]");        // 0x41
            RegisterItem O_CT_RES2_FINE = Parent.RegMgr.GetRegisterItem("O_CT_RES2_FINE[3:0]");            // 0x45
            RegisterItem w_GPIO_WKUP_EN = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_EN");                 // 0x4E
            RegisterItem O_ADC_CM = Parent.RegMgr.GetRegisterItem("O_ADC_CM[1:0]");                        // 0x55
            RegisterItem w_DA_GAIN_MAX = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_MAX[2:0]");              // 0x57
            RegisterItem I_EXT_DA_GAIN_CON = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_CON[1:0]");      // 0x5A
            RegisterItem w_XTAL_CUR_CFG_1 = Parent.RegMgr.GetRegisterItem("w_XTAL_CUR_CFG_1[5:0]");        // 0x5E
            RegisterItem w_DA_GAIN_INTV = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_INTV[1:0]");            // 0x61
            RegisterItem w_PLL_CH_19_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_19_GAIN[2:0]");        // 0x61
            RegisterItem VS_GainMode = Parent.RegMgr.GetRegisterItem("VS_GainMode");                       // 0x61
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL[5:2]");                  // 0x62
            RegisterItem w_TX_SEL_SELECT = Parent.RegMgr.GetRegisterItem("w_TX_SEL_SELECT[1:0]");          // 0x64
            RegisterItem w_PLL_PM_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_PM_GAIN[4:0]");              // 0x64
            RegisterItem w_PLL_CH_0_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_0_GAIN[4:0]");          // 0x64
            RegisterItem w_DA_PEN_SELECT = Parent.RegMgr.GetRegisterItem("w_DA_PEN_SELECT[1:0]");          // 0x65
            RegisterItem w_PLL_CH_12_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_12_GAIN[4:0]");        // 0x66
            RegisterItem w_ADC_SWITCH_MODE = Parent.RegMgr.GetRegisterItem("w_ADC_SWITCH_MODE");           // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_H = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[4]");        // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_L = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[3:0]");      // 0x67

            BGR_TC_CTRL.Read();
            BGR_TC_CTRL.Value = 11;
            BGR_TC_CTRL.Write();

            w_ADC_SAMPLE_CNT_H.Read();
            w_ADC_SAMPLE_CNT_H.Value = 1;
            w_ADC_SAMPLE_CNT_H.Write();

            w_ADC_SAMPLE_CNT_L.Read();
            w_ADC_SAMPLE_CNT_L.Value = 15;
            w_ADC_SAMPLE_CNT_L.Write();

            O_CT_RES2_FINE.Read();
            O_CT_RES2_FINE.Value = 9;
            O_CT_RES2_FINE.Write();

            w_DA_GAIN_MAX.Read();
            w_DA_GAIN_MAX.Value = 6;
            w_DA_GAIN_MAX.Write();

            I_EXT_DA_GAIN_CON.Read();
            I_EXT_DA_GAIN_CON.Value = 2;
            I_EXT_DA_GAIN_CON.Write();

            VS_GainMode.Read();
            VS_GainMode.Value = 0;
            VS_GainMode.Write();

            w_TX_SEL_SELECT.Read();
            w_TX_SEL_SELECT.Value = 2;
            w_TX_SEL_SELECT.Write();

            w_DA_PEN_SELECT.Read();
            w_DA_PEN_SELECT.Value = 2;
            w_DA_PEN_SELECT.Write();
        }

        private bool Run_Cal_BGR(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 300;
            double d_lsl = 295, d_usl = 305;
            uint ldo_val, ldo_val_1;

            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");    // 0x53

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_BGR(true);

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(2, (1 + cnt), (double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));

                ULP_BGR_CONT.Read();
                Parent.xlMgr.Sheet.Select("BGR");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            ldo_val = 15;
            ldo_val_1 = 0;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }

            Set_TestInOut_For_BGR(false);

            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_ALLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 810;
            double d_lsl = 800, d_usl = 820;
            uint ldo_val, ldo_val_1;

            RegisterItem O_ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");        // 0x54
            RegisterItem O_ULP_LDO_LV_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[2:0]");  // 0x61

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            O_ULP_LDO_CONT.Read();
            O_ULP_LDO_LV_CONT.Read();

            O_ULP_LDO_CONT.Value = 8;
            O_ULP_LDO_CONT.Write();
            O_ULP_LDO_LV_CONT.Value = 0;
            O_ULP_LDO_LV_CONT.Write();
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("ALLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv > 816)
            {
                O_ULP_LDO_LV_CONT.Value = 3;
                O_ULP_LDO_LV_CONT.Write();
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (O_ULP_LDO_LV_CONT.Value).ToString());
                Parent.xlMgr.Cell.Write(x_pos + 1, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_MLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 1000;
            double d_lsl = 990, d_usl = 1010;
            uint ldo_val, ldo_val_1;

            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");          // 0x53
            RegisterItem PMU_MLDO_Coarse_L = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[0]");    // 0x62
            RegisterItem PMU_MLDO_Coarse_H = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1]");    // 0x63

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            PMU_MLDO_Coarse_H.Read();
            PMU_MLDO_Coarse_L.Read();
            PMU_LDO_CONT.Read();

            PMU_LDO_CONT.Value = 7;
            PMU_LDO_CONT.Write();
            PMU_MLDO_Coarse_L.Value = 0;
            PMU_MLDO_Coarse_L.Write();

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("MLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv < 1006)
            {
                PMU_MLDO_Coarse_L.Value = 1;
                PMU_MLDO_Coarse_L.Write();
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < 1006)
                {
                    PMU_MLDO_Coarse_L.Value = 0;
                    PMU_MLDO_Coarse_L.Write();
                    PMU_MLDO_Coarse_H.Value = 1;
                    PMU_MLDO_Coarse_H.Write();
                }
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();
            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            if (x_pos != 0)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            }
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ((PMU_MLDO_Coarse_H.Value << 1) | PMU_MLDO_Coarse_L.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_32K_RCOSC(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_khz, d_diff_khz;
            double d_lsl = 32.113, d_usl = 33.423, d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;

            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");      // 0x4D
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");     // 0x4E

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_RCOSC(true);
            if (x_pos != 0)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_khz.ToString("F3"));

                Parent.xlMgr.Sheet.Select("RCOSC");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            // RTC_SCKF[5:0]
            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            for (int val = 4; val >= 0; val--)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
                }
                if (d_freq_khz > d_target_khz)
                {
                    osc_val_l += (uint)(1 << val);
                }
                else
                {
                    osc_val_l -= (uint)(1 << val);
                }
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
            {
                if (osc_val_l != 63) osc_val_l += 1;
            }
            else
            {
                if (osc_val_l != 0) osc_val_l -= 1;
            }
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (RTC_SCKF_H.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), (RTC_SCKF_L.Value).ToString());
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_freq_khz.ToString("F3"));
            }
            Set_TestInOut_For_RCOSC(false);

            if ((d_freq_khz < d_lsl) || (d_freq_khz > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_Temp_Sensor(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 140; // 30deg
            int lsl = 135, usl = 145;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem TEMP_CONT_L = Parent.RegMgr.GetRegisterItem("O_TEMP_CONT[4:0]");  // 0x55
            RegisterItem TEMP_CONT_H = Parent.RegMgr.GetRegisterItem("TEMP_TRIM[5]");      // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
            TEMP_CONT_H.Read();
            TEMP_CONT_L.Read();

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, 1, 63, 31);
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();
            Set_TestInOut_For_VTEMP(false);
#if false // cal
            Parent.xlMgr.Sheet.Select("Temp_Sen");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
                u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
            }
#else // fix
            u_adc_val = 0;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            lsl = 0;
            usl = 255;
#endif
            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

            Set_TestInOut_For_VTEMP(false);
#if false
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_Voltage_Scaler(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 45; // 2.0V
            int lsl = 40, usl = 50;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem VOLSCAL_CON_L = Parent.RegMgr.GetRegisterItem("O_VOLSCAL_CON[3:0]");  // 0x56
            RegisterItem VOLTAGE_CON_H = Parent.RegMgr.GetRegisterItem("VOLTAGE_CON[5:4]");    // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
#if (POWER_SUPPLY_NEW)
			PowerSupply0.Write("VOLT 2.0,(@3)");
#else
            PowerSupply0.Write("VOLT 2.0");
#endif
            VOLTAGE_CON_H.Read();
            VOLSCAL_CON_L.Read();

            Set_TestInOut_For_VS(true);

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, 1, 63, 31);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();

            Parent.xlMgr.Sheet.Select("VS");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
                u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
            }

            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

#if (POWER_SUPPLY_NEW)
			PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif

            Set_TestInOut_For_VS(false);
#if true
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_32M_XTAL_Load_Cap(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_mhz;
            double d_diff_mhz, d_target_mhz = 2402;
            double d_lsl = 2401.9952, d_usl = 2402.0048;
            uint osc_val, osc_val_1;

            RegisterItem XTAL_LOAD_CONT = Parent.RegMgr.GetRegisterItem("O_XTAL_LOAD_CONT[4:0]");  // 0x4F

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            XTAL_LOAD_CONT.Read();
            if (x_pos != 0)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
                Parent.xlMgr.Sheet.Select("XTAL_Cap");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            osc_val = 15;
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            for (int val = 2; val >= 0; val--)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)val), (1 + cnt), d_freq_mhz.ToString("F4"));
                }
                if (d_freq_mhz > d_target_mhz)
                {
                    osc_val += (uint)(1 << val);
                }
                else
                {
                    osc_val -= (uint)(1 << val);
                }
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            osc_val_1 = osc_val;
            d_diff_mhz = Math.Abs(d_freq_mhz - d_target_mhz);

            if (d_freq_mhz > d_target_mhz)
            {
                if (osc_val != 31) osc_val += 1;
            }
            else
            {
                if (osc_val != 0) osc_val -= 1;
            }
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            System.Threading.Thread.Sleep(1);

            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            if (Math.Abs(d_freq_mhz - d_target_mhz) > d_diff_mhz)
            {
                osc_val = osc_val_1;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            }

            if ((d_freq_mhz < d_lsl) || (d_freq_mhz > d_usl))
            {
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return false;
            }
            else
            {
#if true // offset
                osc_val += 3;
                if (osc_val > 31) osc_val = 31;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
#endif
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return true;
            }

        }

        private bool Run_Cal_PLL_2PM(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_fsm;
            uint crc_flag;

            RegisterItem PM_IN = Parent.RegMgr.GetRegisterItem("O_PM_IN[7:0]");                    // 0x58

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            SendCommand("ook 70");
            System.Threading.Thread.Sleep(100);

            u_fsm = Run_Read_FSM_Status() & 0x000f;
            crc_flag = CRC_FLAG_READ();

            if (u_fsm != 7)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "FSM_Error");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) > 1)) // WUR_PKT_END = 1
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CRC_ERROR");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) == 1))
            {
                PM_IN.Read();
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (PM_IN.Value).ToString());
            }
            else
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CHECK_LOG");
                Console.WriteLine("u_fsm : {0}", u_fsm);
                Console.WriteLine("crc_flag : {0}", crc_flag);
                return false;
            }
            return true;
        }

        private void Set_AON_REG_For_BLE(byte[] aon_ble)
        {
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");            // 0x04
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");            // 0x05
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");            // 0x06
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");            // 0x07
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");            // 0x08
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");            // 0x09
            RegisterItem B10_LEN = Parent.RegMgr.GetRegisterItem("B10_LEN[7:0]");              // 0x0A
            RegisterItem B11_MSDF = Parent.RegMgr.GetRegisterItem("B11_MSDF[7:0]");            // 0x0B
            RegisterItem B12_CID_0 = Parent.RegMgr.GetRegisterItem("B12_CID_0[7:0]");          // 0x0C
            RegisterItem B13_CID_1 = Parent.RegMgr.GetRegisterItem("B13_CID_1[7:0]");          // 0x0D
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10
            RegisterItem B17_BA_3 = Parent.RegMgr.GetRegisterItem("B17_BA_3[7:0]");            // 0x11
            RegisterItem B18_BA_4 = Parent.RegMgr.GetRegisterItem("B18_BA_4[7:0]");            // 0x12
            RegisterItem B19_BA_5 = Parent.RegMgr.GetRegisterItem("B19_BA_5[7:0]");            // 0x13
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");            // 0x14
            RegisterItem B21_DC_1 = Parent.RegMgr.GetRegisterItem("B21_DC_1[7:0]");            // 0x15
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");            // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");            // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");            // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");            // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");            // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");            // 0x1B
            RegisterItem B28_RSV_0 = Parent.RegMgr.GetRegisterItem("B28_PUTC_0[7:0]");         // 0x1C
            RegisterItem B29_RSV_1 = Parent.RegMgr.GetRegisterItem("B29_PUTC_1[7:0]");         // 0x1D
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");            // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");            // 0x1F
            RegisterItem B32_PUTC_0 = Parent.RegMgr.GetRegisterItem("B32_CC_0[7:0]");          // 0x20
            RegisterItem B33_PUTC_1 = Parent.RegMgr.GetRegisterItem("B33_CC_1[7:0]");          // 0x21
            RegisterItem B34_PUTC_2 = Parent.RegMgr.GetRegisterItem("B34_CC_2[7:0]");          // 0x22
            RegisterItem B35_PUTC_3 = Parent.RegMgr.GetRegisterItem("B35_CC_3[7:0]");          // 0x23
            RegisterItem B36_BL = Parent.RegMgr.GetRegisterItem("B36_BL[7:0]");                // 0x24
            RegisterItem B37_TP = Parent.RegMgr.GetRegisterItem("B37_TP[7:0]");                // 0x25
            RegisterItem B38_SN_0 = Parent.RegMgr.GetRegisterItem("B38_SN_0[7:0]");            // 0x26
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27
            RegisterItem B40_MIC = Parent.RegMgr.GetRegisterItem("B40_MIC[7:0]");              // 0x28
            RegisterItem B41_advDelay = Parent.RegMgr.GetRegisterItem("B41_ADD_0[7:0]");       // 0x29
            RegisterItem B42_advDelay = Parent.RegMgr.GetRegisterItem("B42_ADD_1[7:0]");       // 0x2A
            RegisterItem B43_advDelay = Parent.RegMgr.GetRegisterItem("B43_ADD_2[7:0]");       // 0x2B

            B04_BA_0.Value = aon_ble[0];
            B04_BA_0.Write();

            B05_BA_1.Value = aon_ble[1];
            B05_BA_1.Write();

            B06_BA_2.Value = aon_ble[2];
            B06_BA_2.Write();

            B07_BA_3.Value = aon_ble[3];
            B07_BA_3.Write();

            B08_BA_4.Value = aon_ble[4];
            B08_BA_4.Write();

            B09_BA_5.Value = aon_ble[5];
            B09_BA_5.Write();

            B10_LEN.Value = aon_ble[6];
            B10_LEN.Write();

            B11_MSDF.Value = aon_ble[7];
            B11_MSDF.Write();

            B12_CID_0.Value = aon_ble[8];
            B12_CID_0.Write();

            B13_CID_1.Value = aon_ble[9];
            B13_CID_1.Write();

            B14_BA_0.Value = aon_ble[10];
            B14_BA_0.Write();

            B15_BA_1.Value = aon_ble[11];
            B15_BA_1.Write();

            B16_BA_2.Value = aon_ble[12];
            B16_BA_2.Write();

            B17_BA_3.Value = aon_ble[13];
            B17_BA_3.Write();

            B18_BA_4.Value = aon_ble[14];
            B18_BA_4.Write();

            B19_BA_5.Value = aon_ble[15];
            B19_BA_5.Write();

            B20_DC_0.Value = aon_ble[16];
            B20_DC_0.Write();

            B21_DC_1.Value = aon_ble[17];
            B21_DC_1.Write();

            B22_WP_0.Value = aon_ble[18];
            B22_WP_0.Write();

            B23_WP_1.Value = aon_ble[19];
            B23_WP_1.Write();

            B24_AI_0.Value = aon_ble[20];
            B24_AI_0.Write();

            B25_AI_1.Value = aon_ble[21];
            B25_AI_1.Write();

            B26_AI_2.Value = aon_ble[22];
            B26_AI_2.Write();

            B27_AI_3.Value = aon_ble[23];
            B27_AI_3.Write();

            B28_RSV_0.Value = aon_ble[24];
            B28_RSV_0.Write();

            B29_RSV_1.Value = aon_ble[25];
            B29_RSV_1.Write();

            B30_AD_0.Value = aon_ble[26];
            B30_AD_0.Write();

            B31_AD_1.Value = aon_ble[27];
            B31_AD_1.Write();

            B32_PUTC_0.Value = aon_ble[28];
            B32_PUTC_0.Write();

            B33_PUTC_1.Value = aon_ble[29];
            B33_PUTC_1.Write();

            B34_PUTC_2.Value = aon_ble[30];
            B34_PUTC_2.Write();

            B35_PUTC_3.Value = aon_ble[31];
            B35_PUTC_3.Write();

            B36_BL.Value = aon_ble[32];
            B36_BL.Write();

            B37_TP.Value = aon_ble[33];
            B37_TP.Write();

            B38_SN_0.Value = aon_ble[34];
            B38_SN_0.Write();

            B39_SN_1.Value = aon_ble[35];
            B39_SN_1.Write();

            B40_MIC.Value = aon_ble[36];
            B40_MIC.Write();

            B41_advDelay.Value = aon_ble[37];
            B41_advDelay.Write();

            B42_advDelay.Value = aon_ble[38];
            B42_advDelay.Write();

            B43_advDelay.Value = aon_ble[39];
            B43_advDelay.Write();
        }

        private void Run_Write_OTP_With_OOK(bool start, uint page)
        {
            string cmd;

            if (start == false)
            {
                return;
            }

            cmd = "ook 41." + page.ToString("X2");
            SendCommand(cmd);

            System.Threading.Thread.Sleep(500);
        }

        private uint Run_Verify_OTP_With_BIST(bool start, uint page)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return 65535;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // set page_flag
            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = (byte)(page & 0xff);
            SendData[5] = (byte)((page >> 8) & 0xff);
            SendData[6] = (byte)(((page >> 16) & 0x3f) | 0x80);
            SendData[7] = 0xBB;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1F;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=0
            SendData[4] = 0x1D;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] == 0x0b)
            {
                return 0x400;
            }
            else
            {
                return RcvData[0];
            }
        }

        private bool Run_Measure_Initial(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }
            // BGR, ALLDO, MLDO
            Set_TestInOut_For_BGR(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                // BGR
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 295) || (d_val > 305))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_12");
                        result = false;
                    }
                }
                // ALLDO
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 3), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 800) || (d_val > 820))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_13");
                        result = false;
                    }
                }

                // MLDO
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 6), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 985) || (d_val > 1015))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_14");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_BGR(false);

            // RCOSC
            Set_TestInOut_For_RCOSC(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + i + 9), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 32.113) || (d_val > 33.423))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_16");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_RCOSC(false);

#if (POWER_SUPPLY_NEW)
			PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_ADC(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            uint u_adc_code, u_temp_code, u_vbat_code;
            uint u_lsl_temp = 130, u_usl_temp = 150, u_lsl_vbat = 255, u_usl_vbat = 0;
            double d_volt_mv;

            if (start == false)
            {
                for (int i = 0; i < 8; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }

            Set_TestInOut_For_VTEMP(true);

            System.Threading.Thread.Sleep(1);
            for (int i = 0; i < 4; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 3.3,(@3)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        u_lsl_vbat = 186;
                        u_usl_vbat = 196;
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        u_lsl_vbat = 96;
                        u_usl_vbat = 106;
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 2.0,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.0");
#endif
                        u_lsl_vbat = 40;
                        u_usl_vbat = 50;
                        break;
                    case 3:
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 1.7,(@3)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        u_lsl_vbat = 6;
                        u_usl_vbat = 16;
                        break;
                    default:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                }
                System.Threading.Thread.Sleep(100);
                u_adc_code = Read_ADC_Result(false, 1, 63, 31);
                u_temp_code = u_adc_code & 0xff;
                u_vbat_code = (u_adc_code >> 8) & 0xff;
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 0), (y_pos + cnt), u_temp_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 1), (y_pos + cnt), u_vbat_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
                Disable_ADC();
#if false
                if ((u_temp_code < u_lsl_temp) || (u_temp_code > u_usl_temp))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_17");
                        result = false;
                    }
                }
#endif
                if ((u_vbat_code < u_lsl_vbat) || (u_vbat_code > u_usl_vbat))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_18");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_VTEMP(false);
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_VCO_Range(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_freq_mhz;
            double d_sl = 2402;

            RegisterItem VCO_TEST = Parent.RegMgr.GetRegisterItem("O_VCO_TEST");           // 0x41
            RegisterItem VCO_CBANK_L = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[7:0]");  // 0x42
            RegisterItem VCO_CBANK_H = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[9:8]");  // 0x43

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            VCO_TEST.Read();
            VCO_CBANK_L.Read();
            VCO_CBANK_H.Read();

            VCO_TEST.Value = 1;
            VCO_TEST.Write();

            // VCO Low = 0
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 0;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:SPAN 500 MHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.2 GHZ");
            System.Threading.Thread.Sleep(400);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 0), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz > d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO High = 1023
            VCO_CBANK_L.Value = 255;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 3;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.7 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz < d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO Mid = 512
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 2;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.5 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(false);
            VCO_TEST.Value = 0;
            VCO_TEST.Write();

            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");

            return result;
        }

        private bool Run_Measure_Tx_Power_Harmonic(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_power_dbm, d_freq_MHz, d_cur_mA;
            double d_lsl = -2, d_usl = 5;
            uint ch;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return result;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;
            PowerSupply0.Write("SENS:CURR:RANG 0.01,(@3)");

            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                    case 1:
                        ch = 12;
                        d_freq_MHz = 2426;
                        break;
                    case 2:
                        ch = 39;
                        d_freq_MHz = 2480;
                        break;
                    default:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                }
                Write_Register_Fractional_Calc_Ch(ch);
                for (int j = 0; j < 2; j++)
                {
                    if (j == 0)
                    {
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                    }
                    else
                    {
#if (POWER_SUPPLY_NEW)
						PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                    }
                    Write_Register_Tx_Tone_Send(true);
                    SpectrumAnalyzer.Write("FREQ:CENT " + d_freq_MHz + " MHZ");
                    System.Threading.Thread.Sleep(400);
                    d_cur_mA = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000.0;
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 2), (y_pos + cnt), d_cur_mA.ToString("F2"));
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < d_lsl) || (d_power_dbm > d_usl))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_21");
                            result = false;
                        }
                    }
                    if ((d_cur_mA < 8) || (d_cur_mA > 13))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_28");
                            result = false;
                        }
                    }

                    SpectrumAnalyzer.Write("FREQ:CENT " + (d_freq_MHz * 2) + " MHZ");
                    System.Threading.Thread.Sleep(50);
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 1), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < (d_lsl - 68)) || (d_power_dbm > (d_usl - 30)))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_22");
                            result = false;
                        }
                    }
                }
                Write_Register_Tx_Tone_Send(false);
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");
            Write_Register_Fractional_Calc_Ch(0);
            PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");

            return result;
        }

        private bool Run_Measure_INTB(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val_H;
            double d_val_L;

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            d_val_H = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
            I2C.GPIOs[3].Direction = GPIO_Direction.Output; // AD7
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].State = GPIO_State.Low;
            System.Threading.Thread.Sleep(115);
            I2C.GPIOs[3].State = GPIO_State.High;
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].Direction = GPIO_Direction.Input;
            System.Threading.Thread.Sleep(100);
            d_val_L = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));

            if (d_val_H > 0.8 && d_val_L < 0.5)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");
            }
            else
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL");
                return false;
            }

            return result;
        }

        private bool Run_Measure_TAMPER(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27 TAMPER dectec : pass = 0 fail = 1

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

#if !(POWER_SUPPLY_NEW)
            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
            return true;
#endif

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.9,(@2)");    // TAMPER 0.9 set
            PowerSupply0.Write("OUTP ON,(@2)");     // TAMPER 0.9 set            
#else
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 0.9");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_H");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW) 
            PowerSupply0.Write("VOLT 0.3,(@2)");   //TAMPER 0.3 set
#else
            PowerSupply0.Write("VOLT 0.3");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_L");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.6,(@2)");   //TAMPER 0.6 set
#else
            PowerSupply0.Write("VOLT 0.6");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x00)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_M");
                result = false;
                return result;
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("OUTP OFF,(@2)");
#else
            PowerSupply0.Write("VOLT 0.0");
            PowerSupply0.Write("INST:NSEL 2");
#endif
            return result;
        }

        private bool Run_Measure_BLE_Packet(bool start, int cnt, int x_pos, int y_pos, bool result, byte[] aon_ble)
        {
            int len;
            byte b;
            byte[] r_data = new byte[37];

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            Serial.RcvQueue.Clear();
            SendCommand("ook 30");
            System.Threading.Thread.Sleep(150);

            len = Serial.RcvQueue.Count;
            Console.WriteLine("UART len : " + len);
            SendCommand("mode ook");
            for (int i = 0; i < 11; i++)
            {
                b = Serial.RcvQueue.Get();
            }

            for (int i = 0; i < 36; i++)
            {
                b = Serial.RcvQueue.Get();

                if (b > 0x60)
                {
                    r_data[i] = (byte)((b - 87) << 4); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] = (byte)((b - 55) << 4); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] = (byte)((b - 48) << 4); // '0' = 0x30 = 48
                }

                b = Serial.RcvQueue.Get();
                if (b > 0x60)
                {
                    r_data[i] += (byte)((b - 87)); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] += (byte)((b - 55)); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] += (byte)((b - 48)); // '0' = 0x30 = 48
                }

                //b = UartRcvQueue.GetByte(); // space

                Console.Write("{0:X}-", r_data[i]);
                if (r_data[i] == aon_ble[i])
                {
                    continue;
                }
                else if ((i == 32) || (i == 33)) // BL, TP
                {
                    if (r_data[i] > 150)
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if (i == 35) // TAMPER
                {
                    if (!((r_data[i] == 0x00) || (r_data[i] == 0x80)))
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if ((i < 3) || (i > 9) && (i < 12))
                {
                    continue;
                }
                else
                {
                    Console.Write("\r\nFail!!{0} read : {1:X}, write : {2:X}\r\n", i, r_data[i], aon_ble[i]);
                    Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), i.ToString());
                    result = false;
                    return result;
                }

                if ((b == '\r') || (b == '\n'))
                {
                    break;
                }
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "P");
            return result;
        }

        private void Test_Good_Chip_Sorting_Rev3(int start_cnt)
        {
            int cnt = 0, pass = 0, fail = 0;
            double d_val;
            int x_pos = 2, y_pos = 12;
            bool result;
            uint mac_code = 0x000000;
            bool OTP_W_Flag = false;

            byte[] aon_ble = { 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0x1E, 0xFF, 0x6D, 0x0B, 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

            Check_Instrument();

            if (I2C.IsOpen == false) // Check I2C connection
            {
                MessageBox.Show("Check I2C");
                return;
            }

            MessageBox.Show("장비 설정을 확인해주세요.\r\n\r\n1.N6705B Power Supply\r\n  - VBAT\r\n  - TEST_IN_OUT(N6705B)\r\n\r\n" +
                "2.SpectrumAnalyzer\r\n  - TX\r\n\r\n3.Digital Multimeter\r\n  - Current mode : VBAT\r\n  - Voltage mode : ALLDO / MLDO / TEST_IN_OUT\r\n\r\n" +
                "4.UM232H\r\n  - AC0_GPIOH0 : I2C EN\r\n  - AD6_GPIOL2 : RX/TX Switch\r\n  - AD7_GPIOL3 : INTB");

            cnt = start_cnt - 1;
            pass = cnt;

            while (true)
            {
                I2C.GPIOs[3].Direction = GPIO_Direction.Input;   // AD7_GPIOL3(INTB)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].Direction = GPIO_Direction.Output;   // AD6_GPIOL2(TRX S/W)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].Direction = GPIO_Direction.Output;   // AC0_GPIOH0(Level Shifter EN)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);

                // Power off
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 0.0,(@3)");
                PowerSupply0.Write("OUTP ON,(@3)");
                PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");
#else
                PowerSupply0.Write("INST:NSEL 1");
                PowerSupply0.Write("VOLT 0.0");
                PowerSupply0.Write("OUTP ON");
#endif
                SendCommand("mode ook\n");
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n\r\nTest\t: " + cnt + "\r\nPass\t: " + pass + "\r\nFail\t: " + fail
                                                        , Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    cnt++;
                    fail++;
                    Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), cnt.ToString());
                    result = true;
                }
                else
                {
                    return;
                }

                DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:AC?"); // ALLDO
                // Power on
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 3.3,(@3)");
                System.Threading.Thread.Sleep(200);
                PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                PowerSupply0.Write("VOLT 3.3");
                System.Threading.Thread.Sleep(500);
                PowerSupply0.Write("VOLT 2.5");
#endif
                // Check MLDO Voltage

                System.Threading.Thread.Sleep(800);
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_val.ToString("F3"));
                if (d_val > 0.2)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                    continue;
                }

                // Seq1 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#else
                d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                d_val = 600;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
#endif
                if ((d_val < 500) || (d_val > 850))
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                    continue;
                }

                I2C.GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                // Wake-up
                WakeUp_I2C();
                System.Threading.Thread.Sleep(5);

                // READ_DIVICE_ID
                if (Check_Revision_Information(0x5F) == false) // N1 = 0x5D, N1B = 0x5E, N1C = 0x5F
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "F");
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_24");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "P");
                }

                mac_code = Read_Mac_Address();
                aon_ble[0] = (byte)(mac_code & 0xff);
                aon_ble[1] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[2] = (byte)((mac_code >> 16) & 0xff);
                aon_ble[10] = (byte)(mac_code & 0xff);
                aon_ble[11] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[12] = (byte)((mac_code >> 16) & 0xff);
                Parent.xlMgr.Cell.Write((x_pos + 81), (y_pos + cnt), (aon_ble[5].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 82), (y_pos + cnt), (aon_ble[4].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 83), (y_pos + cnt), (aon_ble[3].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 84), (y_pos + cnt), (aon_ble[2].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 85), (y_pos + cnt), (aon_ble[1].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 86), (y_pos + cnt), (aon_ble[0].ToString("X")));

                if (Check_OTP_VALID(true, 21) == true)
                {
                    goto CAL_SKIP;
                }

                // Seq1 5.OTP read
                if (Read_FF_NVM(false) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_10");
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "F");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "Skip");
                }

                Write_Register_AON_Fix_Value();

#if true // For Test
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
                Parent.xlMgr.Cell.Write(3, (1 + cnt), (double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Cell.Write(4, (1 + cnt), (double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
#endif
                // Seq2 1.BGR Trim
                if (Run_Cal_BGR(true, cnt, (x_pos + 6), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_4");
                    continue;
                }
                // Seq2 2.ALON LDO Trim
                if (Run_Cal_ALLDO(true, cnt, (x_pos + 8), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_5");
                    continue;
                }
                // Seq2 3.MLDO Trim
                if (Run_Cal_MLDO(true, cnt, (x_pos + 11), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_6");
                    continue;
                }
                // Seq2 4.32K RCOSC Trim
                if (Run_Cal_32K_RCOSC(true, cnt, (x_pos + 14), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_7");
                    continue;
                }
                // Seq2 5.Temp Sensor Trim
                if (Run_Cal_Temp_Sensor(true, cnt, (x_pos + 18), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_8");
                    continue;
                }
                // Seq2 5.Voltage Scaler Trim
                if (Run_Cal_Voltage_Scaler(true, cnt, (x_pos + 22), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_25");
                    continue;
                }
                // Seq2 6.32M X-tal Load Cap Trim
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                if (Run_Cal_32M_XTAL_Load_Cap(true, cnt, (x_pos + 26), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_9");
                    continue;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);
                // Seq2 7.PLL 2PM Trim
                if (Run_Cal_PLL_2PM(true, cnt, (x_pos + 29), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_20");
                    continue;
                }
                // Seq2 8.OTP Trim data write
                // Set AON_REG(BLE) For OTP
                Set_AON_REG_For_BLE(aon_ble);

                Run_Write_OTP_With_OOK(OTP_W_Flag, 0);
                // Run_Write_OTP_With_OOK(OTP_W_Flag, 15);
                Run_Write_OTP_With_OOK(OTP_W_Flag, 21);
                d_val = Run_Verify_OTP_With_BIST(OTP_W_Flag, 0x208001);
                if (d_val == 65535)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "Skip");
                }
                else if (d_val != 0x400)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_11");
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), d_val.ToString());
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "P");
                }

            CAL_SKIP:
                if (OTP_W_Flag) // Reset
                {
                    I2C.GPIOs[4].State = GPIO_State.High;
                    System.Threading.Thread.Sleep(10);
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 0.0,(@3)");
#else
                    PowerSupply0.Write("VOLT 0.0");
#endif
                    System.Threading.Thread.Sleep(500);

#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 3.3,(@3)");
                    System.Threading.Thread.Sleep(200);
                    PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                    PowerSupply0.Write("VOLT 3.3");
                    System.Threading.Thread.Sleep(500);
                    PowerSupply0.Write("VOLT 2.5");
#endif
                    // Check MLDO Voltage                    
                    System.Threading.Thread.Sleep(800);
                    d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), d_val.ToString("F3"));
                    if (d_val > 0.2)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                        continue;
                    }

                    // Seq2 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                    d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), (d_val).ToString("F3"));
#else
                    d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                    d_val = 600;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
#endif
                    if ((d_val < 500) || (d_val > 850))
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                        continue;
                    }

                    I2C.GPIOs[4].State = GPIO_State.Low;
                    System.Threading.Thread.Sleep(10);
                    // Wake-up
                    WakeUp_I2C();
                    System.Threading.Thread.Sleep(5);
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), "Skip");
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
                }

                // Seq3 1.Initial Test
                if (Run_Measure_Initial(true, cnt, (x_pos + 33), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 2.ADC Test
                if (Run_Measure_ADC(true, cnt, (x_pos + 45), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                // Seq3 4.VCO Range Test
                if (Run_Measure_VCO_Range(true, cnt, (x_pos + 57), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 4. Tx output power and Harmonic
                if (Run_Measure_Tx_Power_Harmonic(true, cnt, (x_pos + 60), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);

                // INTB Test
                if (Run_Measure_INTB(true, cnt, (x_pos + 79), y_pos, result) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_27");
                    }
                    result = false;
                }
                // TAMPER Test
                if (Run_Measure_TAMPER(true, cnt, (x_pos + 78), y_pos, result) == false)
                {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("OUTP OFF,(@2)");
#endif
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_26");
                    }
                    result = false;
                }

                // BLE Test
                if (Run_Measure_BLE_Packet(true, cnt, (x_pos + 80), y_pos, result, aon_ble) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_23");
                    }
                    result = false;
                }

                if (result == true)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "PASS");
                    fail--;
                    pass++;
                }
            }
        }

        private void DUT_TempSensor_Test()
        {
            byte Addr;
            byte[] RcvData = new byte[2];

            Parent.xlMgr.Sheet.Add(DateTime.Now.ToString("MMddHHmmss_") + "DUT_TS");
            Parent.xlMgr.Cell.Write(1, 1, "No.");
            Parent.xlMgr.Cell.Write(2, 1, "ADC");
            Parent.xlMgr.Cell.Write(3, 1, "Temp");

            I2C.GPIOs[2].Direction = GPIO_Direction.Output;
            System.Threading.Thread.Sleep(10);

            Addr = 0x00;
#if false            
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(200); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Parent.xlMgr.Cell.Write(1, 2 + i, i.ToString());
                Parent.xlMgr.Cell.Write(2, 2 + i, ((RcvData[0] << 8) | RcvData[1]).ToString());
                Parent.xlMgr.Cell.Write(3, 2 + i, (((RcvData[0] << 8) | RcvData[1]) * 0.0078125).ToString());

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(500);
            }
#else
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(300); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Console.WriteLine("ADC : {0}\tTemp : {1}", (RcvData[0] << 8) | RcvData[1], ((RcvData[0] << 8) | RcvData[1]) * 0.0078125);

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(100);
            }
#endif
        }

        private void Test_THB_SMT_Board(int iVal)
        {
            int cnt = 0;
            double d_val;

            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");            // 0x04
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");            // 0x05
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");            // 0x06
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");            // 0x07
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");            // 0x08
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");            // 0x09

            Check_Instrument();

            Parent.xlMgr.Sheet.Select("THB_SMT");
            Parent.xlMgr.Cell.Write(1, 1, "No.");
            Parent.xlMgr.Cell.Write(2, 1, "Sleep(nA)");
            Parent.xlMgr.Cell.Write(3, 1, "ALLDO(mV)");
            Parent.xlMgr.Cell.Write(4, 1, "MLDO(mV)");
            Parent.xlMgr.Cell.Write(5, 1, "INTB(V)");
            Parent.xlMgr.Cell.Write(6, 1, "MAC5");
            Parent.xlMgr.Cell.Write(7, 1, "MAC4");
            Parent.xlMgr.Cell.Write(8, 1, "MAC3");
            Parent.xlMgr.Cell.Write(9, 1, "MAC2");
            Parent.xlMgr.Cell.Write(10, 1, "MAC1");
            Parent.xlMgr.Cell.Write(11, 1, "MAC0");
            if (iVal == 0)
            {
                cnt = 0;
            }
            else
            {
                cnt = iVal - 1;
            }

            while (true)
            {
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n\r\nTest\t: " + cnt + "\r\n"
                                                        , Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    cnt++;

                    Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
                }
                else
                {
                    return;
                }

                I2C.GPIOs[4].Direction = GPIO_Direction.Output;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);

                PowerSupply0.Write("SOUR:VOLT 3.3"); // BatterySimulator
                PowerSupply0.Write(":OUTP ON");
                System.Threading.Thread.Sleep(100);

                DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:AC?");
                PowerSupply0.Write("SENS:CURR:RANG 0.00001");
                System.Threading.Thread.Sleep(5000);
                d_val = double.Parse(PowerSupply0.WriteAndReadString(":MEAS:CURR? (@1)")) * 1000000000.0;
                for (int i = 0; i < 99; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString(":MEAS:CURR? (@1)")) * 1000000000.0;
                d_val /= 100;
                Parent.xlMgr.Cell.Write(2, (1 + cnt), d_val.ToString("F2"));

                // System.Threading.Thread.Sleep(5000);
                // ALLDO
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write(3, (1 + cnt), d_val.ToString("F2"));

                // INTB
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?"));
                Parent.xlMgr.Cell.Write(5, (1 + cnt), d_val.ToString("F3"));

                PowerSupply0.Write("SENS:CURR:RANG 0.001");
                System.Threading.Thread.Sleep(1000);
                I2C.GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                WakeUp_I2C();
                System.Threading.Thread.Sleep(100);

                Parent.xlMgr.Cell.Write(6, (1 + cnt), B09_BA_5.Read().ToString("X2"));
                Parent.xlMgr.Cell.Write(7, (1 + cnt), B08_BA_4.Read().ToString("X2"));
                Parent.xlMgr.Cell.Write(8, (1 + cnt), B07_BA_3.Read().ToString("X2"));
                Parent.xlMgr.Cell.Write(9, (1 + cnt), B06_BA_2.Read().ToString("X2"));
                Parent.xlMgr.Cell.Write(10, (1 + cnt), B05_BA_1.Read().ToString("X2"));
                Parent.xlMgr.Cell.Write(11, (1 + cnt), B04_BA_0.Read().ToString("X2"));

                // MLDO
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write(4, (1 + cnt), d_val.ToString("F2"));

                PowerSupply0.Write(":OUTP OFF");
            }
        }
        #endregion

        # region Function for DTM
        private void SetBleDTM_AonControlReg()
        {
            WriteRegister(0x33, 0x17);
            WriteRegister(0x34, 0x3F);       //ADC_EOC_SKIP_TEMP
            WriteRegister(0x35, 0x3F);       //ADC_EOC_SKIP_VOLT
            WriteRegister(0x36, 0xFF);
            WriteRegister(0x37, 0x47);
            WriteRegister(0x38, 0x00);
            WriteRegister(0x39, 0x40);
            WriteRegister(0x3A, 0x4A);
            WriteRegister(0x3B, 0x00);
        }

        private void SetBleDTM_AonControlReg_OTP()
        {
            WriteRegister(0x33, 0x17);
            WriteRegister(0x37, 0x47);

            WriteRegister(0x45, 0x7);   //CT_RES32_FINE = 7 (Bandwidth change)
            WriteRegister(0x64, 0xD2);  //TX_SEL = CoraseLock & PLL_PM_GIAN = 18;
            WriteRegister(0x65, 0x50);  //DA_PEN = CoraseLock & CH_0_GAIN = 16;
        }

        private void SetBleDTM_AonAnalogReg()
        {
            WriteRegister(0x62, 0x66);      //r3 ONLY
            WriteRegister(0x61, 0x33);
            WriteRegister(0x53, 0xC7);      //FC_LGA
            WriteRegister(0x54, 0xBC);
            WriteRegister(0x3C, 0x10);
            WriteRegister(0x3D, 0x12);
            WriteRegister(0x3E, 0x00);      //R3 ONLY
            WriteRegister(0x3F, 0x72);
            WriteRegister(0x40, 0x82);
            WriteRegister(0x41, 0x80);
            WriteRegister(0x42, 0x00);
            WriteRegister(0x43, 0xF2);
            WriteRegister(0x44, 0xFF);
            WriteRegister(0x45, 0x02);
            WriteRegister(0x46, 0xFF);
            WriteRegister(0x47, 0x42);
            WriteRegister(0x48, 0x62);
            WriteRegister(0x49, 0xFF);
            WriteRegister(0x4A, 0x3F);
            WriteRegister(0x4B, 0x16);
            WriteRegister(0x4C, 0x00);
            WriteRegister(0x4D, 0x29);
            WriteRegister(0x4E, 0xB2);
            WriteRegister(0x4F, 0x61);
            WriteRegister(0x50, 0xC5);
            WriteRegister(0x51, 0x2F);      //R3 ONLY
            WriteRegister(0x52, 0x22);

            WriteRegister(0x55, 0xCE);
            WriteRegister(0x56, 0x00);  //R3 ONLY DYNAMIC CONTROL FOR BLE LINK
            WriteRegister(0x57, 0x07);  //TX_BUF, DA_PEN, PRE_DA_PEN, DA_GAIN R3 ONLY DYNIMIC CONTROL FOR BLE LINK
            WriteRegister(0x58, 0x33);  //2PM_CAL R3 ONLY  
            WriteRegister(0x59, 0x07);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
            WriteRegister(0x5A, 0x00);  //EXT_DA_GAIN_LSB_BIT
            WriteRegister(0x5B, 0xE2);  //EXT_DA_GAIN_SET = 1; (BLE_LINK) R3                                                       
            WriteRegister(0x5C, 0x00);  //TX_SEL(W),PRE_DA_PEN(W), TX_BUF_PEN(W), DA_PEN(W) R3 ONLY                                         
            WriteRegister(0x5D, 0x60);  //PM_RESETB(W), PLL_PEN(W)
            WriteRegister(0x5E, 0x20);
            WriteRegister(0x5F, 0x45);  //R3 ONLY
            WriteRegister(0x60, 0x08);

            WriteRegister(0x63, 0x00);
            WriteRegister(0x64, 0xB0);  //TX_SEL = DATA_ST

            //FOR R3 ONLY
            WriteRegister(0x65, 0x34);  //DA_PEN = DATA_ST
            WriteRegister(0x66, 0x12);
            WriteRegister(0x67, 0x40);
        }

        private void Set_PLLRESET(int delay)
        {
            WriteRegister(0x59, 0x1);
            System.Threading.Thread.Sleep(delay);
            WriteRegister(0x59, 0x7);
            System.Threading.Thread.Sleep(delay);
        }

        private void Set_BLE_LINK(uint addr, uint data)
        {
            byte[] Addrs = new byte[8];

            Addrs[0] = (byte)(addr & 0xff);
            Addrs[1] = (byte)(addr >> 8 & 0xff);
            Addrs[2] = (byte)(addr >> 16 & 0xff);
            Addrs[3] = 0x00;
            Addrs[4] = (byte)(data & 0xff);
            Addrs[5] = (byte)(data >> 8 & 0xff);
            Addrs[6] = (byte)(data >> 16 & 0xff);
            Addrs[7] = (byte)(data >> 24 & 0xff);

            I2C.WriteBytes(Addrs, 8, true);
        }

        private void Run_BLE_DTM_MODE(uint PayLoadMode)
        {
            uint PmGain = 0;

            PmGain = ReadRegister(0x58);

            if (PmGain == 0x42)
            {
                SetBleDTM_AonControlReg();
                System.Threading.Thread.Sleep(2);
                SetBleDTM_AonAnalogReg();
                System.Threading.Thread.Sleep(2);
            }
            else
            {
                SetBleDTM_AonControlReg_OTP();
                System.Threading.Thread.Sleep(2);
            }

            Set_BLE_LINK(0x058000, 0x80);   //reset
            Set_BLE_LINK(0x058000, 0x48);   //stop dtm
            Set_BLE_LINK(0x0580d8, 0x78);   //bb_clk_freq_minus_1 
            Set_BLE_LINK(0x0581b4, 0x72);   //RADIO_DATA host datain [15:8] / host dataout[15:8]
            Set_BLE_LINK(0x0581ac, 0x9043); //RADIO_ACCESS/RADIO_CNTRL 
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9400);
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9401);
            Set_BLE_LINK(0x0580d8, 0x78);
            Set_BLE_LINK(0x058190, 0x9280);
            Set_BLE_LINK(0x058198, 0x5b5b);
            Set_BLE_LINK(0x0581b0, 0x03);
            Set_BLE_LINK(0x0581b8, 0x05);

            uint pmode = ((PayLoadMode * 128) + 0);
            Set_BLE_LINK(0x058170, pmode);
            Set_BLE_LINK(0x05819c, (uint)DTM_PayLoadLength); //packet length
            Set_BLE_LINK(0x040014, (0x280 + (uint)DTM_Channel)); //Ch Sel Mode & CH 0;
            Set_PLLRESET(2);
            Set_BLE_LINK(0x058000, 0x46); //start dtm
        }

        private void Stop_BLE_DTM_MODE()
        {
            Set_BLE_LINK(0x058000, 0x48);
            WriteRegister(0x59, 0x03);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
        }
        #endregion
    }

    public class SCP1501_R4 : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS_MANUAL
        {
            Write_WURF,
            TX_CH_SEL,
            TX_ON,
            TX_OFF,
            SET_BLE,
            Read_FSM,
            Read_ADC,
            Read_Volt_Temp,
            TESTINOUT_Temp,
            TESTINOUT_Volt,
            TESTINOUT_BGR,
            TESTINOUT_EDOUT,
            TESTINOUT_RCOSC,
            TESTINOUT_Disable,
            NVM_POWER_ON,
            NVM_POWER_OFF,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            LAB_TEST,
            SENSOR_OUTPUT,
            EXT_SENS_Forcing,
            Cal_PMU_All,
            WriteMACBLE00_V1,
            WriteMACBLE00_V2,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_DTM
        {
            SET_CH,
            SET_Length,
            DTM_START_PRNBS9,
            DTM_START_11110000,
            DTM_START_10101010,
            DTM_START_PRNBS15,
            DTM_START_ALL_1,
            DTM_START_ALL_0,
            DTM_START_00001111,
            DTM_START_0101,
            DTM_STOP,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
            DTM,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        public int SlaveAddress { get; private set; } = 0x3A;
        public int DTM_PayLoadLength = 37;
        public int DTM_Channel = 0;

        /* Intrument */
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

        public SCP1501_R4(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "SCP1501_R4");
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

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[8];
            I2C.Config.SlaveAddress = SlaveAddress;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "NVM":
                    // bist_sel=1, bist_type=3, bist_cmd=0, bist_addr=address
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0d);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=3, bist_cmd=1, bist_addr=address
                    SendData[4] = (byte)(0x0f);
                    SendData[5] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_wdata=data
                    SendData[0] = (byte)(0x04);            // ADDR = 0x00060004
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=3, bist_cmd=0, bist_addr=address
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0d);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Dummy
                    System.Threading.Thread.Sleep(5);

                    // bist_sel=0, bist_type=3, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0c);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                case "NVM Controller":
                case "NVM BIST":
                case "I2C Slave":
                case "LPCAL":
                case "ADC Controller":
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                default:
                    // Write Address, Data
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x07);            // AREG_SEL=1, AREG_CE=1, AREG_WE=1
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);            // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);            // AREG_WDATA[7:0]
                    SendData[5] = (byte)(0x00);            // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);            // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
            }
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            uint result = 0xffffffff;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "NVM":
                    // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x11);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
                    SendData[4] = (byte)(0x13);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
                    SendData[4] = (byte)(0x11);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Read bist_rdata
                    SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);

                    // bist_sel=0, bist_type=4, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x10);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                case "NVM Controller":
                case "NVM BIST":
                case "I2C Slave":
                case "LPCAL":
                case "ADC Controller":
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);
                    break;

                default:
                    // Write Address
                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);             // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);   // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    SendData[6] = (byte)(0x05);             // AREG_SEL=1, AREG_CE=1, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Read Data
                    SendData[0] = (byte)(0x10);             // ADDR = 0x00010010
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = RcvData[3];

                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[5] = (byte)(0x00);             // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);             // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
            }

            return result;
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
            Parent.ChipCtrlButtons[4].Text = "GH0_H";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;
            Parent.ChipCtrlButtons[5].Text = "WakeUp";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += WakeUp_I2C;
            Parent.ChipCtrlButtons[6].Text = "MD_OOK";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Set_Gecko_OOK_Mode;
            Parent.ChipCtrlButtons[8].Text = "Manual";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;
            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_Auto_Test_Items;
            Parent.ChipCtrlButtons[10].Text = "DTM";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Change_To_DTM_Test_Items;
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
            Serial.WriteSettingFile(Parent.IniFile, "SCP1501_R4");
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

        private void Change_To_DTM_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.DTM;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_DTM.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_DTM)i).ToString());
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

        private void WakeUp_I2C(object sender, EventArgs e)
        {
            WakeUp_I2C();
        }

        private void Set_Gecko_OOK_Mode(object sender, EventArgs e)
        {
            SendCommand("mode ook");
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
                        case TEST_ITEMS_MANUAL.Write_WURF:
                            if (Arg == "")
                            {
                                Log.WriteLine("WURF에 write할 값을 16진수로 적어주세요.");
                            }
                            else
                            {
                                Log.WriteLine("OOK " + Arg);
                                Write_WURF_AON(Arg);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_CH_SEL:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                Log.WriteLine("Set CH" + Arg);
                                Write_Register_Fractional_Calc_Ch((uint)iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_ON:
                            Log.WriteLine("TX ON");
                            Write_Register_Tx_Tone_Send(true);
                            break;
                        case TEST_ITEMS_MANUAL.TX_OFF:
                            Log.WriteLine("TX OFF");
                            Write_Register_Tx_Tone_Send(false);
                            break;
                        case TEST_ITEMS_MANUAL.SET_BLE:
                            if ((Arg == "") || (iVal < 0) || (iVal > 65535))
                            {
                                Log.WriteLine("WP값을 10진수로 적어주세요. (Range 0~65535)");
                            }
                            else
                            {
                                Log.WriteLine("Write AON for Advertising (WP = )" + Arg);
                                Write_Register_Send_Advertising(iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.Read_FSM:
                            Read_Register_FSM();
                            break;
                        case TEST_ITEMS_MANUAL.Read_ADC:
                            Result = Read_ADC_Result(false, 1, 63, 31);
                            Disable_ADC();
                            Log.WriteLine("Volt : " + (Result >> 8).ToString() + "\tTemp : " + (Result & 0xff).ToString());
                            break;
                        case TEST_ITEMS_MANUAL.Read_Volt_Temp:
                            Calculation_VBAT_Voltage_And_Temperature();
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Temp:
                            Set_TestInOut_For_VTEMP(true);
                            Read_ADC_Result(false, 1, 1, 1);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Volt:
                            Set_TestInOut_For_VS(true);
                            Read_ADC_Result(true, 1, 1, 1);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_BGR:
                            Set_TestInOut_For_BGR(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_EDOUT:
                            Set_TestInOut_For_EDOUT(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_RCOSC:
                            Set_TestInOut_For_RCOSC(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Disable:
                            Set_TestInOut_For_VTEMP(false);
                            Set_TestInOut_For_VS(false);
                            Set_TestInOut_For_BGR(false);
                            Set_TestInOut_For_EDOUT(false);
                            Set_TestInOut_For_RCOSC(false);
                            Disable_ADC();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_ON:
                            Enable_NVM_BIST();
                            Power_On_NVM();
                            Disable_NVM_BIST();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_OFF:
                            Enable_NVM_BIST();
                            Power_Off_NVM();
                            Disable_NVM_BIST();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((TEST_ITEMS_AUTO)TestItemIndex)
                    {
                        case TEST_ITEMS_AUTO.LAB_TEST:
                            Test_Good_Chip_Sorting_Rev3(iVal);
                            break;
                        case TEST_ITEMS_AUTO.Cal_PMU_All:
                            Cal_PMU_All();
                            break;
                        case TEST_ITEMS_AUTO.WriteMACBLE00_V1:
                            WriteMACBLE00_V1(iVal);
                            break;
                        case TEST_ITEMS_AUTO.WriteMACBLE00_V2:
                            WriteMACBLE00_V2(iVal);
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.DTM:
                    switch ((TEST_ITEMS_DTM)TestItemIndex)
                    {
                        case TEST_ITEMS_DTM.SET_CH:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                DTM_Channel = iVal;
                                Log.WriteLine("Channel : " + DTM_Channel.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.SET_Length:
                            if ((Arg == "") || ((iVal != 37) && (iVal != 255)))
                            {
                                Log.WriteLine("Length를 10진수로 적어주세요. (Range 37 또는 255)");
                            }
                            else
                            {
                                DTM_PayLoadLength = iVal;
                                Log.WriteLine("Length : " + DTM_PayLoadLength.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.DTM_START_PRNBS9:
                        case TEST_ITEMS_DTM.DTM_START_11110000:
                        case TEST_ITEMS_DTM.DTM_START_10101010:
                        case TEST_ITEMS_DTM.DTM_START_PRNBS15:
                        case TEST_ITEMS_DTM.DTM_START_ALL_1:
                        case TEST_ITEMS_DTM.DTM_START_ALL_0:
                        case TEST_ITEMS_DTM.DTM_START_00001111:
                        case TEST_ITEMS_DTM.DTM_START_0101:
                            Log.WriteLine("Start DTM\nChannel : " + DTM_Channel.ToString() + "\tLength : " + DTM_PayLoadLength.ToString());
                            Run_BLE_DTM_MODE((uint)(TestItemIndex - TEST_ITEMS_DTM.DTM_START_PRNBS9));
                            break;
                        case TEST_ITEMS_DTM.DTM_STOP:
                            Log.WriteLine("Stop DTM");
                            Stop_BLE_DTM_MODE();
                            break;
                    }
                    break;
                default:
                    break;
            }
        }

        #region Function for Chip Test
        private void WakeUp_I2C()
        {
            byte[] SendData = new byte[2];

            SendData[0] = 0xAA;
            SendData[1] = 0xBB;

            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_CMD_I2C(byte data)
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=1, WURF_WRITE=1
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = data;
            SendData[5] = 0x00;
            SendData[6] = 0x06;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // WURF_WRITE=0
            SendData[6] = 0x02;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_WURF_Sel()
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=0
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_AON(string Arg)
        {
            RegisterItem w_WURF_END = Parent.RegMgr.GetRegisterItem("w_WURF_END");         // 0x38
            byte data;

            switch (Arg)
            {
                case "30":
                    data = 0x30;
                    break;
                case "60":
                    data = 0x60;
                    break;
                case "61":
                    data = 0x61;
                    break;
                case "70":
                    data = 0x70;
                    break;
                case "80":
                    data = 0x80;
                    break;
                case "b0":
                case "B0":
                    data = 0xb0;
                    break;
                case "c0":
                case "C0":
                    data = 0xb0;
                    break;
                default:
                    Log.WriteLine("지원하지 않는 CMD 입니다.");
                    return;
                    break;
            }

            w_WURF_END.Read();
            w_WURF_END.Value = 0;
            w_WURF_END.Write();

            Write_WURF_CMD_I2C(data);

            w_WURF_END.Value = 1;
            w_WURF_END.Write();

            Disable_WURF_Sel();

            w_WURF_END.Value = 0;
            w_WURF_END.Write();
        }

        private void Write_Register_Fractional_Calc_Ch(uint ch)
        {
            uint vco;

            RegisterItem SPI_DSM_F_H = Parent.RegMgr.GetRegisterItem("O_SPI_DSM_F[22:16]"); // 0x3E
            RegisterItem SPI_PS_P = Parent.RegMgr.GetRegisterItem("O_SPI_PS_P[4:0]");       // 0x3F
            RegisterItem SPI_PS_S = Parent.RegMgr.GetRegisterItem("O_SPI_PS_S[2:0]");       // 0x3F                

            vco = 2402 + 2 * ch;
            SPI_PS_P.Value = vco >> 7;
            SPI_PS_S.Value = (vco >> 5) - (SPI_PS_P.Value << 2);
            SPI_DSM_F_H.Value = (vco << 1) - (SPI_PS_P.Value << 8) - (SPI_PS_S.Value << 6);

            SPI_PS_P.Write();
            SPI_DSM_F_H.Write();
        }

        private void Write_Register_Tx_Tone_Send(bool on)
        {
            RegisterItem EXT_CH_MODE = Parent.RegMgr.GetRegisterItem("O_EXT_CH_MODE");             // 0x3E
            RegisterItem TX_SEL = Parent.RegMgr.GetRegisterItem("O_TX_SEL");                       // 0x56
            RegisterItem TX_BUF_PEN = Parent.RegMgr.GetRegisterItem("O_TX_BUF_PEN");               // 0x57
            RegisterItem PRE_DA_PEN = Parent.RegMgr.GetRegisterItem("O_PRE_DA_PEN");               // 0x57
            RegisterItem DA_PEN = Parent.RegMgr.GetRegisterItem("O_DA_PEN");                       // 0x57
            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("O_PLL_PEN");                     // 0x59

            RegisterItem EXT_DA_GAIN_SEL = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_SEL");     // 0x5B
            RegisterItem w_TX_BUF_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_TX_BUF_PEN_MODE");   // 0x5C
            RegisterItem w_PRE_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PRE_DA_PEN_MODE");   // 0x5C
            RegisterItem w_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_DA_PEN_MODE");           // 0x5C
            RegisterItem w_TX_SEL_MODE = Parent.RegMgr.GetRegisterItem("w_TX_SEL_MODE");           // 0x5C
            RegisterItem w_PLL_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PLL_PEN_MODE");         // 0x5D
            RegisterItem XTAL_PLL_CLK_EN_MD = Parent.RegMgr.GetRegisterItem("w_XTAL_PLL_CLK_EN_MD");   // 0x5F

            if (on == true)
            {
                // Tx Active
                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 1;
                EXT_CH_MODE.Write();

                TX_SEL.Read();
                TX_SEL.Value = 1;
                TX_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 1;
                PRE_DA_PEN.Value = 1;
                DA_PEN.Value = 1;
                TX_BUF_PEN.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 0;
                EXT_DA_GAIN_SEL.Write();

                // Dynamic -> Static
                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 1;
                w_PLL_PEN_MODE.Write();

                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 1;
                w_PRE_DA_PEN_MODE.Value = 1;
                w_DA_PEN_MODE.Value = 1;
                w_TX_SEL_MODE.Value = 1;
                w_TX_BUF_PEN_MODE.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 1;
                XTAL_PLL_CLK_EN_MD.Write();

                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                PLL_PEN.Value = 1;
                PLL_PEN.Write();
            }
            else
            {
                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 0;
                XTAL_PLL_CLK_EN_MD.Write();

                // Static -> Dynamic
                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 0;
                w_PRE_DA_PEN_MODE.Value = 0;
                w_DA_PEN_MODE.Value = 0;
                w_TX_SEL_MODE.Value = 0;
                w_TX_BUF_PEN_MODE.Write();

                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 0;
                w_PLL_PEN_MODE.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 1;
                EXT_DA_GAIN_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 0;
                PRE_DA_PEN.Value = 0;
                DA_PEN.Value = 0;
                TX_BUF_PEN.Write();

                TX_SEL.Read();
                TX_SEL.Value = 0;
                TX_SEL.Write();

                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 0;
                EXT_CH_MODE.Write();
            }
        }

        private void Write_Register_Send_Advertising(int iVal)
        {
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");     // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");     // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");     // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");     // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");     // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");     // 0x1B
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");     // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");     // 0x1F

            B22_WP_0.Value = (uint)(iVal & 0xff);
            B22_WP_0.Write();

            B23_WP_1.Value = (uint)((iVal >> 8) & 0xff);
            B23_WP_1.Write();

            B24_AI_0.Value = 210;
            B24_AI_0.Write();

            B25_AI_1.Value = 29;
            B25_AI_1.Write();

            B26_AI_2.Value = 0;
            B26_AI_2.Write();

            B27_AI_3.Value = 0;
            B27_AI_3.Write();

            B30_AD_0.Value = 3;
            B30_AD_0.Write();

            B31_AD_1.Value = 0;
            B31_AD_1.Write();
        }

        private uint Run_Read_FSM_Status()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[1] & 0xf) << 8) | RcvData[0]);
        }

        private uint CRC_FLAG_READ()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[2] & 0x3) << 16) | (RcvData[1] << 8) | RcvData[0]);
        }

        private void Read_Register_FSM()
        {
            uint u_fsm;
            uint crc_flag;
            u_fsm = Run_Read_FSM_Status();
            crc_flag = CRC_FLAG_READ();
            Log.WriteLine("FSM_Status : " + (u_fsm & 0x000f).ToString());
            Log.WriteLine("FSM_State : " + ((u_fsm & 0x0ff0) >> 4).ToString());
            Log.WriteLine("WUR_CRC_ERROR : " + ((crc_flag >> 17) & 0xff).ToString());
            Log.WriteLine("WUR_CRC_DEC : " + (crc_flag & 0xffff).ToString());
        }

        private uint Read_ADC_Result(bool temp_first, uint eoc_skip_temp, uint eoc_skip_volt, uint avg_cnt)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            // Set w_TEST_SELECT
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
#if true // w_TEST_SELECT = 1
            if (temp_first)
            {
                SendData[7] = 0x80; // temp -> volt
            }
            else
            {
                SendData[7] = 0xC0; // volt -> temp
            }
#else // w_TEST_SELECT = 0
            if (temp_first)
            {
                SendData[7] = 0x00; // temp -> volt
            }
            else
            {
                SendData[7] = 0x40; // volt -> temp
            }
#endif
            I2C.WriteBytes(SendData, SendData.Length, true);
            // Enable ADC (TEST_START = 0, I_PEN = 1)
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x01);
            SendData[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            SendData[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            SendData[7] = (byte)((avg_cnt & 0x1f) << 1);
            I2C.WriteBytes(SendData, SendData.Length, true);
            // Enable ADC (TEST_START = 1, I_PEN = 1)
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x01);
            SendData[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            SendData[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            SendData[7] = (byte)(((avg_cnt & 0x1f) << 1) | 0x80);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(10);
            // Read ADC_D_B[7:0], ADC_D_T[7:0]
            SendData[0] = 0x08;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            return (uint)(((RcvData[1] << 8) | RcvData[0]) & 0xffff);
        }

        private void Disable_ADC()
        {
            byte[] SendData = new byte[8];

            // Disable ADC
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private double Calculation_Temperature(uint otp, uint adc)
        {
            double[] lowtempcomp_ADC = { 57.11, 57.23, 57.34, 57.45, 57.56, 57.66, 57.75, 57.84 };
            double temperature;
            uint OTP23p5, OTP85TO23p5;

            OTP23p5 = (otp & 0x1f) + 104;
            OTP85TO23p5 = (otp >> 5) + 35;

            if (adc >= OTP23p5)
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) + 23.5;
            }
            else
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) * ((23.5 - (-40)) / lowtempcomp_ADC[otp >> 5]) + 23.5;
            }

            return temperature;
        }

        private double Calculation_VBAT_Voltage(double temperature, uint adc)
        {
            double voltage;

            if (temperature >= 30)
            {
                voltage = (adc - (11.475 - (-3 / (85 - 30)) * (temperature - 30))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }
            else
            {
                voltage = (adc - (11.475 - (2 / (30 - (-40))) * (30 - temperature))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }

            return voltage;
        }

        private void Calculation_VBAT_Voltage_And_Temperature()
        {
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");       // 0x14
            uint adc;
            double temp, volt;

            B20_DC_0.Read();
            adc = Read_ADC_Result(false, 1, 63, 31);
            Disable_ADC();
            temp = Calculation_Temperature(B20_DC_0.Value, adc & 0xff);
            volt = Calculation_VBAT_Voltage(temp, adc >> 8);
            Log.WriteLine("Volt : " + volt.ToString("F2") + "(" + (adc >> 8).ToString() + ")\tTemp : " + temp.ToString("F3") + "(" + (adc & 0xff).ToString() + ")");
        }

        private void Set_TestInOut_For_VTEMP(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");       // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL");     // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");            // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");            // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BUF_MUX_SEL.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_VS(bool on)
        {
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");        // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");        // 0x4C
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");   // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL"); // 0x5B

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0x20;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();
            }
        }

        private void Set_TestInOut_For_BGR(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");    // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");         // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");         // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_EDOUT(bool on)
        {
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[8]");   // 0x5A
            RegisterItem O_RX_DATAT = Parent.RegMgr.GetRegisterItem("O_RX_DATAT");      // 0x5A

            if (on == true)
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 1;
                O_RX_DATAT.Value = 1;
                ITEST_CONT.Write();
            }
            else
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 0;
                O_RX_DATAT.Value = 0;
                ITEST_CONT.Write();
            }
        }

        private void Set_TestInOut_For_RCOSC(bool on)
        {
            RegisterItem TEST_EN_32K = Parent.RegMgr.GetRegisterItem("O_TEST_EN_32K");     // 0x4C
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");    // 0x4C

            if (on == true)
            {
                TEST_EN_32K.Read();
                TEST_EN_32K.Value = 1;
                TEST_CON_H.Value = 8;
                TEST_EN_32K.Write();
            }
            else
            {
                TEST_EN_32K.Value = 0;
                TEST_CON_H.Value = 0;
                TEST_EN_32K.Write();
            }
        }

        private void Enable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x01;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=0, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_On_NVM()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=1, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x07;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=1, bist_cmd=0
            SendData[4] = 0x05;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_Off_NVM()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=2, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x0B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=2, bist_cmd=0
            SendData[4] = 0x09;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }
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

        private bool Check_Revision_Information(uint value)
        {
            RegisterItem REV_ID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");     // 0x68
            bool result = false;

            REV_ID.Read();

            if (value == REV_ID.Value)
            {
                result = true;
            }

            return result;
        }

        private uint Read_Mac_Address()
        {
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10

            B14_BA_0.Read();
            B15_BA_1.Read();
            B16_BA_2.Read();

            return (B16_BA_2.Value << 16) | (B15_BA_1.Value << 8) | B14_BA_0.Value;
        }

        private bool Check_OTP_VALID(bool start, uint page)
        {
            uint addr_nvm;
            bool result = true;

            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            byte[] Data = new byte[4];

            if (start == false)
            {
                return false;
            }

            Enable_NVM_BIST();
            Power_On_NVM();

            if (page < 19)
            {
                addr_nvm = page * 44; // ble, control
            }
            else
            {
                addr_nvm = page * 44 + (page - 19) * 4; // Analog
            }

            // NVM Verify // VALID
            Data[0] = 0x1D;
            Data[1] = 0xCA;
            Data[2] = 0x1D;
            Data[3] = 0xCA;
            // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
            SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            SendData[4] = (byte)(0x13);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
            SendData[4] = (byte)(0x11);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Read bist_rdata
            SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            for (int j = 0; j < 4; j++)
            {
                if (RcvData[j] != Data[j])
                {
                    result = false;
                }
            }

            Power_Off_NVM();
            Disable_NVM_BIST();

            return result;
        }

        private bool Read_FF_NVM(bool start)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return true;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // bist_sel=1, bist_type=6, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=6, bist_cmd=0
            SendData[4] = 0x19;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] != 0x0b)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void Write_Register_AON_Fix_Value()
        {
            RegisterItem w_ADC_EOC_SKIP_TEMP = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_TEMP[5:0]");  // 0x34
            RegisterItem w_ADC_OFFSET_1 = Parent.RegMgr.GetRegisterItem("w_ADC_OFFSET_1");                 // 0x34
            RegisterItem w_ADC_OFFSET_2 = Parent.RegMgr.GetRegisterItem("w_ADC_OFFSET_2");                 // 0x34
            RegisterItem w_ADC_EOC_SKIP_VOLT = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_VOLT[5:0]");  // 0x35
            RegisterItem w_ADC_OFFSET_3 = Parent.RegMgr.GetRegisterItem("w_ADC_OFFSET_3");                 // 0x35
            RegisterItem w_ADC_OFFSET_4 = Parent.RegMgr.GetRegisterItem("w_ADC_OFFSET_4");                 // 0x35
            RegisterItem w_LPCAL_FINE_EN = Parent.RegMgr.GetRegisterItem("w_LPCAL_FINE_EN");               // 0x37
            RegisterItem w_TAMP_DET_EN = Parent.RegMgr.GetRegisterItem("w_TAMP_DET_EN");                   // 0x39
            RegisterItem w_GPIO_WKUP_MD = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_MD");                 // 0x41
            RegisterItem w_GPIO_WKUP_INTV = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_INTV[1:0]");        // 0x41
            RegisterItem O_CT_RES2_FINE = Parent.RegMgr.GetRegisterItem("O_CT_RES2_FINE[3:0]");            // 0x45
            RegisterItem w_GPIO_WKUP_EN = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_EN");                 // 0x4E
            RegisterItem O_ADC_CM = Parent.RegMgr.GetRegisterItem("O_ADC_CM[1:0]");                        // 0x55
            RegisterItem w_DA_GAIN_MAX = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_MAX[2:0]");              // 0x57
            RegisterItem I_EXT_DA_GAIN_CON = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_CON[1:0]");      // 0x5A
            RegisterItem w_XTAL_CUR_CFG_1 = Parent.RegMgr.GetRegisterItem("w_XTAL_CUR_CFG_1[5:0]");        // 0x5E
            RegisterItem w_DA_GAIN_INTV = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_INTV[1:0]");            // 0x61
            RegisterItem w_PLL_CH_19_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_19_GAIN[2:0]");        // 0x61
            RegisterItem VS_GainMode = Parent.RegMgr.GetRegisterItem("VS_GainMode");                       // 0x61
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL[5:2]");                  // 0x62
            RegisterItem w_TX_SEL_SELECT = Parent.RegMgr.GetRegisterItem("w_TX_SEL_SELECT[1:0]");          // 0x64
            RegisterItem w_PLL_PM_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_PM_GAIN[4:0]");              // 0x64
            RegisterItem w_PLL_CH_0_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_0_GAIN[4:0]");          // 0x65
            RegisterItem w_DA_PEN_SELECT = Parent.RegMgr.GetRegisterItem("w_DA_PEN_SELECT[1:0]");          // 0x65
            RegisterItem w_PLL_CH_12_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_12_GAIN[4:0]");        // 0x66
            RegisterItem w_ADC_SWITCH_MODE = Parent.RegMgr.GetRegisterItem("w_ADC_SWITCH_MODE");           // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_H = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[4]");        // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_L = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[3:0]");      // 0x67
            RegisterItem O_VOLS_SL_SEL = Parent.RegMgr.GetRegisterItem("O_VOLS_SL_SEL[3:0]");              // 0x74

            // w_ADC_OFFSET_1, w_ADC_OFFSET_2, w_ADC_OFFSET_3, w_ADC_OFFSET_4 = 1
            w_ADC_EOC_SKIP_TEMP.Read();
            w_ADC_OFFSET_1.Value = 1;
            w_ADC_OFFSET_2.Value = 1;
            w_ADC_EOC_SKIP_TEMP.Value = 1;
            w_ADC_EOC_SKIP_TEMP.Write();

            w_ADC_EOC_SKIP_VOLT.Read();
            w_ADC_OFFSET_3.Value = 1;
            w_ADC_OFFSET_4.Value = 1;
            w_ADC_EOC_SKIP_VOLT.Value = 63;
            w_ADC_EOC_SKIP_VOLT.Write();

            // VS_SLOPE = 8
            O_VOLS_SL_SEL.Read();
            O_VOLS_SL_SEL.Value = 8;
            O_VOLS_SL_SEL.Write();

            w_LPCAL_FINE_EN.Read();
            w_LPCAL_FINE_EN.Value = 1;
            w_LPCAL_FINE_EN.Write();

            w_TAMP_DET_EN.Read();
            w_TAMP_DET_EN.Value = 1;
            w_TAMP_DET_EN.Write();

            w_GPIO_WKUP_MD.Read();
            w_GPIO_WKUP_MD.Value = 1;
            w_GPIO_WKUP_MD.Write();

            O_CT_RES2_FINE.Read();
            O_CT_RES2_FINE.Value = 9;
            O_CT_RES2_FINE.Write();

            w_GPIO_WKUP_EN.Read();
            w_GPIO_WKUP_EN.Value = 1;
            w_GPIO_WKUP_EN.Write();

            w_DA_GAIN_MAX.Read();
            w_DA_GAIN_MAX.Value = 6;
            w_DA_GAIN_MAX.Write();

            I_EXT_DA_GAIN_CON.Read();
            I_EXT_DA_GAIN_CON.Value = 2;
            I_EXT_DA_GAIN_CON.Write();

            w_XTAL_CUR_CFG_1.Read();
            w_XTAL_CUR_CFG_1.Value = 31;
            w_XTAL_CUR_CFG_1.Write();

            w_DA_GAIN_INTV.Read();
            w_DA_GAIN_INTV.Value = 2;
            w_PLL_CH_19_GAIN.Value = 0;
            w_DA_GAIN_INTV.Write();

            BGR_TC_CTRL.Read();
            BGR_TC_CTRL.Value = 11;
            BGR_TC_CTRL.Write();

            w_TX_SEL_SELECT.Read();
            w_TX_SEL_SELECT.Value = 2;
            w_PLL_PM_GAIN.Value = 18;
            w_TX_SEL_SELECT.Write();

            w_PLL_CH_0_GAIN.Read();
            w_PLL_CH_0_GAIN.Value = 17;
            w_DA_PEN_SELECT.Value = 2;
            w_PLL_CH_0_GAIN.Write();

            w_PLL_CH_12_GAIN.Read();
            w_PLL_CH_12_GAIN.Value = 16;
            w_ADC_SAMPLE_CNT_H.Value = 1;
            w_ADC_SWITCH_MODE.Value = 1;
            w_PLL_CH_12_GAIN.Write();

            w_ADC_SAMPLE_CNT_L.Read();
            w_ADC_SAMPLE_CNT_L.Value = 15;
            w_ADC_SAMPLE_CNT_L.Write();
        }

        private bool Run_Cal_BGR(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 300;
            double d_lsl = 295, d_usl = 305;
            uint ldo_val, ldo_val_1;

            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");    // 0x53

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_BGR(true);

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(2, (1 + cnt), (double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));

                ULP_BGR_CONT.Read();
                Parent.xlMgr.Sheet.Select("BGR");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            ldo_val = 15;
            ldo_val_1 = 0;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }

            Set_TestInOut_For_BGR(false);

            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_ALLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 810;
            double d_lsl = 800, d_usl = 820;
            uint ldo_val, ldo_val_1;

            RegisterItem O_ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");        // 0x54
            RegisterItem O_ULP_LDO_LV_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[2:0]");  // 0x61

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            O_ULP_LDO_CONT.Read();
            O_ULP_LDO_LV_CONT.Read();

            O_ULP_LDO_CONT.Value = 8;
            O_ULP_LDO_CONT.Write();
            O_ULP_LDO_LV_CONT.Value = 0;
            O_ULP_LDO_LV_CONT.Write();
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("ALLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv > 816)
            {
                O_ULP_LDO_LV_CONT.Value = 3;
                O_ULP_LDO_LV_CONT.Write();
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (O_ULP_LDO_LV_CONT.Value).ToString());
                Parent.xlMgr.Cell.Write(x_pos + 1, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_MLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 1000;
            double d_lsl = 990, d_usl = 1010;
            uint ldo_val, ldo_val_1;

            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");          // 0x53
            RegisterItem PMU_MLDO_Coarse_L = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[0]");    // 0x62
            RegisterItem PMU_MLDO_Coarse_H = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1]");    // 0x63

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            PMU_MLDO_Coarse_H.Read();
            PMU_MLDO_Coarse_L.Read();
            PMU_LDO_CONT.Read();

            PMU_LDO_CONT.Value = 7;
            PMU_LDO_CONT.Write();
            PMU_MLDO_Coarse_L.Value = 0;
            PMU_MLDO_Coarse_L.Write();

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("MLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv < 1006)
            {
                PMU_MLDO_Coarse_L.Value = 1;
                PMU_MLDO_Coarse_L.Write();
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < 1006)
                {
                    PMU_MLDO_Coarse_L.Value = 0;
                    PMU_MLDO_Coarse_L.Write();
                    PMU_MLDO_Coarse_H.Value = 1;
                    PMU_MLDO_Coarse_H.Write();
                }
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();
            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            if (x_pos != 0)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            }
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ((PMU_MLDO_Coarse_H.Value << 1) | PMU_MLDO_Coarse_L.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_32K_RCOSC(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_khz, d_diff_khz;
            double d_lsl = 32.113, d_usl = 33.423, d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;

            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");      // 0x4D
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");     // 0x4E

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_RCOSC(true);
            if (x_pos != 0)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_khz.ToString("F3"));

                Parent.xlMgr.Sheet.Select("RCOSC");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            // RTC_SCKF[5:0]
            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            for (int val = 4; val >= 0; val--)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
                }
                if (d_freq_khz > d_target_khz)
                {
                    osc_val_l += (uint)(1 << val);
                }
                else
                {
                    osc_val_l -= (uint)(1 << val);
                }
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
            {
                if (osc_val_l != 63) osc_val_l += 1;
            }
            else
            {
                if (osc_val_l != 0) osc_val_l -= 1;
            }
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (RTC_SCKF_H.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), (RTC_SCKF_L.Value).ToString());
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_freq_khz.ToString("F3"));
            }
            Set_TestInOut_For_RCOSC(false);

            if ((d_freq_khz < d_lsl) || (d_freq_khz > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_Temp_Sensor(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 140; // 30deg
            int lsl = 135, usl = 145;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem TEMP_CONT_L = Parent.RegMgr.GetRegisterItem("O_TEMP_CONT[4:0]");  // 0x55
            RegisterItem TEMP_CONT_H = Parent.RegMgr.GetRegisterItem("TEMP_TRIM[5]");      // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
            TEMP_CONT_H.Read();
            TEMP_CONT_L.Read();

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, 1, 63, 31);
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();
            Set_TestInOut_For_VTEMP(false);
#if false // cal
            Parent.xlMgr.Sheet.Select("Temp_Sen");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
                u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
            }
#else // fix
            u_adc_val = 0;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            lsl = 0;
            usl = 255;
#endif
            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

            Set_TestInOut_For_VTEMP(false);
#if false
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_Voltage_Scaler(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 45; // 2.0V
            int lsl = 40, usl = 50;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem VOLSCAL_CON_L = Parent.RegMgr.GetRegisterItem("O_VOLSCAL_CON[3:0]");  // 0x56
            RegisterItem VOLTAGE_CON_H = Parent.RegMgr.GetRegisterItem("VOLTAGE_CON[5:4]");    // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.0,(@3)");
#else
            PowerSupply0.Write("VOLT 2.0");
#endif
            VOLTAGE_CON_H.Read();
            VOLSCAL_CON_L.Read();

            Set_TestInOut_For_VS(true);

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, 1, 63, 31);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();

            Parent.xlMgr.Sheet.Select("VS");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
                u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
            }

            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif

            Set_TestInOut_For_VS(false);
#if true
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_32M_XTAL_Load_Cap(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_mhz;
            double d_diff_mhz, d_target_mhz = 2402;
            double d_lsl = 2401.9952, d_usl = 2402.0048;
            uint osc_val, osc_val_1;

            RegisterItem XTAL_LOAD_CONT = Parent.RegMgr.GetRegisterItem("O_XTAL_LOAD_CONT[4:0]");  // 0x4F

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            XTAL_LOAD_CONT.Read();
            if (x_pos != 0)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
                Parent.xlMgr.Sheet.Select("XTAL_Cap");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            osc_val = 15;
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            for (int val = 2; val >= 0; val--)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)val), (1 + cnt), d_freq_mhz.ToString("F4"));
                }
                if (d_freq_mhz > d_target_mhz)
                {
                    osc_val += (uint)(1 << val);
                }
                else
                {
                    osc_val -= (uint)(1 << val);
                }
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            osc_val_1 = osc_val;
            d_diff_mhz = Math.Abs(d_freq_mhz - d_target_mhz);

            if (d_freq_mhz > d_target_mhz)
            {
                if (osc_val != 31) osc_val += 1;
            }
            else
            {
                if (osc_val != 0) osc_val -= 1;
            }
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            System.Threading.Thread.Sleep(1);

            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            if (Math.Abs(d_freq_mhz - d_target_mhz) > d_diff_mhz)
            {
                osc_val = osc_val_1;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            }

            if ((d_freq_mhz < d_lsl) || (d_freq_mhz > d_usl))
            {
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return false;
            }
            else
            {
#if true // offset
                osc_val += 3;
                if (osc_val > 31) osc_val = 31;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
#endif
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return true;
            }

        }

        private bool Run_Cal_PLL_2PM(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_fsm;
            uint crc_flag;

            RegisterItem PM_IN = Parent.RegMgr.GetRegisterItem("O_PM_IN[7:0]");                    // 0x58

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            SendCommand("ook 70");
            System.Threading.Thread.Sleep(100);

            u_fsm = Run_Read_FSM_Status() & 0x000f;
            crc_flag = CRC_FLAG_READ();

            if (u_fsm != 7)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "FSM_Error");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) > 1)) // WUR_PKT_END = 1
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CRC_ERROR");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) == 1))
            {
                PM_IN.Read();
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (PM_IN.Value).ToString());
            }
            else
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CHECK_LOG");
                Console.WriteLine("u_fsm : {0}", u_fsm);
                Console.WriteLine("crc_flag : {0}", crc_flag);
                return false;
            }
            return true;
        }

        private void Set_AON_REG_For_BLE(byte[] aon_ble)
        {
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");            // 0x04
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");            // 0x05
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");            // 0x06
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");            // 0x07
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");            // 0x08
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");            // 0x09
            RegisterItem B10_LEN = Parent.RegMgr.GetRegisterItem("B10_LEN[7:0]");              // 0x0A
            RegisterItem B11_MSDF = Parent.RegMgr.GetRegisterItem("B11_MSDF[7:0]");            // 0x0B
            RegisterItem B12_CID_0 = Parent.RegMgr.GetRegisterItem("B12_CID_0[7:0]");          // 0x0C
            RegisterItem B13_CID_1 = Parent.RegMgr.GetRegisterItem("B13_CID_1[7:0]");          // 0x0D
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10
            RegisterItem B17_BA_3 = Parent.RegMgr.GetRegisterItem("B17_BA_3[7:0]");            // 0x11
            RegisterItem B18_BA_4 = Parent.RegMgr.GetRegisterItem("B18_BA_4[7:0]");            // 0x12
            RegisterItem B19_BA_5 = Parent.RegMgr.GetRegisterItem("B19_BA_5[7:0]");            // 0x13
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");            // 0x14
            RegisterItem B21_DC_1 = Parent.RegMgr.GetRegisterItem("B21_DC_1[7:0]");            // 0x15
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");            // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");            // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");            // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");            // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");            // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");            // 0x1B
            RegisterItem B28_RSV_0 = Parent.RegMgr.GetRegisterItem("B28_PUTC_0[7:0]");         // 0x1C
            RegisterItem B29_RSV_1 = Parent.RegMgr.GetRegisterItem("B29_PUTC_1[7:0]");         // 0x1D
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");            // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");            // 0x1F
            RegisterItem B32_PUTC_0 = Parent.RegMgr.GetRegisterItem("B32_CC_0[7:0]");          // 0x20
            RegisterItem B33_PUTC_1 = Parent.RegMgr.GetRegisterItem("B33_CC_1[7:0]");          // 0x21
            RegisterItem B34_PUTC_2 = Parent.RegMgr.GetRegisterItem("B34_CC_2[7:0]");          // 0x22
            RegisterItem B35_PUTC_3 = Parent.RegMgr.GetRegisterItem("B35_CC_3[7:0]");          // 0x23
            RegisterItem B36_BL = Parent.RegMgr.GetRegisterItem("B36_BL[7:0]");                // 0x24
            RegisterItem B37_TP = Parent.RegMgr.GetRegisterItem("B37_TP[7:0]");                // 0x25
            RegisterItem B38_SN_0 = Parent.RegMgr.GetRegisterItem("B38_SN_0[7:0]");            // 0x26
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27
            RegisterItem B40_MIC = Parent.RegMgr.GetRegisterItem("B40_MIC[7:0]");              // 0x28
            RegisterItem B41_advDelay = Parent.RegMgr.GetRegisterItem("B41_ADD_0[7:0]");       // 0x29
            RegisterItem B42_advDelay = Parent.RegMgr.GetRegisterItem("B42_ADD_1[7:0]");       // 0x2A
            RegisterItem B43_advDelay = Parent.RegMgr.GetRegisterItem("B43_ADD_2[7:0]");       // 0x2B

            B04_BA_0.Value = aon_ble[0];
            B04_BA_0.Write();

            B05_BA_1.Value = aon_ble[1];
            B05_BA_1.Write();

            B06_BA_2.Value = aon_ble[2];
            B06_BA_2.Write();

            B07_BA_3.Value = aon_ble[3];
            B07_BA_3.Write();

            B08_BA_4.Value = aon_ble[4];
            B08_BA_4.Write();

            B09_BA_5.Value = aon_ble[5];
            B09_BA_5.Write();

            B10_LEN.Value = aon_ble[6];
            B10_LEN.Write();

            B11_MSDF.Value = aon_ble[7];
            B11_MSDF.Write();

            B12_CID_0.Value = aon_ble[8];
            B12_CID_0.Write();

            B13_CID_1.Value = aon_ble[9];
            B13_CID_1.Write();

            B14_BA_0.Value = aon_ble[10];
            B14_BA_0.Write();

            B15_BA_1.Value = aon_ble[11];
            B15_BA_1.Write();

            B16_BA_2.Value = aon_ble[12];
            B16_BA_2.Write();

            B17_BA_3.Value = aon_ble[13];
            B17_BA_3.Write();

            B18_BA_4.Value = aon_ble[14];
            B18_BA_4.Write();

            B19_BA_5.Value = aon_ble[15];
            B19_BA_5.Write();

            B20_DC_0.Value = aon_ble[16];
            B20_DC_0.Write();

            B21_DC_1.Value = aon_ble[17];
            B21_DC_1.Write();

            B22_WP_0.Value = aon_ble[18];
            B22_WP_0.Write();

            B23_WP_1.Value = aon_ble[19];
            B23_WP_1.Write();

            B24_AI_0.Value = aon_ble[20];
            B24_AI_0.Write();

            B25_AI_1.Value = aon_ble[21];
            B25_AI_1.Write();

            B26_AI_2.Value = aon_ble[22];
            B26_AI_2.Write();

            B27_AI_3.Value = aon_ble[23];
            B27_AI_3.Write();

            B28_RSV_0.Value = aon_ble[24];
            B28_RSV_0.Write();

            B29_RSV_1.Value = aon_ble[25];
            B29_RSV_1.Write();

            B30_AD_0.Value = aon_ble[26];
            B30_AD_0.Write();

            B31_AD_1.Value = aon_ble[27];
            B31_AD_1.Write();

            B32_PUTC_0.Value = aon_ble[28];
            B32_PUTC_0.Write();

            B33_PUTC_1.Value = aon_ble[29];
            B33_PUTC_1.Write();

            B34_PUTC_2.Value = aon_ble[30];
            B34_PUTC_2.Write();

            B35_PUTC_3.Value = aon_ble[31];
            B35_PUTC_3.Write();

            B36_BL.Value = aon_ble[32];
            B36_BL.Write();

            B37_TP.Value = aon_ble[33];
            B37_TP.Write();

            B38_SN_0.Value = aon_ble[34];
            B38_SN_0.Write();

            B39_SN_1.Value = aon_ble[35];
            B39_SN_1.Write();

            B40_MIC.Value = aon_ble[36];
            B40_MIC.Write();

            B41_advDelay.Value = aon_ble[37];
            B41_advDelay.Write();

            B42_advDelay.Value = aon_ble[38];
            B42_advDelay.Write();

            B43_advDelay.Value = aon_ble[39];
            B43_advDelay.Write();
        }

        private void Run_Write_OTP_With_OOK(bool start, uint page)
        {
            string cmd;

            if (start == false)
            {
                return;
            }

            cmd = "ook 41." + page.ToString("X2");
            SendCommand(cmd);

            System.Threading.Thread.Sleep(500);
        }

        private uint Run_Verify_OTP_With_BIST(bool start, uint page)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return 65535;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // set page_flag
            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = (byte)(page & 0xff);
            SendData[5] = (byte)((page >> 8) & 0xff);
            SendData[6] = (byte)(((page >> 16) & 0x3f) | 0x80);
            SendData[7] = 0xBB;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1F;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=0
            SendData[4] = 0x1D;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] == 0x0b)
            {
                return 0x400;
            }
            else
            {
                return RcvData[0];
            }
        }

        private bool Run_Measure_Initial(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }
            // BGR, ALLDO, MLDO
            Set_TestInOut_For_BGR(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                // BGR
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 295) || (d_val > 305))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_12");
                        result = false;
                    }
                }
                // ALLDO
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 3), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 800) || (d_val > 820))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_13");
                        result = false;
                    }
                }

                // MLDO
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 6), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 985) || (d_val > 1015))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_14");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_BGR(false);

            // RCOSC
            Set_TestInOut_For_RCOSC(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + i + 9), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 32.113) || (d_val > 33.423))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_16");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_RCOSC(false);

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_ADC(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            uint u_adc_code, u_temp_code, u_vbat_code;
            uint u_lsl_temp = 130, u_usl_temp = 150, u_lsl_vbat = 255, u_usl_vbat = 0;
            double d_volt_mv;

            if (start == false)
            {
                for (int i = 0; i < 8; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }

            Set_TestInOut_For_VTEMP(true);

            System.Threading.Thread.Sleep(1);
            for (int i = 0; i < 4; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@3)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        u_lsl_vbat = 186;
                        u_usl_vbat = 196;
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        u_lsl_vbat = 96;
                        u_usl_vbat = 106;
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.0,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.0");
#endif
                        u_lsl_vbat = 40;
                        u_usl_vbat = 50;
                        break;
                    case 3:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@3)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        u_lsl_vbat = 6;
                        u_usl_vbat = 16;
                        break;
                    default:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                }
                System.Threading.Thread.Sleep(100);
                u_adc_code = Read_ADC_Result(false, 1, 63, 31);
                u_temp_code = u_adc_code & 0xff;
                u_vbat_code = (u_adc_code >> 8) & 0xff;
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 0), (y_pos + cnt), u_temp_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 1), (y_pos + cnt), u_vbat_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
                Disable_ADC();
#if false
                if ((u_temp_code < u_lsl_temp) || (u_temp_code > u_usl_temp))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_17");
                        result = false;
                    }
                }
#endif
                if ((u_vbat_code < u_lsl_vbat) || (u_vbat_code > u_usl_vbat))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_18");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_VTEMP(false);
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_VCO_Range(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_freq_mhz;
            double d_sl = 2402;

            RegisterItem VCO_TEST = Parent.RegMgr.GetRegisterItem("O_VCO_TEST");           // 0x41
            RegisterItem VCO_CBANK_L = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[7:0]");  // 0x42
            RegisterItem VCO_CBANK_H = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[9:8]");  // 0x43

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            VCO_TEST.Read();
            VCO_CBANK_L.Read();
            VCO_CBANK_H.Read();

            VCO_TEST.Value = 1;
            VCO_TEST.Write();

            // VCO Low = 0
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 0;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:SPAN 500 MHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.2 GHZ");
            System.Threading.Thread.Sleep(400);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 0), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz > d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO High = 1023
            VCO_CBANK_L.Value = 255;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 3;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.7 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz < d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO Mid = 512
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 2;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.5 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(false);
            VCO_TEST.Value = 0;
            VCO_TEST.Write();

            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");

            return result;
        }

        private bool Run_Measure_Tx_Power_Harmonic(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_power_dbm, d_freq_MHz, d_cur_mA;
            double d_lsl = -2, d_usl = 5;
            uint ch;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return result;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;
            PowerSupply0.Write("SENS:CURR:RANG 0.01,(@3)");

            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                    case 1:
                        ch = 12;
                        d_freq_MHz = 2426;
                        break;
                    case 2:
                        ch = 39;
                        d_freq_MHz = 2480;
                        break;
                    default:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                }
                Write_Register_Fractional_Calc_Ch(ch);
                for (int j = 0; j < 2; j++)
                {
                    if (j == 0)
                    {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                    }
                    else
                    {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                    }
                    Write_Register_Tx_Tone_Send(true);
                    SpectrumAnalyzer.Write("FREQ:CENT " + d_freq_MHz + " MHZ");
                    System.Threading.Thread.Sleep(400);
                    d_cur_mA = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000.0;
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 2), (y_pos + cnt), d_cur_mA.ToString("F2"));
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < d_lsl) || (d_power_dbm > d_usl))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_21");
                            result = false;
                        }
                    }
                    if ((d_cur_mA < 8) || (d_cur_mA > 13))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_28");
                            result = false;
                        }
                    }

                    SpectrumAnalyzer.Write("FREQ:CENT " + (d_freq_MHz * 2) + " MHZ");
                    System.Threading.Thread.Sleep(50);
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 1), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < (d_lsl - 68)) || (d_power_dbm > (d_usl - 30)))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_22");
                            result = false;
                        }
                    }
                }
                Write_Register_Tx_Tone_Send(false);
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");
            Write_Register_Fractional_Calc_Ch(0);
            PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");

            return result;
        }

        private bool Run_Measure_INTB(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val_H;
            double d_val_L;

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            d_val_H = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
            I2C.GPIOs[3].Direction = GPIO_Direction.Output; // AD7
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].State = GPIO_State.Low;
            System.Threading.Thread.Sleep(115);
            I2C.GPIOs[3].State = GPIO_State.High;
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].Direction = GPIO_Direction.Input;
            System.Threading.Thread.Sleep(100);
            d_val_L = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));

            if (d_val_H > 0.8 && d_val_L < 0.5)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");
            }
            else
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL");
                return false;
            }

            return result;
        }

        private bool Run_Measure_TAMPER(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27 TAMPER dectec : pass = 0 fail = 1

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

#if !(POWER_SUPPLY_NEW)
            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
            return true;
#endif

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.9,(@2)");    // TAMPER 0.9 set
            PowerSupply0.Write("OUTP ON,(@2)");     // TAMPER 0.9 set            
#else
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 0.9");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_H");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW) 
            PowerSupply0.Write("VOLT 0.3,(@2)");   //TAMPER 0.3 set
#else
            PowerSupply0.Write("VOLT 0.3");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_L");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.6,(@2)");   //TAMPER 0.6 set
#else
            PowerSupply0.Write("VOLT 0.6");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x00)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_M");
                result = false;
                return result;
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("OUTP OFF,(@2)");
#else
            PowerSupply0.Write("VOLT 0.0");
            PowerSupply0.Write("INST:NSEL 2");
#endif
            return result;
        }

        private bool Run_Measure_BLE_Packet(bool start, int cnt, int x_pos, int y_pos, bool result, byte[] aon_ble)
        {
            int len;
            byte b;
            byte[] r_data = new byte[37];

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            Serial.RcvQueue.Clear();
            SendCommand("ook 30");
            System.Threading.Thread.Sleep(150);

            len = Serial.RcvQueue.Count;
            Console.WriteLine("UART len : " + len);
            SendCommand("mode ook");
            for (int i = 0; i < 11; i++)
            {
                b = Serial.RcvQueue.Get();
            }

            for (int i = 0; i < 36; i++)
            {
                b = Serial.RcvQueue.Get();

                if (b > 0x60)
                {
                    r_data[i] = (byte)((b - 87) << 4); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] = (byte)((b - 55) << 4); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] = (byte)((b - 48) << 4); // '0' = 0x30 = 48
                }

                b = Serial.RcvQueue.Get();
                if (b > 0x60)
                {
                    r_data[i] += (byte)((b - 87)); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] += (byte)((b - 55)); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] += (byte)((b - 48)); // '0' = 0x30 = 48
                }

                //b = UartRcvQueue.GetByte(); // space

                Console.Write("{0:X}-", r_data[i]);
                if (r_data[i] == aon_ble[i])
                {
                    continue;
                }
                else if ((i == 32) || (i == 33)) // BL, TP
                {
                    if (r_data[i] > 150)
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if (i == 35) // TAMPER
                {
                    if (!((r_data[i] == 0x00) || (r_data[i] == 0x80)))
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if ((i < 3) || (i > 9) && (i < 12))
                {
                    continue;
                }
                else
                {
                    Console.Write("\r\nFail!!{0} read : {1:X}, write : {2:X}\r\n", i, r_data[i], aon_ble[i]);
                    Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), i.ToString());
                    result = false;
                    return result;
                }

                if ((b == '\r') || (b == '\n'))
                {
                    break;
                }
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "P");
            return result;
        }

        private void Test_Good_Chip_Sorting_Rev3(int start_cnt)
        {
            int cnt = 0, pass = 0, fail = 0;
            double d_val;
            int x_pos = 2, y_pos = 12;
            bool result;
            uint mac_code = 0x000000;
            bool OTP_W_Flag = false;

            byte[] aon_ble = { 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0x1E, 0xFF, 0x6D, 0x0B, 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

            Check_Instrument();

            if (I2C.IsOpen == false) // Check I2C connection
            {
                MessageBox.Show("Check I2C");
                return;
            }

            MessageBox.Show("장비 설정을 확인해주세요.\r\n\r\n1.N6705B Power Supply\r\n  - VBAT\r\n  - TEST_IN_OUT(N6705B)\r\n\r\n" +
                "2.SpectrumAnalyzer\r\n  - TX\r\n\r\n3.Digital Multimeter\r\n  - Current mode : VBAT\r\n  - Voltage mode : ALLDO / MLDO / TEST_IN_OUT\r\n\r\n" +
                "4.UM232H\r\n  - AC0_GPIOH0 : I2C EN\r\n  - AD6_GPIOL2 : RX/TX Switch\r\n  - AD7_GPIOL3 : INTB");

            cnt = start_cnt - 1;
            pass = cnt;

            while (true)
            {
                I2C.GPIOs[3].Direction = GPIO_Direction.Input;   // AD7_GPIOL3(INTB)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].Direction = GPIO_Direction.Output;   // AD6_GPIOL2(TRX S/W)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].Direction = GPIO_Direction.Output;   // AC0_GPIOH0(Level Shifter EN)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);

                // Power off
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 0.0,(@3)");
                PowerSupply0.Write("OUTP ON,(@3)");
                PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");
#else
                PowerSupply0.Write("INST:NSEL 1");
                PowerSupply0.Write("VOLT 0.0");
                PowerSupply0.Write("OUTP ON");
#endif
                SendCommand("mode ook\n");
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n\r\nTest\t: " + cnt + "\r\nPass\t: " + pass + "\r\nFail\t: " + fail
                                                        , Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    cnt++;
                    fail++;
                    Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), cnt.ToString());
                    result = true;
                }
                else
                {
                    return;
                }

                DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:AC?"); // ALLDO
                // Power on
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 3.3,(@3)");
                System.Threading.Thread.Sleep(200);
                PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                PowerSupply0.Write("VOLT 3.3");
                System.Threading.Thread.Sleep(500);
                PowerSupply0.Write("VOLT 2.5");
#endif
                // Check MLDO Voltage

                System.Threading.Thread.Sleep(800);
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_val.ToString("F3"));
                if (d_val > 0.2)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                    continue;
                }

                // Seq1 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#else
                d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                d_val = 600;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
#endif
                if ((d_val < 500) || (d_val > 850))
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                    continue;
                }

                I2C.GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                // Wake-up
                WakeUp_I2C();
                System.Threading.Thread.Sleep(5);

                // READ_DIVICE_ID
                if (Check_Revision_Information(0x5F) == false) // N1 = 0x5D, N1B = 0x5E, N1C = 0x5F
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "F");
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_24");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "P");
                }

                mac_code = Read_Mac_Address();
                aon_ble[0] = (byte)(mac_code & 0xff);
                aon_ble[1] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[2] = (byte)((mac_code >> 16) & 0xff);
                aon_ble[10] = (byte)(mac_code & 0xff);
                aon_ble[11] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[12] = (byte)((mac_code >> 16) & 0xff);
                Parent.xlMgr.Cell.Write((x_pos + 81), (y_pos + cnt), (aon_ble[5].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 82), (y_pos + cnt), (aon_ble[4].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 83), (y_pos + cnt), (aon_ble[3].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 84), (y_pos + cnt), (aon_ble[2].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 85), (y_pos + cnt), (aon_ble[1].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 86), (y_pos + cnt), (aon_ble[0].ToString("X")));

                if (Check_OTP_VALID(true, 21) == true)
                {
                    goto CAL_SKIP;
                }

                // Seq1 5.OTP read
                if (Read_FF_NVM(false) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_10");
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "F");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "Skip");
                }

                Write_Register_AON_Fix_Value();

#if true // For Test
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
                Parent.xlMgr.Cell.Write(3, (1 + cnt), (double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Cell.Write(4, (1 + cnt), (double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
#endif
                // Seq2 1.BGR Trim
                if (Run_Cal_BGR(true, cnt, (x_pos + 6), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_4");
                    continue;
                }
                // Seq2 2.ALON LDO Trim
                if (Run_Cal_ALLDO(true, cnt, (x_pos + 8), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_5");
                    continue;
                }
                // Seq2 3.MLDO Trim
                if (Run_Cal_MLDO(true, cnt, (x_pos + 11), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_6");
                    continue;
                }
                // Seq2 4.32K RCOSC Trim
                if (Run_Cal_32K_RCOSC(true, cnt, (x_pos + 14), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_7");
                    continue;
                }
                // Seq2 5.Temp Sensor Trim
                if (Run_Cal_Temp_Sensor(true, cnt, (x_pos + 18), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_8");
                    continue;
                }
                // Seq2 5.Voltage Scaler Trim
                if (Run_Cal_Voltage_Scaler(true, cnt, (x_pos + 22), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_25");
                    continue;
                }
                // Seq2 6.32M X-tal Load Cap Trim
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                if (Run_Cal_32M_XTAL_Load_Cap(true, cnt, (x_pos + 26), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_9");
                    continue;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);
                // Seq2 7.PLL 2PM Trim
                if (Run_Cal_PLL_2PM(true, cnt, (x_pos + 29), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_20");
                    continue;
                }
                // Seq2 8.OTP Trim data write
                // Set AON_REG(BLE) For OTP
                Set_AON_REG_For_BLE(aon_ble);

                Run_Write_OTP_With_OOK(OTP_W_Flag, 0);
                // Run_Write_OTP_With_OOK(OTP_W_Flag, 15);
                Run_Write_OTP_With_OOK(OTP_W_Flag, 21);
                d_val = Run_Verify_OTP_With_BIST(OTP_W_Flag, 0x208001);
                if (d_val == 65535)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "Skip");
                }
                else if (d_val != 0x400)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_11");
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), d_val.ToString());
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "P");
                }

            CAL_SKIP:
                if (OTP_W_Flag) // Reset
                {
                    I2C.GPIOs[4].State = GPIO_State.High;
                    System.Threading.Thread.Sleep(10);
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 0.0,(@3)");
#else
                    PowerSupply0.Write("VOLT 0.0");
#endif
                    System.Threading.Thread.Sleep(500);

#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 3.3,(@3)");
                    System.Threading.Thread.Sleep(200);
                    PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                    PowerSupply0.Write("VOLT 3.3");
                    System.Threading.Thread.Sleep(500);
                    PowerSupply0.Write("VOLT 2.5");
#endif
                    // Check MLDO Voltage                    
                    System.Threading.Thread.Sleep(800);
                    d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), d_val.ToString("F3"));
                    if (d_val > 0.2)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                        continue;
                    }

                    // Seq2 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                    d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), (d_val).ToString("F3"));
#else
                    d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                    d_val = 600;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
#endif
                    if ((d_val < 500) || (d_val > 850))
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                        continue;
                    }

                    I2C.GPIOs[4].State = GPIO_State.Low;
                    System.Threading.Thread.Sleep(10);
                    // Wake-up
                    WakeUp_I2C();
                    System.Threading.Thread.Sleep(5);
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), "Skip");
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
                }

                // Seq3 1.Initial Test
                if (Run_Measure_Initial(true, cnt, (x_pos + 33), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 2.ADC Test
                if (Run_Measure_ADC(true, cnt, (x_pos + 45), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                // Seq3 4.VCO Range Test
                if (Run_Measure_VCO_Range(true, cnt, (x_pos + 57), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 4. Tx output power and Harmonic
                if (Run_Measure_Tx_Power_Harmonic(true, cnt, (x_pos + 60), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);

                // INTB Test
                if (Run_Measure_INTB(true, cnt, (x_pos + 79), y_pos, result) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_27");
                    }
                    result = false;
                }
                // TAMPER Test
                if (Run_Measure_TAMPER(true, cnt, (x_pos + 78), y_pos, result) == false)
                {
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("OUTP OFF,(@2)");
#endif
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_26");
                    }
                    result = false;
                }

                // BLE Test
                if (Run_Measure_BLE_Packet(true, cnt, (x_pos + 80), y_pos, result, aon_ble) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_23");
                    }
                    result = false;
                }

                if (result == true)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "PASS");
                    fail--;
                    pass++;
                }
            }
        }

        private void DUT_TempSensor_Test()
        {
            byte Addr;
            byte[] RcvData = new byte[2];

            Parent.xlMgr.Sheet.Add(DateTime.Now.ToString("MMddHHmmss_") + "DUT_TS");
            Parent.xlMgr.Cell.Write(1, 1, "No.");
            Parent.xlMgr.Cell.Write(2, 1, "ADC");
            Parent.xlMgr.Cell.Write(3, 1, "Temp");

            I2C.GPIOs[2].Direction = GPIO_Direction.Output;
            System.Threading.Thread.Sleep(10);

            Addr = 0x00;
#if false            
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(200); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Parent.xlMgr.Cell.Write(1, 2 + i, i.ToString());
                Parent.xlMgr.Cell.Write(2, 2 + i, ((RcvData[0] << 8) | RcvData[1]).ToString());
                Parent.xlMgr.Cell.Write(3, 2 + i, (((RcvData[0] << 8) | RcvData[1]) * 0.0078125).ToString());

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(500);
            }
#else
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(300); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Console.WriteLine("ADC : {0}\tTemp : {1}", (RcvData[0] << 8) | RcvData[1], ((RcvData[0] << 8) | RcvData[1]) * 0.0078125);

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(100);
            }
#endif
        }
        #endregion

        # region Function for DTM
        private void SetBleDTM_AonControlReg()
        {
            WriteRegister(0x33, 0x17);
            WriteRegister(0x34, 0x3F);       //ADC_EOC_SKIP_TEMP
            WriteRegister(0x35, 0x3F);       //ADC_EOC_SKIP_VOLT
            WriteRegister(0x36, 0xFF);
            WriteRegister(0x37, 0x47);
            WriteRegister(0x38, 0x00);
            WriteRegister(0x39, 0x40);
            WriteRegister(0x3A, 0x4A);
            WriteRegister(0x3B, 0x00);
        }

        private void SetBleDTM_AonControlReg_OTP()
        {
            WriteRegister(0x33, 0x17);
            WriteRegister(0x37, 0x47);

            WriteRegister(0x45, 0x7);   //CT_RES32_FINE = 7 (Bandwidth change)
            WriteRegister(0x64, 0xD2);  //TX_SEL = CoraseLock & PLL_PM_GIAN = 18;
            WriteRegister(0x65, 0x50);  //DA_PEN = CoraseLock & CH_0_GAIN = 16;
        }

        private void SetBleDTM_AonAnalogReg()
        {
            WriteRegister(0x62, 0x66);      //r3 ONLY
            WriteRegister(0x61, 0x33);
            WriteRegister(0x53, 0xC7);      //FC_LGA
            WriteRegister(0x54, 0xBC);
            WriteRegister(0x3C, 0x10);
            WriteRegister(0x3D, 0x12);
            WriteRegister(0x3E, 0x00);      //R3 ONLY
            WriteRegister(0x3F, 0x72);
            WriteRegister(0x40, 0x82);
            WriteRegister(0x41, 0x80);
            WriteRegister(0x42, 0x00);
            WriteRegister(0x43, 0xF2);
            WriteRegister(0x44, 0xFF);
            WriteRegister(0x45, 0x02);
            WriteRegister(0x46, 0xFF);
            WriteRegister(0x47, 0x42);
            WriteRegister(0x48, 0x62);
            WriteRegister(0x49, 0xFF);
            WriteRegister(0x4A, 0x3F);
            WriteRegister(0x4B, 0x16);
            WriteRegister(0x4C, 0x00);
            WriteRegister(0x4D, 0x29);
            WriteRegister(0x4E, 0xB2);
            WriteRegister(0x4F, 0x61);
            WriteRegister(0x50, 0xC5);
            WriteRegister(0x51, 0x2F);      //R3 ONLY
            WriteRegister(0x52, 0x22);

            WriteRegister(0x55, 0xCE);
            WriteRegister(0x56, 0x00);  //R3 ONLY DYNAMIC CONTROL FOR BLE LINK
            WriteRegister(0x57, 0x07);  //TX_BUF, DA_PEN, PRE_DA_PEN, DA_GAIN R3 ONLY DYNIMIC CONTROL FOR BLE LINK
            WriteRegister(0x58, 0x33);  //2PM_CAL R3 ONLY  
            WriteRegister(0x59, 0x07);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
            WriteRegister(0x5A, 0x00);  //EXT_DA_GAIN_LSB_BIT
            WriteRegister(0x5B, 0xE2);  //EXT_DA_GAIN_SET = 1; (BLE_LINK) R3                                                       
            WriteRegister(0x5C, 0x00);  //TX_SEL(W),PRE_DA_PEN(W), TX_BUF_PEN(W), DA_PEN(W) R3 ONLY                                         
            WriteRegister(0x5D, 0x60);  //PM_RESETB(W), PLL_PEN(W)
            WriteRegister(0x5E, 0x20);
            WriteRegister(0x5F, 0x45);  //R3 ONLY
            WriteRegister(0x60, 0x08);

            WriteRegister(0x63, 0x00);
            WriteRegister(0x64, 0xB0);  //TX_SEL = DATA_ST

            //FOR R3 ONLY
            WriteRegister(0x65, 0x34);  //DA_PEN = DATA_ST
            WriteRegister(0x66, 0x12);
            WriteRegister(0x67, 0x40);
        }

        private void Set_PLLRESET(int delay)
        {
            WriteRegister(0x59, 0x1);
            System.Threading.Thread.Sleep(delay);
            WriteRegister(0x59, 0x7);
            System.Threading.Thread.Sleep(delay);
        }

        private void Set_BLE_LINK(uint addr, uint data)
        {
            byte[] Addrs = new byte[8];

            Addrs[0] = (byte)(addr & 0xff);
            Addrs[1] = (byte)(addr >> 8 & 0xff);
            Addrs[2] = (byte)(addr >> 16 & 0xff);
            Addrs[3] = 0x00;
            Addrs[4] = (byte)(data & 0xff);
            Addrs[5] = (byte)(data >> 8 & 0xff);
            Addrs[6] = (byte)(data >> 16 & 0xff);
            Addrs[7] = (byte)(data >> 24 & 0xff);

            I2C.WriteBytes(Addrs, 8, true);
        }

        private void Run_BLE_DTM_MODE(uint PayLoadMode)
        {
            uint PmGain = 0;

            PmGain = ReadRegister(0x58);

            if (PmGain == 0x42)
            {
                SetBleDTM_AonControlReg();
                System.Threading.Thread.Sleep(2);
                SetBleDTM_AonAnalogReg();
                System.Threading.Thread.Sleep(2);
            }
            else
            {
                SetBleDTM_AonControlReg_OTP();
                System.Threading.Thread.Sleep(2);
            }

            Set_BLE_LINK(0x058000, 0x80);   //reset
            Set_BLE_LINK(0x058000, 0x48);   //stop dtm
            Set_BLE_LINK(0x0580d8, 0x78);   //bb_clk_freq_minus_1 
            Set_BLE_LINK(0x0581b4, 0x72);   //RADIO_DATA host datain [15:8] / host dataout[15:8]
            Set_BLE_LINK(0x0581ac, 0x9043); //RADIO_ACCESS/RADIO_CNTRL 
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9400);
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9401);
            Set_BLE_LINK(0x0580d8, 0x78);
            Set_BLE_LINK(0x058190, 0x9280);
            Set_BLE_LINK(0x058198, 0x5b5b);
            Set_BLE_LINK(0x0581b0, 0x03);
            Set_BLE_LINK(0x0581b8, 0x05);

            uint pmode = ((PayLoadMode * 128) + 0);
            Set_BLE_LINK(0x058170, pmode);
            Set_BLE_LINK(0x05819c, (uint)DTM_PayLoadLength); //packet length
            Set_BLE_LINK(0x040014, (0x280 + (uint)DTM_Channel)); //Ch Sel Mode & CH 0;
            Set_PLLRESET(2);
            Set_BLE_LINK(0x058000, 0x46); //start dtm
        }

        private void Stop_BLE_DTM_MODE()
        {
            Set_BLE_LINK(0x058000, 0x48);
            WriteRegister(0x59, 0x03);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
        }

        private void Cal_PMU_All()
        {
            // AL BGR(300mV), ALLDO(810mV), MLDO(1V), RCOSC(32.768kHz)
            double[] dVal = new double[5];
            double d_alldo_mv, d_mldo_mv, d_volt_mv;
            double d_diff_mv, d_target_bgr_mv = 300, d_target_alldo_mv = 810, d_target_mlod_mv = 1000;
            double d_freq_khz, d_diff_khz;
            double d_lsl = 32.113, d_usl = 33.423, d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;
            uint alldo_val, mldo_val, bgr_val, alc = 0, mlc0 = 0, mlc1 = 0, ul = 0;
            int y_pos = 2;

            // R4
            // RCOSC
            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");                 // 0x4D
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");                // 0x4E
            // AL BGR
            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");           // 0x53
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL[5:2]");               // 0x62
            // ALLDO
            RegisterItem ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");           // 0x54
            RegisterItem ULP_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[2:0]");      // 0x61
            // MLDO
            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");           // 0x53
            RegisterItem PMU_LDO_Coarse0 = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[0]");       // 0x62
            RegisterItem PMU_LDO_Coarse1 = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1]");       // 0x63

            Parent.xlMgr.Sheet.Add(DateTime.Now.ToString("MMddHHmm_") + "R4_CAL_LDO");
            Parent.xlMgr.Cell.Write(1, 1, "O_ULP_BGR_CONT[3:0]");
            Parent.xlMgr.Cell.Write(2, 1, "BGR(mV)");
            Parent.xlMgr.Cell.Write(4, 1, "O_ULP_LDO_LV_CONT[2:0]");
            Parent.xlMgr.Cell.Write(5, 1, "O_ULP_LDO_CONT[3:0]");
            Parent.xlMgr.Cell.Write(6, 1, "ALLDO(mV)");
            Parent.xlMgr.Cell.Write(8, 1, "O_PMU_MLDO_Coarse[1:0]");
            Parent.xlMgr.Cell.Write(9, 1, "O_PMU_LDO_CONT[3:0]");
            Parent.xlMgr.Cell.Write(10, 1, "MLDO(mV)");
            Parent.xlMgr.Cell.Write(12, 1, "O_RTC_SCKF[5:0]");
            Parent.xlMgr.Cell.Write(13, 1, "RCOSC(kHz)");

            Check_Instrument();
            PowerSupply0.Write("VOLT 2.5,(@2)");
            System.Threading.Thread.Sleep(1000);

            Write_Register_AON_Fix_Value();

            // CAL BGR
            Set_TestInOut_For_BGR(true);

            ULP_BGR_CONT.Read();
            bgr_val = 0;
            d_diff_mv = 5000;
            System.Threading.Thread.Sleep(500);
            for (uint val = 0; val < 16; val++)
            {
                System.Threading.Thread.Sleep(500);
                ULP_BGR_CONT.Value = val;
                ULP_BGR_CONT.Write();

                System.Threading.Thread.Sleep(500);
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write(1, (2 + (int)val), val.ToString());
                Parent.xlMgr.Cell.Write(2, (2 + (int)val), d_volt_mv.ToString("F3"));
                if (Math.Abs(d_volt_mv - d_target_bgr_mv) < d_diff_mv)
                {
                    d_diff_mv = Math.Abs(d_volt_mv - d_target_bgr_mv);
                    bgr_val = val;
                }
            }
            ULP_BGR_CONT.Value = bgr_val;
            ULP_BGR_CONT.Write();
            System.Threading.Thread.Sleep(500);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(1, (3 + 16), bgr_val.ToString());
            Parent.xlMgr.Cell.Write(2, (3 + 16), d_volt_mv.ToString("F3"));
            Set_TestInOut_For_BGR(false);

            // Cal ALLDO
            ULP_LDO_CONT.Read();
            ULP_LDO_Coarse.Read();
            d_diff_mv = 5000.0;
            alldo_val = 0;
            for (uint j = 0; j < 4; j++)
            {
                if (j == 0)
                    ul = 1;
                else if (j == 1)
                    ul = 3;
                else if (j == 2)
                    ul = 0;
                else ul = 2;

                System.Threading.Thread.Sleep(500);
                ULP_LDO_Coarse.Value = ul;
                ULP_LDO_Coarse.Write();

                for (uint i = 0; i < 16; i++)
                {
                    System.Threading.Thread.Sleep(500);
                    Parent.xlMgr.Cell.Write(4, ((int)ul * 16 + (int)i + 2), ul.ToString());
                    ULP_LDO_CONT.Value = i;
                    ULP_LDO_CONT.Write();
                    Parent.xlMgr.Cell.Write(5, ((int)ul * 16 + (int)i + 2), (i).ToString());
                    System.Threading.Thread.Sleep(100);
                    d_alldo_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(6, ((int)ul * 16 + (int)i + 2), d_alldo_mv.ToString());
                    if (Math.Abs(d_alldo_mv - d_target_alldo_mv) < d_diff_mv)
                    {
                        d_diff_mv = Math.Abs(d_alldo_mv - d_target_alldo_mv);
                        alldo_val = i;
                        alc = ul;
                    }
                }
            }

            ULP_LDO_CONT.Value = alldo_val;
            ULP_LDO_CONT.Write();
            ULP_LDO_Coarse.Value = alc;
            ULP_LDO_Coarse.Write();
            Parent.xlMgr.Cell.Write(4, (3 * 16 + 16 + 3), alc.ToString());
            Parent.xlMgr.Cell.Write(5, (3 * 16 + 16 + 3), alldo_val.ToString());
            System.Threading.Thread.Sleep(100);
            d_alldo_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(6, (3 * 16 + 16 + 3), d_alldo_mv.ToString());

            // Cal MLDO
            PMU_LDO_Coarse0.Read();
            PMU_LDO_Coarse1.Read();
            PMU_LDO_CONT.Read();
            d_diff_mv = 5000.0;
            mldo_val = 0;
            for (uint j = 0; j < 4; j++)
            {
                System.Threading.Thread.Sleep(100);
                PMU_LDO_Coarse0.Value = j % 2;
                PMU_LDO_Coarse0.Write();
                PMU_LDO_Coarse1.Value = j / 2;
                PMU_LDO_Coarse1.Write();
                for (uint i = 0; i < 16; i++)
                {
                    System.Threading.Thread.Sleep(100);
                    Parent.xlMgr.Cell.Write(8, ((int)j * 16 + (int)i + 2), j.ToString());
                    PMU_LDO_CONT.Value = i;
                    PMU_LDO_CONT.Write();
                    Parent.xlMgr.Cell.Write(9, ((int)j * 16 + (int)i + 2), (i).ToString());
                    System.Threading.Thread.Sleep(100);
                    d_mldo_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(10, ((int)j * 16 + (int)i + 2), d_mldo_mv.ToString());
                    if (Math.Abs(d_mldo_mv - d_target_mlod_mv) < d_diff_mv)
                    {
                        d_diff_mv = Math.Abs(d_mldo_mv - d_target_mlod_mv);
                        mldo_val = i;
                        mlc0 = j % 2;
                        mlc1 = j / 2;
                    }
                }

            }
            PMU_LDO_CONT.Value = mldo_val;
            PMU_LDO_CONT.Write();
            PMU_LDO_Coarse0.Value = mlc0;
            PMU_LDO_Coarse0.Write();
            PMU_LDO_Coarse1.Value = mlc1;
            PMU_LDO_Coarse1.Write();
            Parent.xlMgr.Cell.Write(8, (3 * 16 + 16 + 3), ((mlc1 << 1) | mlc0).ToString());
            Parent.xlMgr.Cell.Write(9, (3 * 16 + 16 + 3), mldo_val.ToString());
            System.Threading.Thread.Sleep(100);
            d_mldo_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(10, (3 * 16 + 16 + 3), d_mldo_mv.ToString());

            // Cal OSC
            Set_TestInOut_For_RCOSC(true);

            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            Parent.xlMgr.Cell.Write(13, y_pos, d_freq_khz.ToString("F3"));
            Parent.xlMgr.Cell.Write(12, y_pos++, osc_val_l.ToString());

            for (int val = 4; val >= 0; val--, y_pos++)
            {
                if (d_freq_khz > d_target_khz)
                {
                    osc_val_l += (uint)(1 << val);
                }
                else
                {
                    osc_val_l -= (uint)(1 << val);
                }
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();

                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

                Parent.xlMgr.Cell.Write(13, y_pos, d_freq_khz.ToString("F3"));
                Parent.xlMgr.Cell.Write(12, y_pos, osc_val_l.ToString());
            }

            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
            {
                if (osc_val_l != 63) osc_val_l += 1;
            }
            else
            {
                if (osc_val_l != 0) osc_val_l -= 1;
            }
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            Parent.xlMgr.Cell.Write(13, y_pos, d_freq_khz.ToString("F3"));
            Parent.xlMgr.Cell.Write(12, y_pos++, osc_val_l.ToString());

            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            Parent.xlMgr.Cell.Write(13, y_pos + 1, d_freq_khz.ToString("F3"));
            Parent.xlMgr.Cell.Write(12, y_pos + 1, osc_val_l.ToString());

            Set_TestInOut_For_RCOSC(false);
        }

        private void WriteMACBLE00_V1(int num)
        {
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");                 // 0x09
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");                 // 0x08
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");                 // 0x07
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");                 // 0x06
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");                 // 0x05
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                 // 0x04

            SendCommand("mode ook\n");
            DialogResult dialog = MessageBox.Show(
                "정말로 OTP를 Write 하시겠습니까?\n" +
                "BLE PAGE 00\n" +
                "MAC : 78 46 FF FF 01 " + num.ToString("X2"),
                Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (dialog == DialogResult.OK)
            {
                B09_BA_5.Read();
                B09_BA_5.Value = 0x78;
                B09_BA_5.Write();

                B08_BA_4.Read();
                B08_BA_4.Value = 0x46;
                B08_BA_4.Write();

                B05_BA_1.Read();
                B05_BA_1.Value = 0x01;
                B05_BA_1.Write();

                B04_BA_0.Read();
                B04_BA_0.Value = (uint)num;
                B04_BA_0.Write();

                SendCommand("ook 41.00\n");
            }
            else
            {
                return;
            }
        }

        private void WriteMACBLE00_V2(int num)
        {
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");                 // 0x09
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");                 // 0x08
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");                 // 0x07
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");                 // 0x06
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");                 // 0x05
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                 // 0x04

            SendCommand("mode ook\n");
            DialogResult dialog = MessageBox.Show(
                "정말로 OTP를 Write 하시겠습니까?\n" +
                "BLE PAGE 00\n" +
                "MAC : 78 46 FF FF 02 " + num.ToString("X2"),
                Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dialog == DialogResult.OK)
            {
                B09_BA_5.Read();
                B09_BA_5.Value = 0x78;
                B09_BA_5.Write();

                B08_BA_4.Read();
                B08_BA_4.Value = 0x46;
                B08_BA_4.Write();

                B05_BA_1.Read();
                B05_BA_1.Value = 0x02;
                B05_BA_1.Write();

                B04_BA_0.Read();
                B04_BA_0.Value = (uint)num;
                B04_BA_0.Write();

                SendCommand("ook 41.00\n");
            }
            else
            {
                return;
            }
        }
        #endregion
    }

    public class SCP1501_R5 : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS_MANUAL
        {
            Write_WURF,
            TX_CH_SEL,
            TX_ON,
            TX_OFF,
            SET_BLE,
            SET_AON_OOK_TEST,
            SET_WP_ADV,
            Read_FSM,
            Set_AON_ADC,
            Read_AON_ADC,
            Test_ADC_Read,
            Test_ADC_Clear,
            Read_ADC,
            Read_ADC_Ext,
            Read_Volt_Temp,
            TestInOut_ADC_In,
            TestInOut_Temp,
            TestInOut_Volt,
            TestInOut_BGR,
            TestInOut_EXT_SENS,
            TestInOut_ADC_VREF,
            TestInOut_EDOUT,
            TestInOut_RCOSC,
            TestInOut_Disable,
            NVM_POWER_ON,
            NVM_POWER_OFF,
            OTP_PWR_ON_DIR,
            OTPw_DIRECT,
            SET_OTP_CTRL_ANA,
            SMPL_I2C,
            SET_SIM,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            Cal_PMU_SweepCal,
            Cal_PMU_FastCal,
            EXT_SENS_ADC_In,
            ADC_EXT_SENS_Forcing,
            ADC_TestIO_Forcing,
            ADC_Top_Bot_Sweep,
            ADC_VREF_Cal_EXT_Forcing,
            ADC_VREF_Calibration,
            ook_80_ADC_Read,
            I_PEN_On,
            I_PEN_Off,
            WakeUp_I2C_Multi,
            BGR_Temp_Sweep_Multi,
            PMU_Temp_Sweep_Multi,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_DTM
        {
            SET_CH,
            SET_Length,
            START_PRNBS9,
            START_11110000,
            START_10101010,
            START_PRNBS15,
            START_ALL_1,
            START_ALL_0,
            START_00001111,
            START_0101,
            STOP,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
            DTM,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Custom.I2C I2C1 { get; set; }
        private JLcLib.Custom.I2C I2C2 { get; set; }
        private JLcLib.Custom.I2C I2C3 { get; set; }
        private JLcLib.Custom.I2C I2C4 { get; set; }
        private JLcLib.Custom.I2C I2C5 { get; set; }

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        public int SlaveAddress { get; private set; } = 0x3A;
        public int DTM_PayLoadLength = 37;
        public int DTM_Channel = 0;

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
        private const UInt32 VALID_CODE = 0xCA1DCA1D;
        private const UInt16 BLE_BASE = 0x0;
        private const UInt16 CTRL_BASE = 0x264;
        private const UInt16 ANA_BASE = 0x2F4;

        private uint[] eoc_vte = { 1, 63, 31 };     // volt, temp, ext
        private uint[] vref_bct = { 64, 128, 64 };    // bottom, center, top
        #endregion Variable and declaration

        public SCP1501_R5(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            I2C1 = form.I2C1;
            I2C2 = form.I2C2;
            I2C3 = form.I2C3;
            I2C4 = form.I2C4;
            I2C5 = form.I2C5;
            Serial.ReadSettingFile(form.IniFile, "SCP1501_R5");
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

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[8];
            uint temp;
            I2C.Config.SlaveAddress = SlaveAddress;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "NVM":
                    // bist_type=3, bist_cmd=0, bist_sel=1, bist_addr=address
                    temp = (Address << 16) | 0xD;
                    Write_I2C_Data(0x00060000, temp, 10);

                    // bist_type=3, bist_cmd=1, bist_sel=1, bist_addr=address
                    temp = (Address << 16) | 0xF;
                    Write_I2C_Data(0x00060000, temp, 10);

                    // bist_wdata=data
                    Write_I2C_Data(0x00060004, Data, 10);

                    temp = (Address << 16) | 0xD;
                    Write_I2C_Data(0x00060000, temp, 10);

                    // Dummy
                    /*
                    // bist_type=3, bist_cmd=0, bist_sel=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0c);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    */
                    break;

                case "NVM Controller":
                case "NVM BIST":
                case "I2C Slave":
                case "LPCAL":
                case "ADC Controller":
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;

                default:
                    // Write Address, Data
                    temp = (0x7 << 16) | (Address << 8) | Data;     // AREG_SEL=1, AREG_CE=1, AREG_WE=1
                    Write_I2C_Data(0x00010008, temp, 10);

                    temp = (0x4 << 16) | (Address << 8) | Data;     // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    Write_I2C_Data(0x00010008, temp, 10);

                    // AREG_SEL Disable
                    temp = (0x0 << 16) | (0 << 8) | 0;              // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    Write_I2C_Data(0x00010008, temp, 10);
                    break;
            }
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            uint result = 0xffffffff;
            uint temp;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "NVM":
                    // bist_type=4, bist_cmd=0, bist_sel=1, bist_addr=0
                    temp = 0x00000011;
                    Write_I2C_Data(0x00060000, temp, 10);

                    // bist_type=4, bist_cmd=1, bist_sel=1, bist_addr=address
                    temp = (Address << 16) | 0x0013;
                    Write_I2C_Data(0x00060000, temp, 10);

                    // bist_type=4, bist_cmd=0, bist_sel=1, bist_addr=address
                    temp = (Address << 16) | 0x11;
                    Write_I2C_Data(0x00060000, temp, 10);

                    RcvData = Read_I2C_Data(0x00060018, 10);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);
                    /*
                                        // bist_type=4, bist_cmd=0, bist_sel=0, bist_addr=0
                                        SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                                        SendData[1] = (byte)(0x00);
                                        SendData[2] = (byte)(0x06);
                                        SendData[3] = (byte)(0x00);
                                        SendData[4] = (byte)(0x10);
                                        SendData[5] = (byte)(0x00);
                                        SendData[6] = (byte)(0x00);
                                        SendData[7] = (byte)(0x00);
                                        I2C.WriteBytes(SendData, SendData.Length, true);
                    */
                    break;

                case "NVM Controller":
                case "NVM BIST":
                case "I2C Slave":
                case "LPCAL":
                case "ADC Controller":
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);
                    break;

                default:
                    // Write Address
                    temp = (0x4 << 16) | (Address << 8);
                    Write_I2C_Data(0x00010008, temp, 10);   // AREG_SEL=1, AREG_CE=0, AREG_WE=0

                    temp = (0x5 << 16) | (Address << 8);
                    Write_I2C_Data(0x00010008, temp, 10);   // AREG_SEL=1, AREG_CE=1, AREG_WE=0

                    // Read Data
                    RcvData = Read_I2C_Data(0x00010010, 10);
                    result = RcvData[3];

                    temp = (0x4 << 16) | (Address << 8);
                    Write_I2C_Data(0x00010008, temp, 10);   // AREG_SEL=1, AREG_CE=0, AREG_WE=0

                    // AREG_SEL Disable
                    Address = 0;
                    temp = (0x0 << 16) | (Address << 8);
                    Write_I2C_Data(0x00010008, temp, 10);   // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    break;
            }

            return result;
        }

        private uint ReadRegister_Multi(uint Address, JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            uint result = 0xffffffff;

            switch (Parent.xlMgr.Sheet.Name)
            {
                default:
                    // Write Address
                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);             // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);   // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);

                    SendData[6] = (byte)(0x05);             // AREG_SEL=1, AREG_CE=1, AREG_WE=0
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);

                    // Read Data
                    SendData[0] = (byte)(0x10);             // ADDR = 0x00010010
                    I2C_Sel.WriteBytes(SendData, 4, false);
                    RcvData = I2C_Sel.ReadBytes(RcvData.Length);
                    result = RcvData[3];

                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[5] = (byte)(0x00);             // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);             // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);
                    break;
            }
            return result;
        }

        private void WriteRegister_Multi(uint Address, uint Data, JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[8];
            I2C.Config.SlaveAddress = SlaveAddress;

            switch (Parent.xlMgr.Sheet.Name)
            {
                default:
                    // Write Address, Data
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x07);            // AREG_SEL=1, AREG_CE=1, AREG_WE=1
                    SendData[7] = (byte)(0x00);
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);

                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);            // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);            // AREG_WDATA[7:0]
                    SendData[5] = (byte)(0x00);            // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);            // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C_Sel.WriteBytes(SendData, SendData.Length, true);
                    break;
            }
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
            Parent.ChipCtrlButtons[4].Text = "GH0_H";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;
            Parent.ChipCtrlButtons[5].Text = "WakeUp";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += WakeUp_I2C;
            Parent.ChipCtrlButtons[6].Text = "MD_OOK";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Set_Gecko_OOK_Mode;
            Parent.ChipCtrlButtons[8].Text = "Manual";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;
            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_Auto_Test_Items;
            Parent.ChipCtrlButtons[10].Text = "DTM";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Change_To_DTM_Test_Items;
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

        private void Change_To_DTM_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.DTM;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_DTM.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_DTM)i).ToString());
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

        private void WakeUp_I2C(object sender, EventArgs e)
        {
            WakeUp_I2C();
        }

        private void Set_Gecko_OOK_Mode(object sender, EventArgs e)
        {
            SendCommand("mode ook");
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
                        case TEST_ITEMS_MANUAL.Write_WURF:
                            if (Arg == "")
                            {
                                Log.WriteLine("WURF에 write할 값을 16진수로 적어주세요.");
                            }
                            else
                            {
                                Log.WriteLine("OOK " + Arg);
                                Write_WURF_AON(Arg);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_CH_SEL:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                Log.WriteLine("Set CH" + Arg);
                                Write_Register_Fractional_Calc_Ch((uint)iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_ON:
                            Log.WriteLine("TX ON");
                            Write_Register_Tx_Tone_Send(true);
                            break;
                        case TEST_ITEMS_MANUAL.TX_OFF:
                            Log.WriteLine("TX OFF");
                            Write_Register_Tx_Tone_Send(false);
                            break;
                        case TEST_ITEMS_MANUAL.SET_BLE:
                            if ((Arg == "") || (iVal < 0) || (iVal > 65535))
                            {
                                Log.WriteLine("WP값을 10진수로 적어주세요. (Range 0~65535)");
                            }
                            else
                            {
                                Log.WriteLine("Write AON for Advertising (WP = )" + Arg);
                                Write_Register_Send_Advertising(iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.SET_AON_OOK_TEST:
                            Set_AON_OOK_Test();
                            break;
                        case TEST_ITEMS_MANUAL.Read_FSM:
                            Read_Register_FSM();
                            break;
                        case TEST_ITEMS_MANUAL.Test_ADC_Read:
                            Result = Test_Adc_Read();
                            //Disable_ADC();
                            Log.WriteLine("TEST Read Volt: " + ((Result >> 8) & 0xff).ToString() + "\tTemp: " + (Result & 0xff).ToString() + "\tExt. Temp: " + ((Result >> 16) & 0xff).ToString());
                            break;
                        case TEST_ITEMS_MANUAL.Test_ADC_Clear:
                            Disable_ADC();
                            break;
                        case TEST_ITEMS_MANUAL.Read_ADC:
                            Clear_ADC_Ext();
                            Result = Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct);
                            Disable_ADC();
                            Log.WriteLine("Volt: " + ((Result >> 8) & 0xff).ToString() + "\tTemp: " + (Result & 0xff).ToString());
                            break;
                        case TEST_ITEMS_MANUAL.Read_ADC_Ext:
                            Result = Read_ADC_Ext_VT(true);      // External temp
                            Disable_ADC();
                            Log.WriteLine("Volt: " + ((Result >> 8) & 0xff).ToString() + "\tTemp: " + (Result & 0xff).ToString() + "\tExt. Temp: " + ((Result >> 16) & 0xff).ToString());
                            break;
                        case TEST_ITEMS_MANUAL.Read_Volt_Temp:
                            Calculation_VBAT_Voltage_And_Temperature();
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_ADC_In:
                            TestInOut_ADC_In(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_Temp:
                            TestInOut_Tempsensor(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_Volt:
                            TestInOut_Voltagesensor(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_BGR:
                            Set_TestInOut_For_BGR(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_EXT_SENS:
                            TestInOut_EXT_SENS(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_ADC_VREF:
                            TestInOut_ADC_VREF(iVal);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_EDOUT:
                            Set_TestInOut_For_EDOUT(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_RCOSC:
                            Set_TestInOut_For_RCOSC(true);
                            break;
                        case TEST_ITEMS_MANUAL.TestInOut_Disable:
                            TestInOut_Voltagesensor(false);
                            TestInOut_Tempsensor(false);
                            TestInOut_ADC_In(false);
                            Set_TestInOut_For_BGR(false);
                            Set_TestInOut_For_EDOUT(false);
                            Set_TestInOut_For_RCOSC(false);
                            TestInOut_EXT_SENS(false);
                            TestInOut_ADC_VREF(0);
                            Disable_ADC();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_ON:
                            Enable_NVM_BIST();
                            Power_On_NVM();
                            //Disable_NVM_BIST();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_OFF:
                            Enable_NVM_BIST();
                            Power_Off_NVM();
                            Disable_NVM_BIST();
                            break;
                        case TEST_ITEMS_MANUAL.SMPL_I2C:
                            if (Arg == "")
                            {
                                Log.WriteLine("Simple I2C로 write할 값을 16진수로 적어주세요.(30,60,61,70,80,90,A0,B0,C0");
                            }
                            else
                            {
                                Log.WriteLine("Simple I2C Cmd" + Arg);
                                Simple_I2C_Cmd(Arg);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.SET_SIM:
                            Set_Sim_Val();
                            break;
                        case TEST_ITEMS_MANUAL.OTP_PWR_ON_DIR:
                            Otp_Pwr_On_Direct();
                            break;
                        case TEST_ITEMS_MANUAL.OTPw_DIRECT:
                            Otp_Wr_Direct();
                            break;
                        case TEST_ITEMS_MANUAL.SET_OTP_CTRL_ANA:
                            Set_Otp_ctrl_ana();
                            break;
                        case TEST_ITEMS_MANUAL.SET_WP_ADV:
                            Set_Wp_Adv();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((TEST_ITEMS_AUTO)TestItemIndex)
                    {
                        case TEST_ITEMS_AUTO.Cal_PMU_SweepCal:
                            Cal_PMU_SweepCal(iVal);
                            break;
                        case TEST_ITEMS_AUTO.Cal_PMU_FastCal:
                            Cal_PMU_FastCal(iVal);
                            break;
                        case TEST_ITEMS_AUTO.ADC_EXT_SENS_Forcing:
                            ADC_EXT_SENS_Forcing();
                            break;
                        case TEST_ITEMS_AUTO.ADC_TestIO_Forcing:
                            TestInOut_EXT_ADC_In();
                            break;
                        case TEST_ITEMS_AUTO.EXT_SENS_ADC_In:
                            EXT_SENS_ADC_In();
                            break;
                        case TEST_ITEMS_AUTO.ADC_Top_Bot_Sweep:
                            ADC_Top_Bot_Sweep();
                            break;
                        case TEST_ITEMS_AUTO.ADC_VREF_Cal_EXT_Forcing:
                            ADC_VREF_Cal_EXT_Forcing();
                            break;
                        case TEST_ITEMS_AUTO.ADC_VREF_Calibration:
                            ADC_VREF_Calibration();
                            break;
                        case TEST_ITEMS_AUTO.ook_80_ADC_Read:
                            ook_80_ADC_Read(true);
                            break;
                        case TEST_ITEMS_AUTO.I_PEN_On:
                            I_PEN(true);
                            break;
                        case TEST_ITEMS_AUTO.I_PEN_Off:
                            I_PEN(false);
                            break;
                        case TEST_ITEMS_AUTO.WakeUp_I2C_Multi:
                            WakeUp_I2C_Multi_Test(iVal);
                            break;
                        case TEST_ITEMS_AUTO.BGR_Temp_Sweep_Multi:
                            BGR_Temp_Sweep_Multi();
                            break;
                        case TEST_ITEMS_AUTO.PMU_Temp_Sweep_Multi:
                            PMU_Temp_Sweep_Multi();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.DTM:
                    switch ((TEST_ITEMS_DTM)TestItemIndex)
                    {
                        case TEST_ITEMS_DTM.SET_CH:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                DTM_Channel = iVal;
                                Log.WriteLine("Channel : " + DTM_Channel.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.SET_Length:
                            if ((Arg == "") || ((iVal != 37) && (iVal != 255)))
                            {
                                Log.WriteLine("Length를 10진수로 적어주세요. (Range 37 또는 255)");
                            }
                            else
                            {
                                DTM_PayLoadLength = iVal;
                                Log.WriteLine("Length : " + DTM_PayLoadLength.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.START_PRNBS9:
                        case TEST_ITEMS_DTM.START_11110000:
                        case TEST_ITEMS_DTM.START_10101010:
                        case TEST_ITEMS_DTM.START_PRNBS15:
                        case TEST_ITEMS_DTM.START_ALL_1:
                        case TEST_ITEMS_DTM.START_ALL_0:
                        case TEST_ITEMS_DTM.START_00001111:
                        case TEST_ITEMS_DTM.START_0101:
                            Log.WriteLine("Start DTM\nChannel : " + DTM_Channel.ToString() + "\tLength : " + DTM_PayLoadLength.ToString());
                            Run_BLE_DTM_MODE((uint)(TestItemIndex - TEST_ITEMS_DTM.START_PRNBS9));
                            break;
                        case TEST_ITEMS_DTM.STOP:
                            Log.WriteLine("Stop DTM");
                            Stop_BLE_DTM_MODE();
                            break;
                    }
                    break;
                default:
                    break;
            }
        }

        #region Function for Chip Test
        private void WakeUp_I2C()
        {
            byte[] SendData = new byte[2];

            SendData[0] = 0xAA;
            SendData[1] = 0xBB;

            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void WakeUp_I2C_Multi(JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[2];

            SendData[0] = 0xAA;
            SendData[1] = 0xBB;

            I2C_Sel.WriteBytes(SendData, SendData.Length, true);
        }

        private void WakeUp_I2C_Multi_Test(int BoardNUM)
        {
            JLcLib.Custom.I2C[] i2c = { I2C, I2C1, I2C2, I2C3 };

            for (int i = 0; i < i2c.Length; i++)
            {
                i2c[i].GPIOs[4].Direction = GPIO_Direction.Output;
                i2c[i].GPIOs[4].State = GPIO_State.High;
            }

            i2c[BoardNUM].GPIOs[4].State = GPIO_State.Low;

            WakeUp_I2C_Multi(i2c[BoardNUM]);

            Set_TestInOut_For_BGR_Multi(true, i2c[BoardNUM]);
        }

        private void Write_WURF_CMD_I2C(byte data)
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=1, WURF_WRITE=1
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = data;
            SendData[5] = 0x00;
            SendData[6] = 0x06;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // WURF_WRITE=0
            SendData[6] = 0x02;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_CMD_I2C_Multi(byte data, JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=1, WURF_WRITE=1
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = data;
            SendData[5] = 0x00;
            SendData[6] = 0x06;
            SendData[7] = 0x00;
            I2C_Sel.WriteBytes(SendData, SendData.Length, true);

            // WURF_WRITE=0
            SendData[6] = 0x02;
            I2C_Sel.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_WURF_Sel()
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=0
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_WURF_Sel_Multi(JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=0
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C_Sel.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_AON(string Arg)
        {
            RegisterItem w_WURF_END = Parent.RegMgr.GetRegisterItem("w_WURF_END");         // 0x38
            byte data;

            switch (Arg)
            {
                case "30":
                    data = 0x30;
                    break;
                case "60":
                    data = 0x60;
                    break;
                case "61":
                    data = 0x61;
                    break;
                case "70":
                    data = 0x70;
                    break;
                case "80":
                    data = 0x80;
                    break;
                case "b0":
                case "B0":
                    data = 0xb0;
                    break;
                case "c0":
                case "C0":
                    data = 0xb0;
                    break;
                default:
                    Log.WriteLine("지원하지 않는 CMD 입니다.");
                    return;
                    break;
            }

            w_WURF_END.Read();
            w_WURF_END.Value = 0;
            w_WURF_END.Write();

            Write_WURF_CMD_I2C(data);

            w_WURF_END.Value = 1;
            w_WURF_END.Write();

            Disable_WURF_Sel();

            w_WURF_END.Value = 0;
            w_WURF_END.Write();
        }

        private void Write_WURF_AON_Multi(string Arg, JLcLib.Custom.I2C I2C_Sel)
        {
            RegisterItem w_WURF_END = Parent.RegMgr.GetRegisterItem("w_WURF_END");         // 0x26
            byte data;
            uint rData;

            switch (Arg)
            {
                case "30":
                    data = 0x30;
                    break;
                case "60":
                    data = 0x60;
                    break;
                case "61":
                    data = 0x61;
                    break;
                case "70":
                    data = 0x70;
                    break;
                case "80":
                    data = 0x80;
                    break;
                case "b0":
                case "B0":
                    data = 0xb0;
                    break;
                case "c0":
                case "C0":
                    data = 0xb0;
                    break;
                default:
                    Log.WriteLine("지원하지 않는 CMD 입니다.");
                    return;
                    break;
            }

            //w_WURF_END.Read();
            //w_WURF_END.Value = 0;
            //w_WURF_END.Write();
            rData = ReadRegister_Multi(0x26, I2C_Sel);
            WriteRegister_Multi(0x26, rData & 0x7F, I2C_Sel);

            Write_WURF_CMD_I2C_Multi(data, I2C_Sel);

            //w_WURF_END.Value = 1;
            //w_WURF_END.Write();
            rData = ReadRegister_Multi(0x26, I2C_Sel);
            WriteRegister_Multi(0x26, rData | 0x80, I2C_Sel);

            Disable_WURF_Sel_Multi(I2C_Sel);

            //w_WURF_END.Value = 0;
            //w_WURF_END.Write();
            rData = ReadRegister_Multi(0x26, I2C_Sel);
            WriteRegister_Multi(0x26, rData & 0x7F, I2C_Sel);
        }

        private void Write_Register_Fractional_Calc_Ch(uint ch)
        {
            uint vco;

            RegisterItem SPI_DSM_F_H = Parent.RegMgr.GetRegisterItem("O_SPI_DSM_F[22:16]"); // 0x3E
            RegisterItem SPI_PS_P = Parent.RegMgr.GetRegisterItem("O_SPI_PS_P[4:0]");       // 0x3F
            RegisterItem SPI_PS_S = Parent.RegMgr.GetRegisterItem("O_SPI_PS_S[2:0]");       // 0x3F                

            vco = 2402 + 2 * ch;
            SPI_PS_P.Value = vco >> 7;
            SPI_PS_S.Value = (vco >> 5) - (SPI_PS_P.Value << 2);
            SPI_DSM_F_H.Value = (vco << 1) - (SPI_PS_P.Value << 8) - (SPI_PS_S.Value << 6);

            SPI_PS_P.Write();
            SPI_DSM_F_H.Write();
        }

        private void Write_Register_Tx_Tone_Send(bool on)
        {
            RegisterItem EXT_CH_MODE = Parent.RegMgr.GetRegisterItem("O_EXT_CH_MODE");             // 0x3E
            RegisterItem TX_SEL = Parent.RegMgr.GetRegisterItem("O_TX_SEL");                       // 0x56
            RegisterItem TX_BUF_PEN = Parent.RegMgr.GetRegisterItem("O_TX_BUF_PEN");               // 0x57
            RegisterItem PRE_DA_PEN = Parent.RegMgr.GetRegisterItem("O_PRE_DA_PEN");               // 0x57
            RegisterItem DA_PEN = Parent.RegMgr.GetRegisterItem("O_DA_PEN");                       // 0x57
            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("O_PLL_PEN");                     // 0x59

            RegisterItem EXT_DA_GAIN_SEL = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_SEL");     // 0x5B
            RegisterItem w_TX_BUF_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_TX_BUF_PEN_MODE");   // 0x5C
            RegisterItem w_PRE_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PRE_DA_PEN_MODE");   // 0x5C
            RegisterItem w_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_DA_PEN_MODE");           // 0x5C
            RegisterItem w_TX_SEL_MODE = Parent.RegMgr.GetRegisterItem("w_TX_SEL_MODE");           // 0x5C
            RegisterItem w_PLL_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PLL_PEN_MODE");         // 0x5D
            RegisterItem XTAL_PLL_CLK_EN_MD = Parent.RegMgr.GetRegisterItem("w_XTAL_PLL_CLK_EN_MD");   // 0x5F

            if (on == true)
            {
                // Tx Active
                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 1;
                EXT_CH_MODE.Write();

                TX_SEL.Read();
                TX_SEL.Value = 1;
                TX_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 1;
                PRE_DA_PEN.Value = 1;
                DA_PEN.Value = 1;
                TX_BUF_PEN.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 0;
                EXT_DA_GAIN_SEL.Write();

                // Dynamic -> Static
                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 1;
                w_PLL_PEN_MODE.Write();

                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 1;
                w_PRE_DA_PEN_MODE.Value = 1;
                w_DA_PEN_MODE.Value = 1;
                w_TX_SEL_MODE.Value = 1;
                w_TX_BUF_PEN_MODE.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 1;
                XTAL_PLL_CLK_EN_MD.Write();

                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                PLL_PEN.Value = 1;
                PLL_PEN.Write();
            }
            else
            {
                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 0;
                XTAL_PLL_CLK_EN_MD.Write();

                // Static -> Dynamic
                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 0;
                w_PRE_DA_PEN_MODE.Value = 0;
                w_DA_PEN_MODE.Value = 0;
                w_TX_SEL_MODE.Value = 0;
                w_TX_BUF_PEN_MODE.Write();

                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 0;
                w_PLL_PEN_MODE.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 1;
                EXT_DA_GAIN_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 0;
                PRE_DA_PEN.Value = 0;
                DA_PEN.Value = 0;
                TX_BUF_PEN.Write();

                TX_SEL.Read();
                TX_SEL.Value = 0;
                TX_SEL.Write();

                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 0;
                EXT_CH_MODE.Write();
            }
        }

        private void Write_Register_Send_Advertising(int iVal)
        {
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");     // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");     // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");     // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");     // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");     // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");     // 0x1B
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");     // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");     // 0x1F

            B22_WP_0.Value = (uint)(iVal & 0xff);
            B22_WP_0.Write();

            B23_WP_1.Value = (uint)((iVal >> 8) & 0xff);
            B23_WP_1.Write();

            B24_AI_0.Value = 210;
            B24_AI_0.Write();

            B25_AI_1.Value = 29;
            B25_AI_1.Write();

            B26_AI_2.Value = 0;
            B26_AI_2.Write();

            B27_AI_3.Value = 0;
            B27_AI_3.Write();

            B30_AD_0.Value = 3;
            B30_AD_0.Write();

            B31_AD_1.Value = 0;
            B31_AD_1.Write();
        }

        private uint Run_Read_FSM_Status()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[1] & 0xf) << 8) | RcvData[0]);
        }

        private uint CRC_FLAG_READ()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[2] & 0x3) << 16) | (RcvData[1] << 8) | RcvData[0]);
        }

        private void Read_Register_FSM()
        {
            uint u_fsm;
            uint crc_flag;
            u_fsm = Run_Read_FSM_Status();
            crc_flag = CRC_FLAG_READ();
            Log.WriteLine("FSM_Status : " + (u_fsm & 0x000f).ToString());
            Log.WriteLine("FSM_State : " + ((u_fsm & 0x0ff0) >> 4).ToString());
            Log.WriteLine("WUR_CRC_ERROR : " + ((crc_flag >> 17) & 0xff).ToString());
            Log.WriteLine("WUR_CRC_DEC : " + (crc_flag & 0xffff).ToString());
        }

        private void Set_AON_ADC()
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            byte data = 0x80;
            uint result = 0xffffffff;
            uint Address1 = 0x1A;
            uint Address2 = 0x1B;

            // set ADC
            WriteRegister(0x2E, 175);
            WriteRegister(0x2F, 47);
            WriteRegister(0x30, 255);
            WriteRegister(0x31, 67);
            WriteRegister(0x32, 191);
            WriteRegister(0x33, 215);
            WriteRegister(0x34, 57);
            WriteRegister(0x35, 176);
            WriteRegister(0x36, 127);
            WriteRegister(0x37, 127);
            WriteRegister(0x38, 58);
            WriteRegister(0x39, 194);
            WriteRegister(0x3A, 122);
            WriteRegister(0x3B, 11);
            WriteRegister(0x3C, (1 << 5) + (1 << 4) + (1 << 2) + (1 << 0));
            WriteRegister(0x3D, 128);

            Read_AON_ADC();
        }

        private void Read_AON_ADC()
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            byte data = 0x80;
            uint result = 0xffffffff;
            uint Address1 = 0x1A;
            uint Address2 = 0x1B;

            WriteRegister(0x26, 0x20);  // WURF_SEL = 0
            Write_WURF_CMD_I2C(0x80);   // Run ADC
            WriteRegister(0x26, 0xA0);  // WURF_SEL = 1
            Disable_WURF_Sel();
            WriteRegister(0x26, 0x20);  // WURF_SEL = 0

            // Read ADC
            // Write Address1
            SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x01);
            SendData[3] = (byte)(0x00);
            SendData[4] = (byte)(0x00);             // AREG_WDATA[7:0]
            SendData[5] = (byte)(Address1 & 0xff);   // AREG_ADDR[7:0]
            SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
            SendData[7] = (byte)(0x00);
            I2C.WriteBytes(SendData, SendData.Length, true);

            SendData[6] = (byte)(0x05);             // AREG_SEL=1, AREG_CE=1, AREG_WE=0
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Read Data
            SendData[0] = (byte)(0x10);             // ADDR = 0x00010010
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            result = RcvData[3];
            Log.WriteLine("0x1A: " + (result & 0xff).ToString());

            SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
            SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
            I2C.WriteBytes(SendData, SendData.Length, true);

            // AREG_SEL Disable
            SendData[5] = (byte)(0x00);             // AREG_ADDR[7:0]
            SendData[6] = (byte)(0x00);             // AREG_SEL=0, AREG_CE=0, AREG_WE=0
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Write Address2
            SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x01);
            SendData[3] = (byte)(0x00);
            SendData[4] = (byte)(0x00);             // AREG_WDATA[7:0]
            SendData[5] = (byte)(Address2 & 0xff);   // AREG_ADDR[7:0]
            SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
            SendData[7] = (byte)(0x00);
            I2C.WriteBytes(SendData, SendData.Length, true);

            SendData[6] = (byte)(0x05);             // AREG_SEL=1, AREG_CE=1, AREG_WE=0
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Read Data
            SendData[0] = (byte)(0x10);             // ADDR = 0x00010010
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            result = RcvData[3];

            SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
            SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
            I2C.WriteBytes(SendData, SendData.Length, true);

            // AREG_SEL Disable
            SendData[5] = (byte)(0x00);             // AREG_ADDR[7:0]
            SendData[6] = (byte)(0x00);             // AREG_SEL=0, AREG_CE=0, AREG_WE=0
            I2C.WriteBytes(SendData, SendData.Length, true);

            Log.WriteLine("0x1B: " + (result & 0xff).ToString());
        }

        private uint Test_Adc_Read()
        {
            byte[] RcvData = new byte[4];
            RcvData = Read_I2C_Data(0x0003000C, 1);
            return BitConverter.ToUInt32(RcvData, 0);
        }

        private uint Read_ADC_Result(bool temp_first, ref uint[] eoc_skip_vte, uint sample_cnt, ref uint[] vref_bct)
        {
            byte[] RcvData = new byte[4];
            byte[] temp = new byte[4];
            UInt32 addr = 0;
            UInt32 wdata = 0, rdata = 0;
            UInt32 adc_switch = 0;

            if (temp_first)
            {
                adc_switch = 0;     // temp.-> battery
            }
            else
            {
                adc_switch = 1;     // battery -> temp.
            }
            addr = 0x00030004;
            RcvData = Read_I2C_Data(addr, 1);
            //rdata = BitConverter.ToUInt32(RcvData, 0);
            //rdata &= ~(1U << 30);
            //wdata = rdata + (adc_switch << 30);
            wdata = (adc_switch << 30) + (eoc_skip_vte[2] << 24) + (vref_bct[2] << 16) + (vref_bct[1] << 8) + (vref_bct[0]);
            Write_I2C_Data(addr, wdata, 1);

            UInt32 pen = 1;
            addr = 0x00030000;
            RcvData = Read_I2C_Data(addr, 5);
            rdata = BitConverter.ToUInt32(RcvData, 0);
            //Array.Reverse(RcvData);
            //Log.WriteLine("0x3_0004: " + BitConverter.ToString(RcvData).Replace("-", ""));
            rdata &= ~((1U << 31) + (0x1F << 25) + (0x3F << 14) + (0x3F << 8));
            wdata = rdata + (pen << 31) + (sample_cnt << 25) + (eoc_skip_vte[0] << 14) + (eoc_skip_vte[1] << 8);
            //temp = BitConverter.GetBytes(wdata);
            //Array.Reverse(temp);
            //Log.WriteLine("0x3_0000(w): " + BitConverter.ToString(temp).Replace("-", ""));
            Write_I2C_Data(addr, wdata, 1);

            addr = 0x00030000;
            RcvData = Read_I2C_Data(addr, 5);
            rdata = BitConverter.ToUInt32(RcvData, 0);
            rdata &= ~((1U << 31) + (0x1F << 25) + (0x3F << 14) + (0x3F << 8));
            pen = 0;
            wdata = rdata + (pen << 31) + (sample_cnt << 25) + (eoc_skip_vte[0] << 14) + (eoc_skip_vte[1] << 8);
            Write_I2C_Data(addr, wdata, 1);

            // Read ADC_D_T[23:16], ADC_D_B[15:8],ADC_D_T[7:0]
            RcvData = Read_I2C_Data(0x0003000C, 1);

            return BitConverter.ToUInt32(RcvData, 0);
        }

        private void Clear_ADC_Ext()
        {
            uint atemp = 0;
            // clear AON volt, temp EN
            atemp = (0x7 << 16) | (0x2A << 8) | 0;     // AREG_SEL=1, AREG_CE=1, AREG_WE=1
            Write_I2C_Data(0x00010008, atemp, 10);

            atemp = (0x4 << 16) | (0x2A << 8) | 0;     // AREG_SEL=1, AREG_CE=0, AREG_WE=0
            Write_I2C_Data(0x00010008, atemp, 10);

            atemp = (0x0 << 16) | (0 << 8) | 0;              // AREG_SEL=0, AREG_CE=0, AREG_WE=0
            Write_I2C_Data(0x00010008, atemp, 10);
        }

        private uint Read_ADC_Ext_VT(bool sel_temp)
        {
            UInt32 atemp, data;
            UInt32 ret;

            if (sel_temp)
            {
                data = (1U << 6);    // temperature
            }
            else
            {
                data = (1U << 7);
            }
            data += eoc_vte[2];

            // Set volt_en or temp_en
            atemp = (0x7 << 16) | (0x2A << 8) | data;     // AREG_SEL=1, AREG_CE=1, AREG_WE=1
            Write_I2C_Data(0x00010008, atemp, 10);

            atemp = (0x4 << 16) | (0x2A << 8) | data;     // AREG_SEL=1, AREG_CE=0, AREG_WE=0
            Write_I2C_Data(0x00010008, atemp, 10);

            atemp = (0x0 << 16) | (0 << 8) | 0;              // AREG_SEL=0, AREG_CE=0, AREG_WE=0
            Write_I2C_Data(0x00010008, atemp, 10);

            // Read ADC
            ret = Read_ADC_Result(sel_temp, ref eoc_vte, 5, ref vref_bct);

            return ret;
        }

        private void Disable_ADC()
        {
            byte[] RcvData = new byte[4];
            UInt32 addr, wdata = 0;

            addr = 0x00030000;
            wdata = 0;
            Write_I2C_Data(addr, wdata, 10);
        }

        private double Calculation_Temperature(uint otp, uint adc)
        {
            double[] lowtempcomp_ADC = { 57.11, 57.23, 57.34, 57.45, 57.56, 57.66, 57.75, 57.84 };
            double temperature;
            uint OTP23p5, OTP85TO23p5;

            OTP23p5 = (otp & 0x1f) + 104;
            OTP85TO23p5 = (otp >> 5) + 35;

            if (adc >= OTP23p5)
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) + 23.5;
            }
            else
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) * ((23.5 - (-40)) / lowtempcomp_ADC[otp >> 5]) + 23.5;
            }

            return temperature;
        }

        private double Calculation_VBAT_Voltage(double temperature, uint adc)
        {
            double voltage;

            if (temperature >= 30)
            {
                voltage = (adc - (11.475 - (-3 / (85 - 30)) * (temperature - 30))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }
            else
            {
                voltage = (adc - (11.475 - (2 / (30 - (-40))) * (30 - temperature))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }

            return voltage;
        }

        private void Calculation_VBAT_Voltage_And_Temperature()
        {
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");       // 0x14
            uint adc;
            double temp, volt;

            B20_DC_0.Read();
            adc = Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct);
            Disable_ADC();
            temp = Calculation_Temperature(B20_DC_0.Value, adc & 0xff);
            volt = Calculation_VBAT_Voltage(temp, adc >> 8);
            Log.WriteLine("Volt : " + volt.ToString("F2") + "(" + (adc >> 8).ToString() + ")\tTemp : " + temp.ToString("F3") + "(" + (adc & 0xff).ToString() + ")");
        }

        private void Set_TestInOut_For_VTEMP(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");       // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL");     // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");            // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");            // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BUF_MUX_SEL.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_VS(bool on)
        {
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");        // 0x61
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");   // 0x5F
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL"); // 0x5F

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();
            }
            else
            {

            }
        }

        private void Set_TestInOut_For_BGR(bool on)
        {
            RegisterItem TEST_BGR_MUX_SEL = Parent.RegMgr.GetRegisterItem("O_TEST_BGR_MUX_SEL");    // 0x4A
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");    // 0x5F
            RegisterItem TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");         // 0x61

            if (on == true)
            {
                TEST_BGR_MUX_SEL.Read();
                TEST_BGR_MUX_SEL.Value = 1;
                TEST_BGR_MUX_SEL.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON.Read();
                TEST_CON.Value = 2;
                TEST_CON.Write();

            }
            else
            {
                TEST_CON.Read();
                TEST_CON.Value = 0;
                TEST_CON.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_BGR_Multi(bool on, JLcLib.Custom.I2C I2C_Sel)
        {
            uint rData;

            RegisterItem O_TEST_BGR_MUX_SEL = Parent.RegMgr.GetRegisterItem("O_TEST_BGR_MUX_SEL");      // 0x4A
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");            // 0x5F
            RegisterItem TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");                   // 0x61

            if (on == true)
            {
                rData = ReadRegister_Multi(0x4A, I2C_Sel);
                WriteRegister_Multi(0x4A, rData | 0x20, I2C_Sel);
                //O_TEST_BGR_MUX_SEL.Read();
                //O_TEST_BGR_MUX_SEL.Value = 1;
                //O_TEST_BGR_MUX_SEL.Write();

                rData = ReadRegister_Multi(0x5F, I2C_Sel);
                WriteRegister_Multi(0x5F, rData | 0x04, I2C_Sel);
                //TEST_BGR_BUF_EN.Read();
                //TEST_BGR_BUF_EN.Value = 1;
                //TEST_BGR_BUF_EN.Write();

                rData = ReadRegister_Multi(0x61, I2C_Sel);
                WriteRegister_Multi(0x61, rData | 0x02, I2C_Sel);
                //TEST_CON.Read();
                //TEST_CON.Value = 2;
                //TEST_CON.Write();
            }
            else
            {
                rData = ReadRegister_Multi(0x4A, I2C_Sel);
                WriteRegister_Multi(0x4A, rData & 0xDF, I2C_Sel);
                //O_TEST_BGR_MUX_SEL.Read();
                //O_TEST_BGR_MUX_SEL.Value = 0;
                //O_TEST_BGR_MUX_SEL.Write();

                rData = ReadRegister_Multi(0x5F, I2C_Sel);
                WriteRegister_Multi(0x5F, rData & 0xFB, I2C_Sel);
                //TEST_BGR_BUF_EN.Read();
                //TEST_BGR_BUF_EN.Value = 0;
                //TEST_BGR_BUF_EN.Write();

                rData = ReadRegister_Multi(0x61, I2C_Sel);
                WriteRegister_Multi(0x61, rData & 0xFD, I2C_Sel);
                //TEST_CON.Read();
                //TEST_CON.Value = 0;
                //TEST_CON.Write();
            }
        }

        private void Set_TestInOut_For_EDOUT(bool on)
        {
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[8]");   // 0x5A
            RegisterItem O_RX_DATAT = Parent.RegMgr.GetRegisterItem("O_RX_DATAT");      // 0x5A

            if (on == true)
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 1;
                O_RX_DATAT.Value = 1;
                ITEST_CONT.Write();
            }
            else
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 0;
                O_RX_DATAT.Value = 0;
                ITEST_CONT.Write();
            }
        }

        private void Set_TestInOut_For_RCOSC(bool on)
        {
            RegisterItem TEST_EN_32K = Parent.RegMgr.GetRegisterItem("O_TEST_EN_32K");     // 0x5D
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");    // 0x61

            if (on == true)
            {
                TEST_EN_32K.Read();
                TEST_EN_32K.Value = 1;
                TEST_CON_H.Value = 8;
                TEST_EN_32K.Write();
            }
            else
            {
                TEST_EN_32K.Value = 0;
                TEST_CON_H.Value = 0;
                TEST_EN_32K.Write();
            }
        }

        private void Set_TestInOut_For_RCOSC_Multi(bool on, JLcLib.Custom.I2C I2C_Sel)
        {
            uint rData;

            RegisterItem TEST_EN_32K = Parent.RegMgr.GetRegisterItem("O_TEST_EN_32K");      // 0x5D
            RegisterItem TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");       // 0x61

            if (on == true)
            {
                rData = ReadRegister_Multi(0x5D, I2C_Sel);
                WriteRegister_Multi(0x5D, rData | 0x80, I2C_Sel);
                //TEST_EN_32K.Read();
                //TEST_EN_32K.Value = 1;
                //TEST_EN_32K.Write();

                rData = ReadRegister_Multi(0x61, I2C_Sel);
                WriteRegister_Multi(0x61, rData | 0x20, I2C_Sel);
                //TEST_CON.Read();
                //TEST_CON.Value = 32;
                //TEST_CON.Write();
            }
            else
            {
                rData = ReadRegister_Multi(0x5D, I2C_Sel);
                WriteRegister_Multi(0x5D, rData & 0x7F, I2C_Sel);
                //TEST_EN_32K.Read();
                //TEST_EN_32K.Value = 0;
                //TEST_EN_32K.Write();

                rData = ReadRegister_Multi(0x61, I2C_Sel);
                WriteRegister_Multi(0x61, rData & 0xDF, I2C_Sel);
                //TEST_CON.Read();
                //TEST_CON.Value = 0;
                //TEST_CON.Write();
            }
        }

        private void Enable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x01;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=0, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_On_NVM()
        {
            Write_I2C_Data(0x00060000, 0x00000007, 10);     // bist_type=1, bist_cmd=1, bist_sel=1
            Write_I2C_Data(0x00060000, 0x00000005, 10);     // bist_type=1, bist_cmd=0, bist_sel=1
        }

        private void Power_Off_NVM()
        {
            Write_I2C_Data(0x00060000, 0x0000000B, 10);     // bist_type=2, bist_cmd=1, bist_sel=1
            Write_I2C_Data(0x00060000, 0x00000009, 10);     // bist_type=2, bist_cmd=0, bist_sel=1
        }

        private void Write_I2C_Data(uint addr, uint data, UInt16 delay)
        {
            byte[] buf = new byte[8];
            buf[0] = (byte)(addr & 0xFF);
            buf[1] = (byte)((addr >> 8) & 0xFF);
            buf[2] = (byte)((addr >> 16) & 0xFF);
            buf[3] = (byte)((addr >> 24) & 0xFF);
            buf[4] = (byte)((data) & 0xFF);
            buf[5] = (byte)((data >> 8) & 0xFF);
            buf[6] = (byte)((data >> 16) & 0xFF);
            buf[7] = (byte)((data >> 24) & 0xFF);
            I2C.WriteBytes(buf, buf.Length, true);
            System.Threading.Thread.Sleep(delay);
        }

        private byte[] Read_I2C_Data(uint addr, UInt16 delay)
        {
            byte[] buf = new byte[4];
            byte[] recv_data = new byte[4];
            buf[0] = (byte)(addr & 0xFF);
            buf[1] = (byte)((addr >> 8) & 0xFF);
            buf[2] = (byte)((addr >> 16) & 0xFF);
            buf[3] = (byte)((addr >> 24) & 0xFF);
            I2C.WriteBytes(buf, buf.Length, false);
            recv_data = I2C.ReadBytes(recv_data.Length);
            System.Threading.Thread.Sleep(delay);

            return recv_data;
        }

        private void Otp_Pwr_On_Direct()
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            UInt32 addr = 0;
            UInt32 wdata = 0;

            // Check read
            // 0x0003'0004 - data 7 (default)
            addr = 0x00030004;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);

            // *** Power-on
            // Set PENVDD2 0x0000'8000 - 0x0000'0200
            addr = 0x00008000;
            wdata = (0 << 10) + (1 << 9) + (0 << 4) + (0 << 0);
            Write_I2C_Data(addr, wdata, 50);

            // Set PLDO 0x0000'8000 - 0x0000'0600
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (0 << 4) + (0 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);


            // Set PDRSTB 0x0000'8000 - 0x0000'0601
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);


            // Set PTM 0x0000'8000 -
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (2 << 5) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);
        }

        private void Otp_Wr_Direct()
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            UInt32 addr = 0;
            UInt32 wdata = 0;

            // Check read
            // 0x0003'0004 - data 7 (default)
            addr = 0x00030004;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);
            /*
            // *** Power-on
            // Set PENVDD2 0x0000'8000 - 0x0000'0200
            addr = 0x00008000;
            wdata = (0 << 10) + (1 << 9) + (0 << 4) + (0 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);
            
            // read
//            SendData[3] = 0x00; SendData[2] = 0x00; SendData[1] = 0x80; SendData[0] = 0x00;
//            I2C.WriteBytes(SendData, 4, false);
//            RcvData = I2C.ReadBytes(RcvData.Length);
//            System.Threading.Thread.Sleep(50);
            
            // Set PLDO 0x0000'8000 - 0x0000'0600
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (0 << 4) + (0 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);


            // Set PDRSTB 0x0000'8000 - 0x0000'0601
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);
            */
            /*
            // Set PTM 0x0000'8000 -
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (2 << 5) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);
            */

            // Hidden code
            addr = 0x00008000;
            wdata = ((UInt32)0x5A << 24) + (1 << 19) + (1 << 10) + (1 << 9) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);

            SendData[3] = 0x00; SendData[2] = 0x00; SendData[1] = 0x80; SendData[0] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);

            addr = 0x00008000;
            wdata = ((UInt32)0xA5 << 24) + (1 << 19) + (1 << 10) + (1 << 9) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);

            SendData[3] = 0x00; SendData[2] = 0x00; SendData[1] = 0x80; SendData[0] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);

            addr = 0x00008000;
            wdata = ((UInt32)0x3C << 24) + (1 << 19) + (1 << 10) + (1 << 9) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);

            SendData[3] = 0x00; SendData[2] = 0x00; SendData[1] = 0x80; SendData[0] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);

            addr = 0x00008000;
            wdata = ((UInt32)0xC3 << 24) + (1 << 19) + (1 << 10) + (1 << 9) + (0 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);

            SendData[3] = 0x00; SendData[2] = 0x00; SendData[1] = 0x80; SendData[0] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(100);

#if false
            // for test.
            // OTP data @0x0000'0000
            addr = 0x00000020;
            wdata = 0xFaFbFcFd;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(100);
#else
            addr = 0x00008014;
            wdata = 0x00000008;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);

            addr = 0x00008018;
            wdata = 0x89ABCDEF;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(100);
#endif

            // *** Read OTP
            // PCE
            addr = 0x00008000;
            wdata = (1 << 10) + (1 << 9) + (1 << 4) + (1 << 0);
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            SendData[4] = (byte)((wdata) & 0xFF);
            SendData[5] = (byte)((wdata >> 8) & 0xFF);
            SendData[6] = (byte)((wdata >> 16) & 0xFF);
            SendData[7] = (byte)((wdata >> 24) & 0xFF);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(50);

#if false
            // Read data
            //byte[] RcvData = new byte[4];
            addr = 0x00000020;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);
#else
            addr = 0x00008018;
            SendData[0] = (byte)(addr & 0xFF);
            SendData[1] = (byte)((addr >> 8) & 0xFF);
            SendData[2] = (byte)((addr >> 16) & 0xFF);
            SendData[3] = (byte)((addr >> 24) & 0xFF);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            System.Threading.Thread.Sleep(50);
#endif
        }

        private void Set_Sim_Val()
        {
            uint data = 0;
            /********** BLE **********/
            data = 0x02DE0001;
            WriteRegister(0xC, (data & 0xFF));
            WriteRegister(0xD, (data >> 8) & 0xFF);
            WriteRegister(0xE, (data >> 16) & 0xFF);
            WriteRegister(0xF, (data >> 24) & 0xFF);

            data = 0x0001FFFF;
            WriteRegister(0x10, (data & 0xFF));
            WriteRegister(0x11, (data >> 8) & 0xFF);
            WriteRegister(0x12, (data >> 16) & 0xFF);
            WriteRegister(0x13, (data >> 24) & 0xFF);

            data = 0x00001C4E;
            WriteRegister(0x14, (data & 0xFF));
            WriteRegister(0x15, (data >> 8) & 0xFF);
            WriteRegister(0x16, (data >> 16) & 0xFF);
            WriteRegister(0x17, (data >> 24) & 0xFF);

            data = 0xFFFFFFFF;
            WriteRegister(0x18, (data & 0xFF));
            WriteRegister(0x19, (data >> 8) & 0xFF);
            WriteRegister(0x1A, (data >> 16) & 0xFF);
            WriteRegister(0x1B, (data >> 24) & 0xFF);

            data = 0x000380D3;
            WriteRegister(0x1C, (data & 0xFF));
            WriteRegister(0x1D, (data >> 8) & 0xFF);
            WriteRegister(0x1E, (data >> 16) & 0xFF);
            WriteRegister(0x1F, (data >> 24) & 0xFF);

            /********** CTRL **********/
            data = 0x8E89BED6;
            WriteRegister(0x20, (data & 0xFF));
            WriteRegister(0x21, (data >> 8) & 0xFF);
            WriteRegister(0x22, (data >> 16) & 0xFF);
            WriteRegister(0x23, (data >> 24) & 0xFF);

            data = 0x00000000;
            WriteRegister(0x24, (data & 0xFF));
            WriteRegister(0x25, (data >> 8) & 0xFF);
            WriteRegister(0x26, (data >> 16) & 0xFF);
            WriteRegister(0x27, (data >> 24) & 0xFF);

            data = 0x04FF0707;
            WriteRegister(0x28, (data & 0xFF));
            WriteRegister(0x29, (data >> 8) & 0xFF);
            WriteRegister(0x2A, (data >> 16) & 0xFF);
            WriteRegister(0x2B, (data >> 24) & 0xFF);

            data = 0x24000008;
            WriteRegister(0x2C, (data & 0xFF));
            WriteRegister(0x2D, (data >> 8) & 0xFF);
            WriteRegister(0x2E, (data >> 16) & 0xFF);
            WriteRegister(0x2F, (data >> 24) & 0xFF);

            data = 0x00000000;
            WriteRegister(0x30, (data & 0xFF));
            WriteRegister(0x31, (data >> 8) & 0xFF);
            WriteRegister(0x32, (data >> 16) & 0xFF);
            WriteRegister(0x33, (data >> 24) & 0xFF);

            data = 0x00000000;
            WriteRegister(0x34, (data & 0xFF));
            WriteRegister(0x35, (data >> 8) & 0xFF);
            WriteRegister(0x36, (data >> 16) & 0xFF);
            WriteRegister(0x37, (data >> 24) & 0xFF);

            data = 0x00000000;
            WriteRegister(0x38, (data & 0xFF));
            WriteRegister(0x39, (data >> 8) & 0xFF);
            WriteRegister(0x3A, (data >> 16) & 0xFF);
            WriteRegister(0x3B, (data >> 24) & 0xFF);

            data = 0x000C0003;
            WriteRegister(0x3C, (data & 0xFF));
            WriteRegister(0x3D, (data >> 8) & 0xFF);
            WriteRegister(0x3E, (data >> 16) & 0xFF);
            WriteRegister(0x3F, (data >> 24) & 0xFF);
        }

        private void Simple_I2C_Cmd(string cmd)
        {
            //RegisterItem w_WURF_END = Parent.RegMgr.GetRegisterItem("w_WURF_END");         // 0x38
            byte[] SendData = new byte[2];

            SendData[0] = 0xAA;
            SendData[1] = System.Convert.ToByte(cmd, 16);

            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private uint Get_AON_OOK_Test()
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            UInt32 ret = 0;

            SendData[3] = 0x00; SendData[2] = 0x00; SendData[1] = 0x80; SendData[0] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            ret = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);

            return ret;
        }

        private void Set_AON_OOK_Test()
        {
            uint data = 0;

            /********** BLE **********/
            //data = 0x1EFF6D0B;
            data = 0x0B6DFF1E;
            WriteRegister(0x0, (data & 0xFF));
            WriteRegister(0x1, (data >> 8) & 0xFF);
            WriteRegister(0x2, (data >> 16) & 0xFF);
            WriteRegister(0x3, (data >> 24) & 0xFF);

            data = 0x6EB60000;
            //data = 0x0000B66E;
            WriteRegister(0x4, (data & 0xFF));
            WriteRegister(0x5, (data >> 8) & 0xFF);
            WriteRegister(0x6, (data >> 16) & 0xFF);
            WriteRegister(0x7, (data >> 24) & 0xFF);

            data = 0x02DEE781;
            //data = 0x81E7DE02;
            WriteRegister(0x8, (data & 0xFF));
            WriteRegister(0x9, (data >> 8) & 0xFF);
            WriteRegister(0xA, (data >> 16) & 0xFF);
            WriteRegister(0xB, (data >> 24) & 0xFF);

            data = 0x0001FFFF;
            //data = 0xFFFF0100;
            WriteRegister(0xC, (data & 0xFF));
            WriteRegister(0xD, (data >> 8) & 0xFF);
            WriteRegister(0xE, (data >> 16) & 0xFF);
            WriteRegister(0xF, (data >> 24) & 0xFF);

            data = 0x00001C4E;
            //data = 0x4E1C0000;
            WriteRegister(0x10, (data & 0xFF));
            WriteRegister(0x11, (data >> 8) & 0xFF);
            WriteRegister(0x12, (data >> 16) & 0xFF);
            WriteRegister(0x13, (data >> 24) & 0xFF);

            data = 0x0002FFFF;
            //data = 0xFFFF0200;
            WriteRegister(0x14, (data & 0xFF));
            WriteRegister(0x15, (data >> 8) & 0xFF);
            WriteRegister(0x16, (data >> 16) & 0xFF);
            WriteRegister(0x17, (data >> 24) & 0xFF);

            data = 0xFFFFFFFF;
            //data = 0xFFFFFFFF;
            WriteRegister(0x18, (data & 0xFF));
            WriteRegister(0x19, (data >> 8) & 0xFF);
            WriteRegister(0x1A, (data >> 16) & 0xFF);
            WriteRegister(0x1B, (data >> 24) & 0xFF);

            /********** CTRL **********/
            data = 0x8E89BED6;
            //data = 0xD6BE898E;
            WriteRegister(0x20, (data & 0xFF));
            WriteRegister(0x21, (data >> 8) & 0xFF);
            WriteRegister(0x22, (data >> 16) & 0xFF);
            WriteRegister(0x23, (data >> 24) & 0xFF);
            data = 0x05000000;
            //data = 0x00000005;
            WriteRegister(0x24, (data & 0xFF));
            WriteRegister(0x25, (data >> 8) & 0xFF);
            WriteRegister(0x26, (data >> 16) & 0xFF);
            WriteRegister(0x27, (data >> 24) & 0xFF);
            data = 0x07FF0000;
            //data = 0x0000FF07;
            WriteRegister(0x28, (data & 0xFF));
            WriteRegister(0x29, (data >> 8) & 0xFF);
            WriteRegister(0x2A, (data >> 16) & 0xFF);
            WriteRegister(0x2B, (data >> 24) & 0xFF);
            data = 0x024A8008;
            //data = 0x08804A02;
            WriteRegister(0x2C, (data & 0xFF));
            WriteRegister(0x2D, (data >> 8) & 0xFF);
            WriteRegister(0x2E, (data >> 16) & 0xFF);
            WriteRegister(0x2F, (data >> 24) & 0xFF);
#if false
            /********** ANALOG **********/
            data = 0x72111210;
            //data = 0x10121172;
            WriteRegister(0x40, (data & 0xFF));
            WriteRegister(0x41, (data >> 8) & 0xFF);
            WriteRegister(0x42, (data >> 16) & 0xFF);
            WriteRegister(0x43, (data >> 24) & 0xFF);

            data = 0xAA008482;
            //data = 0x828400AA;
            WriteRegister(0x44, (data & 0xFF));
            WriteRegister(0x45, (data >> 8) & 0xFF);
            WriteRegister(0x46, (data >> 16) & 0xFF);
            WriteRegister(0x47, (data >> 24) & 0xFF);

            data = 0x42FF7676;
            //data = 0x7676FF42;
            WriteRegister(0x48, (data & 0xFF));
            WriteRegister(0x49, (data >> 8) & 0xFF);
            WriteRegister(0x4A, (data >> 16) & 0xFF);
            WriteRegister(0x4B, (data >> 24) & 0xFF);

            data = 0x173FFF62;
            //data = 0x62FF3F17;
            WriteRegister(0x4C, (data & 0xFF));
            WriteRegister(0x4D, (data >> 8) & 0xFF);
            WriteRegister(0x4E, (data >> 16) & 0xFF);
            WriteRegister(0x4F, (data >> 24) & 0xFF);

            data = 0x01402900;
            //data = 0x00294001;
            WriteRegister(0x50, (data & 0xFF));
            WriteRegister(0x51, (data >> 8) & 0xFF);
            WriteRegister(0x52, (data >> 16) & 0xFF);
            WriteRegister(0x53, (data >> 24) & 0xFF);

            data = 0x004444C5;
            //data = 0xC5444400;
            //data = 0xC5444445;
            WriteRegister(0x54, (data & 0xFF));
            WriteRegister(0x55, (data >> 8) & 0xFF);
            WriteRegister(0x56, (data >> 16) & 0xFF);
            WriteRegister(0x57, (data >> 24) & 0xFF);

            data = 0x052ACEB0;
            //data = 0xB0CE2A05;
            WriteRegister(0x58, (data & 0xFF));
            WriteRegister(0x59, (data >> 8) & 0xFF);
            WriteRegister(0x5A, (data >> 16) & 0xFF);
            WriteRegister(0x5B, (data >> 24) & 0xFF);

            data = 0x00000142;
            //data = 0x42010000;
            WriteRegister(0x5C, (data & 0xFF));
            WriteRegister(0x5D, (data >> 8) & 0xFF);
            WriteRegister(0x5E, (data >> 16) & 0xFF);
            WriteRegister(0x5F, (data >> 24) & 0xFF);

            data = 0x05206000;
            //data = 0x00602005;
            WriteRegister(0x60, (data & 0xFF));
            WriteRegister(0x61, (data >> 8) & 0xFF);
            WriteRegister(0x62, (data >> 16) & 0xFF);
            WriteRegister(0x63, (data >> 24) & 0xFF);

            data = 0x00000008;
            //data = 0x08000000;
            WriteRegister(0x64, (data & 0xFF));
            WriteRegister(0x65, (data >> 8) & 0xFF);
            WriteRegister(0x66, (data >> 16) & 0xFF);
            WriteRegister(0x67, (data >> 24) & 0xFF);

            data = 0x80325450;
            //data = 0x50543280;
            WriteRegister(0x68, (data & 0xFF));
            WriteRegister(0x69, (data >> 8) & 0xFF);
            WriteRegister(0x6A, (data >> 16) & 0xFF);
            WriteRegister(0x6B, (data >> 24) & 0xFF);
#endif
        }

        private void Set_Otp_ctrl_ana()
        {
            // BLE p.xx
            uint ble_pnum = 0;     // 주의: 0번부터 시작
            WriteRegister(BLE_BASE + 0x0 + (ble_pnum * 36), VALID_CODE);
            //WriteRegister(BLE_BASE + 0x4 +  (ble_pnum * 36), );
            //WriteRegister(BLE_BASE + 0x8 +  (ble_pnum * 36), 0x);
            //WriteRegister(BLE_BASE + 0xC +  (ble_pnum * 36), 0x);
            //WriteRegister(BLE_BASE + 0x10 + (ble_pnum * 36), 0x);
            //WriteRegister(BLE_BASE + 0x14 + (ble_pnum * 36), 0x);
            //WriteRegister(BLE_BASE + 0x18 + (ble_pnum * 36), 0x);
            //WriteRegister(BLE_BASE + 0x1C + (ble_pnum * 36), 0x);
            //WriteRegister(BLE_BASE + 0x20 + (ble_pnum * 36), 0x);


            // control p.xx
            uint crtl_pnum = 0;     // 주의: 0번부터 시작
            WriteRegister(CTRL_BASE + 0x0 + (crtl_pnum * 36), VALID_CODE);
            WriteRegister(CTRL_BASE + 0x4 + (crtl_pnum * 36), 0x8E89BED6);
            WriteRegister(CTRL_BASE + 0x8 + (crtl_pnum * 36), 0x05000000);
            WriteRegister(CTRL_BASE + 0xC + (crtl_pnum * 36), 0x07FF0000);
            WriteRegister(CTRL_BASE + 0x10 + (crtl_pnum * 36), 0x024A8008);


            // ananlog p.xx
            uint ana_pnum = 0;      // 주의: 0번부터 시작
            WriteRegister(ANA_BASE + 0x0 + (ana_pnum * 48), VALID_CODE);
            WriteRegister(ANA_BASE + 0x4 + (ana_pnum * 48), 0x72111210);
            WriteRegister(ANA_BASE + 0x8 + (ana_pnum * 48), 0xAA008482);
            WriteRegister(ANA_BASE + 0xC + (ana_pnum * 48), 0x42FF7676);
            WriteRegister(ANA_BASE + 0x10 + (ana_pnum * 48), 0x173FFF62);
            WriteRegister(ANA_BASE + 0x14 + (ana_pnum * 48), 0x01402900);
            WriteRegister(ANA_BASE + 0x18 + (ana_pnum * 48), 0x004444C5);
            WriteRegister(ANA_BASE + 0x1C + (ana_pnum * 48), 0x052ACEB0);
            WriteRegister(ANA_BASE + 0x20 + (ana_pnum * 48), 0x00000142);
            WriteRegister(ANA_BASE + 0x24 + (ana_pnum * 48), 0x05206000);
            WriteRegister(ANA_BASE + 0x28 + (ana_pnum * 48), 0x00000008);
            WriteRegister(ANA_BASE + 0x2C + (ana_pnum * 48), 0x80325450);
        }

        private void Set_Wp_Adv()
        {
            // Set advertised mode, Wake-up period
            WriteRegister(0xC, 1);
            WriteRegister(0xD, 0);
            WriteRegister(0xE, 210);
            WriteRegister(0xF, 29);
            WriteRegister(0x10, 0);
            WriteRegister(0x11, 0);
            WriteRegister(0x14, 3);
            WriteRegister(0x15, 0);

            //temp
            WriteRegister(0x2B, 0x40);
        }
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

        private bool Check_Revision_Information(uint value)
        {
            RegisterItem REV_ID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");     // 0x68
            bool result = false;

            REV_ID.Read();

            if (value == REV_ID.Value)
            {
                result = true;
            }

            return result;
        }

        private uint Read_Mac_Address()
        {
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10

            B14_BA_0.Read();
            B15_BA_1.Read();
            B16_BA_2.Read();

            return (B16_BA_2.Value << 16) | (B15_BA_1.Value << 8) | B14_BA_0.Value;
        }

        private bool Check_OTP_VALID(bool start, uint page)
        {
            uint addr_nvm;
            bool result = true;

            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            byte[] Data = new byte[4];

            if (start == false)
            {
                return false;
            }

            Enable_NVM_BIST();
            Power_On_NVM();

            if (page < 19)
            {
                addr_nvm = page * 44; // ble, control
            }
            else
            {
                addr_nvm = page * 44 + (page - 19) * 4; // Analog
            }

            // NVM Verify // VALID
            Data[0] = 0x1D;
            Data[1] = 0xCA;
            Data[2] = 0x1D;
            Data[3] = 0xCA;
            // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
            SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            SendData[4] = (byte)(0x13);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
            SendData[4] = (byte)(0x11);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Read bist_rdata
            SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            for (int j = 0; j < 4; j++)
            {
                if (RcvData[j] != Data[j])
                {
                    result = false;
                }
            }

            Power_Off_NVM();
            Disable_NVM_BIST();

            return result;
        }

        private bool Read_FF_NVM(bool start)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return true;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // bist_sel=1, bist_type=6, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=6, bist_cmd=0
            SendData[4] = 0x19;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] != 0x0b)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void Write_Register_AON_Fix_Value()
        {
            RegisterItem w_ADC_EOC_SKIP_TEMP = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_TEMP[5:0]");  // 0x34
            RegisterItem w_ADC_EOC_SKIP_VOLT = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_VOLT[5:0]");  // 0x35
            RegisterItem w_LPCAL_FINE_EN = Parent.RegMgr.GetRegisterItem("w_LPCAL_FINE_EN");               // 0x37
            RegisterItem w_TAMP_DET_EN = Parent.RegMgr.GetRegisterItem("w_TAMP_DET_EN");                   // 0x39
            RegisterItem w_GPIO_WKUP_MD = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_MD");                 // 0x41
            RegisterItem w_GPIO_WKUP_INTV = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_INTV[1:0]");        // 0x41
            RegisterItem O_CT_RES2_FINE = Parent.RegMgr.GetRegisterItem("O_CT_RES2_FINE[3:0]");            // 0x45
            RegisterItem w_GPIO_WKUP_EN = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_EN");                 // 0x4E
            RegisterItem O_ADC_CM = Parent.RegMgr.GetRegisterItem("O_ADC_CM[1:0]");                        // 0x55
            RegisterItem w_DA_GAIN_MAX = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_MAX[2:0]");              // 0x57
            RegisterItem I_EXT_DA_GAIN_CON = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_CON[1:0]");      // 0x5A
            RegisterItem w_XTAL_CUR_CFG_1 = Parent.RegMgr.GetRegisterItem("w_XTAL_CUR_CFG_1[5:0]");        // 0x5E
            RegisterItem w_DA_GAIN_INTV = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_INTV[1:0]");            // 0x61
            RegisterItem w_PLL_CH_19_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_19_GAIN[2:0]");        // 0x61
            RegisterItem VS_GainMode = Parent.RegMgr.GetRegisterItem("VS_GainMode");                       // 0x61
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL[5:2]");                  // 0x62
            RegisterItem w_TX_SEL_SELECT = Parent.RegMgr.GetRegisterItem("w_TX_SEL_SELECT[1:0]");          // 0x64
            RegisterItem w_PLL_PM_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_PM_GAIN[4:0]");              // 0x64
            RegisterItem w_PLL_CH_0_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_0_GAIN[4:0]");          // 0x64
            RegisterItem w_DA_PEN_SELECT = Parent.RegMgr.GetRegisterItem("w_DA_PEN_SELECT[1:0]");          // 0x65
            RegisterItem w_PLL_CH_12_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_12_GAIN[4:0]");        // 0x66
            RegisterItem w_ADC_SWITCH_MODE = Parent.RegMgr.GetRegisterItem("w_ADC_SWITCH_MODE");           // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_H = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[4]");        // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_L = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[3:0]");      // 0x67

            BGR_TC_CTRL.Read();
            BGR_TC_CTRL.Value = 11;
            BGR_TC_CTRL.Write();

            w_ADC_SAMPLE_CNT_H.Read();
            w_ADC_SAMPLE_CNT_H.Value = 1;
            w_ADC_SAMPLE_CNT_H.Write();

            w_ADC_SAMPLE_CNT_L.Read();
            w_ADC_SAMPLE_CNT_L.Value = 15;
            w_ADC_SAMPLE_CNT_L.Write();

            O_CT_RES2_FINE.Read();
            O_CT_RES2_FINE.Value = 9;
            O_CT_RES2_FINE.Write();

            w_DA_GAIN_MAX.Read();
            w_DA_GAIN_MAX.Value = 6;
            w_DA_GAIN_MAX.Write();

            I_EXT_DA_GAIN_CON.Read();
            I_EXT_DA_GAIN_CON.Value = 2;
            I_EXT_DA_GAIN_CON.Write();

            VS_GainMode.Read();
            VS_GainMode.Value = 0;
            VS_GainMode.Write();

            w_TX_SEL_SELECT.Read();
            w_TX_SEL_SELECT.Value = 2;
            w_TX_SEL_SELECT.Write();

            w_DA_PEN_SELECT.Read();
            w_DA_PEN_SELECT.Value = 2;
            w_DA_PEN_SELECT.Write();
        }

        private bool Run_Cal_BGR(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 300;
            double d_lsl = 295, d_usl = 305;
            uint ldo_val, ldo_val_1;

            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");    // 0x57

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_BGR(true);

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(2, (1 + cnt), (double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));

                ULP_BGR_CONT.Read();
                Parent.xlMgr.Sheet.Select("BGR");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            ldo_val = 15;
            ldo_val_1 = 0;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }

            Set_TestInOut_For_BGR(false);

            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_ALLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 810;
            double d_lsl = 800, d_usl = 820;
            uint ldo_val, ldo_val_1;

            RegisterItem O_ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");        // 0x54
            RegisterItem O_ULP_LDO_LV_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[2:0]");  // 0x61

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            O_ULP_LDO_CONT.Read();
            O_ULP_LDO_LV_CONT.Read();

            O_ULP_LDO_CONT.Value = 8;
            O_ULP_LDO_CONT.Write();
            O_ULP_LDO_LV_CONT.Value = 0;
            O_ULP_LDO_LV_CONT.Write();
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("ALLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv > 816)
            {
                O_ULP_LDO_LV_CONT.Value = 3;
                O_ULP_LDO_LV_CONT.Write();
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (O_ULP_LDO_LV_CONT.Value).ToString());
                Parent.xlMgr.Cell.Write(x_pos + 1, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_MLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 1000;
            double d_lsl = 990, d_usl = 1010;
            uint ldo_val, ldo_val_1;

            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");          // 0x53
            RegisterItem PMU_MLDO_Coarse_L = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[0]");    // 0x62
            RegisterItem PMU_MLDO_Coarse_H = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1]");    // 0x63

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            PMU_MLDO_Coarse_H.Read();
            PMU_MLDO_Coarse_L.Read();
            PMU_LDO_CONT.Read();

            PMU_LDO_CONT.Value = 7;
            PMU_LDO_CONT.Write();
            PMU_MLDO_Coarse_L.Value = 0;
            PMU_MLDO_Coarse_L.Write();

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("MLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv < 1006)
            {
                PMU_MLDO_Coarse_L.Value = 1;
                PMU_MLDO_Coarse_L.Write();
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < 1006)
                {
                    PMU_MLDO_Coarse_L.Value = 0;
                    PMU_MLDO_Coarse_L.Write();
                    PMU_MLDO_Coarse_H.Value = 1;
                    PMU_MLDO_Coarse_H.Write();
                }
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();
            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            if (x_pos != 0)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            }
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ((PMU_MLDO_Coarse_H.Value << 1) | PMU_MLDO_Coarse_L.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_32K_RCOSC(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_khz, d_diff_khz;
            double d_lsl = 32.113, d_usl = 33.423, d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;

            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");      // 0x4D
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");     // 0x4E

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_RCOSC(true);
            if (x_pos != 0)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_khz.ToString("F3"));

                Parent.xlMgr.Sheet.Select("RCOSC");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            // RTC_SCKF[5:0]
            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            for (int val = 4; val >= 0; val--)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
                }
                if (d_freq_khz > d_target_khz)
                {
                    osc_val_l += (uint)(1 << val);
                }
                else
                {
                    osc_val_l -= (uint)(1 << val);
                }
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
            {
                if (osc_val_l != 63) osc_val_l += 1;
            }
            else
            {
                if (osc_val_l != 0) osc_val_l -= 1;
            }
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (RTC_SCKF_H.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), (RTC_SCKF_L.Value).ToString());
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_freq_khz.ToString("F3"));
            }
            Set_TestInOut_For_RCOSC(false);

            if ((d_freq_khz < d_lsl) || (d_freq_khz > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_Temp_Sensor(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 140; // 30deg
            int lsl = 135, usl = 145;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem TEMP_CONT_L = Parent.RegMgr.GetRegisterItem("O_TEMP_CONT[4:0]");  // 0x55
            RegisterItem TEMP_CONT_H = Parent.RegMgr.GetRegisterItem("TEMP_TRIM[5]");      // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
            TEMP_CONT_H.Read();
            TEMP_CONT_L.Read();

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct);
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();
            Set_TestInOut_For_VTEMP(false);
#if false // cal
            Parent.xlMgr.Sheet.Select("Temp_Sen");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
                u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
            }
#else // fix
            u_adc_val = 0;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            lsl = 0;
            usl = 255;
#endif
            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

            Set_TestInOut_For_VTEMP(false);
#if false
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_Voltage_Scaler(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 45; // 2.0V
            int lsl = 40, usl = 50;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem VOLSCAL_CON_L = Parent.RegMgr.GetRegisterItem("O_VOLSCAL_CON[3:0]");  // 0x56
            RegisterItem VOLTAGE_CON_H = Parent.RegMgr.GetRegisterItem("VOLTAGE_CON[5:4]");    // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.0,(@3)");
#else
            PowerSupply0.Write("VOLT 2.0");
#endif
            VOLTAGE_CON_H.Read();
            VOLSCAL_CON_L.Read();

            Set_TestInOut_For_VS(true);

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();

            Parent.xlMgr.Sheet.Select("VS");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
                u_adc_code = (Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct) >> 8) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
            }

            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = (Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct) & 0xff);
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif

            Set_TestInOut_For_VS(false);
#if true
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_32M_XTAL_Load_Cap(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_mhz;
            double d_diff_mhz, d_target_mhz = 2402;
            double d_lsl = 2401.9952, d_usl = 2402.0048;
            uint osc_val, osc_val_1;

            RegisterItem XTAL_LOAD_CONT = Parent.RegMgr.GetRegisterItem("O_XTAL_LOAD_CONT[4:0]");  // 0x4F

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            XTAL_LOAD_CONT.Read();
            if (x_pos != 0)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
                Parent.xlMgr.Sheet.Select("XTAL_Cap");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            osc_val = 15;
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            for (int val = 2; val >= 0; val--)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)val), (1 + cnt), d_freq_mhz.ToString("F4"));
                }
                if (d_freq_mhz > d_target_mhz)
                {
                    osc_val += (uint)(1 << val);
                }
                else
                {
                    osc_val -= (uint)(1 << val);
                }
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            osc_val_1 = osc_val;
            d_diff_mhz = Math.Abs(d_freq_mhz - d_target_mhz);

            if (d_freq_mhz > d_target_mhz)
            {
                if (osc_val != 31) osc_val += 1;
            }
            else
            {
                if (osc_val != 0) osc_val -= 1;
            }
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            System.Threading.Thread.Sleep(1);

            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            if (Math.Abs(d_freq_mhz - d_target_mhz) > d_diff_mhz)
            {
                osc_val = osc_val_1;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            }

            if ((d_freq_mhz < d_lsl) || (d_freq_mhz > d_usl))
            {
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return false;
            }
            else
            {
#if true // offset
                osc_val += 3;
                if (osc_val > 31) osc_val = 31;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
#endif
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return true;
            }

        }

        private bool Run_Cal_PLL_2PM(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_fsm;
            uint crc_flag;

            RegisterItem PM_IN = Parent.RegMgr.GetRegisterItem("O_PM_IN[7:0]");                    // 0x58

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            SendCommand("ook 70");
            System.Threading.Thread.Sleep(100);

            u_fsm = Run_Read_FSM_Status() & 0x000f;
            crc_flag = CRC_FLAG_READ();

            if (u_fsm != 7)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "FSM_Error");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) > 1)) // WUR_PKT_END = 1
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CRC_ERROR");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) == 1))
            {
                PM_IN.Read();
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (PM_IN.Value).ToString());
            }
            else
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CHECK_LOG");
                Console.WriteLine("u_fsm : {0}", u_fsm);
                Console.WriteLine("crc_flag : {0}", crc_flag);
                return false;
            }
            return true;
        }

        private void Set_AON_REG_For_BLE(byte[] aon_ble)
        {
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");            // 0x04
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");            // 0x05
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");            // 0x06
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");            // 0x07
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");            // 0x08
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");            // 0x09
            RegisterItem B10_LEN = Parent.RegMgr.GetRegisterItem("B10_LEN[7:0]");              // 0x0A
            RegisterItem B11_MSDF = Parent.RegMgr.GetRegisterItem("B11_MSDF[7:0]");            // 0x0B
            RegisterItem B12_CID_0 = Parent.RegMgr.GetRegisterItem("B12_CID_0[7:0]");          // 0x0C
            RegisterItem B13_CID_1 = Parent.RegMgr.GetRegisterItem("B13_CID_1[7:0]");          // 0x0D
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10
            RegisterItem B17_BA_3 = Parent.RegMgr.GetRegisterItem("B17_BA_3[7:0]");            // 0x11
            RegisterItem B18_BA_4 = Parent.RegMgr.GetRegisterItem("B18_BA_4[7:0]");            // 0x12
            RegisterItem B19_BA_5 = Parent.RegMgr.GetRegisterItem("B19_BA_5[7:0]");            // 0x13
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");            // 0x14
            RegisterItem B21_DC_1 = Parent.RegMgr.GetRegisterItem("B21_DC_1[7:0]");            // 0x15
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");            // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");            // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");            // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");            // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");            // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");            // 0x1B
            RegisterItem B28_RSV_0 = Parent.RegMgr.GetRegisterItem("B28_PUTC_0[7:0]");         // 0x1C
            RegisterItem B29_RSV_1 = Parent.RegMgr.GetRegisterItem("B29_PUTC_1[7:0]");         // 0x1D
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");            // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");            // 0x1F
            RegisterItem B32_PUTC_0 = Parent.RegMgr.GetRegisterItem("B32_CC_0[7:0]");          // 0x20
            RegisterItem B33_PUTC_1 = Parent.RegMgr.GetRegisterItem("B33_CC_1[7:0]");          // 0x21
            RegisterItem B34_PUTC_2 = Parent.RegMgr.GetRegisterItem("B34_CC_2[7:0]");          // 0x22
            RegisterItem B35_PUTC_3 = Parent.RegMgr.GetRegisterItem("B35_CC_3[7:0]");          // 0x23
            RegisterItem B36_BL = Parent.RegMgr.GetRegisterItem("B36_BL[7:0]");                // 0x24
            RegisterItem B37_TP = Parent.RegMgr.GetRegisterItem("B37_TP[7:0]");                // 0x25
            RegisterItem B38_SN_0 = Parent.RegMgr.GetRegisterItem("B38_SN_0[7:0]");            // 0x26
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27
            RegisterItem B40_MIC = Parent.RegMgr.GetRegisterItem("B40_MIC[7:0]");              // 0x28
            RegisterItem B41_advDelay = Parent.RegMgr.GetRegisterItem("B41_ADD_0[7:0]");       // 0x29
            RegisterItem B42_advDelay = Parent.RegMgr.GetRegisterItem("B42_ADD_1[7:0]");       // 0x2A
            RegisterItem B43_advDelay = Parent.RegMgr.GetRegisterItem("B43_ADD_2[7:0]");       // 0x2B

            B04_BA_0.Value = aon_ble[0];
            B04_BA_0.Write();

            B05_BA_1.Value = aon_ble[1];
            B05_BA_1.Write();

            B06_BA_2.Value = aon_ble[2];
            B06_BA_2.Write();

            B07_BA_3.Value = aon_ble[3];
            B07_BA_3.Write();

            B08_BA_4.Value = aon_ble[4];
            B08_BA_4.Write();

            B09_BA_5.Value = aon_ble[5];
            B09_BA_5.Write();

            B10_LEN.Value = aon_ble[6];
            B10_LEN.Write();

            B11_MSDF.Value = aon_ble[7];
            B11_MSDF.Write();

            B12_CID_0.Value = aon_ble[8];
            B12_CID_0.Write();

            B13_CID_1.Value = aon_ble[9];
            B13_CID_1.Write();

            B14_BA_0.Value = aon_ble[10];
            B14_BA_0.Write();

            B15_BA_1.Value = aon_ble[11];
            B15_BA_1.Write();

            B16_BA_2.Value = aon_ble[12];
            B16_BA_2.Write();

            B17_BA_3.Value = aon_ble[13];
            B17_BA_3.Write();

            B18_BA_4.Value = aon_ble[14];
            B18_BA_4.Write();

            B19_BA_5.Value = aon_ble[15];
            B19_BA_5.Write();

            B20_DC_0.Value = aon_ble[16];
            B20_DC_0.Write();

            B21_DC_1.Value = aon_ble[17];
            B21_DC_1.Write();

            B22_WP_0.Value = aon_ble[18];
            B22_WP_0.Write();

            B23_WP_1.Value = aon_ble[19];
            B23_WP_1.Write();

            B24_AI_0.Value = aon_ble[20];
            B24_AI_0.Write();

            B25_AI_1.Value = aon_ble[21];
            B25_AI_1.Write();

            B26_AI_2.Value = aon_ble[22];
            B26_AI_2.Write();

            B27_AI_3.Value = aon_ble[23];
            B27_AI_3.Write();

            B28_RSV_0.Value = aon_ble[24];
            B28_RSV_0.Write();

            B29_RSV_1.Value = aon_ble[25];
            B29_RSV_1.Write();

            B30_AD_0.Value = aon_ble[26];
            B30_AD_0.Write();

            B31_AD_1.Value = aon_ble[27];
            B31_AD_1.Write();

            B32_PUTC_0.Value = aon_ble[28];
            B32_PUTC_0.Write();

            B33_PUTC_1.Value = aon_ble[29];
            B33_PUTC_1.Write();

            B34_PUTC_2.Value = aon_ble[30];
            B34_PUTC_2.Write();

            B35_PUTC_3.Value = aon_ble[31];
            B35_PUTC_3.Write();

            B36_BL.Value = aon_ble[32];
            B36_BL.Write();

            B37_TP.Value = aon_ble[33];
            B37_TP.Write();

            B38_SN_0.Value = aon_ble[34];
            B38_SN_0.Write();

            B39_SN_1.Value = aon_ble[35];
            B39_SN_1.Write();

            B40_MIC.Value = aon_ble[36];
            B40_MIC.Write();

            B41_advDelay.Value = aon_ble[37];
            B41_advDelay.Write();

            B42_advDelay.Value = aon_ble[38];
            B42_advDelay.Write();

            B43_advDelay.Value = aon_ble[39];
            B43_advDelay.Write();
        }

        private void Run_Write_OTP_With_OOK(bool start, uint page)
        {
            string cmd;

            if (start == false)
            {
                return;
            }

            cmd = "ook 41." + page.ToString("X2");
            SendCommand(cmd);

            System.Threading.Thread.Sleep(500);
        }

        private uint Run_Verify_OTP_With_BIST(bool start, uint page)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return 65535;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // set page_flag
            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = (byte)(page & 0xff);
            SendData[5] = (byte)((page >> 8) & 0xff);
            SendData[6] = (byte)(((page >> 16) & 0x3f) | 0x80);
            SendData[7] = 0xBB;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1F;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=0
            SendData[4] = 0x1D;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] == 0x0b)
            {
                return 0x400;
            }
            else
            {
                return RcvData[0];
            }
        }

        private bool Run_Measure_Initial(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }
            // BGR, ALLDO, MLDO
            Set_TestInOut_For_BGR(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                // BGR
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 295) || (d_val > 305))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_12");
                        result = false;
                    }
                }
                // ALLDO
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 3), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 800) || (d_val > 820))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_13");
                        result = false;
                    }
                }

                // MLDO
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 6), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 985) || (d_val > 1015))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_14");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_BGR(false);

            // RCOSC
            Set_TestInOut_For_RCOSC(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + i + 9), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 32.113) || (d_val > 33.423))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_16");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_RCOSC(false);

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_ADC(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            uint u_adc_code, u_temp_code, u_vbat_code;
            uint u_lsl_temp = 130, u_usl_temp = 150, u_lsl_vbat = 255, u_usl_vbat = 0;
            double d_volt_mv;

            if (start == false)
            {
                for (int i = 0; i < 8; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }

            Set_TestInOut_For_VTEMP(true);

            System.Threading.Thread.Sleep(1);
            for (int i = 0; i < 4; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@3)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        u_lsl_vbat = 186;
                        u_usl_vbat = 196;
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        u_lsl_vbat = 96;
                        u_usl_vbat = 106;
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.0,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.0");
#endif
                        u_lsl_vbat = 40;
                        u_usl_vbat = 50;
                        break;
                    case 3:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@3)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        u_lsl_vbat = 6;
                        u_usl_vbat = 16;
                        break;
                    default:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                }
                System.Threading.Thread.Sleep(100);
                u_adc_code = Read_ADC_Result(false, ref eoc_vte, 5, ref vref_bct);
                u_temp_code = u_adc_code & 0xff;
                u_vbat_code = (u_adc_code >> 8) & 0xff;
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 0), (y_pos + cnt), u_temp_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 1), (y_pos + cnt), u_vbat_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
                Disable_ADC();
#if false
                if ((u_temp_code < u_lsl_temp) || (u_temp_code > u_usl_temp))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_17");
                        result = false;
                    }
                }
#endif
                if ((u_vbat_code < u_lsl_vbat) || (u_vbat_code > u_usl_vbat))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_18");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_VTEMP(false);
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_VCO_Range(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_freq_mhz;
            double d_sl = 2402;

            RegisterItem VCO_TEST = Parent.RegMgr.GetRegisterItem("O_VCO_TEST");           // 0x41
            RegisterItem VCO_CBANK_L = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[7:0]");  // 0x42
            RegisterItem VCO_CBANK_H = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[9:8]");  // 0x43

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            VCO_TEST.Read();
            VCO_CBANK_L.Read();
            VCO_CBANK_H.Read();

            VCO_TEST.Value = 1;
            VCO_TEST.Write();

            // VCO Low = 0
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 0;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:SPAN 500 MHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.2 GHZ");
            System.Threading.Thread.Sleep(400);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 0), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz > d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO High = 1023
            VCO_CBANK_L.Value = 255;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 3;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.7 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz < d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO Mid = 512
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 2;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.5 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(false);
            VCO_TEST.Value = 0;
            VCO_TEST.Write();

            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");

            return result;
        }

        private bool Run_Measure_Tx_Power_Harmonic(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_power_dbm, d_freq_MHz, d_cur_mA;
            double d_lsl = -2, d_usl = 5;
            uint ch;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return result;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;
            PowerSupply0.Write("SENS:CURR:RANG 0.01,(@3)");

            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                    case 1:
                        ch = 12;
                        d_freq_MHz = 2426;
                        break;
                    case 2:
                        ch = 39;
                        d_freq_MHz = 2480;
                        break;
                    default:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                }
                Write_Register_Fractional_Calc_Ch(ch);
                for (int j = 0; j < 2; j++)
                {
                    if (j == 0)
                    {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                    }
                    else
                    {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                    }
                    Write_Register_Tx_Tone_Send(true);
                    SpectrumAnalyzer.Write("FREQ:CENT " + d_freq_MHz + " MHZ");
                    System.Threading.Thread.Sleep(400);
                    d_cur_mA = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000.0;
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 2), (y_pos + cnt), d_cur_mA.ToString("F2"));
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < d_lsl) || (d_power_dbm > d_usl))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_21");
                            result = false;
                        }
                    }
                    if ((d_cur_mA < 8) || (d_cur_mA > 13))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_28");
                            result = false;
                        }
                    }

                    SpectrumAnalyzer.Write("FREQ:CENT " + (d_freq_MHz * 2) + " MHZ");
                    System.Threading.Thread.Sleep(50);
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 1), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < (d_lsl - 68)) || (d_power_dbm > (d_usl - 30)))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_22");
                            result = false;
                        }
                    }
                }
                Write_Register_Tx_Tone_Send(false);
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");
            Write_Register_Fractional_Calc_Ch(0);
            PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");

            return result;
        }

        private bool Run_Measure_INTB(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val_H;
            double d_val_L;

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            d_val_H = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
            I2C.GPIOs[3].Direction = GPIO_Direction.Output; // AD7
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].State = GPIO_State.Low;
            System.Threading.Thread.Sleep(115);
            I2C.GPIOs[3].State = GPIO_State.High;
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].Direction = GPIO_Direction.Input;
            System.Threading.Thread.Sleep(100);
            d_val_L = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));

            if (d_val_H > 0.8 && d_val_L < 0.5)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");
            }
            else
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL");
                return false;
            }

            return result;
        }

        private bool Run_Measure_TAMPER(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27 TAMPER dectec : pass = 0 fail = 1

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

#if !(POWER_SUPPLY_NEW)
            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
            return true;
#endif

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.9,(@2)");    // TAMPER 0.9 set
            PowerSupply0.Write("OUTP ON,(@2)");     // TAMPER 0.9 set            
#else
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 0.9");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_H");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.3,(@2)");   //TAMPER 0.3 set
#else
            PowerSupply0.Write("VOLT 0.3");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_L");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.6,(@2)");   //TAMPER 0.6 set
#else
            PowerSupply0.Write("VOLT 0.6");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x00)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_M");
                result = false;
                return result;
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("OUTP OFF,(@2)");
#else
            PowerSupply0.Write("VOLT 0.0");
            PowerSupply0.Write("INST:NSEL 2");
#endif
            return result;
        }

        private bool Run_Measure_BLE_Packet(bool start, int cnt, int x_pos, int y_pos, bool result, byte[] aon_ble)
        {
            int len;
            byte b;
            byte[] r_data = new byte[37];

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            Serial.RcvQueue.Clear();
            SendCommand("ook 30");
            System.Threading.Thread.Sleep(150);

            len = Serial.RcvQueue.Count;
            Console.WriteLine("UART len : " + len);
            SendCommand("mode ook");
            for (int i = 0; i < 11; i++)
            {
                b = Serial.RcvQueue.Get();
            }

            for (int i = 0; i < 36; i++)
            {
                b = Serial.RcvQueue.Get();

                if (b > 0x60)
                {
                    r_data[i] = (byte)((b - 87) << 4); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] = (byte)((b - 55) << 4); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] = (byte)((b - 48) << 4); // '0' = 0x30 = 48
                }

                b = Serial.RcvQueue.Get();
                if (b > 0x60)
                {
                    r_data[i] += (byte)((b - 87)); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] += (byte)((b - 55)); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] += (byte)((b - 48)); // '0' = 0x30 = 48
                }

                //b = UartRcvQueue.GetByte(); // space

                Console.Write("{0:X}-", r_data[i]);
                if (r_data[i] == aon_ble[i])
                {
                    continue;
                }
                else if ((i == 32) || (i == 33)) // BL, TP
                {
                    if (r_data[i] > 150)
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if (i == 35) // TAMPER
                {
                    if (!((r_data[i] == 0x00) || (r_data[i] == 0x80)))
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if ((i < 3) || (i > 9) && (i < 12))
                {
                    continue;
                }
                else
                {
                    Console.Write("\r\nFail!!{0} read : {1:X}, write : {2:X}\r\n", i, r_data[i], aon_ble[i]);
                    Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), i.ToString());
                    result = false;
                    return result;
                }

                if ((b == '\r') || (b == '\n'))
                {
                    break;
                }
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "P");
            return result;
        }

        private void Test_Good_Chip_Sorting_Rev3(int start_cnt)
        {
            int cnt = 0, pass = 0, fail = 0;
            double d_val;
            int x_pos = 2, y_pos = 12;
            bool result;
            uint mac_code = 0x000000;
            bool OTP_W_Flag = false;

            byte[] aon_ble = { 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0x1E, 0xFF, 0x6D, 0x0B, 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

            Check_Instrument();

            if (I2C.IsOpen == false) // Check I2C connection
            {
                MessageBox.Show("Check I2C");
                return;
            }

            MessageBox.Show("장비 설정을 확인해주세요.\r\n\r\n1.N6705B Power Supply\r\n  - VBAT\r\n  - TEST_IN_OUT(N6705B)\r\n\r\n" +
                "2.SpectrumAnalyzer\r\n  - TX\r\n\r\n3.Digital Multimeter\r\n  - Current mode : VBAT\r\n  - Voltage mode : ALLDO / MLDO / TEST_IN_OUT\r\n\r\n" +
                "4.UM232H\r\n  - AC0_GPIOH0 : I2C EN\r\n  - AD6_GPIOL2 : RX/TX Switch\r\n  - AD7_GPIOL3 : INTB");

            cnt = start_cnt - 1;
            pass = cnt;

            while (true)
            {
                I2C.GPIOs[3].Direction = GPIO_Direction.Input;   // AD7_GPIOL3(INTB)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].Direction = GPIO_Direction.Output;   // AD6_GPIOL2(TRX S/W)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].Direction = GPIO_Direction.Output;   // AC0_GPIOH0(Level Shifter EN)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);

                // Power off
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 0.0,(@3)");
                PowerSupply0.Write("OUTP ON,(@3)");
                PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");
#else
                PowerSupply0.Write("INST:NSEL 1");
                PowerSupply0.Write("VOLT 0.0");
                PowerSupply0.Write("OUTP ON");
#endif
                SendCommand("mode ook\n");
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n\r\nTest\t: " + cnt + "\r\nPass\t: " + pass + "\r\nFail\t: " + fail
                                                        , Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    cnt++;
                    fail++;
                    Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), cnt.ToString());
                    result = true;
                }
                else
                {
                    return;
                }

                DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:AC?"); // ALLDO
                // Power on
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 3.3,(@3)");
                System.Threading.Thread.Sleep(200);
                PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                PowerSupply0.Write("VOLT 3.3");
                System.Threading.Thread.Sleep(500);
                PowerSupply0.Write("VOLT 2.5");
#endif
                // Check MLDO Voltage

                System.Threading.Thread.Sleep(800);
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_val.ToString("F3"));
                if (d_val > 0.2)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                    continue;
                }

                // Seq1 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#else
                d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                d_val = 600;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
#endif
                if ((d_val < 500) || (d_val > 850))
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                    continue;
                }

                I2C.GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                // Wake-up
                WakeUp_I2C();
                System.Threading.Thread.Sleep(5);

                // READ_DIVICE_ID
                if (Check_Revision_Information(0x5F) == false) // N1 = 0x5D, N1B = 0x5E, N1C = 0x5F
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "F");
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_24");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "P");
                }

                mac_code = Read_Mac_Address();
                aon_ble[0] = (byte)(mac_code & 0xff);
                aon_ble[1] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[2] = (byte)((mac_code >> 16) & 0xff);
                aon_ble[10] = (byte)(mac_code & 0xff);
                aon_ble[11] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[12] = (byte)((mac_code >> 16) & 0xff);
                Parent.xlMgr.Cell.Write((x_pos + 81), (y_pos + cnt), (aon_ble[5].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 82), (y_pos + cnt), (aon_ble[4].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 83), (y_pos + cnt), (aon_ble[3].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 84), (y_pos + cnt), (aon_ble[2].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 85), (y_pos + cnt), (aon_ble[1].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 86), (y_pos + cnt), (aon_ble[0].ToString("X")));

                if (Check_OTP_VALID(true, 21) == true)
                {
                    goto CAL_SKIP;
                }

                // Seq1 5.OTP read
                if (Read_FF_NVM(false) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_10");
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "F");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "Skip");
                }

                Write_Register_AON_Fix_Value();

#if true // For Test
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
                Parent.xlMgr.Cell.Write(3, (1 + cnt), (double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Cell.Write(4, (1 + cnt), (double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
#endif
                // Seq2 1.BGR Trim
                if (Run_Cal_BGR(true, cnt, (x_pos + 6), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_4");
                    continue;
                }
                // Seq2 2.ALON LDO Trim
                if (Run_Cal_ALLDO(true, cnt, (x_pos + 8), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_5");
                    continue;
                }
                // Seq2 3.MLDO Trim
                if (Run_Cal_MLDO(true, cnt, (x_pos + 11), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_6");
                    continue;
                }
                // Seq2 4.32K RCOSC Trim
                if (Run_Cal_32K_RCOSC(true, cnt, (x_pos + 14), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_7");
                    continue;
                }
                // Seq2 5.Temp Sensor Trim
                if (Run_Cal_Temp_Sensor(true, cnt, (x_pos + 18), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_8");
                    continue;
                }
                // Seq2 5.Voltage Scaler Trim
                if (Run_Cal_Voltage_Scaler(true, cnt, (x_pos + 22), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_25");
                    continue;
                }
                // Seq2 6.32M X-tal Load Cap Trim
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                if (Run_Cal_32M_XTAL_Load_Cap(true, cnt, (x_pos + 26), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_9");
                    continue;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);
                // Seq2 7.PLL 2PM Trim
                if (Run_Cal_PLL_2PM(true, cnt, (x_pos + 29), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_20");
                    continue;
                }
                // Seq2 8.OTP Trim data write
                // Set AON_REG(BLE) For OTP
                Set_AON_REG_For_BLE(aon_ble);

                Run_Write_OTP_With_OOK(OTP_W_Flag, 0);
                // Run_Write_OTP_With_OOK(OTP_W_Flag, 15);
                Run_Write_OTP_With_OOK(OTP_W_Flag, 21);
                d_val = Run_Verify_OTP_With_BIST(OTP_W_Flag, 0x208001);
                if (d_val == 65535)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "Skip");
                }
                else if (d_val != 0x400)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_11");
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), d_val.ToString());
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "P");
                }

            CAL_SKIP:
                if (OTP_W_Flag) // Reset
                {
                    I2C.GPIOs[4].State = GPIO_State.High;
                    System.Threading.Thread.Sleep(10);
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 0.0,(@3)");
#else
                    PowerSupply0.Write("VOLT 0.0");
#endif
                    System.Threading.Thread.Sleep(500);

#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 3.3,(@3)");
                    System.Threading.Thread.Sleep(200);
                    PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                    PowerSupply0.Write("VOLT 3.3");
                    System.Threading.Thread.Sleep(500);
                    PowerSupply0.Write("VOLT 2.5");
#endif
                    // Check MLDO Voltage                    
                    System.Threading.Thread.Sleep(800);
                    d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), d_val.ToString("F3"));
                    if (d_val > 0.2)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                        continue;
                    }

                    // Seq2 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                    d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), (d_val).ToString("F3"));
#else
                    d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                    d_val = 600;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
#endif
                    if ((d_val < 500) || (d_val > 850))
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                        continue;
                    }

                    I2C.GPIOs[4].State = GPIO_State.Low;
                    System.Threading.Thread.Sleep(10);
                    // Wake-up
                    WakeUp_I2C();
                    System.Threading.Thread.Sleep(5);
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), "Skip");
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
                }

                // Seq3 1.Initial Test
                if (Run_Measure_Initial(true, cnt, (x_pos + 33), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 2.ADC Test
                if (Run_Measure_ADC(true, cnt, (x_pos + 45), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                // Seq3 4.VCO Range Test
                if (Run_Measure_VCO_Range(true, cnt, (x_pos + 57), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 4. Tx output power and Harmonic
                if (Run_Measure_Tx_Power_Harmonic(true, cnt, (x_pos + 60), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);

                // INTB Test
                if (Run_Measure_INTB(true, cnt, (x_pos + 79), y_pos, result) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_27");
                    }
                    result = false;
                }
                // TAMPER Test
                if (Run_Measure_TAMPER(true, cnt, (x_pos + 78), y_pos, result) == false)
                {
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("OUTP OFF,(@2)");
#endif
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_26");
                    }
                    result = false;
                }

                // BLE Test
                if (Run_Measure_BLE_Packet(true, cnt, (x_pos + 80), y_pos, result, aon_ble) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_23");
                    }
                    result = false;
                }

                if (result == true)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "PASS");
                    fail--;
                    pass++;
                }
            }
        }

        private void DUT_TempSensor_Test()
        {
            byte Addr;
            byte[] RcvData = new byte[2];

            Parent.xlMgr.Sheet.Add(DateTime.Now.ToString("MMddHHmmss_") + "DUT_TS");
            Parent.xlMgr.Cell.Write(1, 1, "No.");
            Parent.xlMgr.Cell.Write(2, 1, "ADC");
            Parent.xlMgr.Cell.Write(3, 1, "Temp");

            I2C.GPIOs[2].Direction = GPIO_Direction.Output;
            System.Threading.Thread.Sleep(10);

            Addr = 0x00;
#if false
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(200); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Parent.xlMgr.Cell.Write(1, 2 + i, i.ToString());
                Parent.xlMgr.Cell.Write(2, 2 + i, ((RcvData[0] << 8) | RcvData[1]).ToString());
                Parent.xlMgr.Cell.Write(3, 2 + i, (((RcvData[0] << 8) | RcvData[1]) * 0.0078125).ToString());

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(500);
            }
#else
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(300); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Console.WriteLine("ADC : {0}\tTemp : {1}", (RcvData[0] << 8) | RcvData[1], ((RcvData[0] << 8) | RcvData[1]) * 0.0078125);

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(100);
            }
#endif
        }
        #endregion

        #region Function for DTM
        private void SetBleDTM_AonControlReg()
        {
            //WriteRegister(0x33, 0x17);      //AON_REG_CTRL_7
            WriteRegister(0x28, 0x3F);       //ADC_EOC_SKIP_TEMP
            WriteRegister(0x29, 0x3F);       //ADC_EOC_SKIP_VOLT
            //WriteRegister(0x36, 0xFF);     // pin pull-up en
            WriteRegister(0x2B, (1 << 5) | (0 << 4) | (1 << 3));       // BLE Tx mode set, byte swap/order
            //WriteRegister(0x38, 0x00);        //w_BLE_TX_OFF
            //WriteRegister(0x39, 0x40);        //BLE_BSWAP
            WriteRegister(0x3A, 0x4A);      //w_MIC_INIT[7:0]
            WriteRegister(0x2D, 0x00);      //MLDO/XTAL Wait
        }

        private void SetBleDTM_AonControlReg_OTP()
        {
            WriteRegister(0x26, 0x60);  // WUR FIFO End , w_WUR_WR_FORCE	w_WUR_THR[5:0]
            WriteRegister(0x2B, 0x8);  // w_FIXED_ADV_DELAY w_BLE_TX_OFF w_BLE_TX_MODE w_BIST_WDATA_SWAP[1:0]	BLE_BSWAP

            WriteRegister(0x4F, 0x7);   //CT_RES32_FINE = 7 (Bandwidth change)
            //WriteRegister(0x64, 0xD2);  //0x1,w_TX_SEL_SELECT[1:0] 0x10,w_PLL_PM_GAIN[4:0] 0x12 TX_SEL = CoraseLock & PLL_PM_GIAN = 18;
            //WriteRegister(0x65, (0x2<<6)|(0x12));  //O_VCO_CBANK[9:8](default), PLL_PM_GAIN=0x12
        }

        private void SetBleDTM_AonAnalogReg()
        {
            //WriteRegister(0x62, 0x66);      //r3 ONLY
            WriteRegister(0x67, 0x66);      //
            //WriteRegister(0x61, 0x33);
            WriteRegister(0x50, (0 << 6) | (2 << 4) | (3 << 0));       // TEST_ADC_MUX_SEL, O_ULP_LDO_LV_CONT, O_ULP_LDO_CONT[3:0]

            WriteRegister(0x53, 0xC7);      //FC_LGA
            WriteRegister(0x54, 0xBC);
            WriteRegister(0x3C, 0x10);
            WriteRegister(0x3D, 0x12);
            WriteRegister(0x3E, 0x00);      //R3 ONLY
            WriteRegister(0x3F, 0x72);
            WriteRegister(0x40, 0x82);
            WriteRegister(0x41, 0x80);
            WriteRegister(0x42, 0x00);
            WriteRegister(0x43, 0xF2);
            WriteRegister(0x44, 0xFF);
            WriteRegister(0x45, 0x02);
            WriteRegister(0x46, 0xFF);
            WriteRegister(0x47, 0x42);
            WriteRegister(0x48, 0x62);
            WriteRegister(0x49, 0xFF);
            WriteRegister(0x4A, 0x3F);
            WriteRegister(0x4B, 0x16);
            WriteRegister(0x4C, 0x00);
            WriteRegister(0x4D, 0x29);
            WriteRegister(0x4E, 0xB2);
            WriteRegister(0x4F, 0x61);
            WriteRegister(0x50, 0xC5);
            WriteRegister(0x51, 0x2F);      //R3 ONLY
            WriteRegister(0x52, 0x22);

            WriteRegister(0x55, 0xCE);
            WriteRegister(0x56, 0x00);  //R3 ONLY DYNAMIC CONTROL FOR BLE LINK
            //WriteRegister(0x57, 0x07);  //TX_BUF, DA_PEN, PRE_DA_PEN, DA_GAIN R3 ONLY DYNIMIC CONTROL FOR BLE LINK
            WriteRegister(0x5B, (2 << 3) | (6 << 0));

            WriteRegister(0x58, 0x33);  //2PM_CAL R3 ONLY  
            WriteRegister(0x59, 0x07);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
            WriteRegister(0x5A, 0x00);  //EXT_DA_GAIN_LSB_BIT
            WriteRegister(0x5B, 0xE2);  //EXT_DA_GAIN_SET = 1; (BLE_LINK) R3                                                       
            WriteRegister(0x5C, 0x00);  //TX_SEL(W),PRE_DA_PEN(W), TX_BUF_PEN(W), DA_PEN(W) R3 ONLY                                         
            WriteRegister(0x5D, 0x60);  //PM_RESETB(W), PLL_PEN(W)
            WriteRegister(0x5E, 0x20);
            WriteRegister(0x5F, 0x45);  //R3 ONLY
            WriteRegister(0x60, 0x08);

            WriteRegister(0x63, 0x00);
            WriteRegister(0x64, 0xB0);  //TX_SEL = DATA_ST

            //FOR R3 ONLY
            WriteRegister(0x65, 0x34);  //DA_PEN = DATA_ST
            WriteRegister(0x66, 0x12);
            WriteRegister(0x67, 0x40);
        }

        private void Set_PLLRESET(int delay)
        {
#if false
            WriteRegister(0x59, 0x1);   //O_PLL_2PM_CAL_HOLD
            System.Threading.Thread.Sleep(delay);
            WriteRegister(0x59, 0x7);   //O_PLL_PEN	O_PM_RESETB	
            System.Threading.Thread.Sleep(delay);
#else
            WriteRegister(0x5D, 0x6);   //O_PLL_PEN	O_PM_RESETB	
            System.Threading.Thread.Sleep(delay);
#endif
        }

        private void Set_BLE_LINK(uint addr, uint data)
        {
            byte[] Addrs = new byte[8];

            Addrs[0] = (byte)(addr & 0xff);
            Addrs[1] = (byte)(addr >> 8 & 0xff);
            Addrs[2] = (byte)(addr >> 16 & 0xff);
            Addrs[3] = 0x00;
            Addrs[4] = (byte)(data & 0xff);
            Addrs[5] = (byte)(data >> 8 & 0xff);
            Addrs[6] = (byte)(data >> 16 & 0xff);
            Addrs[7] = (byte)(data >> 24 & 0xff);

            I2C.WriteBytes(Addrs, 8, true);
        }

        private void Run_BLE_DTM_MODE(uint PayLoadMode)
        {
            uint PmGain = 0;

            PmGain = ReadRegister(0x5C);            // AON_REG_ANA_28(PLL)

            if (PmGain == 0x42)                     // Did not Cal. (Default value)
            {
                // Please update below functions if you needed.
                //SetBleDTM_AonControlReg();
                //System.Threading.Thread.Sleep(2);
                //SetBleDTM_AonAnalogReg();
                //System.Threading.Thread.Sleep(2);
            }
            else
            {
                SetBleDTM_AonControlReg_OTP();
                System.Threading.Thread.Sleep(2);
            }

            Set_BLE_LINK(0x058000, 0x80);   //reset
            Set_BLE_LINK(0x058000, 0x48);   //stop dtm
            Set_BLE_LINK(0x0580d8, 0x78);   //bb_clk_freq_minus_1 
            Set_BLE_LINK(0x0581b4, 0x72);   //RADIO_DATA host datain [15:8] / host dataout[15:8]
            Set_BLE_LINK(0x0581ac, 0x9043); //RADIO_ACCESS/RADIO_CNTRL 
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9400);
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9401);
            Set_BLE_LINK(0x0580d8, 0x78);
            Set_BLE_LINK(0x058190, 0x9280);
            Set_BLE_LINK(0x058198, 0x5b5b);
            Set_BLE_LINK(0x0581b0, 0x03);
            Set_BLE_LINK(0x0581b8, 0x05);

            uint pmode = ((PayLoadMode * 128) + 0);
            Set_BLE_LINK(0x058170, pmode);
            Set_BLE_LINK(0x05819c, (uint)DTM_PayLoadLength); //packet length
            Set_BLE_LINK(0x040014, (0x280 + (uint)DTM_Channel)); //Ch Sel Mode & CH 0;
            Set_PLLRESET(2);
            Set_BLE_LINK(0x058000, 0x46); //start dtm
        }

        private void Stop_BLE_DTM_MODE()
        {
            Set_BLE_LINK(0x058000, 0x48);
            WriteRegister(0x59, 0x03);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
        }

        private uint Read_ADC_Result_Multi(bool temp_first, uint eoc_skip_temp, uint eoc_skip_volt, uint avg_cnt, JLcLib.Custom.I2C I2c_Sel)
        {
            byte[] Addrs = new byte[8];
            byte[] Bytes = new byte[4];

            // Set w_TEST_SELECT
            Addrs[0] = 0x04;
            Addrs[1] = 0x00;
            Addrs[2] = 0x03;
            Addrs[3] = 0x00;
            Addrs[4] = 0x00;
            Addrs[5] = 0x00;
            Addrs[6] = 0x00;
#if false // w_TEST_SELECT = 1
            if(temp_first)
            {
                Addrs[7] = 0x80; // temp -> volt
            }
            else
            {
                Addrs[7] = 0xC0; // volt -> temp
            }
#else // w_TEST_SELECT = 0
            if (temp_first)
            {
                Addrs[7] = 0x00; // temp -> volt
            }
            else
            {
                Addrs[7] = 0x40; // volt -> temp
            }
#endif
            I2c_Sel.WriteBytes(Addrs, 8, true);

            // Set ADC Register (I_PEN = 0, w_TEST_START = 0)
            Addrs[0] = 0x00;
            Addrs[1] = 0x00;
            Addrs[2] = 0x03;
            Addrs[3] = 0x00;
            Addrs[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x3C); // OFFSET = 0xF
            Addrs[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            Addrs[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            Addrs[7] = (byte)(((avg_cnt & 0x1f) << 1) | 0x00);
            I2c_Sel.WriteBytes(Addrs, 8, true);

            // Enable ADC (I_PEN = 1, w_TEST_START = 0)
            Addrs[0] = 0x00;
            Addrs[1] = 0x00;
            Addrs[2] = 0x03;
            Addrs[3] = 0x00;
            Addrs[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x3C); // OFFSET = 0xF
            Addrs[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            Addrs[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            Addrs[7] = (byte)(((avg_cnt & 0x1f) << 1) | 0x80);
            I2c_Sel.WriteBytes(Addrs, 8, true);
            System.Threading.Thread.Sleep(10);

            Addrs[0] = 0x0C;
            Addrs[1] = 0x00;
            Addrs[2] = 0x03;
            Addrs[3] = 0x00;
            I2C.WriteBytes(Addrs, 4, false);
            Bytes = I2C.ReadBytes(Bytes.Length);

            return (uint)(((Bytes[1] << 8) | Bytes[0]) & 0xffff);
        }

        private void TestInOut_ADC_In(bool on)
        {
            RegisterItem TEST_ADC_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_ADC_MUX_SEL");      // 0x50
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[9]");               // 0x5E
            RegisterItem O_TEST_BGR_MUX_SEL = Parent.RegMgr.GetRegisterItem("O_TEST_BGR_MUX_SEL");  // 0x5F
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");        // 0x5F
            RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");             // 0x61

            if (on == true)
            {
                TEST_ADC_MUX_SEL.Read();
                TEST_ADC_MUX_SEL.Value = 1;
                TEST_ADC_MUX_SEL.Write();

                ITEST_CONT.Read();
                ITEST_CONT.Value = 1;
                ITEST_CONT.Write();

                O_TEST_BGR_MUX_SEL.Read();
                O_TEST_BGR_MUX_SEL.Value = 0;
                O_TEST_BGR_MUX_SEL.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BGR_BUF_EN.Write();

                O_TEST_CON.Read();
                O_TEST_CON.Value = 2;
                O_TEST_CON.Write();
            }
            else
            {
                TEST_ADC_MUX_SEL.Read();
                TEST_ADC_MUX_SEL.Value = 0;
                TEST_ADC_MUX_SEL.Write();

                ITEST_CONT.Read();
                ITEST_CONT.Value = 0;
                ITEST_CONT.Write();

                O_TEST_BGR_MUX_SEL.Read();
                O_TEST_BGR_MUX_SEL.Value = 0;
                O_TEST_BGR_MUX_SEL.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BGR_BUF_EN.Write();

                O_TEST_CON.Read();
                O_TEST_CON.Value = 0;
                O_TEST_CON.Write();
            }

        }

        private void TestInOut_ADC_In_Multi(bool on, JLcLib.Custom.I2C I2C_Sel)
        {
            uint rData;

            RegisterItem TEST_ADC_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_ADC_MUX_SEL");      // 0x50
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[9]");               // 0x5E
            RegisterItem O_TEST_BGR_MUX_SEL = Parent.RegMgr.GetRegisterItem("O_TEST_BGR_MUX_SEL");  // 0x4A
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");        // 0x5F
            RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");             // 0x61

            if (on == true)
            {
                //O_TEST_BGR_MUX_SEL.Read();
                //O_TEST_BGR_MUX_SEL.Value = 0;
                //O_TEST_BGR_MUX_SEL.Write();
                rData = ReadRegister_Multi(0x4A, I2C_Sel);
                WriteRegister_Multi(0x4A, rData & 0xDF, I2C_Sel);

                //TEST_ADC_MUX_SEL.Read();
                //TEST_ADC_MUX_SEL.Value = 1;
                //TEST_ADC_MUX_SEL.Write();
                rData = ReadRegister_Multi(0x50, I2C_Sel);
                WriteRegister_Multi(0x50, rData | 0x40, I2C_Sel);

                //ITEST_CONT.Read();
                //ITEST_CONT.Value = 1;
                //ITEST_CONT.Write();
                rData = ReadRegister_Multi(0x5E, I2C_Sel);
                WriteRegister_Multi(0x5E, rData | 0x01, I2C_Sel);

                //TEST_BGR_BUF_EN.Read();
                //TEST_BGR_BUF_EN.Value = 1;
                //TEST_BGR_BUF_EN.Write();
                rData = ReadRegister_Multi(0x5F, I2C_Sel);
                WriteRegister_Multi(0x5F, rData | 0x04, I2C_Sel);

                //O_TEST_CON.Read();
                //O_TEST_CON.Value = 2;
                //O_TEST_CON.Write();
                rData = ReadRegister_Multi(0x61, I2C_Sel);
                WriteRegister_Multi(0x61, rData | 0x02, I2C_Sel);
            }
            else
            {
                //TEST_ADC_MUX_SEL.Read();
                //TEST_ADC_MUX_SEL.Value = 0;
                //TEST_ADC_MUX_SEL.Write();
                rData = ReadRegister_Multi(0x50, I2C_Sel);
                WriteRegister_Multi(0x50, rData & 0xBF, I2C_Sel);

                //ITEST_CONT.Read();
                //ITEST_CONT.Value = 0;
                //ITEST_CONT.Write();
                rData = ReadRegister_Multi(0x5E, I2C_Sel);
                WriteRegister_Multi(0x5E, rData & 0xFE, I2C_Sel);

                //TEST_BGR_BUF_EN.Read();
                //TEST_BGR_BUF_EN.Value = 0;
                //TEST_BGR_BUF_EN.Write();
                rData = ReadRegister_Multi(0x5F, I2C_Sel);
                WriteRegister_Multi(0x5F, rData & 0xFB, I2C_Sel);

                //O_TEST_CON.Read();
                //O_TEST_CON.Value = 0;
                //O_TEST_CON.Write();
                rData = ReadRegister_Multi(0x61, I2C_Sel);
                WriteRegister_Multi(0x61, rData & 0xFD, I2C_Sel);
            }

        }

        private void Cal_PMU_SweepCal(int num)
        {
            // AL BGR(300mV), ALLDO(810mV), MLDO(1V), RCOSC(32.768kHz)
            double[] dVal = new double[5];
            double d_alldo_mv, d_mldo_mv, d_volt_mv;
            double d_diff_mv, d_target_bgr_mv = 300, d_target_alldo_mv = 810, d_target_mlod_mv = 1000;
            double d_freq_khz, d_diff_khz;
            double d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;
            uint alldo_val, mldo_val, bgr_val, alc = 0, mlc = 0, ul = 0;
            int y_pos = 2;
            string sTempVal;
            bool bRetrunTemp = true;
            uint rev_id;

            // R5
            // RCOSC
            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");               // 0x51
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");              // 0x52
            // AL BGR
            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");         // 0x57
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL[5:2]");             // 0x67
            // ALLDO
            RegisterItem ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");         // 0x50
            RegisterItem ULP_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[1:0]");    // 0x50
            // MLDO
            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");         // 0x57
            RegisterItem PMU_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1:0]");    // 0x67
            // Rev_ID
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                     // 0x6C

            I_DEVID.Read();
            rev_id = I_DEVID.Value;
            if (rev_id == 180)
            {
                Parent.xlMgr.Sheet.Add("R5V1_" + num.ToString("X2") + "_CAL_LDO");
            }

            else if (rev_id == 212)
            {
                Parent.xlMgr.Sheet.Add("R5V2_" + num.ToString("X2") + "_CAL_LDO");
            }

            else return;

            Parent.xlMgr.Cell.Write(1, 1, "O_ULP_BGR_CONT[3:0]");
            Parent.xlMgr.Cell.Write(2, 1, "BGR(mV)");
            Parent.xlMgr.Cell.Write(4, 1, "O_ULP_LDO_LV_CONT[2:0]");
            Parent.xlMgr.Cell.Write(5, 1, "O_ULP_LDO_CONT[3:0]");
            Parent.xlMgr.Cell.Write(6, 1, "ALLDO(mV)");
            Parent.xlMgr.Cell.Write(8, 1, "O_PMU_MLDO_Coarse[1:0]");
            Parent.xlMgr.Cell.Write(9, 1, "O_PMU_LDO_CONT[3:0]");
            Parent.xlMgr.Cell.Write(10, 1, "MLDO(mV)");
            Parent.xlMgr.Cell.Write(12, 1, "O_RTC_SCKF[5:0]");
            Parent.xlMgr.Cell.Write(13, 1, "RCOSC(kHz)");

            Check_Instrument();
            PowerSupply0.Write("VOLT 2.5,(@2)");
            System.Threading.Thread.Sleep(1000);

            BGR_TC_CTRL.Read();
            BGR_TC_CTRL.Value = 9;
            BGR_TC_CTRL.Write();

#if true   // TempChamber 23.5℃ Set
            sTempVal = TempChamber.WriteAndReadString("01,TEMP,S23.5");
            System.Threading.Thread.Sleep(1000);
            sTempVal = TempChamber.WriteAndReadString("TEMP?");
            System.Threading.Thread.Sleep(100);
            Console.WriteLine("Run Chamber : " + 23.5);
            sTempVal = TempChamber.WriteAndReadString("TEMP?");
            System.Threading.Thread.Sleep(100);
            bRetrunTemp = true;
            while (bRetrunTemp)
            {
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                for (int split = 0; split < 4; split++)
                    dVal[split] = double.Parse(ArrBuf[split]);

                double dTHVal = 0.1;

                if (dVal[0] + dTHVal == 23.5 || dVal[0] - dTHVal == 23.5 || dVal[0] == 23.5)
                {
                    bRetrunTemp = false;
                    Log.WriteLine("Done SetTemp!");
                }
                else
                    bRetrunTemp = true;
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                System.Threading.Thread.Sleep(1000 * 5);
            }
            System.Threading.Thread.Sleep(1000 * 10);
            Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
            // CAL BGR
            Set_TestInOut_For_BGR(true);

            ULP_BGR_CONT.Read();
            bgr_val = 0;
            d_diff_mv = 5000;
            System.Threading.Thread.Sleep(500);
            for (uint val = 0; val < 16; val++)
            {
                System.Threading.Thread.Sleep(500);
                ULP_BGR_CONT.Value = val;
                ULP_BGR_CONT.Write();

                System.Threading.Thread.Sleep(500);
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write(1, (2 + (int)val), val.ToString());
                Parent.xlMgr.Cell.Write(2, (2 + (int)val), d_volt_mv.ToString("F3"));
                if (Math.Abs(d_volt_mv - d_target_bgr_mv) < d_diff_mv)
                {
                    d_diff_mv = Math.Abs(d_volt_mv - d_target_bgr_mv);
                    bgr_val = val;
                }
            }
            ULP_BGR_CONT.Value = bgr_val;
            ULP_BGR_CONT.Write();
            System.Threading.Thread.Sleep(500);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(1, (2 + 16) + 1, bgr_val.ToString());
            Parent.xlMgr.Cell.Write(2, (2 + 16) + 1, d_volt_mv.ToString("F3"));
            Set_TestInOut_For_BGR(false);

            BGR_TC_CTRL.Read();
            Parent.xlMgr.Cell.Write(1, (2 + 16) + 2, "BGR_TC_CTRL<5:2>");
            Parent.xlMgr.Cell.Write(2, (2 + 16) + 2, (BGR_TC_CTRL.Value).ToString());


            // Cal ALLDO
            ULP_LDO_CONT.Read();
            ULP_LDO_Coarse.Read();
            d_diff_mv = 5000.0;
            alldo_val = 0;
            for (uint j = 0; j < 4; j++)
            {
                if (j == 0)
                    ul = 1;
                else if (j == 1)
                    ul = 3;
                else if (j == 2)
                    ul = 0;
                else ul = 2;

                System.Threading.Thread.Sleep(500);
                ULP_LDO_Coarse.Value = ul;
                ULP_LDO_Coarse.Write();

                for (uint i = 0; i < 16; i++)
                {
                    System.Threading.Thread.Sleep(500);
                    Parent.xlMgr.Cell.Write(4, ((int)ul * 16 + (int)i + 2), ul.ToString());
                    ULP_LDO_CONT.Value = i;
                    ULP_LDO_CONT.Write();
                    Parent.xlMgr.Cell.Write(5, ((int)ul * 16 + (int)i + 2), (i).ToString());
                    System.Threading.Thread.Sleep(100);
                    d_alldo_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(6, ((int)ul * 16 + (int)i + 2), d_alldo_mv.ToString("F3"));
                    if (Math.Abs(d_alldo_mv - d_target_alldo_mv) < d_diff_mv)
                    {
                        d_diff_mv = Math.Abs(d_alldo_mv - d_target_alldo_mv);
                        alldo_val = i;
                        alc = ul;
                    }
                }
            }

            ULP_LDO_CONT.Value = alldo_val;
            ULP_LDO_CONT.Write();
            ULP_LDO_Coarse.Value = alc;
            ULP_LDO_Coarse.Write();
            Parent.xlMgr.Cell.Write(4, (3 * 16 + 16 + 2) + 1, alc.ToString());
            Parent.xlMgr.Cell.Write(5, (3 * 16 + 16 + 2) + 1, alldo_val.ToString());
            System.Threading.Thread.Sleep(100);
            d_alldo_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(6, (3 * 16 + 16 + 2) + 1, d_alldo_mv.ToString("F3"));

            // Cal MLDO
            PMU_LDO_Coarse.Read();
            PMU_LDO_CONT.Read();
            d_diff_mv = 5000.0;
            mldo_val = 0;
            for (uint j = 0; j < 4; j++)
            {
                System.Threading.Thread.Sleep(100);
                PMU_LDO_Coarse.Value = j;
                PMU_LDO_Coarse.Write();
                for (uint i = 0; i < 16; i++)
                {
                    System.Threading.Thread.Sleep(100);
                    Parent.xlMgr.Cell.Write(8, ((int)j * 16 + (int)i + 2), j.ToString());
                    PMU_LDO_CONT.Value = i;
                    PMU_LDO_CONT.Write();
                    Parent.xlMgr.Cell.Write(9, ((int)j * 16 + (int)i + 2), (i).ToString());
                    System.Threading.Thread.Sleep(100);
                    d_mldo_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(10, ((int)j * 16 + (int)i + 2), d_mldo_mv.ToString("F3"));
                    if (Math.Abs(d_mldo_mv - d_target_mlod_mv) < d_diff_mv)
                    {
                        d_diff_mv = Math.Abs(d_mldo_mv - d_target_mlod_mv);
                        mldo_val = i;
                        mlc = j;
                    }
                }
            }
            PMU_LDO_CONT.Value = mldo_val;
            PMU_LDO_CONT.Write();
            PMU_LDO_Coarse.Value = mlc;
            PMU_LDO_Coarse.Write();
            Parent.xlMgr.Cell.Write(8, (3 * 16 + 16 + 2) + 1, mlc.ToString());
            Parent.xlMgr.Cell.Write(9, (3 * 16 + 16 + 2) + 1, mldo_val.ToString());
            System.Threading.Thread.Sleep(100);
            d_mldo_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(10, (3 * 16 + 16 + 2) + 1, d_mldo_mv.ToString("F3"));

            // Cal OSC
            Set_TestInOut_For_RCOSC(true);

            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            Parent.xlMgr.Cell.Write(13, y_pos, d_freq_khz.ToString("F3"));
            Parent.xlMgr.Cell.Write(12, y_pos++, osc_val_l.ToString());

            for (int val = 4; val >= 0; val--, y_pos++)
            {
                if (d_freq_khz > d_target_khz)
                {
                    osc_val_l += (uint)(1 << val);
                }
                else
                {
                    osc_val_l -= (uint)(1 << val);
                }
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();

                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

                Parent.xlMgr.Cell.Write(13, y_pos, d_freq_khz.ToString("F3"));
                Parent.xlMgr.Cell.Write(12, y_pos, osc_val_l.ToString());
            }

            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
                if (osc_val_l != 63) osc_val_l += 1;
                else
                if (osc_val_l != 0) osc_val_l -= 1;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            Parent.xlMgr.Cell.Write(13, y_pos, d_freq_khz.ToString("F3"));
            Parent.xlMgr.Cell.Write(12, y_pos++, osc_val_l.ToString());

            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            Parent.xlMgr.Cell.Write(13, y_pos + 1, d_freq_khz.ToString("F3"));
            Parent.xlMgr.Cell.Write(12, y_pos + 1, osc_val_l.ToString());

            Set_TestInOut_For_RCOSC(false);
        }

        private void LDO_Temp_Sweep(int num)
        {
            string sTempVal;
            bool bRetrunTemp = true;
            double[] dVal = new double[5];
            int x_pos = 1, y_pos = 1, y_Increment;
            string[] labels_n = { "ALLDO[mV]", "MLDO[mV]", "ALBGR[mV]", "RCOSC[kHz]" };
            double[] temps =
            {
                87.5,   85,     80,     75,     70,
                65,     60,     55,     50,     45,
                40,     35,     30,     25,     20,
                15,     10,     5,      0,      -5,
                -10,    -15,    -20,    -25,    -30,
                -35,    -40,    23.5
            };
            y_Increment = (temps.Length) + 3;

            Check_Instrument();
            // 87.5 to -40
            Parent.xlMgr.Sheet.Add("R5_" + num.ToString("X2") + "_LDOTempSweep");

            for (int i = 0; i < labels_n.Length; i++)
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + y_Increment * i), labels_n[i]);

            x_pos = 2;

            for (double k = 1.7; k < 3.4; k += 0.1)
            {
                for (int q = 0; q < labels_n.Length; q++)
                    Parent.xlMgr.Cell.Write(x_pos, y_pos + (y_Increment * q), k.ToString());
                x_pos++;
            }

            x_pos = 1;
            y_pos = y_Increment - 1;

            for (int i = 0; i < temps.Length; i++)
            {
#if true   // TempChamber temps[i] ℃ Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[i].ToString());
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + sTempVal);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == temps[i] || dVal[0] - dTHVal == temps[i] || dVal[0] == temps[i])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 60);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                for (double volt = 1.7; volt < 3.4; volt += 0.1)
                {
                    if (volt == 1.7)
                    {
                        PowerSupply0.Write("VOLT 2.5,(@2)");
                        System.Threading.Thread.Sleep(100);
                        PowerSupply0.Write("VOLT 1.7,(@2)");
                        System.Threading.Thread.Sleep(100);
                    }
                    else
                    {
                        PowerSupply0.Write("VOLT " + volt.ToString() + ",(@2)");
                        System.Threading.Thread.Sleep(200);
                    }
                    if (volt == 1.7)
                    {
                        for (int q = 0; q < labels_n.Length; q++)
                            Parent.xlMgr.Cell.Write(x_pos, (y_pos + y_Increment * q), temps[i].ToString());
                        x_pos++;
                    }
                    // Measure ALLDO
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + y_Increment * 0), (double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F3"));

                    // Measure MLDO
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + y_Increment * 1), (double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F3"));

                    // Measure BGR
                    Set_TestInOut_For_BGR(true);
                    System.Threading.Thread.Sleep(100);
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + y_Increment * 2), (double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F3"));
                    Set_TestInOut_For_BGR(false);

                    // Measure RCOSC
                    Set_TestInOut_For_RCOSC(true);
                    System.Threading.Thread.Sleep(100);
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + y_Increment * 3), (double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0).ToString("F3"));
                    Set_TestInOut_For_RCOSC(false);

                    x_pos++;
                }
                x_pos = 1;
                y_pos--;
            }
        }

        private void TestInOut_Tempsensor(bool on)
        {
            byte[] SendData = new byte[8];
            if (on == true)
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x80;
                SendData[5] = 0x00;
                SendData[6] = 0x20;
                SendData[7] = 0x8A;
            }
            else
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x00;
                SendData[5] = 0x00;
                SendData[6] = 0x00;
                SendData[7] = 0x0A;
            }
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void TestInOut_Voltagesensor(bool on)
        {
            byte[] SendData = new byte[8];

            if (on == true)
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x80;
                SendData[5] = 0x00;
                SendData[6] = 0xC0;
                SendData[7] = 0x8A;
            }
            else
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x00;
                SendData[5] = 0x00;
                SendData[6] = 0x00;
                SendData[7] = 0x0A;
            }
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void TestInOut_Tempsensor_Multi(bool on, JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[8];
            if (on == true)
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x80;
                SendData[5] = 0x00;
                SendData[6] = 0x20;
                SendData[7] = 0x8A;
            }
            else
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x00;
                SendData[5] = 0x00;
                SendData[6] = 0x00;
                SendData[7] = 0x0A;
            }
            I2C_Sel.WriteBytes(SendData, SendData.Length, true);
        }

        private void TestInOut_Voltagesensor_Multi(bool on, JLcLib.Custom.I2C I2C_Sel)
        {
            byte[] SendData = new byte[8];

            if (on == true)
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x80;
                SendData[5] = 0x00;
                SendData[6] = 0xC0;
                SendData[7] = 0x8A;
            }
            else
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x00;
                SendData[5] = 0x00;
                SendData[6] = 0x00;
                SendData[7] = 0x0A;
            }
            I2C_Sel.WriteBytes(SendData, SendData.Length, true);
        }

        private void TestInOut_EXT_SENS(bool on)
        {
            RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A

            if (on == true)
            {
                TestInOut_ADC_In(true);

                O_EXT_VOLT_EN.Read();
                O_EXT_VOLT_EN.Value = 1;
                O_EXT_VOLT_EN.Write();

                I_PEN(true);
            }
            else
            {
                TestInOut_ADC_In(false);

                O_EXT_VOLT_EN.Read();
                O_EXT_VOLT_EN.Value = 0;
                O_EXT_VOLT_EN.Write();

                I_PEN(false);
            }
        }

        private void ook_80_ADC_Read(bool on)
        {
            RegisterItem B17_ES = Parent.RegMgr.GetRegisterItem("B17_ES[7:0]");                         // 0x11
            RegisterItem B26_BL = Parent.RegMgr.GetRegisterItem("B26_BL[7:0]");                         // 0x1A
            RegisterItem B27_TP = Parent.RegMgr.GetRegisterItem("B27_TP[7:0]");                         // 0x1B

            SendCommand("ook 80\n");
            System.Threading.Thread.Sleep(100);

            B17_ES.Read();
            B26_BL.Read();
            B27_TP.Read();

            Log.WriteLine("Volt : " + (B26_BL.Value).ToString());
            Log.WriteLine("Temp : " + (B27_TP.Value).ToString());
            Log.WriteLine("EXT_SENS : " + (B17_ES.Value).ToString());
        }

        private void TestInOut_ADC_VREF(int select)
        {
            RegisterItem O_ADC_TEST = Parent.RegMgr.GetRegisterItem("O_ADC_TEST<2:0>");                 // 0x4A
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[9]");                   // 0x5E

            O_ADC_TEST.Read();
            ITEST_CONT.Read();

            if (select == 1)
            {
                I_PEN(false);

                O_ADC_TEST.Value = 4;
                O_ADC_TEST.Write();

                ITEST_CONT.Value = 1;
                ITEST_CONT.Write();

                I_PEN(true);
            }
            else if (select == 2)
            {
                I_PEN(false);

                O_ADC_TEST.Value = 2;
                O_ADC_TEST.Write();

                ITEST_CONT.Value = 1;
                ITEST_CONT.Write();

                I_PEN(true);
            }
            else if (select == 3)
            {
                I_PEN(false);

                O_ADC_TEST.Value = 1;
                O_ADC_TEST.Write();

                ITEST_CONT.Value = 1;
                ITEST_CONT.Write();

                I_PEN(true);
            }
            else
            {
                O_ADC_TEST.Value = 0;
                O_ADC_TEST.Write();

                ITEST_CONT.Value = 0;
                ITEST_CONT.Write();

                I_PEN(false);
            }
        }

        private void ADC_Result_Read()
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x0A;
            I2C.WriteBytes(SendData, SendData.Length, true);

            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x8A;
            I2C.WriteBytes(SendData, SendData.Length, true);

            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x0A;
            I2C.WriteBytes(SendData, SendData.Length, true);

            SendCommand("ook 80");

            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            Log.WriteLine("EXT : " + RcvData[2].ToString());
            Log.WriteLine("VBAT : " + RcvData[1].ToString());
            Log.WriteLine("Temp : " + RcvData[0].ToString());
        }

        private void BGR_TC_CTRL_Sweep(int num)
        {
            double alldo_target_mv = 810;
            double mldo_target_mv = 1000;
            double rcosc_target_khz = 32.768;
            double d_lsl = 800, d_usl = 820;
            uint ldo_val, ldo_val_1;
            double[] dVal = new double[5];
            double d_alldo_mv, d_mldo_mv, d_volt_mv;
            double d_diff_mv, d_target_bgr_mv = 300, d_target_alldo_mv = 810, d_target_mlod_mv = 1000;
            double d_freq_khz, d_diff_khz;
            double d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;
            uint alldo_val, mldo_val, bgr_val, alc = 0, mlc = 0, ul = 0;
            int x_pos = 1, y_pos = 1, y_Increment;
            string sTempVal, start_time;
            bool bRetrunTemp = true;
            uint rev_id;
            uint[,] reg_val = new uint[16, 8];
            string[] labels_cal =
            {
                "BGR_TC_CTRL<5:2>",         "O_ULP_BGR_CONT[3:0]",  "BGR[mV]",
                "O_ULP_LDO_LV_CONT[1:0]",   "O_ULP_LDO_CONT[3:0]",  "ALLDO[mV]",
                "O_PMU_MLDO_Coarse[1:0]",   "O_PMU_LDO_CONT[3:0]",  "MLDO[mV]",
                "O_RTC_SCKF[10:6]",         "O_RTC_SCKF[5:0]",      "RCOSC[kHz]"
            };
            double[] temps =
            {
                87.5,   85,     80,     75,     70,
                65,     60,     55,     50,     45,
                40,     35,     30,     25,     20,
                15,     10,     5,      0,      -5,
                -10,    -15,    -20,    -25,    -30,
                -35,    -40,    23.5
            };
            string[] volts =
            {
                "2.5", "3.0", "3.3"
            };
            y_Increment = (temps.Length) + 3;

            // RCOSC
            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");               // 0x51
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");              // 0x52
            // AL BGR
            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");         // 0x57
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL<5:2>");             // 0x67
            // ALLDO
            RegisterItem ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");         // 0x50
            RegisterItem ULP_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[1:0]");    // 0x50
            // MLDO
            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");         // 0x57
            RegisterItem PMU_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1:0]");    // 0x67
            // Rev_ID
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                     // 0x6C

            Check_Instrument();
            start_time = DateTime.Now.ToString("MM/dd/HH:mm");

            // Device ID Check & Excel Sheet Add
            I_DEVID.Read();
            rev_id = I_DEVID.Value;
            if (rev_id == 180)      // R5V1
            {
                Log.WriteLine("Device : IRIS R5V1");
                Parent.xlMgr.Sheet.Add("R5V1_" + num.ToString("X2") + "_CAL_LDO");
                Parent.xlMgr.Sheet.Add("R5V1_" + num.ToString("X2") + "_TC_Sweep");
                Parent.xlMgr.Sheet.Select("R5V1_" + num.ToString("X2") + "_CAL_LDO");
            }
            else if (rev_id == 212) // R5V2
            {
                Log.WriteLine("Device : IRIS R5V2");
                Parent.xlMgr.Sheet.Add("R5V2_" + num.ToString("X2") + "_CAL_LDO");
                Parent.xlMgr.Sheet.Add("R5V2_" + num.ToString("X2") + "_TC_Sweep");
                Parent.xlMgr.Sheet.Select("R5V2_" + num.ToString("X2") + "_CAL_LDO");
            }
            else
            {
                Log.WriteLine("Device : Null");
                return;
            }

            for (int i = 0; i < labels_cal.Length; i++)
            {
                Parent.xlMgr.Cell.Write(x_pos + i, y_pos, labels_cal[i]);
                if (labels_cal[i] == "BGR[mV]" || labels_cal[i] == "ALLDO[mV]" || labels_cal[i] == "MLDO[mV]")
                    x_pos++;
            }
            x_pos = 1;
            y_pos = 2;

            PowerSupply0.Write("VOLT 2.5,(@2)");

#if true   // 1. TempChamber 23.5℃ Set
            sTempVal = TempChamber.WriteAndReadString("01,TEMP,S23.5");
            System.Threading.Thread.Sleep(1000);
            sTempVal = TempChamber.WriteAndReadString("TEMP?");
            System.Threading.Thread.Sleep(100);
            Console.WriteLine("Run Chamber : " + 23.5);
            sTempVal = TempChamber.WriteAndReadString("TEMP?");
            System.Threading.Thread.Sleep(100);
            bRetrunTemp = true;
            while (bRetrunTemp)
            {
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                for (int split = 0; split < 4; split++)
                    dVal[split] = double.Parse(ArrBuf[split]);

                double dTHVal = 0.1;

                if (dVal[0] + dTHVal == 23.5 || dVal[0] - dTHVal == 23.5 || dVal[0] == 23.5)
                {
                    bRetrunTemp = false;
                    Log.WriteLine("Done SetTemp!");
                }
                else
                    bRetrunTemp = true;
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                System.Threading.Thread.Sleep(1000 * 5);
            }
            System.Threading.Thread.Sleep(1000 * 10);
            Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
            // 2. for TC 0~15 Sweep
            //      - for BGR CONT Sweep
            //      - BGR       300mV       Calibration
            //      - ALLDO     810mV       Calibration
            //      - MLDO      1000mV      Calibration
            //      - RCOSC     32.768kHz   Calibration
            //      - Register Value Array Save
            for (uint tc_val = 0; tc_val < 16; tc_val++)
            {
                // Cal AL BGR
                BGR_TC_CTRL.Read();
                BGR_TC_CTRL.Value = tc_val;
                BGR_TC_CTRL.Write();

                Set_TestInOut_For_BGR(true);

                ULP_BGR_CONT.Read();
                bgr_val = 0;
                d_diff_mv = 5000;
                System.Threading.Thread.Sleep(500);
                for (uint val = 0; val < 16; val++)
                {
                    System.Threading.Thread.Sleep(500);
                    ULP_BGR_CONT.Value = val;
                    ULP_BGR_CONT.Write();

                    System.Threading.Thread.Sleep(500);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    if (Math.Abs(d_volt_mv - d_target_bgr_mv) < d_diff_mv)
                    {
                        d_diff_mv = Math.Abs(d_volt_mv - d_target_bgr_mv);
                        bgr_val = val;
                    }
                }
                ULP_BGR_CONT.Value = bgr_val;
                ULP_BGR_CONT.Write();
                System.Threading.Thread.Sleep(500);
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                BGR_TC_CTRL.Read();

                Parent.xlMgr.Cell.Write(1, y_pos, (BGR_TC_CTRL.Value).ToString());
                Parent.xlMgr.Cell.Write(2, y_pos, bgr_val.ToString());
                Parent.xlMgr.Cell.Write(3, y_pos, d_volt_mv.ToString("F3"));

                reg_val[tc_val, 0] = tc_val;
                reg_val[tc_val, 1] = bgr_val;
                Set_TestInOut_For_BGR(false);
                // Cal AL BGR

                // Cal ALLDO
                ULP_LDO_CONT.Read();
                ULP_LDO_Coarse.Read();

                ULP_LDO_CONT.Value = 8;
                ULP_LDO_CONT.Write();
                ULP_LDO_Coarse.Value = 0;
                ULP_LDO_Coarse.Write();

                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv > 816)
                {
                    ULP_LDO_Coarse.Value = 3;
                    ULP_LDO_Coarse.Write();
                }

                ldo_val = 15;
                ldo_val_1 = 0;
                ULP_LDO_CONT.Value = ldo_val;
                ULP_LDO_CONT.Write();

                for (int val = 2; val >= 0; val--)
                {
                    d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    if (d_volt_mv < alldo_target_mv)
                    {
                        ldo_val += (uint)(1 << val);
                    }
                    else
                    {
                        ldo_val -= (uint)(1 << val);
                    }
                    ldo_val = ldo_val & 0xf;
                    ULP_LDO_CONT.Value = ldo_val;
                    ULP_LDO_CONT.Write();
                }
                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                ldo_val_1 = ldo_val;
                d_diff_mv = Math.Abs(d_volt_mv - alldo_target_mv);

                if (d_volt_mv < alldo_target_mv)
                {
                    if (ldo_val != 7) ldo_val += 1;
                }
                else
                {
                    if (ldo_val != 8) ldo_val -= 1;
                }
                ldo_val = ldo_val & 0xf;
                ULP_LDO_CONT.Value = ldo_val;
                ULP_LDO_CONT.Write();

                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (Math.Abs(d_volt_mv - alldo_target_mv) > d_diff_mv)
                {
                    ldo_val = ldo_val_1;
                    ULP_LDO_CONT.Value = ldo_val;
                    ULP_LDO_CONT.Write();
                }

                reg_val[tc_val, 2] = ULP_LDO_Coarse.Value;
                reg_val[tc_val, 3] = ldo_val;

                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write(5, y_pos, (ULP_LDO_Coarse.Value).ToString());
                Parent.xlMgr.Cell.Write(6, y_pos, ldo_val.ToString());
                Parent.xlMgr.Cell.Write(7, y_pos, d_volt_mv.ToString("F3"));
                // Cal ALLDO

                // Cal MLDO
                PMU_LDO_Coarse.Read();
                PMU_LDO_CONT.Read();

                PMU_LDO_CONT.Value = 7;
                PMU_LDO_CONT.Write();
                PMU_LDO_Coarse.Value = 0;
                PMU_LDO_Coarse.Write();

                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < 1006)
                {
                    PMU_LDO_Coarse.Value = 1;
                    PMU_LDO_Coarse.Write();
                    d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    if (d_volt_mv < 1006)
                    {
                        PMU_LDO_Coarse.Value = 2;
                        PMU_LDO_Coarse.Write();
                    }
                }

                ldo_val = 15;
                ldo_val_1 = 0;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();

                for (int val = 2; val >= 0; val--)
                {
                    d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    if (d_volt_mv < mldo_target_mv)
                    {
                        ldo_val += (uint)(1 << val);
                    }
                    else
                    {
                        ldo_val -= (uint)(1 << val);
                    }
                    ldo_val = ldo_val & 0xf;
                    PMU_LDO_CONT.Value = ldo_val;
                    PMU_LDO_CONT.Write();
                }

                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                ldo_val_1 = ldo_val;
                d_diff_mv = Math.Abs(d_volt_mv - mldo_target_mv);

                if (d_volt_mv < mldo_target_mv)
                {
                    if (ldo_val != 7) ldo_val += 1;
                }
                else
                {
                    if (ldo_val != 8) ldo_val -= 1;
                }
                ldo_val = ldo_val & 0xf;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (Math.Abs(d_volt_mv - mldo_target_mv) > d_diff_mv)
                {
                    ldo_val = ldo_val_1;
                    PMU_LDO_CONT.Value = ldo_val;
                    PMU_LDO_CONT.Write();
                }

                reg_val[tc_val, 4] = PMU_LDO_Coarse.Value;
                reg_val[tc_val, 5] = ldo_val;

                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write(9, y_pos, (PMU_LDO_Coarse.Value).ToString());
                Parent.xlMgr.Cell.Write(10, y_pos, ldo_val.ToString());
                Parent.xlMgr.Cell.Write(11, y_pos, d_volt_mv.ToString("F3"));
                // Cal MLDO

                // Cal OSC
                Set_TestInOut_For_RCOSC(true);

                RTC_SCKF_L.Read();
                RTC_SCKF_H.Read();

                osc_val_l = 31;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();

                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

                for (int val = 4; val >= 0; val--)
                {
                    if (d_freq_khz > d_target_khz)
                        osc_val_l += (uint)(1 << val);
                    else
                        osc_val_l -= (uint)(1 << val);
                    RTC_SCKF_L.Value = osc_val_l;
                    RTC_SCKF_L.Write();

                    d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                }

                osc_val_l_1 = osc_val_l;
                d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

                if (d_freq_khz > d_target_khz)
                    if (osc_val_l != 63) osc_val_l += 1;
                    else
                    if (osc_val_l != 0) osc_val_l -= 1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();

                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

                if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
                {
                    osc_val_l = osc_val_l_1;
                    RTC_SCKF_L.Value = osc_val_l;
                    RTC_SCKF_L.Write();
                }

                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

                RTC_SCKF_H.Read();
                Parent.xlMgr.Cell.Write(13, y_pos, (RTC_SCKF_H.Value).ToString("F3"));
                Parent.xlMgr.Cell.Write(14, y_pos, osc_val_l.ToString());
                Parent.xlMgr.Cell.Write(15, y_pos, d_freq_khz.ToString("F3"));

                reg_val[tc_val, 6] = RTC_SCKF_H.Value;
                reg_val[tc_val, 7] = osc_val_l;

                Set_TestInOut_For_RCOSC(false);
                y_pos++;
            }

            Log.WriteLine("Calibration Complete");
            Log.WriteLine("TempSweep Start.");
            Set_TestInOut_For_BGR(true);

            if (rev_id == 180)      // R5V1
            {
                Parent.xlMgr.Sheet.Select("R5V1_" + num.ToString("X2") + "_TC_Sweep");
            }
            else if (rev_id == 212) // R5V2
            {
                Parent.xlMgr.Sheet.Select("R5V2_" + num.ToString("X2") + "_TC_Sweep");
            }

            x_pos = 2;
            y_pos = 1;

            Parent.xlMgr.Cell.Write(x_pos++, y_pos, "BGR_TC_CTRL<5:2>");
            Parent.xlMgr.Cell.Write(x_pos++, y_pos, "O_ULP_BGR_CONT[3:0]");

            for (int j = 0; j < temps.Length; j++)
            {
                Parent.xlMgr.Cell.Write(x_pos + j, y_pos, temps[j].ToString());
            }

            x_pos = 1;
            y_pos = 2;

            // 3. Temp Sweep 87.5 ~ -40℃
            //      - TC 0~15 Sweep
            //      - Array Save Value Init
            //      - LDO Voltage Level Measure.
            for (int j = 0; j < temps.Length; j++)
            {
#if true    // 1. TempChamber Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[j]);
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + temps[j]);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == temps[j] || dVal[0] - dTHVal == temps[j] || dVal[0] == temps[j])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 60);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                for (int tc_cnt = 0; tc_cnt < 16; tc_cnt++)
                {
                    for (int k = 0; k < volts.Length; k++)
                    {
                        PowerSupply0.Write("VOLT " + volts[k] + ",(@2)");
                        System.Threading.Thread.Sleep(500);
                        if (k == 0)
                        {
                            ULP_LDO_Coarse.Value = reg_val[tc_cnt, 2];
                            ULP_LDO_CONT.Value = reg_val[tc_cnt, 3];
                            ULP_LDO_CONT.Write();
                            System.Threading.Thread.Sleep(10);

                            BGR_TC_CTRL.Value = reg_val[tc_cnt, 0];
                            ULP_BGR_CONT.Value = reg_val[tc_cnt, 1];

                            PMU_LDO_Coarse.Value = reg_val[tc_cnt, 4];
                            PMU_LDO_Coarse.Write();
                            System.Threading.Thread.Sleep(10);
                            PMU_LDO_CONT.Value = reg_val[tc_cnt, 5];
                            PMU_LDO_CONT.Write();
                            System.Threading.Thread.Sleep(10);

                            RTC_SCKF_H.Value = reg_val[tc_cnt, 6];
                            RTC_SCKF_H.Write();
                            System.Threading.Thread.Sleep(10);
                            RTC_SCKF_L.Value = reg_val[tc_cnt, 7];
                            RTC_SCKF_L.Write();
                            System.Threading.Thread.Sleep(10);
                        }

                        System.Threading.Thread.Sleep(50);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;

                        if (j == 0)
                        {
                            Parent.xlMgr.Cell.Write(x_pos + 0, 2 + (17 * k), "VBAT : " + volts[k].ToString() + "[V]");
                            Parent.xlMgr.Cell.Write(x_pos + 1, y_pos + (17 * k), tc_cnt.ToString());
                            Parent.xlMgr.Cell.Write(x_pos + 2, y_pos + (17 * k), (ULP_BGR_CONT.Value).ToString());
                        }
                        Parent.xlMgr.Cell.Write(x_pos + 3 + j, y_pos + (17 * k), d_volt_mv.ToString());
                    }
                    y_pos++;
                }
                y_pos = 2;
            }
            Set_TestInOut_For_BGR(false);
            Log.WriteLine("Test Start : " + start_time);
            Log.WriteLine("TC Sweep Complete : " + DateTime.Now.ToString("MM/dd/HH:mm"));
        }

        private void Write_BLE00_OTP(int num)
        {
            uint device_code;
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");                 // 0x09
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");                 // 0x08
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");                 // 0x07
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");                 // 0x06
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");                 // 0x05
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                 // 0x04
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");               // 0x6C

            // Devicd ID Check
            I_DEVID.Read();
            if (I_DEVID.Value == 180)      // R5V1
            {
                Log.WriteLine("Device Check : IRIS R5V1");
                device_code = 1;
            }
            else if (I_DEVID.Value == 212) // R5V2
            {
                Log.WriteLine("Device Check : IRIS R5V2");
                device_code = 2;
            }
            else
            {
                Log.WriteLine("Device Check : Fail");
                return;
            }

            SendCommand("mode ook\n");
            DialogResult dialog = MessageBox.Show(
                "\"확인\"을 누르면 OTP를 작성합니다.\n" +
                "작성 BLE PAGE : 00\n" +
                "MAC : 78 46 FF FF " + device_code.ToString("X2") + " " + num.ToString("X2"),
                Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (dialog == DialogResult.OK)
            {
                B09_BA_5.Read();
                B09_BA_5.Value = 0x78;
                B09_BA_5.Write();

                B08_BA_4.Read();
                B08_BA_4.Value = 0x46;
                B08_BA_4.Write();

                B05_BA_1.Read();
                B05_BA_1.Value = device_code;
                B05_BA_1.Write();

                B04_BA_0.Read();
                B04_BA_0.Value = (uint)num;
                B04_BA_0.Write();
                System.Threading.Thread.Sleep(10);

                SendCommand("ook 41.00\n");
                System.Threading.Thread.Sleep(1000);

                PowerSupply0.Write("OUTP OFF");
                System.Threading.Thread.Sleep(1000);
                PowerSupply0.Write("OUTP ON");
                System.Threading.Thread.Sleep(1000);

                WakeUp_I2C();
                System.Threading.Thread.Sleep(500);

                B05_BA_1.Read();
                B04_BA_0.Read();
                if ((B05_BA_1.Value == device_code) && (B04_BA_0.Value == (uint)num))
                {
                    MessageBox.Show("BLE PAGE OTP Write Complete");
                }
                else
                {
                    MessageBox.Show("BLE PAGE OTP Write Fail");
                }
            }
            else return;
        }

        private void Cal_PMU_FastCal(int num)
        {
            double alldo_target_mv = 810;
            double mldo_target_mv = 1000;
            uint ldo_val, ldo_val_1;
            double[] dVal = new double[5];
            double d_volt_mv;
            double d_diff_mv, d_target_bgr_mv = 300;
            double d_freq_khz, d_diff_khz;
            double d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;
            uint bgr_val;
            int y_pos = 1;
            string sTempVal;
            bool bRetrunTemp = true;
            uint rev_id;

            // RCOSC
            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");               // 0x51
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");              // 0x52
            // AL BGR
            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");         // 0x57
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL<5:2>");             // 0x67
            // ALLDO
            RegisterItem ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");         // 0x50
            RegisterItem ULP_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[1:0]");    // 0x50
            // MLDO
            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");         // 0x57
            RegisterItem PMU_LDO_Coarse = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1:0]");    // 0x67
            // Rev_ID
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                     // 0x6C

            Check_Instrument();
            I_DEVID.Read();
            rev_id = I_DEVID.Value;
            if (rev_id == 180)      // R5V1
            {
                Log.WriteLine("Device : IRIS R5V1");
                Parent.xlMgr.Sheet.Select("R5V1_CAL_LDO");
                Parent.xlMgr.Cell.Write(1, y_pos + num, num.ToString());
            }
            else if (rev_id == 212) // R5V2
            {
                Log.WriteLine("Device : IRIS R5V2");
                Parent.xlMgr.Sheet.Select("R5V2_CAL_LDO");
                Parent.xlMgr.Cell.Write(1, y_pos + num, num.ToString());
            }
            else
            {
                Log.WriteLine("Device : Null");
                return;
            }

            PowerSupply0.Write("VOLT 2.5,(@2)");

#if true   // 1. TempChamber 23.5℃ Set
            sTempVal = TempChamber.WriteAndReadString("01,TEMP,S23.5");
            System.Threading.Thread.Sleep(1000);
            sTempVal = TempChamber.WriteAndReadString("TEMP?");
            System.Threading.Thread.Sleep(100);
            Console.WriteLine("Run Chamber : " + 23.5);
            sTempVal = TempChamber.WriteAndReadString("TEMP?");
            System.Threading.Thread.Sleep(100);
            bRetrunTemp = true;
            while (bRetrunTemp)
            {
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                for (int split = 0; split < 4; split++)
                    dVal[split] = double.Parse(ArrBuf[split]);

                if (dVal[0] >= 23.4 && dVal[0] <= 23.6)
                {
                    bRetrunTemp = false;
                    Log.WriteLine("Done SetTemp!");
                }
                else
                    bRetrunTemp = true;
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                System.Threading.Thread.Sleep(1000 * 5);
            }
            System.Threading.Thread.Sleep(1000 * 60);
            Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
            // Cal AL BGR
            BGR_TC_CTRL.Read();
            BGR_TC_CTRL.Value = 9;
            BGR_TC_CTRL.Write();

            Set_TestInOut_For_BGR(true);

            ldo_val = 15;
            ldo_val_1 = 0;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < d_target_bgr_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_bgr_mv);

            if (d_volt_mv < d_target_bgr_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (Math.Abs(d_volt_mv - d_target_bgr_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(2, y_pos + num, (BGR_TC_CTRL.Value).ToString());
            Parent.xlMgr.Cell.Write(3, y_pos + num, ldo_val.ToString());
            Parent.xlMgr.Cell.Write(4, y_pos + num, d_volt_mv.ToString("F3"));

            Set_TestInOut_For_BGR(false);
            // Cal AL BGR

            // Cal ALLDO
            ULP_LDO_CONT.Read();
            ULP_LDO_Coarse.Read();

            ULP_LDO_CONT.Value = 8;
            ULP_LDO_CONT.Write();
            ULP_LDO_Coarse.Value = 0;
            ULP_LDO_Coarse.Write();

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv > 816)
            {
                ULP_LDO_Coarse.Value = 3;
                ULP_LDO_Coarse.Write();
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            ULP_LDO_CONT.Value = ldo_val;
            ULP_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < alldo_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                ULP_LDO_CONT.Value = ldo_val;
                ULP_LDO_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - alldo_target_mv);

            if (d_volt_mv < alldo_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            ULP_LDO_CONT.Value = ldo_val;
            ULP_LDO_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (Math.Abs(d_volt_mv - alldo_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                ULP_LDO_CONT.Value = ldo_val;
                ULP_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(5, y_pos + num, (ULP_LDO_Coarse.Value).ToString());
            Parent.xlMgr.Cell.Write(6, y_pos + num, ldo_val.ToString());
            Parent.xlMgr.Cell.Write(7, y_pos + num, d_volt_mv.ToString("F3"));
            // Cal ALLDO

            // Cal MLDO
            PMU_LDO_Coarse.Read();
            PMU_LDO_CONT.Read();

            PMU_LDO_CONT.Value = 7;
            PMU_LDO_CONT.Write();
            PMU_LDO_Coarse.Value = 0;
            PMU_LDO_Coarse.Write();

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv < 1006)
            {
                PMU_LDO_Coarse.Value = 1;
                PMU_LDO_Coarse.Write();
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < 1006)
                {
                    PMU_LDO_Coarse.Value = 2;
                    PMU_LDO_Coarse.Write();
                }
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < mldo_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - mldo_target_mv);

            if (d_volt_mv < mldo_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();
            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (Math.Abs(d_volt_mv - mldo_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write(8, y_pos + num, (PMU_LDO_Coarse.Value).ToString());
            Parent.xlMgr.Cell.Write(9, y_pos + num, ldo_val.ToString());
            Parent.xlMgr.Cell.Write(10, y_pos + num, d_volt_mv.ToString("F3"));
            // Cal MLDO

            // Cal OSC
            Set_TestInOut_For_RCOSC(true);

            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            for (int val = 4; val >= 0; val--)
            {
                if (d_freq_khz > d_target_khz)
                    osc_val_l += (uint)(1 << val);
                else
                    osc_val_l -= (uint)(1 << val);
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();

                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            }

            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
                if (osc_val_l != 63) osc_val_l += 1;
                else
                if (osc_val_l != 0) osc_val_l -= 1;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;

            RTC_SCKF_H.Read();
            Parent.xlMgr.Cell.Write(11, y_pos + num, osc_val_l.ToString());
            Parent.xlMgr.Cell.Write(12, y_pos + num, d_freq_khz.ToString("F3"));

            Set_TestInOut_For_RCOSC(false);
            Log.WriteLine("Calibration Complete");

#if true
            SendCommand("mode ook\n");
            DialogResult dialog = MessageBox.Show(
                "\"확인\"을 누르면 OTP를 작성합니다.\n" +
                "작성 ANA PAGE : 21\n",
                Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (dialog == DialogResult.OK)
            {
                SendCommand("ook 41.15\n");
                System.Threading.Thread.Sleep(1000);

                PowerSupply0.Write("OUTP OFF");
                System.Threading.Thread.Sleep(1000);
                PowerSupply0.Write("OUTP ON");
                System.Threading.Thread.Sleep(1000);

                WakeUp_I2C();
                System.Threading.Thread.Sleep(500);

                ULP_BGR_CONT.Read();
                BGR_TC_CTRL.Read();

                if (BGR_TC_CTRL.Value == 9)
                {
                    MessageBox.Show("ANA PAGE OTP Write Complete");
                }
                else
                {
                    MessageBox.Show("ANA PAGE OTP Write Fail");
                }
            }
#endif
        }
        private void ADC_EXT_SENS_Forcing()
        {
            for (int cnt = 1; cnt <= 10; cnt++)
            {
                double[] dVal = new double[5];
                double d_volt_mv;
                string sTempVal;
                bool bRetrunTemp = true;
                uint rev_id;
                int x_pos = 1, y_pos = 1;
                byte[] SendData = new byte[8];
                byte[] RcvData = new byte[4];
                string[] labels =
                {
                    "VIN[mV]",  "ADC code", "BGR", "ADC_TOP", "ADC_CM", "ADC_BOT"
                },
                VBAT_volts =
                {
                    "1.7", "2.5", "3.3"
                };
                double[] temps =
                {
                    87.5, 23.5, -40
                };
                string[] EXT_volts = { "0.15", "0.2", "0.250", "0.300", "0.350", "0.400", "0.450", "0.500", "0.550", "0.600", "0.650", "0.700", "0.750" };

                int y_increment = 1 + VBAT_volts.Length * labels.Length + 1;

                Check_Instrument();

                RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                     // 0x04
                RegisterItem B17_ES = Parent.RegMgr.GetRegisterItem("B17_ES[7:0]");                         // 0x1B

                RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A

                RegisterItem O_ADC_TEST = Parent.RegMgr.GetRegisterItem("O_ADC_TEST<2:0>");                 // 0x4A
                RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");                 // 0x61

                RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                       // 0x6C

                I_DEVID.Read();
                B04_BA_0.Read();

                rev_id = I_DEVID.Value;
                if (rev_id == 180)      // R5V1
                {
                    Log.WriteLine("Device : IRIS R5V1");
                    Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_EXT_Forcing" + cnt.ToString());
                    Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_EXT_Forcing" + cnt.ToString());
                }
                else if (rev_id == 212) // R5V2
                {
                    Log.WriteLine("Device : IRIS R5V2");
                    Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_EXT_Forcing" + cnt.ToString());
                    Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_EXT_Forcing" + cnt.ToString());
                }
                else
                {
                    Log.WriteLine("Device : Null");
                    return;
                }

                for (int i = 0; i < temps.Length; i++)
                {
                    int current_y = y_pos + (y_increment * i);

                    Parent.xlMgr.Cell.Write(x_pos, current_y, "Chamber Temp[℃]");
                    Parent.xlMgr.Cell.Write(x_pos + 1, current_y, temps[i].ToString());
                    Parent.xlMgr.Cell.Write(x_pos, current_y + 1, "VBAT[V]");

                    for (int c = 0; c < EXT_volts.Length; c++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos + 2 + c, current_y + 1, EXT_volts[c]);
                    }

                    for (int j = 0; j < VBAT_volts.Length; j++)
                    {
                        int vbat_y = current_y + 2 + j * labels.Length;

                        Parent.xlMgr.Cell.Write(x_pos, vbat_y, VBAT_volts[j].ToString());

                        for (int k = 0; k < labels.Length; k++)
                        {
                            Parent.xlMgr.Cell.Write(x_pos + 1, vbat_y + k, labels[k]);
                        }
                    }
                }
                x_pos = 3; y_pos = 3;
                for (int j = 0; j < temps.Length; j++)
                {
#if true    // 1. TempChamber Set
                    sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[j]);
                    System.Threading.Thread.Sleep(1000);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    Console.WriteLine("Run Chamber : " + temps[j]);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    bRetrunTemp = true;
                    while (bRetrunTemp)
                    {
                        sTempVal = TempChamber.WriteAndReadString("TEMP?");
                        System.Threading.Thread.Sleep(100);
                        sTempVal = TempChamber.WriteAndReadString("TEMP?");
                        System.Threading.Thread.Sleep(100);
                        string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                        for (int split = 0; split < 4; split++)
                            dVal[split] = double.Parse(ArrBuf[split]);

                        double dTHVal = 0.1;

                        if (dVal[0] + dTHVal == temps[j] || dVal[0] - dTHVal == temps[j] || dVal[0] == temps[j])
                        {
                            bRetrunTemp = false;
                            Log.WriteLine("Done SetTemp!");
                        }
                        else
                            bRetrunTemp = true;
                        Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                        System.Threading.Thread.Sleep(1000 * 5);
                    }
                    System.Threading.Thread.Sleep(1000 * 60);
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                    for (int i = 0; i < VBAT_volts.Length; i++)
                    {
                        PowerSupply0.Write("VOLT " + VBAT_volts[i] + ",(@2)");
                        System.Threading.Thread.Sleep(700);
                        for (int k = 0; k < EXT_volts.Length; k++)
                        {
                            PowerSupply0.Write("VOLT " + EXT_volts[k] + ",(@3)");
                            System.Threading.Thread.Sleep(700);

                            TestInOut_ADC_In(true);
                            O_EXT_VOLT_EN.Read();
                            O_EXT_VOLT_EN.Value = 1;
                            O_EXT_VOLT_EN.Write();
                            System.Threading.Thread.Sleep(100);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(100);

                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + (y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);
                            TestInOut_ADC_In(false);

                            System.Threading.Thread.Sleep(100);

                            SendCommand("ook 80\n");
                            System.Threading.Thread.Sleep(100);

                            B17_ES.Read();
                            System.Threading.Thread.Sleep(10);
                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + 1 + (y_increment * j), (B17_ES.Value).ToString());
                            O_EXT_VOLT_EN.Value = 0;
                            O_EXT_VOLT_EN.Write();

                            Set_TestInOut_For_BGR(true);
                            System.Threading.Thread.Sleep(500);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + 2 + (y_increment * j), d_volt_mv.ToString("F3"));
                            Set_TestInOut_For_BGR(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 4;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + 3 + (y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 2;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + 4 + (y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 1;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + 5 + (y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 0;
                            O_ADC_TEST.Write();
                        }
                        y_pos += labels.Length;
                    }
                    y_pos = 3;
                }
            }
            TempChamber.WriteAndReadString("01,TEMP,S23.5");
        }

        private void TestInOut_EXT_ADC_In()
        {
            for (int cnt = 1; cnt <= 10; cnt++)
            {
                double[] dVal = new double[5];
                string sTempVal;
                bool bRetrunTemp = true;
                uint rev_id;
                int x_pos = 1, y_pos = 1;
                byte[] SendData = new byte[8];
                byte[] RcvData = new byte[4];
                string[] labels =
                {
                    "ADC code"
                },
                VBAT_volts =
                {
                    "1.7", "2.5", "3.3"
                };
                double[] temps =
                {
                    87.5, 23.5, -40
                };
                string[] EXT_volts = { "0.15", "0.2", "0.250", "0.300", "0.350", "0.400", "0.450", "0.500", "0.550", "0.600", "0.650", "0.700", "0.750" };

                int y_increment = 1 + VBAT_volts.Length * labels.Length + 1;

                Check_Instrument();

                RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                     // 0x04

                RegisterItem B26_BL = Parent.RegMgr.GetRegisterItem("B26_BL[7:0]");                         // 0x1A

                RegisterItem O_ADC_VREFT_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_TEMP<6:0>");     // 0x44
                RegisterItem O_ADC_VREFB_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_TEMP<6:0>");     // 0x46
                RegisterItem O_ADC_VREFT_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_VOLT<6:0>");     // 0x47
                RegisterItem O_ADC_VREFB_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_VOLT<6:0>");     // 0x49

                RegisterItem ADC_T_EN = Parent.RegMgr.GetRegisterItem("ADC_T_EN");                          // 0x5E
                RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");                 // 0x61

                RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                       // 0x6C

                I_DEVID.Read();
                B04_BA_0.Read();

                O_ADC_VREFT_TEMP.Read();
                O_ADC_VREFT_TEMP.Value = 32;
                O_ADC_VREFT_TEMP.Write();

                O_ADC_VREFB_TEMP.Read();
                O_ADC_VREFB_TEMP.Value = 32;
                O_ADC_VREFB_TEMP.Write();

                O_ADC_VREFT_VOLT.Read();
                O_ADC_VREFT_VOLT.Value = 32;
                O_ADC_VREFT_VOLT.Write();

                O_ADC_VREFB_VOLT.Read();
                O_ADC_VREFB_VOLT.Value = 32;
                O_ADC_VREFB_VOLT.Write();

                ADC_T_EN.Read();
                ADC_T_EN.Value = 1;
                ADC_T_EN.Write();

                O_TEST_CON.Read();
                O_TEST_CON.Value = 1;
                O_TEST_CON.Write();

                rev_id = I_DEVID.Value;
                if (rev_id == 180)      // R5V1
                {
                    Log.WriteLine("Device : IRIS R5V1");
                    Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_TestIO_EXT_ADC" + cnt.ToString());
                    Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_TestIO_EXT_ADC" + cnt.ToString());
                }
                else if (rev_id == 212) // R5V2
                {
                    Log.WriteLine("Device : IRIS R5V2");
                    Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_TestIO_EXT_ADC" + cnt.ToString());
                    Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_TestIO_EXT_ADC" + cnt.ToString());
                }
                else
                {
                    Log.WriteLine("Device : Null");
                    return;
                }

                for (int i = 0; i < temps.Length; i++)
                {
                    int current_y = y_pos + (y_increment * i);

                    Parent.xlMgr.Cell.Write(x_pos, current_y, "Chamber Temp[℃]");
                    Parent.xlMgr.Cell.Write(x_pos + 1, current_y, temps[i].ToString());
                    Parent.xlMgr.Cell.Write(x_pos, current_y + 1, "VBAT[V]");

                    for (int c = 0; c < EXT_volts.Length; c++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos + 2 + c, current_y + 1, EXT_volts[c]);
                    }

                    for (int j = 0; j < VBAT_volts.Length; j++)
                    {
                        int vbat_y = current_y + 2 + j * labels.Length;

                        Parent.xlMgr.Cell.Write(x_pos, vbat_y, VBAT_volts[j].ToString());

                        for (int k = 0; k < labels.Length; k++)
                        {
                            Parent.xlMgr.Cell.Write(x_pos + 1, vbat_y + k, labels[k]);
                        }
                    }
                }
                x_pos = 3; y_pos = 3;
                for (int j = 0; j < temps.Length; j++)
                {
#if true    // 1. TempChamber Set
                    sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[j]);
                    System.Threading.Thread.Sleep(1000);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    Console.WriteLine("Run Chamber : " + temps[j]);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    bRetrunTemp = true;
                    while (bRetrunTemp)
                    {
                        sTempVal = TempChamber.WriteAndReadString("TEMP?");
                        System.Threading.Thread.Sleep(100);
                        sTempVal = TempChamber.WriteAndReadString("TEMP?");
                        System.Threading.Thread.Sleep(100);
                        string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                        for (int split = 0; split < 4; split++)
                            dVal[split] = double.Parse(ArrBuf[split]);

                        double dTHVal = 0.1;

                        if (dVal[0] + dTHVal == temps[j] || dVal[0] - dTHVal == temps[j] || dVal[0] == temps[j])
                        {
                            bRetrunTemp = false;
                            Log.WriteLine("Done SetTemp!");
                        }
                        else
                            bRetrunTemp = true;
                        Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                        System.Threading.Thread.Sleep(1000 * 5);
                    }
                    System.Threading.Thread.Sleep(1000 * 60);
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                    for (int i = 0; i < VBAT_volts.Length; i++)
                    {
                        PowerSupply0.Write("VOLT " + VBAT_volts[i] + ",(@2)");
                        System.Threading.Thread.Sleep(700);
                        for (int k = 0; k < EXT_volts.Length; k++)
                        {
                            PowerSupply0.Write("VOLT " + EXT_volts[k] + ",(@3)");
                            System.Threading.Thread.Sleep(700);

                            SendCommand("ook 80\n");
                            System.Threading.Thread.Sleep(100);

                            B26_BL.Read();

                            Parent.xlMgr.Cell.Write(x_pos + k, y_pos + (y_increment * j), (B26_BL.Value).ToString("F3"));
                        }
                        y_pos += labels.Length;
                    }
                    y_pos = 3;
                }
            }
            TempChamber.WriteAndReadString("01,TEMP,S23.5");
        }

        private void ADC_Top_Bot_Sweep()
        {
            double[] dVal = new double[5];
            double d_volt_mv;
            string sTempVal;
            bool bRetrunTemp = true;
            uint rev_id;
            int x_pos = 1, y_pos = 1;
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            string[] labels =
            {
                "ADC code"
            },
            VBAT_volts =
            {
                "1.7", "2.5", "3.3"
            };
            double[] temps =
            {
                23.5
            };

            int y_increment = 1 + VBAT_volts.Length * labels.Length + 1;

            Check_Instrument();

            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                     // 0x04
            RegisterItem B17_ES = Parent.RegMgr.GetRegisterItem("B17_ES[7:0]");                         // 0x15
            RegisterItem B26_BL = Parent.RegMgr.GetRegisterItem("B26_BL[7:0]");                         // 0x1A
            RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A
            RegisterItem O_ADC_VREFT_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_TEMP<6:0>");     // 0x44
            RegisterItem O_ADC_VREFB_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_TEMP<6:0>");     // 0x46
            RegisterItem O_ADC_VREFT_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_VOLT<6:0>");     // 0x47
            RegisterItem O_ADC_VREFB_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_VOLT<6:0>");     // 0x49
            RegisterItem O_ADC_TEST = Parent.RegMgr.GetRegisterItem("O_ADC_TEST<2:0>");                 // 0x4A
            RegisterItem ADC_T_EN = Parent.RegMgr.GetRegisterItem("ADC_T_EN");                          // 0x5E
            RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");                 // 0x61
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                       // 0x6C

            O_ADC_VREFT_TEMP.Read();
            O_ADC_VREFT_TEMP.Value = 32;
            O_ADC_VREFT_TEMP.Write();

            O_ADC_VREFB_TEMP.Read();
            O_ADC_VREFB_TEMP.Value = 32;
            O_ADC_VREFB_TEMP.Write();

            O_ADC_VREFT_VOLT.Read();
            O_ADC_VREFT_VOLT.Value = 32;
            O_ADC_VREFT_VOLT.Write();

            O_ADC_VREFB_VOLT.Read();
            O_ADC_VREFB_VOLT.Value = 32;
            O_ADC_VREFB_VOLT.Write();

            O_TEST_CON.Read();
            O_TEST_CON.Value = 1;
            O_TEST_CON.Write();

            ADC_T_EN.Read();
            ADC_T_EN.Value = 1;
            ADC_T_EN.Write();

            I_DEVID.Read();
            B04_BA_0.Read();

            rev_id = I_DEVID.Value;
            if (rev_id == 180)      // R5V1
            {
                Log.WriteLine("Device : IRIS R5V1");
                Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_ADC_Top_Bot_Sweep");
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_ADC_Top_Bot_Sweep");
            }
            else if (rev_id == 212) // R5V2
            {
                Log.WriteLine("Device : IRIS R5V2");
                Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_ADC_Top_Bot_Sweep");
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_ADC_Top_Bot_Sweep");
            }
            else
            {
                Log.WriteLine("Device : Null");
                return;
            }

            for (int i = 0; i < temps.Length; i++)
            {
                int current_y = y_pos + (y_increment * i);

                Parent.xlMgr.Cell.Write(x_pos, current_y, "Chamber Temp[℃]");
                Parent.xlMgr.Cell.Write(x_pos + 1, current_y, temps[i].ToString());
                Parent.xlMgr.Cell.Write(x_pos, current_y + 1, "VBAT[V]");

                for (double c = 0.150; c < 0.751; c += 0.001)
                {
                    Parent.xlMgr.Cell.Write(x_pos + 2, current_y + 1, c.ToString());
                    x_pos++;
                }
                x_pos = 1;

                for (int j = 0; j < VBAT_volts.Length; j++)
                {
                    int vbat_y = current_y + 2 + j * labels.Length;

                    Parent.xlMgr.Cell.Write(x_pos, vbat_y, VBAT_volts[j].ToString());

                    for (int k = 0; k < labels.Length; k++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos + 1, vbat_y + k, labels[k]);
                    }
                }
            }
            x_pos = 3; y_pos = 3;
            for (int j = 0; j < temps.Length; j++)
            {
#if true    // 1. TempChamber Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[j]);
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + temps[j]);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == temps[j] || dVal[0] - dTHVal == temps[j] || dVal[0] == temps[j])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 60);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                for (int i = 0; i < VBAT_volts.Length; i++)
                {
                    PowerSupply0.Write("VOLT " + VBAT_volts[i] + ",(@2)");
                    System.Threading.Thread.Sleep(700);
                    for (double k = 0.150; k < 0.751; k += 0.001)
                    {
                        PowerSupply0.Write("VOLT " + k.ToString() + ",(@3)");
                        System.Threading.Thread.Sleep(700);

                        SendCommand("ook 80\n");
                        System.Threading.Thread.Sleep(100);

                        B26_BL.Read();
                        System.Threading.Thread.Sleep(10);
                        Parent.xlMgr.Cell.Write(x_pos, y_pos + (y_increment * j), (B26_BL.Value).ToString());

                        x_pos++;
                    }
                    x_pos = 3;
                    y_pos += labels.Length;
                }
                y_pos = 3;
            }
            TempChamber.WriteAndReadString("01,TEMP,S23.5");
        }

        private void ADC_VREF_Cal_EXT_Forcing()
        {
            int cal_y_increment, forcing_y_increment;
            double top_target_mv = 750, vcm_target_mv = 450, bot_target_mv = 150;
            double d_diff_mv, d_volt_mv;
            double[] dVal = new double[5];
            string sTempVal;
            bool bRetrunTemp = true;
            uint rev_id, top_val, bot_val, vcm_val;
            int cal_x_pos = 1, cal_y_pos = 1, forcing_x_pos = 1, forcing_y_pos = 1;
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            string[] cal_labels =
            {
                "Volt[mV]",  "Trim Value"
            };
            string[] forcing_labels =
            {
                "VIN[mV]",  "ADC code", "BGR", "ADC_TOP", "ADC_CM", "ADC_BOT"
            },
            cal_vbats =
            {
                "2.5"
            },
            forcing_vbats =
            {
                "1.7", "2.5", "3.3"
            };
            double[] cal_temps =
            {
                23.5, 23.5
            };
            double[] forcing_temps =
            {
                23.5, 23.5
            };

            string[] vrefs = { "VREF_TOP", "VREF_BOT", "VREF_CM" };
            string[] EXT_volts1 = { "0.15", "0.2", "0.250", "0.300", "0.350", "0.400", "0.450", "0.500", "0.550", "0.600", "0.650", "0.700", "0.750" };
            string[] EXT_volts2 = { "0.2", "0.250", "0.300", "0.350", "0.400", "0.450", "0.500", "0.550", "0.600", "0.650", "0.700" };

            Check_Instrument();

            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                     // 0x04
            RegisterItem B17_ES = Parent.RegMgr.GetRegisterItem("B17_ES[7:0]");                         // 0x1B

            RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A

            RegisterItem O_ADC_VREFT_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_VOLT<6:0>");     // 0x47
            RegisterItem O_ADC_VCM_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VCM_VOLT<7:0>");         // 0x48
            RegisterItem O_ADC_VREFB_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_VOLT<6:0>");     // 0x49

            RegisterItem O_ADC_TEST = Parent.RegMgr.GetRegisterItem("O_ADC_TEST<2:0>");                 // 0x4A
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[9]");                   // 0x5E
            RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");                 // 0x61

            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                       // 0x6C

            I_DEVID.Read();
            B04_BA_0.Read();
            rev_id = I_DEVID.Value;
            if (rev_id == 180)      // R5V1
            {
                Log.WriteLine("Device : IRIS R5V1");
                Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
                Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_ADC_EXT_Forcing");
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
            }
            else if (rev_id == 212) // R5V2
            {
                Log.WriteLine("Device : IRIS R5V2");
                Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
                Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_ADC_EXT_Forcing");
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
            }
            else
            {
                Log.WriteLine("Device : Null");
                return;
            }

            // Calibration Sheet
            cal_y_increment = (cal_vbats.Length * cal_labels.Length) + 2;

            if (rev_id == 180)      // R5V1
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
            else                    // R5V2
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");

            for (int i = 0; i < cal_temps.Length; i++)
            {
                int current_y = cal_y_pos + (cal_y_increment * i);

                Parent.xlMgr.Cell.Write(cal_x_pos, current_y, "Chamber Temp[℃]");
                Parent.xlMgr.Cell.Write(cal_x_pos + 1, current_y, cal_temps[i].ToString());
                Parent.xlMgr.Cell.Write(cal_x_pos, current_y + 1, "VBAT[V]");

                for (int c = 0; c < vrefs.Length; c++)
                {
                    Parent.xlMgr.Cell.Write(cal_x_pos + 2 + c, current_y + 1, vrefs[c]);
                }

                for (int j = 0; j < cal_vbats.Length; j++)
                {
                    int vbat_y = current_y + 2 + j * cal_labels.Length;

                    Parent.xlMgr.Cell.Write(cal_x_pos, vbat_y, cal_vbats[j].ToString());

                    for (int k = 0; k < cal_labels.Length; k++)
                    {
                        Parent.xlMgr.Cell.Write(cal_x_pos + 1, vbat_y + k, cal_labels[k]);
                    }
                }
            }

            // Forcing Sheet
            forcing_y_increment = 1 + forcing_vbats.Length * forcing_labels.Length + 1;

            if (rev_id == 180)      // R5V1
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_ADC_EXT_Forcing");
            else                    // R5V2
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_ADC_EXT_Forcing");


            for (int i = 0; i < forcing_temps.Length; i++)
            {
                int current_y = forcing_y_pos + (forcing_y_increment * i);

                Parent.xlMgr.Cell.Write(forcing_x_pos, current_y, "Chamber Temp[℃]");
                Parent.xlMgr.Cell.Write(forcing_x_pos + 1, current_y, forcing_temps[i].ToString());
                Parent.xlMgr.Cell.Write(forcing_x_pos, current_y + 1, "VBAT[V]");

                for (int c = 0; c < EXT_volts1.Length; c++)
                {
                    Parent.xlMgr.Cell.Write(forcing_x_pos + 2 + c, current_y + 1, EXT_volts1[c]);
                }

                for (int j = 0; j < forcing_vbats.Length; j++)
                {
                    int vbat_y = current_y + 2 + j * forcing_labels.Length;

                    Parent.xlMgr.Cell.Write(forcing_x_pos, vbat_y, forcing_vbats[j].ToString());

                    for (int k = 0; k < forcing_labels.Length; k++)
                    {
                        Parent.xlMgr.Cell.Write(forcing_x_pos + 1, vbat_y + k, forcing_labels[k]);
                    }
                }
            }

            cal_x_pos = 3; cal_y_pos = 3;
            forcing_x_pos = 3; forcing_y_pos = 3;

            ITEST_CONT.Value = 1;
            ITEST_CONT.Write();
            O_TEST_CON.Value = 0;
            O_TEST_CON.Write();

            for (int j = 0; j < forcing_temps.Length; j++)
            {
#if true    // 1. TempChamber Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + forcing_temps[j]);
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + forcing_temps[j]);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == forcing_temps[j] || dVal[0] - dTHVal == forcing_temps[j] || dVal[0] == forcing_temps[j])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 60);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                // VREF Calibration
                if (j == 0)
                {
                    if (rev_id == 180)      // R5V1
                        Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
                    else                    // R5V2
                        Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");

                    PowerSupply0.Write("VOLT " + cal_vbats[0] + ",(@2)");
                    System.Threading.Thread.Sleep(700);

                    O_ADC_TEST.Value = 4;
                    O_ADC_TEST.Write();
                    d_diff_mv = 5000;
                    top_val = 0;
                    for (uint val = 10; val < 51; val++)
                    {
                        O_ADC_VREFT_VOLT.Value = val;
                        O_ADC_VREFT_VOLT.Write();
                        System.Threading.Thread.Sleep(100);
                        I_PEN(true);
                        System.Threading.Thread.Sleep(100);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        if (Math.Abs(d_volt_mv - top_target_mv) < d_diff_mv)
                        {
                            d_diff_mv = Math.Abs(d_volt_mv - top_target_mv);
                            top_val = val;
                        }
                        I_PEN(false);
                    }
                    O_ADC_VREFT_VOLT.Value = top_val;
                    O_ADC_VREFT_VOLT.Write();
                    System.Threading.Thread.Sleep(100);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(100);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 0, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 0, cal_y_pos + 1 + (cal_y_increment * j), top_val.ToString());

                    // BOT Calibration
                    O_ADC_TEST.Value = 1;
                    O_ADC_TEST.Write();
                    d_diff_mv = 5000;
                    bot_val = 0;
                    for (uint val = 10; val < 51; val++)
                    {
                        O_ADC_VREFB_VOLT.Value = val;
                        O_ADC_VREFB_VOLT.Write();
                        System.Threading.Thread.Sleep(100);
                        I_PEN(true);
                        System.Threading.Thread.Sleep(100);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        if (Math.Abs(d_volt_mv - bot_target_mv) < d_diff_mv)
                        {
                            d_diff_mv = Math.Abs(d_volt_mv - bot_target_mv);
                            bot_val = val;
                        }
                        I_PEN(false);
                    }
                    O_ADC_VREFB_VOLT.Value = bot_val;
                    O_ADC_VREFB_VOLT.Write();
                    System.Threading.Thread.Sleep(100);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(100);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 1, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 1, cal_y_pos + 1 + (cal_y_increment * j), bot_val.ToString());

                    // VCM Calibration
                    O_ADC_TEST.Value = 2;
                    O_ADC_TEST.Write();
                    d_diff_mv = 5000;
                    vcm_val = 0;
                    for (uint val = 110; val < 201; val++)
                    {
                        O_ADC_VCM_VOLT.Value = val;
                        O_ADC_VCM_VOLT.Write();
                        System.Threading.Thread.Sleep(100);
                        I_PEN(true);
                        System.Threading.Thread.Sleep(100);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        if (Math.Abs(d_volt_mv - vcm_target_mv) < d_diff_mv)
                        {
                            d_diff_mv = Math.Abs(d_volt_mv - vcm_target_mv);
                            vcm_val = val;
                        }
                        I_PEN(false);
                    }
                    O_ADC_VCM_VOLT.Value = vcm_val;
                    O_ADC_VCM_VOLT.Write();
                    System.Threading.Thread.Sleep(100);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(100);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 2, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 2, cal_y_pos + 1 + (cal_y_increment * j), vcm_val.ToString());

                    cal_y_pos += cal_labels.Length;
                    O_ADC_TEST.Value = 0;
                    O_ADC_TEST.Write();
                }
                for (int i = 0; i < forcing_vbats.Length; i++)
                {
                    PowerSupply0.Write("VOLT " + forcing_vbats[i] + ",(@2)");
                    System.Threading.Thread.Sleep(700);

                    // Forcing Measure
                    if (rev_id == 180)      // R5V1
                        Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_ADC_EXT_Forcing");
                    else                    // R5V2
                        Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_ADC_EXT_Forcing");

                    if (j == 0)
                    {
                        for (int k = 0; k < EXT_volts1.Length; k++)
                        {
                            PowerSupply0.Write("VOLT " + EXT_volts1[k] + ",(@3)");
                            System.Threading.Thread.Sleep(700);

                            TestInOut_ADC_In(true);
                            O_EXT_VOLT_EN.Read();
                            O_EXT_VOLT_EN.Value = 1;
                            O_EXT_VOLT_EN.Write();
                            System.Threading.Thread.Sleep(100);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(100);

                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);
                            TestInOut_ADC_In(false);

                            System.Threading.Thread.Sleep(100);

                            SendCommand("ook 80\n");
                            System.Threading.Thread.Sleep(100);

                            B17_ES.Read();
                            System.Threading.Thread.Sleep(10);
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 1 + (forcing_y_increment * j), (B17_ES.Value).ToString());
                            O_EXT_VOLT_EN.Value = 0;
                            O_EXT_VOLT_EN.Write();

                            Set_TestInOut_For_BGR(true);
                            System.Threading.Thread.Sleep(500);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 2 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            Set_TestInOut_For_BGR(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 4;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 3 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 2;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 4 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 1;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 5 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 0;
                            O_ADC_TEST.Write();
                        }
                    }
                    else if (j == 1)
                    {
                        forcing_x_pos += 1;
                        for (int k = 0; k < EXT_volts2.Length; k++)
                        {
                            PowerSupply0.Write("VOLT " + EXT_volts2[k] + ",(@3)");
                            System.Threading.Thread.Sleep(700);

                            TestInOut_ADC_In(true);
                            O_EXT_VOLT_EN.Read();
                            O_EXT_VOLT_EN.Value = 1;
                            O_EXT_VOLT_EN.Write();
                            System.Threading.Thread.Sleep(100);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(100);

                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);
                            TestInOut_ADC_In(false);

                            System.Threading.Thread.Sleep(100);

                            SendCommand("ook 80\n");
                            System.Threading.Thread.Sleep(100);

                            B17_ES.Read();
                            System.Threading.Thread.Sleep(10);
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 1 + (forcing_y_increment * j), (B17_ES.Value).ToString());
                            O_EXT_VOLT_EN.Value = 0;
                            O_EXT_VOLT_EN.Write();

                            Set_TestInOut_For_BGR(true);
                            System.Threading.Thread.Sleep(500);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 2 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            Set_TestInOut_For_BGR(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 4;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 3 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 2;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 4 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 1;
                            O_ADC_TEST.Write();
                            System.Threading.Thread.Sleep(10);
                            I_PEN(true);
                            System.Threading.Thread.Sleep(10);
                            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            Parent.xlMgr.Cell.Write(forcing_x_pos + k, forcing_y_pos + 5 + (forcing_y_increment * j), d_volt_mv.ToString("F3"));
                            I_PEN(false);

                            O_ADC_TEST.Read();
                            O_ADC_TEST.Value = 0;
                            O_ADC_TEST.Write();
                        }

                    }
                    forcing_y_pos += forcing_labels.Length;
                }
                cal_y_pos = 3;
                forcing_y_pos = 3;
            }
            TempChamber.WriteAndReadString("01,TEMP,S23.5");
        }

        private void ADC_VREF_Calibration()
        {
            int cal_y_increment;
            double top_target_mv = 750, vcm_target_mv = 450, bot_target_mv = 150;
            double d_diff_mv, d_volt_mv;
            double[] dVal = new double[5];
            string sTempVal;
            bool bRetrunTemp = true;
            uint rev_id, top_val, bot_val, vcm_val;
            int cal_x_pos = 1, cal_y_pos = 1;
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            string[] cal_labels =
            {
                "Volt[mV]",  "Trim Value"
            },
            cal_vbats =
            {
                "2.5"
            };
            double[] cal_temps =
{
                23.5
            };

            string[] vrefs = { "VOLT_TOP", "VOLT_BOT", "VOLT_CM", "TEMP_TOP", "TEMP_BOT", "TEMP_CM" };

            Check_Instrument();

            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                     // 0x04
            RegisterItem B17_ES = Parent.RegMgr.GetRegisterItem("B17_ES[7:0]");                         // 0x1B

            RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A

            RegisterItem O_ADC_VREFT_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_TEMP<6:0>");     // 0x44
            RegisterItem O_ADC_VCM_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VCM_TEMP<7:0>");         // 0x45
            RegisterItem O_ADC_VREFB_TEMP = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_TEMP<6:0>");     // 0x46

            RegisterItem w_ADC_SWITCH_MODE = Parent.RegMgr.GetRegisterItem("w_ADC_SWITCH_MODE");        // 0x46

            RegisterItem O_ADC_VREFT_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFT_VOLT<6:0>");     // 0x47
            RegisterItem O_ADC_VCM_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VCM_VOLT<7:0>");         // 0x48
            RegisterItem O_ADC_VREFB_VOLT = Parent.RegMgr.GetRegisterItem("O_ADC_VREFB_VOLT<6:0>");     // 0x49

            RegisterItem O_ADC_TEST = Parent.RegMgr.GetRegisterItem("O_ADC_TEST<2:0>");                 // 0x4A
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[9]");                   // 0x5E
            RegisterItem O_TEST_CON = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:0]");                 // 0x61

            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                       // 0x6C

            I_DEVID.Read();
            B04_BA_0.Read();
            rev_id = I_DEVID.Value;
            if (rev_id == 180)      // R5V1
            {
                Log.WriteLine("Device : IRIS R5V1");
                Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
            }
            else if (rev_id == 212) // R5V2
            {
                Log.WriteLine("Device : IRIS R5V2");
                Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
            }
            else
            {
                Log.WriteLine("Device : Null");
                return;
            }

            // Calibration Sheet
            cal_y_increment = (cal_vbats.Length * cal_labels.Length) + 2;

            if (rev_id == 180)      // R5V1
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
            else                    // R5V2
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");

            for (int i = 0; i < cal_temps.Length; i++)
            {
                int current_y = cal_y_pos + (cal_y_increment * i);

                Parent.xlMgr.Cell.Write(cal_x_pos, current_y, "Chamber Temp[℃]");
                Parent.xlMgr.Cell.Write(cal_x_pos + 1, current_y, cal_temps[i].ToString());
                Parent.xlMgr.Cell.Write(cal_x_pos, current_y + 1, "VBAT[V]");

                for (int c = 0; c < vrefs.Length; c++)
                {
                    Parent.xlMgr.Cell.Write(cal_x_pos + 2 + c, current_y + 1, vrefs[c]);
                }

                for (int j = 0; j < cal_vbats.Length; j++)
                {
                    int vbat_y = current_y + 2 + j * cal_labels.Length;

                    Parent.xlMgr.Cell.Write(cal_x_pos, vbat_y, cal_vbats[j].ToString());

                    for (int k = 0; k < cal_labels.Length; k++)
                    {
                        Parent.xlMgr.Cell.Write(cal_x_pos + 1, vbat_y + k, cal_labels[k]);
                    }
                }
            }

            cal_x_pos = 3; cal_y_pos = 3;

            ITEST_CONT.Value = 1;
            ITEST_CONT.Write();
            O_TEST_CON.Value = 0;
            O_TEST_CON.Write();

            for (int j = 0; j < cal_temps.Length; j++)
            {
#if false    // 1. TempChamber Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + cal_temps[j]);
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + cal_temps[j]);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == cal_temps[j] || dVal[0] - dTHVal == cal_temps[j] || dVal[0] == cal_temps[j])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 60);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                // VREF Calibration
                if (j == 0)
                {
                    if (rev_id == 180)      // R5V1
                        Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");
                    else                    // R5V2
                        Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_VREF_Calibration");

                    PowerSupply0.Write("VOLT " + cal_vbats[0] + ",(@2)");
                    System.Threading.Thread.Sleep(700);

                    // TOP Calibration
                    O_ADC_TEST.Value = 4;
                    O_ADC_TEST.Write();
                    d_diff_mv = 5000;
                    top_val = 0;
                    for (uint val = 0; val < 51; val++)
                    {
                        O_ADC_VREFT_VOLT.Value = val;
                        O_ADC_VREFT_VOLT.Write();
                        System.Threading.Thread.Sleep(50);
                        I_PEN(true);
                        System.Threading.Thread.Sleep(150);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        if (Math.Abs(d_volt_mv - top_target_mv) < d_diff_mv)
                        {
                            d_diff_mv = Math.Abs(d_volt_mv - top_target_mv);
                            top_val = val;
                        }
                        I_PEN(false);
                    }
                    O_ADC_VREFT_VOLT.Value = top_val;
                    O_ADC_VREFT_TEMP.Value = top_val;
                    O_ADC_VREFT_VOLT.Write();
                    O_ADC_VREFT_TEMP.Write();
                    System.Threading.Thread.Sleep(50);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(150);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 0, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 0, cal_y_pos + 1 + (cal_y_increment * j), top_val.ToString());
                    I_PEN(false);
                    w_ADC_SWITCH_MODE.Value = 0;
                    w_ADC_SWITCH_MODE.Write();
                    System.Threading.Thread.Sleep(50);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(150);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 3, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 3, cal_y_pos + 1 + (cal_y_increment * j), top_val.ToString());
                    I_PEN(false);

                    // BOT Calibration
                    O_ADC_TEST.Value = 1;
                    O_ADC_TEST.Write();
                    d_diff_mv = 5000;
                    bot_val = 0;
                    for (uint val = 0; val < 51; val++)
                    {
                        O_ADC_VREFB_VOLT.Value = val;
                        O_ADC_VREFB_VOLT.Write();
                        System.Threading.Thread.Sleep(50);
                        I_PEN(true);
                        System.Threading.Thread.Sleep(150);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        if (Math.Abs(d_volt_mv - bot_target_mv) < d_diff_mv)
                        {
                            d_diff_mv = Math.Abs(d_volt_mv - bot_target_mv);
                            bot_val = val;
                        }
                        I_PEN(false);
                    }
                    O_ADC_VREFB_VOLT.Value = bot_val;
                    O_ADC_VREFB_TEMP.Value = bot_val;
                    O_ADC_VREFB_VOLT.Write();
                    O_ADC_VREFB_TEMP.Write();
                    System.Threading.Thread.Sleep(50);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(150);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 1, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 1, cal_y_pos + 1 + (cal_y_increment * j), bot_val.ToString());
                    I_PEN(false);
                    w_ADC_SWITCH_MODE.Value = 0;
                    w_ADC_SWITCH_MODE.Write();
                    System.Threading.Thread.Sleep(50);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(150);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 4, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 4, cal_y_pos + 1 + (cal_y_increment * j), bot_val.ToString());
                    I_PEN(false);


                    // VCM Calibration
                    O_ADC_TEST.Value = 2;
                    O_ADC_TEST.Write();
                    d_diff_mv = 5000;
                    vcm_val = 0;
                    for (uint val = 110; val < 201; val++)
                    {
                        O_ADC_VCM_VOLT.Value = val;
                        O_ADC_VCM_VOLT.Write();
                        System.Threading.Thread.Sleep(50);
                        I_PEN(true);
                        System.Threading.Thread.Sleep(150);
                        d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        if (Math.Abs(d_volt_mv - vcm_target_mv) < d_diff_mv)
                        {
                            d_diff_mv = Math.Abs(d_volt_mv - vcm_target_mv);
                            vcm_val = val;
                        }
                        I_PEN(false);
                    }
                    O_ADC_VCM_VOLT.Value = vcm_val;
                    O_ADC_VCM_TEMP.Value = vcm_val;
                    O_ADC_VCM_VOLT.Write();
                    O_ADC_VCM_TEMP.Write();
                    System.Threading.Thread.Sleep(50);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(150);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 2, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 2, cal_y_pos + 1 + (cal_y_increment * j), vcm_val.ToString());
                    I_PEN(false);
                    w_ADC_SWITCH_MODE.Value = 0;
                    w_ADC_SWITCH_MODE.Write();
                    System.Threading.Thread.Sleep(50);
                    I_PEN(true);
                    System.Threading.Thread.Sleep(150);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(cal_x_pos + 5, cal_y_pos + 0 + (cal_y_increment * j), d_volt_mv.ToString("F3"));
                    Parent.xlMgr.Cell.Write(cal_x_pos + 5, cal_y_pos + 1 + (cal_y_increment * j), vcm_val.ToString());
                    I_PEN(false);

                    cal_y_pos += cal_labels.Length;
                    O_ADC_TEST.Value = 0;
                    O_ADC_TEST.Write();
                }
                cal_y_pos = 3;
            }
            ITEST_CONT.Value = 0;
            ITEST_CONT.Write();
            w_ADC_SWITCH_MODE.Value = 1;
            w_ADC_SWITCH_MODE.Write();
            System.Threading.Thread.Sleep(50);
            TempChamber.WriteAndReadString("01,TEMP,S23.5");
        }

        private void EXT_SENS_ADC_In()
        {
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                     // 0x04
            RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                       // 0x6C

            string[] vbat =
            {
                "1.7", "2.5", "3.3"
            };
            string[] ext_sens = { "0.15", "0.2", "0.250", "0.300", "0.350", "0.400", "0.450", "0.500", "0.550", "0.600", "0.650", "0.700", "0.750" };
            double d_volt_mv;

            Check_Instrument();

            TestInOut_ADC_In(true);

            O_EXT_VOLT_EN.Read();
            O_EXT_VOLT_EN.Value = 1;
            O_EXT_VOLT_EN.Write();

            I_PEN(true);

            I_DEVID.Read();
            B04_BA_0.Read();

            uint rev_id = I_DEVID.Value;
            if (rev_id == 180)      // R5V1
            {
                Log.WriteLine("Device : IRIS R5V1");
                Parent.xlMgr.Sheet.Add("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_EXT_ADC_In");
                Parent.xlMgr.Sheet.Select("R5V1_" + (B04_BA_0.Value).ToString("X2") + "_EXT_ADC_In");
            }
            else if (rev_id == 212) // R5V2
            {
                Log.WriteLine("Device : IRIS R5V2");
                Parent.xlMgr.Sheet.Add("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_EXT_ADC_In");
                Parent.xlMgr.Sheet.Select("R5V2_" + (B04_BA_0.Value).ToString("X2") + "_EXT_ADC_In");
            }
            else
            {
                Log.WriteLine("Device : Null");
                return;
            }

            for (int i = 0; i < vbat.Length; i++)
            {
                PowerSupply0.Write("VOLT " + vbat[i] + ",(@2)");
                for (int j = 0; j < ext_sens.Length; j++)
                {
                    PowerSupply0.Write("VOLT " + ext_sens[j] + ",(@3)");
                    System.Threading.Thread.Sleep(50);
                    d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Parent.xlMgr.Cell.Write(j + 1, i + 1, d_volt_mv.ToString("F3"));
                }
            }
        }

        private void I_PEN(bool on)
        {
            RegisterItem O_EXT_VOLT_EN = Parent.RegMgr.GetRegisterItem("O_EXT_VOLT_EN");                // 0x2A

            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (on == true)
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x00;
                SendData[5] = 0x00;
                SendData[6] = 0x00;
                SendData[7] = 0x8A;
                I2C.WriteBytes(SendData, SendData.Length, true);
            }
            else
            {
                SendData[0] = 0x00;
                SendData[1] = 0x00;
                SendData[2] = 0x03;
                SendData[3] = 0x00;
                SendData[4] = 0x00;
                SendData[5] = 0x00;
                SendData[6] = 0x00;
                SendData[7] = 0x0A;
                I2C.WriteBytes(SendData, SendData.Length, true);

                O_EXT_VOLT_EN.Read();
                O_EXT_VOLT_EN.Value = 0;
                O_EXT_VOLT_EN.Write();
            }
        }

        private void BGR_Temp_Sweep_Multi()
        {
            double d_volt_mv;
            uint volt_adc, temp_adc;
            double[] dVal = new double[5];
            int x_pos = 1, y_pos = 1;
            uint[] chip_num = { 0, 1, 2, 3 };
            string sTempVal, start_time;
            bool bRetrunTemp = true;
            uint[,] reg_val = new uint[16, 8];
            double[] temps =
            {
                87.5,   85,     80,     75,     70,
                65,     60,     56.5,     55,     50,     45,
                40,     35,     30,     25,     20,
                15,     10,     5,      0,      -5,     -8.5,
                -10,    -15,    -20,    -25,    -30,
                -35,    -40,    23.5
            };
            string[] volts =
            {
                "1.7", "2.5", "3.3"
            };
            string[] labels =
            {
                "BGR[㎷]", "Volt[㎷]", "VoltADC", "Temp[㎷]", "TempADC"
            };
            JLcLib.Custom.I2C[] i2c = { I2C, I2C1, I2C2, I2C3 };
            /*
            // MAC5
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");                   // 0x04
            // AL BGR
            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");         // 0x57
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL<5:2>");             // 0x67
            // Rev_ID
            RegisterItem I_DEVID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");                     // 0x6C
            */

            Check_Instrument();
            JLcLib.Instrument.SCPI[] DMM = { DigitalMultimeter0, DigitalMultimeter1, DigitalMultimeter2, DigitalMultimeter3, DigitalMultimeter4, DigitalMultimeter5 };
            start_time = DateTime.Now.ToString("MMddHHmm");
            Log.WriteLine("Test : IRIS R5 ALBGR Temp Measure");
            Parent.xlMgr.Sheet.Add(start_time + "_ADC_Temp");

            // MAC Check & Excel Sheet Add
            for (int i = 0; i < i2c.Length; i++)
            {
                for (int j = 0; j < i2c.Length; j++)
                {
                    i2c[j].GPIOs[4].Direction = GPIO_Direction.Output;
                    i2c[j].GPIOs[4].State = GPIO_State.High;
                }
                i2c[i].GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(100);
                WakeUp_I2C_Multi(i2c[i]);
                System.Threading.Thread.Sleep(1000);
                chip_num[i] = ReadRegister_Multi(0x04, i2c[i]);

                x_pos = 1; y_pos = 2;
                Parent.xlMgr.Cell.Write(x_pos, 1 + (17 * i), "Chip " + chip_num[i].ToString("D0"));
                Parent.xlMgr.Cell.Write(x_pos + 1, 1 + (17 * i), "MAC : 0x" + chip_num[i].ToString("X0"));

                for (int k = 0; k < i2c.Length; k++)
                {
                    for (int j = 0; j < volts.Length; j++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos, y_pos, volts[j].ToString());
                        for (int q = 0; q < labels.Length; q++)
                        {
                            Parent.xlMgr.Cell.Write(x_pos + 1, y_pos++, labels[q].ToString());
                        }
                    }
                    y_pos += 2;
                }

                x_pos = 3; y_pos = 1;
                for (int k = 0; k < i2c.Length; k++)
                {
                    for (int j = 0; j < temps.Length; j++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos + j, y_pos, temps[j].ToString());
                    }
                    y_pos += 17;
                }
                i2c[i].GPIOs[4].State = GPIO_State.High;
            }
            for (int j = 0; j < temps.Length; j++)
            {
#if true    // 1. TempChamber Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[j]);
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + temps[j]);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == temps[j] || dVal[0] - dTHVal == temps[j] || dVal[0] == temps[j])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 300);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                for (int Boardcnt = 0; Boardcnt < chip_num.Length; Boardcnt++)
                {
                    for (int k = 0; k < i2c.Length; k++)
                    {
                        i2c[k].GPIOs[4].State = GPIO_State.High;
                    }
                    i2c[Boardcnt].GPIOs[4].State = GPIO_State.Low;

                    for (int k = 0; k < volts.Length; k++)
                    {
                        if (k == 0)
                        {
                            PowerSupply0.Write("VOLT 2.5,(@2)");
                            System.Threading.Thread.Sleep(250);
                            PowerSupply0.Write("VOLT 1.7,(@2)");
                            System.Threading.Thread.Sleep(250);
                        }
                        else
                        {
                            PowerSupply0.Write("VOLT " + volts[k] + ",(@2)");
                            System.Threading.Thread.Sleep(500);

                        }
                        Write_WURF_AON_Multi("80", i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(100);

                        Set_TestInOut_For_BGR_Multi(true, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);

                        d_volt_mv = double.Parse(DMM[Boardcnt].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        Parent.xlMgr.Cell.Write(3 + j, 2 + (Boardcnt * 17) + (k * 5), d_volt_mv.ToString("F2"));

                        Set_TestInOut_For_BGR_Multi(false, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);

                        TestInOut_ADC_In_Multi(true, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);

                        TestInOut_Voltagesensor_Multi(true, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);
                        d_volt_mv = double.Parse(DMM[Boardcnt].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        Parent.xlMgr.Cell.Write(3 + j, 2 + (Boardcnt * 17) + (k * 5) + 1, d_volt_mv.ToString("F2"));
                        TestInOut_Voltagesensor_Multi(false, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);

                        volt_adc = ReadRegister_Multi(0x1A, i2c[Boardcnt]);
                        Parent.xlMgr.Cell.Write(3 + j, 2 + (Boardcnt * 17) + (k * 5) + 2, volt_adc.ToString("D0"));

                        TestInOut_Tempsensor_Multi(true, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);
                        d_volt_mv = double.Parse(DMM[Boardcnt].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        Parent.xlMgr.Cell.Write(3 + j, 2 + (Boardcnt * 17) + (k * 5) + 3, d_volt_mv.ToString("F2"));
                        TestInOut_Tempsensor_Multi(false, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);

                        TestInOut_ADC_In_Multi(false, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(200);

                        temp_adc = ReadRegister_Multi(0x1B, i2c[Boardcnt]);
                        Parent.xlMgr.Cell.Write(3 + j, 2 + (Boardcnt * 17) + (k * 5) + 4, temp_adc.ToString("D0"));
                    }
                }
            }
            Log.WriteLine("Test Complete : " + DateTime.Now.ToString("MM/dd/HH:mm"));
        }

        private void PMU_Temp_Sweep_Multi()
        {
            double d_volt_mv;
            double[] dVal = new double[5];
            int x_pos = 1, y_pos = 1;
            uint[] chip_num = { 0, 1 };
            string sTempVal, start_time;
            bool bRetrunTemp = true;
            uint[,] reg_val = new uint[16, 8];
            double[] temps =
            {
                87.5,   85,     80,     75,     70,
                65,     60,     55,     50,     45,
                40,     35,     30,     25,     20,
                15,     10,     5,      0,      -5,
                -10,    -15,    -20,    -25,    -30,
                -35,    -40
            };
            double[] volts =
            {
                1.7, 1.8, 1.9, 2, 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 3, 3.1, 3.2, 3.3, 3.4, 3.5,
            };
            string[] labels =
            {
                "ALLDO[㎷]", "MLDO[㎷]", "BGR[㎷]", "RCOSC[㎑]"
            };
            JLcLib.Custom.I2C[] i2c = { I2C, I2C1 };

            Check_Instrument();
            JLcLib.Instrument.SCPI[] DMM = { DigitalMultimeter0, DigitalMultimeter1, DigitalMultimeter2, DigitalMultimeter3, DigitalMultimeter4, DigitalMultimeter5 };
            start_time = DateTime.Now.ToString("MMddHHmm");
            Log.WriteLine("Test : IRIS R5 PMU Temp Measure");

            PowerSupply0.Write("VOLT 2.5,(@2)");
            PowerSupply0.Write("OUTP ON,(@2)");

            for (int i = 0; i < i2c.Length; i++)
            {
                for (int j = 0; j < i2c.Length; j++)
                {
                    i2c[j].GPIOs[4].Direction = GPIO_Direction.Output;
                    i2c[j].GPIOs[4].State = GPIO_State.High;
                }
                i2c[i].GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(100);
                WakeUp_I2C_Multi(i2c[i]);
                System.Threading.Thread.Sleep(1000);
                chip_num[i] = ReadRegister_Multi(0x04, i2c[i]);
                Parent.xlMgr.Sheet.Add(start_time + "_PMU_Temp_" + chip_num[i].ToString("D0"));

                x_pos = 1; y_pos = 1;
                Parent.xlMgr.Cell.Write(x_pos + 0, y_pos, "Chip " + chip_num[i].ToString("D0"));
                Parent.xlMgr.Cell.Write(x_pos + 1, y_pos, "MAC : 0x" + chip_num[i].ToString("X0"));
                y_pos = 2;
                for (int k = 0; k < labels.Length; k++)
                {
                    Parent.xlMgr.Cell.Write(x_pos, y_pos + (k * 29), labels[k]);
                    for (int k2 = 0; k2 < temps.Length; k2++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos, y_pos + 1 + k2 + (k * 29), temps[k2].ToString());
                    }
                    for (int k3 = 0; k3 < volts.Length; k3++)
                    {
                        Parent.xlMgr.Cell.Write(x_pos + 1 + k3, y_pos + (k * 29), volts[k3].ToString());
                    }
                }
            }

            x_pos = 2; y_pos = 3;


            for (int j = 0; j < temps.Length; j++)
            {
#if true    // 1. TempChamber Set
                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temps[j]);
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Console.WriteLine("Run Chamber : " + temps[j]);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                bRetrunTemp = true;
                while (bRetrunTemp)
                {
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    System.Threading.Thread.Sleep(100);
                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    double dTHVal = 0.1;

                    if (dVal[0] + dTHVal == temps[j] || dVal[0] - dTHVal == temps[j] || dVal[0] == temps[j])
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                        bRetrunTemp = true;
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 300);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
#endif
                for (int Boardcnt = 0; Boardcnt < chip_num.Length; Boardcnt++)
                {
                    Parent.xlMgr.Sheet.Select(start_time + "_PMU_Temp_" + chip_num[Boardcnt].ToString("D0"));

                    for (int k = 0; k < i2c.Length; k++)
                    {
                        i2c[k].GPIOs[4].State = GPIO_State.High;
                    }
                    i2c[Boardcnt].GPIOs[4].State = GPIO_State.Low;

                    for (int k = 0; k < volts.Length; k++)
                    {
                        if (k == 0)
                        {
                            PowerSupply0.Write("VOLT 2.5,(@2)");
                            System.Threading.Thread.Sleep(250);
                            PowerSupply0.Write("VOLT 1.7,(@2)");
                            System.Threading.Thread.Sleep(250);
                        }
                        else
                        {
                            PowerSupply0.Write("VOLT " + volts[k].ToString() + ",(@2)");
                            System.Threading.Thread.Sleep(500);
                        }

                        d_volt_mv = double.Parse(DMM[0 + (3 * Boardcnt)].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        Parent.xlMgr.Cell.Write(x_pos + k, y_pos + j + (29 * 0), d_volt_mv.ToString("F2"));

                        d_volt_mv = double.Parse(DMM[1 + (3 * Boardcnt)].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        Parent.xlMgr.Cell.Write(x_pos + k, y_pos + j + (29 * 1), d_volt_mv.ToString("F2"));

                        Set_TestInOut_For_BGR_Multi(true, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(100);

                        d_volt_mv = double.Parse(DMM[2 + (3 * Boardcnt)].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                        Parent.xlMgr.Cell.Write(x_pos + k, y_pos + j + (29 * 2), d_volt_mv.ToString("F2"));

                        Set_TestInOut_For_BGR_Multi(false, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(100);

                        Set_TestInOut_For_RCOSC_Multi(true, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(100);

                        d_volt_mv = double.Parse(DMM[2 + (3 * Boardcnt)].WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                        Parent.xlMgr.Cell.Write(x_pos + k, y_pos + j + (29 * 3), d_volt_mv.ToString("F3"));

                        Set_TestInOut_For_RCOSC_Multi(false, i2c[Boardcnt]);
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            Log.WriteLine("Test Complete : " + DateTime.Now.ToString("MM/dd/HH:mm"));
            TempChamber.WriteAndReadString("01,TEMP,S23.5");
        }

        private void Calibration_OTP_Write(int start_cnt)
        {
            int cnt;
            if (start_cnt == 0)
            {
                cnt = 50;
            }
            else
            {
                cnt = start_cnt;
            }
            Check_Instrument();
            I2C.GPIOs[4].Direction = GPIO_Direction.Output;
            System.Threading.Thread.Sleep(10);
            I2C.GPIOs[4].State = GPIO_State.Low;
            System.Threading.Thread.Sleep(10);

            PowerSupply0.Write("VOLT 2.5,(@2)");
            PowerSupply0.Write("OUTP ON,(@2)");
            System.Threading.Thread.Sleep(500);

            SendCommand("mode ook\n");
            DialogResult dialog = MessageBox.Show("IRIS R5 MAC Write/LDO Calibration.\r\n\r\nChip Num\t: " + cnt
                                                    , Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dialog != DialogResult.OK)
            {
                return;
            }

            WakeUp_I2C();
            System.Threading.Thread.Sleep(200);

            Write_BLE00_OTP(cnt);

            Cal_PMU_FastCal(cnt);
        }
        #endregion
    }

    public class SCP1501_N2 : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS_MANUAL
        {
            Write_WURF,
            TX_CH_SEL,
            TX_ON,
            TX_OFF,
            SET_BLE,
            Read_FSM,
            Read_ADC,
            Read_Volt_Temp,
            TESTINOUT_Temp,
            TESTINOUT_Volt,
            TESTINOUT_BGR,
            TESTINOUT_EDOUT,
            TESTINOUT_RCOSC,
            TESTINOUT_Disable,
            NVM_POWER_ON,
            NVM_POWER_OFF,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_AUTO
        {
            TEST,
            NUM_TEST_ITEMS,
        }

        public enum TEST_ITEMS_DTM
        {
            SET_CH,
            SET_Length,
            DTM_START_PRNBS9,
            DTM_START_11110000,
            DTM_START_10101010,
            DTM_START_PRNBS15,
            DTM_START_ALL_1,
            DTM_START_ALL_0,
            DTM_START_00001111,
            DTM_START_0101,
            DTM_STOP,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            MANUAL,
            AUTO,
            DTM,
        }

        private JLcLib.Custom.I2C I2C { get; set; }

        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        public int SlaveAddress { get; private set; } = 0x3A;
        public int DTM_PayLoadLength = 37;
        public int DTM_Channel = 0;

        /* Intrument */
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

        public SCP1501_N2(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "SCP1501_N2");
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

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[8];
            I2C.Config.SlaveAddress = SlaveAddress;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "AON_REG(BLE)":
                case "AON_REG(Control)":
                case "AON_REG(Analog)":
                    // Write Address, Data
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x07);            // AREG_SEL=1, AREG_CE=1, AREG_WE=1
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);     // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);  // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);            // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[0] = (byte)(0x08);            // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);            // AREG_WDATA[7:0]
                    SendData[5] = (byte)(0x00);            // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);            // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
                case "NVM":
                    // bist_sel=1, bist_type=3, bist_cmd=0, bist_addr=address
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0d);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=3, bist_cmd=1, bist_addr=address
                    SendData[4] = (byte)(0x0f);
                    SendData[5] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_wdata=data
                    SendData[0] = (byte)(0x04);            // ADDR = 0x00060004
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=3, bist_cmd=0, bist_addr=address
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0d);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Dummy
                    System.Threading.Thread.Sleep(5);

                    // bist_sel=0, bist_type=3, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x0c);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
                default:
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    SendData[4] = (byte)(Data & 0xff);
                    SendData[5] = (byte)((Data >> 8) & 0xff);
                    SendData[6] = (byte)((Data >> 16) & 0xff);
                    SendData[7] = (byte)((Data >> 24) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
            }
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            uint result = 0xffffffff;

            switch (Parent.xlMgr.Sheet.Name)
            {
                case "AON_REG(BLE)":
                case "AON_REG(Control)":
                case "AON_REG(Analog)":
                    // Write Address
                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x01);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x00);             // AREG_WDATA[7:0]
                    SendData[5] = (byte)(Address & 0xff);   // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    SendData[6] = (byte)(0x05);             // AREG_SEL=1, AREG_CE=1, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Read Data
                    SendData[0] = (byte)(0x10);             // ADDR = 0x00010010
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = RcvData[3];

                    SendData[0] = (byte)(0x08);             // ADDR = 0x00010008
                    SendData[6] = (byte)(0x04);             // AREG_SEL=1, AREG_CE=0, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // AREG_SEL Disable
                    SendData[5] = (byte)(0x00);             // AREG_ADDR[7:0]
                    SendData[6] = (byte)(0x00);             // AREG_SEL=0, AREG_CE=0, AREG_WE=0
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
                case "NVM":
                    // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x11);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
                    SendData[4] = (byte)(0x13);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
                    SendData[4] = (byte)(0x11);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(Address & 0xff);
                    SendData[7] = (byte)((Address >> 8) & 0xff);
                    I2C.WriteBytes(SendData, SendData.Length, true);

                    // Read bist_rdata
                    SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);

                    // bist_sel=0, bist_type=4, bist_cmd=0, bist_addr=0
                    SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
                    SendData[1] = (byte)(0x00);
                    SendData[2] = (byte)(0x06);
                    SendData[3] = (byte)(0x00);
                    SendData[4] = (byte)(0x10);
                    SendData[5] = (byte)(0x00);
                    SendData[6] = (byte)(0x00);
                    SendData[7] = (byte)(0x00);
                    I2C.WriteBytes(SendData, SendData.Length, true);
                    break;
                default:
                    SendData[0] = (byte)(Address & 0xff);
                    SendData[1] = (byte)((Address >> 8) & 0xff);
                    SendData[2] = (byte)((Address >> 16) & 0xff);
                    SendData[3] = (byte)((Address >> 24) & 0xff);
                    I2C.WriteBytes(SendData, 4, false);
                    RcvData = I2C.ReadBytes(RcvData.Length);
                    result = (uint)(((RcvData[3] << 24)) | ((RcvData[2] << 16)) | ((RcvData[1] << 8)) | RcvData[0]);
                    break;
            }

            return result;
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
            Parent.ChipCtrlButtons[4].Text = "GH0_H";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;
            Parent.ChipCtrlButtons[5].Text = "WakeUp";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += WakeUp_I2C;
            Parent.ChipCtrlButtons[6].Text = "MD_OOK";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Set_Gecko_OOK_Mode;
            Parent.ChipCtrlButtons[8].Text = "Manual";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Change_To_Manual_Test_Items;
            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Change_To_Auto_Test_Items;
            Parent.ChipCtrlButtons[10].Text = "DTM";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Change_To_DTM_Test_Items;
            Parent.ChipCtrlButtons[10].Text = "VCO_T";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += VCO_TEST_autotest;

        }

        private void VCO_TEST_autotest(object sender, EventArgs e)
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
            Serial.WriteSettingFile(Parent.IniFile, "SCP1501_N2");
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

        private void Change_To_DTM_Test_Items(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.DTM;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS_DTM.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS_DTM)i).ToString());
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

        private void WakeUp_I2C(object sender, EventArgs e)
        {
            WakeUp_I2C();
        }

        private void Set_Gecko_OOK_Mode(object sender, EventArgs e)
        {
            SendCommand("mode ook");
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
                        case TEST_ITEMS_MANUAL.Write_WURF:
                            if (Arg == "")
                            {
                                Log.WriteLine("WURF에 write할 값을 16진수로 적어주세요.");
                            }
                            else
                            {
                                Log.WriteLine("OOK " + Arg);
                                Write_WURF_AON(Arg);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_CH_SEL:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                Log.WriteLine("Set CH" + Arg);
                                Write_Register_Fractional_Calc_Ch((uint)iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.TX_ON:
                            Log.WriteLine("TX ON");
                            Write_Register_Tx_Tone_Send(true);
                            break;
                        case TEST_ITEMS_MANUAL.TX_OFF:
                            Log.WriteLine("TX OFF");
                            Write_Register_Tx_Tone_Send(false);
                            break;
                        case TEST_ITEMS_MANUAL.SET_BLE:
                            if ((Arg == "") || (iVal < 0) || (iVal > 65535))
                            {
                                Log.WriteLine("WP값을 10진수로 적어주세요. (Range 0~65535)");
                            }
                            else
                            {
                                Log.WriteLine("Write AON for Advertising (WP = )" + Arg);
                                Write_Register_Send_Advertising(iVal);
                            }
                            break;
                        case TEST_ITEMS_MANUAL.Read_FSM:
                            Read_Register_FSM();
                            break;
                        case TEST_ITEMS_MANUAL.Read_ADC:
                            Result = Read_ADC_Result(false, 1, 63, 31);
                            Disable_ADC();
                            Log.WriteLine("Volt : " + (Result >> 8).ToString() + "\tTemp : " + (Result & 0xff).ToString());
                            break;
                        case TEST_ITEMS_MANUAL.Read_Volt_Temp:
                            Calculation_VBAT_Voltage_And_Temperature();
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Temp:
                            Set_TestInOut_For_VTEMP(true);
                            Read_ADC_Result(false, 1, 1, 1);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Volt:
                            Set_TestInOut_For_VS(true);
                            Read_ADC_Result(true, 1, 1, 1);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_BGR:
                            Set_TestInOut_For_BGR(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_EDOUT:
                            Set_TestInOut_For_EDOUT(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_RCOSC:
                            Set_TestInOut_For_RCOSC(true);
                            break;
                        case TEST_ITEMS_MANUAL.TESTINOUT_Disable:
                            Set_TestInOut_For_VTEMP(false);
                            Set_TestInOut_For_VS(false);
                            Set_TestInOut_For_BGR(false);
                            Set_TestInOut_For_EDOUT(false);
                            Set_TestInOut_For_RCOSC(false);
                            Disable_ADC();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_ON:
                            Enable_NVM_BIST();
                            Power_On_NVM();
                            Disable_NVM_BIST();
                            break;
                        case TEST_ITEMS_MANUAL.NVM_POWER_OFF:
                            Enable_NVM_BIST();
                            Power_Off_NVM();
                            Disable_NVM_BIST();
                            break;
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
                case COMBOBOX_ITEMS.DTM:
                    switch ((TEST_ITEMS_DTM)TestItemIndex)
                    {
                        case TEST_ITEMS_DTM.SET_CH:
                            if ((Arg == "") || (iVal < 0) || (iVal > 39))
                            {
                                Log.WriteLine("Channel을 10진수로 적어주세요. (Range 0~39)");
                            }
                            else
                            {
                                DTM_Channel = iVal;
                                Log.WriteLine("Channel : " + DTM_Channel.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.SET_Length:
                            if ((Arg == "") || ((iVal != 37) && (iVal != 255)))
                            {
                                Log.WriteLine("Length를 10진수로 적어주세요. (Range 37 또는 255)");
                            }
                            else
                            {
                                DTM_PayLoadLength = iVal;
                                Log.WriteLine("Length : " + DTM_PayLoadLength.ToString());
                            }
                            break;
                        case TEST_ITEMS_DTM.DTM_START_PRNBS9:
                        case TEST_ITEMS_DTM.DTM_START_11110000:
                        case TEST_ITEMS_DTM.DTM_START_10101010:
                        case TEST_ITEMS_DTM.DTM_START_PRNBS15:
                        case TEST_ITEMS_DTM.DTM_START_ALL_1:
                        case TEST_ITEMS_DTM.DTM_START_ALL_0:
                        case TEST_ITEMS_DTM.DTM_START_00001111:
                        case TEST_ITEMS_DTM.DTM_START_0101:
                            Log.WriteLine("Start DTM\nChannel : " + DTM_Channel.ToString() + "\tLength : " + DTM_PayLoadLength.ToString());
                            Run_BLE_DTM_MODE((uint)(TestItemIndex - TEST_ITEMS_DTM.DTM_START_PRNBS9));
                            break;
                        case TEST_ITEMS_DTM.DTM_STOP:
                            Log.WriteLine("Stop DTM");
                            Stop_BLE_DTM_MODE();
                            break;
                    }
                    break;
                default:
                    break;
            }
        }
        #region Function for Chip Test
        private void WakeUp_I2C()
        {
            byte[] SendData = new byte[2];

            SendData[0] = 0xAA;
            SendData[1] = 0xBB;

            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_CMD_I2C(byte data)
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=1, WURF_WRITE=1
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = data;
            SendData[5] = 0x00;
            SendData[6] = 0x06;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // WURF_WRITE=0
            SendData[6] = 0x02;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_WURF_Sel()
        {
            byte[] SendData = new byte[8];

            // WURF_SEL=0
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Write_WURF_AON(string Arg)
        {
            RegisterItem w_WURF_END = Parent.RegMgr.GetRegisterItem("w_WURF_END");         // 0x38
            byte data;

            switch (Arg)
            {
                case "30":
                    data = 0x30;
                    break;
                case "60":
                    data = 0x60;
                    break;
                case "61":
                    data = 0x61;
                    break;
                case "70":
                    data = 0x70;
                    break;
                case "80":
                    data = 0x80;
                    break;
                case "b0":
                case "B0":
                    data = 0xb0;
                    break;
                case "c0":
                case "C0":
                    data = 0xb0;
                    break;
                default:
                    Log.WriteLine("지원하지 않는 CMD 입니다.");
                    return;
                    break;
            }

            w_WURF_END.Read();
            w_WURF_END.Value = 0;
            w_WURF_END.Write();

            Write_WURF_CMD_I2C(data);

            w_WURF_END.Value = 1;
            w_WURF_END.Write();

            Disable_WURF_Sel();

            w_WURF_END.Value = 0;
            w_WURF_END.Write();
        }

        private void Write_Register_Fractional_Calc_Ch(uint ch)
        {
            uint vco;

            RegisterItem SPI_DSM_F_H = Parent.RegMgr.GetRegisterItem("O_SPI_DSM_F[22:16]"); // 0x3E
            RegisterItem SPI_PS_P = Parent.RegMgr.GetRegisterItem("O_SPI_PS_P[4:0]");       // 0x3F
            RegisterItem SPI_PS_S = Parent.RegMgr.GetRegisterItem("O_SPI_PS_S[2:0]");       // 0x3F                

            vco = 2402 + 2 * ch;
            SPI_PS_P.Value = vco >> 7;
            SPI_PS_S.Value = (vco >> 5) - (SPI_PS_P.Value << 2);
            SPI_DSM_F_H.Value = (vco << 1) - (SPI_PS_P.Value << 8) - (SPI_PS_S.Value << 6);

            SPI_PS_P.Write();
            SPI_DSM_F_H.Write();
        }

        private void Write_Register_Tx_Tone_Send(bool on)
        {
            RegisterItem EXT_CH_MODE = Parent.RegMgr.GetRegisterItem("O_EXT_CH_MODE");             // 0x3E
            RegisterItem TX_SEL = Parent.RegMgr.GetRegisterItem("O_TX_SEL");                       // 0x56
            RegisterItem TX_BUF_PEN = Parent.RegMgr.GetRegisterItem("O_TX_BUF_PEN");               // 0x57
            RegisterItem PRE_DA_PEN = Parent.RegMgr.GetRegisterItem("O_PRE_DA_PEN");               // 0x57
            RegisterItem DA_PEN = Parent.RegMgr.GetRegisterItem("O_DA_PEN");                       // 0x57
            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("O_PLL_PEN");                     // 0x59

            RegisterItem EXT_DA_GAIN_SEL = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_SEL");     // 0x5B
            RegisterItem w_TX_BUF_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_TX_BUF_PEN_MODE");   // 0x5C
            RegisterItem w_PRE_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PRE_DA_PEN_MODE");   // 0x5C
            RegisterItem w_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_DA_PEN_MODE");           // 0x5C
            RegisterItem w_TX_SEL_MODE = Parent.RegMgr.GetRegisterItem("w_TX_SEL_MODE");           // 0x5C
            RegisterItem w_PLL_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_PLL_PEN_MODE");         // 0x5D
            RegisterItem XTAL_PLL_CLK_EN_MD = Parent.RegMgr.GetRegisterItem("w_XTAL_PLL_CLK_EN_MD");   // 0x5F

            if (on == true)
            {
                // Tx Active
                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 1;
                EXT_CH_MODE.Write();

                TX_SEL.Read();
                TX_SEL.Value = 1;
                TX_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 1;
                PRE_DA_PEN.Value = 1;
                DA_PEN.Value = 1;
                TX_BUF_PEN.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 0;
                EXT_DA_GAIN_SEL.Write();

                // Dynamic -> Static
                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 1;
                w_PLL_PEN_MODE.Write();

                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 1;
                w_PRE_DA_PEN_MODE.Value = 1;
                w_DA_PEN_MODE.Value = 1;
                w_TX_SEL_MODE.Value = 1;
                w_TX_BUF_PEN_MODE.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 1;
                XTAL_PLL_CLK_EN_MD.Write();

                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                PLL_PEN.Value = 1;
                PLL_PEN.Write();
            }
            else
            {
                PLL_PEN.Read();
                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                XTAL_PLL_CLK_EN_MD.Read();
                XTAL_PLL_CLK_EN_MD.Value = 0;
                XTAL_PLL_CLK_EN_MD.Write();

                // Static -> Dynamic
                w_TX_BUF_PEN_MODE.Read();
                w_TX_BUF_PEN_MODE.Value = 0;
                w_PRE_DA_PEN_MODE.Value = 0;
                w_DA_PEN_MODE.Value = 0;
                w_TX_SEL_MODE.Value = 0;
                w_TX_BUF_PEN_MODE.Write();

                w_PLL_PEN_MODE.Read();
                w_PLL_PEN_MODE.Value = 0;
                w_PLL_PEN_MODE.Write();

                EXT_DA_GAIN_SEL.Read();
                EXT_DA_GAIN_SEL.Value = 1;
                EXT_DA_GAIN_SEL.Write();

                TX_BUF_PEN.Read();
                TX_BUF_PEN.Value = 0;
                PRE_DA_PEN.Value = 0;
                DA_PEN.Value = 0;
                TX_BUF_PEN.Write();

                TX_SEL.Read();
                TX_SEL.Value = 0;
                TX_SEL.Write();

                EXT_CH_MODE.Read();
                EXT_CH_MODE.Value = 0;
                EXT_CH_MODE.Write();
            }
        }

        private void Write_Register_Send_Advertising(int iVal)
        {
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");     // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");     // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");     // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");     // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");     // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");     // 0x1B
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");     // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");     // 0x1F

            B22_WP_0.Value = (uint)(iVal & 0xff);
            B22_WP_0.Write();

            B23_WP_1.Value = (uint)((iVal >> 8) & 0xff);
            B23_WP_1.Write();

            B24_AI_0.Value = 210;
            B24_AI_0.Write();

            B25_AI_1.Value = 29;
            B25_AI_1.Write();

            B26_AI_2.Value = 0;
            B26_AI_2.Write();

            B27_AI_3.Value = 0;
            B27_AI_3.Write();

            B30_AD_0.Value = 3;
            B30_AD_0.Write();

            B31_AD_1.Value = 0;
            B31_AD_1.Write();
        }

        private uint Run_Read_FSM_Status()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[1] & 0xf) << 8) | RcvData[0]);
        }

        private uint CRC_FLAG_READ()
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[4];

            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x01;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            return (uint)(((RcvData[2] & 0x3) << 16) | (RcvData[1] << 8) | RcvData[0]);
        }

        private void Read_Register_FSM()
        {
            uint u_fsm;
            uint crc_flag;
            u_fsm = Run_Read_FSM_Status();
            crc_flag = CRC_FLAG_READ();
            Log.WriteLine("FSM_Status : " + (u_fsm & 0x000f).ToString());
            Log.WriteLine("FSM_State : " + ((u_fsm & 0x0ff0) >> 4).ToString());
            Log.WriteLine("WUR_CRC_ERROR : " + ((crc_flag >> 17) & 0xff).ToString());
            Log.WriteLine("WUR_CRC_DEC : " + (crc_flag & 0xffff).ToString());
        }

        private uint Read_ADC_Result(bool temp_first, uint eoc_skip_temp, uint eoc_skip_volt, uint avg_cnt)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            // Set w_TEST_SELECT
            SendData[0] = 0x04;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
#if true // w_TEST_SELECT = 1
            if (temp_first)
            {
                SendData[7] = 0x80; // temp -> volt
            }
            else
            {
                SendData[7] = 0xC0; // volt -> temp
            }
#else // w_TEST_SELECT = 0
            if (temp_first)
            {
                SendData[7] = 0x00; // temp -> volt
            }
            else
            {
                SendData[7] = 0x40; // volt -> temp
            }
#endif
            I2C.WriteBytes(SendData, SendData.Length, true);
            // Enable ADC (TEST_START = 0, I_PEN = 1)
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x01);
            SendData[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            SendData[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            SendData[7] = (byte)((avg_cnt & 0x1f) << 1);
            I2C.WriteBytes(SendData, SendData.Length, true);
            // Enable ADC (TEST_START = 1, I_PEN = 1)
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = (byte)(((eoc_skip_temp & 0x01) << 7) | 0x01);
            SendData[5] = (byte)(((eoc_skip_volt & 0x07) << 5) | ((eoc_skip_temp >> 1) & 0x1f));
            SendData[6] = (byte)((eoc_skip_volt >> 3) & 0x07);
            SendData[7] = (byte)(((avg_cnt & 0x1f) << 1) | 0x80);
            I2C.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(10);
            // Read ADC_D_B[7:0], ADC_D_T[7:0]
            SendData[0] = 0x08;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);
            return (uint)(((RcvData[1] << 8) | RcvData[0]) & 0xffff);
        }

        private void Disable_ADC()
        {
            byte[] SendData = new byte[8];

            // Disable ADC
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x03;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private double Calculation_Temperature(uint otp, uint adc)
        {
            double[] lowtempcomp_ADC = { 57.11, 57.23, 57.34, 57.45, 57.56, 57.66, 57.75, 57.84 };
            double temperature;
            uint OTP23p5, OTP85TO23p5;

            OTP23p5 = (otp & 0x1f) + 104;
            OTP85TO23p5 = (otp >> 5) + 35;

            if (adc >= OTP23p5)
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) + 23.5;
            }
            else
            {
                temperature = (adc - OTP23p5) * ((85 - 23.5) / OTP85TO23p5) * ((23.5 - (-40)) / lowtempcomp_ADC[otp >> 5]) + 23.5;
            }

            return temperature;
        }

        private double Calculation_VBAT_Voltage(double temperature, uint adc)
        {
            double voltage;

            if (temperature >= 30)
            {
                voltage = (adc - (11.475 - (-3 / (85 - 30)) * (temperature - 30))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }
            else
            {
                voltage = (adc - (11.475 - (2 / (30 - (-40))) * (30 - temperature))) * ((3.3 - 1.7) / (194.225 - 11.475)) + 1.7;
            }

            return voltage;
        }

        private void Calculation_VBAT_Voltage_And_Temperature()
        {
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");       // 0x14
            uint adc;
            double temp, volt;

            B20_DC_0.Read();
            adc = Read_ADC_Result(false, 1, 63, 31);
            Disable_ADC();
            temp = Calculation_Temperature(B20_DC_0.Value, adc & 0xff);
            volt = Calculation_VBAT_Voltage(temp, adc >> 8);
            Log.WriteLine("Volt : " + volt.ToString("F2") + "(" + (adc >> 8).ToString() + ")\tTemp : " + temp.ToString("F3") + "(" + (adc & 0xff).ToString() + ")");
        }

        private void Set_TestInOut_For_VTEMP(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");       // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL");     // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");            // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");            // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BUF_MUX_SEL.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_VS(bool on)
        {
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");        // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");        // 0x4C
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");   // 0x5B
            RegisterItem TEST_BUF_MUX_SEL = Parent.RegMgr.GetRegisterItem("TEST_BUF_MUX_SEL"); // 0x5B

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BUF_MUX_SEL.Value = 0;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0x20;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();
            }
        }

        private void Set_TestInOut_For_BGR(bool on)
        {
            RegisterItem TEST_BGR_BUF_EN = Parent.RegMgr.GetRegisterItem("TEST_BGR_BUF_EN");    // 0x5B
            RegisterItem TEST_CON_L = Parent.RegMgr.GetRegisterItem("O_TEST_CON[1:0]");         // 0x4B
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");         // 0x4C

            if (on == true)
            {
                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 1;
                TEST_BGR_BUF_EN.Write();

                TEST_CON_H.Read();
                TEST_CON_H.Value = 0;
                TEST_CON_H.Write();

                TEST_CON_L.Read();
                TEST_CON_L.Value = 2;
                TEST_CON_L.Write();
            }
            else
            {
                TEST_CON_L.Read();
                TEST_CON_L.Value = 0;
                TEST_CON_L.Write();

                TEST_BGR_BUF_EN.Read();
                TEST_BGR_BUF_EN.Value = 0;
                TEST_BGR_BUF_EN.Write();
            }
        }

        private void Set_TestInOut_For_EDOUT(bool on)
        {
            RegisterItem ITEST_CONT = Parent.RegMgr.GetRegisterItem("ITEST_CONT[8]");   // 0x5A
            RegisterItem O_RX_DATAT = Parent.RegMgr.GetRegisterItem("O_RX_DATAT");      // 0x5A

            if (on == true)
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 1;
                O_RX_DATAT.Value = 1;
                ITEST_CONT.Write();
            }
            else
            {
                ITEST_CONT.Read();
                ITEST_CONT.Value = 0;
                O_RX_DATAT.Value = 0;
                ITEST_CONT.Write();
            }
        }

        private void Set_TestInOut_For_RCOSC(bool on)
        {
            RegisterItem TEST_EN_32K = Parent.RegMgr.GetRegisterItem("O_TEST_EN_32K");     // 0x4C
            RegisterItem TEST_CON_H = Parent.RegMgr.GetRegisterItem("O_TEST_CON[7:2]");    // 0x4C

            if (on == true)
            {
                TEST_EN_32K.Read();
                TEST_EN_32K.Value = 1;
                TEST_CON_H.Value = 8;
                TEST_EN_32K.Write();
            }
            else
            {
                TEST_EN_32K.Value = 0;
                TEST_CON_H.Value = 0;
                TEST_EN_32K.Write();
            }
        }

        private void Enable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x01;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Disable_NVM_BIST()
        {
            byte[] SendData = new byte[8];

            // bist_sel=0, bist_type=0, bist_cmd=0
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x00;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_On_NVM()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=1, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x07;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=1, bist_cmd=0
            SendData[4] = 0x05;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }

        private void Power_Off_NVM()
        {
            byte[] SendData = new byte[8];

            // bist_sel=1, bist_type=2, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x0B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=2, bist_cmd=0
            SendData[4] = 0x09;
            I2C.WriteBytes(SendData, SendData.Length, true);
        }
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

        private bool Check_Revision_Information(uint value)
        {
            RegisterItem REV_ID = Parent.RegMgr.GetRegisterItem("I_DEVID[7:0]");     // 0x68
            bool result = false;

            REV_ID.Read();

            if (value == REV_ID.Value)
            {
                result = true;
            }

            return result;
        }

        private uint Read_Mac_Address()
        {
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10

            B14_BA_0.Read();
            B15_BA_1.Read();
            B16_BA_2.Read();

            return (B16_BA_2.Value << 16) | (B15_BA_1.Value << 8) | B14_BA_0.Value;
        }

        private bool Check_OTP_VALID(bool start, uint page)
        {
            uint addr_nvm;
            bool result = true;

            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];
            byte[] Data = new byte[4];

            if (start == false)
            {
                return false;
            }

            Enable_NVM_BIST();
            Power_On_NVM();

            if (page < 19)
            {
                addr_nvm = page * 44; // ble, control
            }
            else
            {
                addr_nvm = page * 44 + (page - 19) * 4; // Analog
            }

            // NVM Verify // VALID
            Data[0] = 0x1D;
            Data[1] = 0xCA;
            Data[2] = 0x1D;
            Data[3] = 0xCA;
            // bist_sel=1, bist_type=4, bist_cmd=1, bist_addr=address
            SendData[0] = (byte)(0x00);            // ADDR = 0x00060000
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            SendData[4] = (byte)(0x13);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=4, bist_cmd=0, bist_addr=address
            SendData[4] = (byte)(0x11);
            SendData[5] = (byte)(0x00);
            SendData[6] = (byte)(addr_nvm & 0xff);
            SendData[7] = (byte)((addr_nvm >> 8) & 0xff);
            I2C.WriteBytes(SendData, SendData.Length, true);

            // Read bist_rdata
            SendData[0] = (byte)(0x14);            // ADDR = 0x00060014
            SendData[1] = (byte)(0x00);
            SendData[2] = (byte)(0x06);
            SendData[3] = (byte)(0x00);
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            for (int j = 0; j < 4; j++)
            {
                if (RcvData[j] != Data[j])
                {
                    result = false;
                }
            }

            Power_Off_NVM();
            Disable_NVM_BIST();

            return result;
        }

        private bool Read_FF_NVM(bool start)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return true;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // bist_sel=1, bist_type=6, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1B;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=6, bist_cmd=0
            SendData[4] = 0x19;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] != 0x0b)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void Write_Register_AON_Fix_Value()
        {
            RegisterItem w_ADC_EOC_SKIP_TEMP = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_TEMP[5:0]");  // 0x34
            RegisterItem w_ADC_EOC_SKIP_VOLT = Parent.RegMgr.GetRegisterItem("w_ADC_EOC_SKIP_VOLT[5:0]");  // 0x35
            RegisterItem w_LPCAL_FINE_EN = Parent.RegMgr.GetRegisterItem("w_LPCAL_FINE_EN");               // 0x37
            RegisterItem w_TAMP_DET_EN = Parent.RegMgr.GetRegisterItem("w_TAMP_DET_EN");                   // 0x39
            RegisterItem w_GPIO_WKUP_MD = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_MD");                 // 0x41
            RegisterItem w_GPIO_WKUP_INTV = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_INTV[1:0]");        // 0x41
            RegisterItem O_CT_RES2_FINE = Parent.RegMgr.GetRegisterItem("O_CT_RES2_FINE[3:0]");            // 0x45
            RegisterItem w_GPIO_WKUP_EN = Parent.RegMgr.GetRegisterItem("w_GPIO_WKUP_EN");                 // 0x4E
            RegisterItem O_ADC_CM = Parent.RegMgr.GetRegisterItem("O_ADC_CM[1:0]");                        // 0x55
            RegisterItem w_DA_GAIN_MAX = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_MAX[2:0]");              // 0x57
            RegisterItem I_EXT_DA_GAIN_CON = Parent.RegMgr.GetRegisterItem("I_EXT_DA_GAIN_CON[1:0]");      // 0x5A
            RegisterItem w_XTAL_CUR_CFG_1 = Parent.RegMgr.GetRegisterItem("w_XTAL_CUR_CFG_1[5:0]");        // 0x5E
            RegisterItem w_DA_GAIN_INTV = Parent.RegMgr.GetRegisterItem("w_DA_GAIN_INTV[1:0]");            // 0x61
            RegisterItem w_PLL_CH_19_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_19_GAIN[2:0]");        // 0x61
            RegisterItem VS_GainMode = Parent.RegMgr.GetRegisterItem("VS_GainMode");                       // 0x61
            RegisterItem BGR_TC_CTRL = Parent.RegMgr.GetRegisterItem("BGR_TC_CTRL[5:2]");                  // 0x62
            RegisterItem w_TX_SEL_SELECT = Parent.RegMgr.GetRegisterItem("w_TX_SEL_SELECT[1:0]");          // 0x64
            RegisterItem w_PLL_PM_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_PM_GAIN[4:0]");              // 0x64
            RegisterItem w_PLL_CH_0_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_0_GAIN[4:0]");          // 0x64
            RegisterItem w_DA_PEN_SELECT = Parent.RegMgr.GetRegisterItem("w_DA_PEN_SELECT[1:0]");          // 0x65
            RegisterItem w_PLL_CH_12_GAIN = Parent.RegMgr.GetRegisterItem("w_PLL_CH_12_GAIN[4:0]");        // 0x66
            RegisterItem w_ADC_SWITCH_MODE = Parent.RegMgr.GetRegisterItem("w_ADC_SWITCH_MODE");           // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_H = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[4]");        // 0x66
            RegisterItem w_ADC_SAMPLE_CNT_L = Parent.RegMgr.GetRegisterItem("w_ADC_SAMPLE_CNT[3:0]");      // 0x67

            BGR_TC_CTRL.Read();
            BGR_TC_CTRL.Value = 11;
            BGR_TC_CTRL.Write();

            w_ADC_SAMPLE_CNT_H.Read();
            w_ADC_SAMPLE_CNT_H.Value = 1;
            w_ADC_SAMPLE_CNT_H.Write();

            w_ADC_SAMPLE_CNT_L.Read();
            w_ADC_SAMPLE_CNT_L.Value = 15;
            w_ADC_SAMPLE_CNT_L.Write();

            O_CT_RES2_FINE.Read();
            O_CT_RES2_FINE.Value = 9;
            O_CT_RES2_FINE.Write();

            w_DA_GAIN_MAX.Read();
            w_DA_GAIN_MAX.Value = 6;
            w_DA_GAIN_MAX.Write();

            I_EXT_DA_GAIN_CON.Read();
            I_EXT_DA_GAIN_CON.Value = 2;
            I_EXT_DA_GAIN_CON.Write();

            VS_GainMode.Read();
            VS_GainMode.Value = 0;
            VS_GainMode.Write();

            w_TX_SEL_SELECT.Read();
            w_TX_SEL_SELECT.Value = 2;
            w_TX_SEL_SELECT.Write();

            w_DA_PEN_SELECT.Read();
            w_DA_PEN_SELECT.Value = 2;
            w_DA_PEN_SELECT.Write();
        }

        private bool Run_Cal_BGR(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 300;
            double d_lsl = 295, d_usl = 305;
            uint ldo_val, ldo_val_1;

            RegisterItem ULP_BGR_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_BGR_CONT[3:0]");    // 0x53

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_BGR(true);

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(2, (1 + cnt), (double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));

                ULP_BGR_CONT.Read();
                Parent.xlMgr.Sheet.Select("BGR");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            ldo_val = 15;
            ldo_val_1 = 0;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            ULP_BGR_CONT.Value = ldo_val;
            ULP_BGR_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)ULP_BGR_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                ULP_BGR_CONT.Value = ldo_val;
                ULP_BGR_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }

            Set_TestInOut_For_BGR(false);

            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_ALLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 810;
            double d_lsl = 800, d_usl = 820;
            uint ldo_val, ldo_val_1;

            RegisterItem O_ULP_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_CONT[3:0]");        // 0x54
            RegisterItem O_ULP_LDO_LV_CONT = Parent.RegMgr.GetRegisterItem("O_ULP_LDO_LV_CONT[2:0]");  // 0x61

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            O_ULP_LDO_CONT.Read();
            O_ULP_LDO_LV_CONT.Read();

            O_ULP_LDO_CONT.Value = 8;
            O_ULP_LDO_CONT.Write();
            O_ULP_LDO_LV_CONT.Value = 0;
            O_ULP_LDO_LV_CONT.Write();
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("ALLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv > 816)
            {
                O_ULP_LDO_LV_CONT.Value = 3;
                O_ULP_LDO_LV_CONT.Write();
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }
            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            O_ULP_LDO_CONT.Value = ldo_val;
            O_ULP_LDO_CONT.Write();

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)O_ULP_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                O_ULP_LDO_CONT.Value = ldo_val;
                O_ULP_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (O_ULP_LDO_LV_CONT.Value).ToString());
                Parent.xlMgr.Cell.Write(x_pos + 1, (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_MLDO(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_volt_mv;
            double d_diff_mv, d_target_mv = 1000;
            double d_lsl = 990, d_usl = 1010;
            uint ldo_val, ldo_val_1;

            RegisterItem PMU_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_PMU_LDO_CONT[3:0]");          // 0x53
            RegisterItem PMU_MLDO_Coarse_L = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[0]");    // 0x62
            RegisterItem PMU_MLDO_Coarse_H = Parent.RegMgr.GetRegisterItem("O_PMU_MLDO_Coarse[1]");    // 0x63

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            PMU_MLDO_Coarse_H.Read();
            PMU_MLDO_Coarse_L.Read();
            PMU_LDO_CONT.Read();

            PMU_LDO_CONT.Value = 7;
            PMU_LDO_CONT.Write();
            PMU_MLDO_Coarse_L.Value = 0;
            PMU_MLDO_Coarse_L.Write();

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("MLDO");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (d_volt_mv < 1006)
            {
                PMU_MLDO_Coarse_L.Value = 1;
                PMU_MLDO_Coarse_L.Write();
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (d_volt_mv < 1006)
                {
                    PMU_MLDO_Coarse_L.Value = 0;
                    PMU_MLDO_Coarse_L.Write();
                    PMU_MLDO_Coarse_H.Value = 1;
                    PMU_MLDO_Coarse_H.Write();
                }
            }

            ldo_val = 15;
            ldo_val_1 = 0;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();

            for (int val = 2; val >= 0; val--)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
                }
                if (d_volt_mv < d_target_mv)
                {
                    ldo_val += (uint)(1 << val);
                }
                else
                {
                    ldo_val -= (uint)(1 << val);
                }
                ldo_val = ldo_val & 0xf;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            ldo_val_1 = ldo_val;
            d_diff_mv = Math.Abs(d_volt_mv - d_target_mv);

            if (d_volt_mv < d_target_mv)
            {
                if (ldo_val != 7) ldo_val += 1;
            }
            else
            {
                if (ldo_val != 8) ldo_val -= 1;
            }
            ldo_val = ldo_val & 0xf;
            PMU_LDO_CONT.Value = ldo_val;
            PMU_LDO_CONT.Write();
            d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)PMU_LDO_CONT.Value), (1 + cnt), d_volt_mv.ToString("F3"));
            }
            if (Math.Abs(d_volt_mv - d_target_mv) > d_diff_mv)
            {
                ldo_val = ldo_val_1;
                PMU_LDO_CONT.Value = ldo_val;
                PMU_LDO_CONT.Write();
            }

            if (x_pos != 0)
            {
                d_volt_mv = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            }
            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), ((PMU_MLDO_Coarse_H.Value << 1) | PMU_MLDO_Coarse_L.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), ldo_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            }
            if ((d_volt_mv < d_lsl) || (d_volt_mv > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_32K_RCOSC(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_khz, d_diff_khz;
            double d_lsl = 32.113, d_usl = 33.423, d_target_khz = 32.768;
            uint osc_val_l, osc_val_l_1;

            RegisterItem RTC_SCKF_L = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[5:0]");      // 0x4D
            RegisterItem RTC_SCKF_H = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:6]");     // 0x4E

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }

            Set_TestInOut_For_RCOSC(true);
            if (x_pos != 0)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_khz.ToString("F3"));

                Parent.xlMgr.Sheet.Select("RCOSC");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }
            RTC_SCKF_L.Read();
            RTC_SCKF_H.Read();

            // RTC_SCKF[5:0]
            osc_val_l = 31;
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            for (int val = 4; val >= 0; val--)
            {
                d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
                }
                if (d_freq_khz > d_target_khz)
                {
                    osc_val_l += (uint)(1 << val);
                }
                else
                {
                    osc_val_l -= (uint)(1 << val);
                }
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            osc_val_l_1 = osc_val_l;
            d_diff_khz = Math.Abs(d_freq_khz - d_target_khz);

            if (d_freq_khz > d_target_khz)
            {
                if (osc_val_l != 63) osc_val_l += 1;
            }
            else
            {
                if (osc_val_l != 0) osc_val_l -= 1;
            }
            RTC_SCKF_L.Value = osc_val_l;
            RTC_SCKF_L.Write();
            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val_l), (1 + cnt), d_freq_khz.ToString("F3"));
            }
            if (Math.Abs(d_freq_khz - d_target_khz) > d_diff_khz)
            {
                osc_val_l = osc_val_l_1;
                RTC_SCKF_L.Value = osc_val_l;
                RTC_SCKF_L.Write();
            }

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (RTC_SCKF_H.Value).ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), (RTC_SCKF_L.Value).ToString());
            }

            d_freq_khz = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_freq_khz.ToString("F3"));
            }
            Set_TestInOut_For_RCOSC(false);

            if ((d_freq_khz < d_lsl) || (d_freq_khz > d_usl))
                return false;
            else
                return true;
        }

        private bool Run_Cal_Temp_Sensor(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 140; // 30deg
            int lsl = 135, usl = 145;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem TEMP_CONT_L = Parent.RegMgr.GetRegisterItem("O_TEMP_CONT[4:0]");  // 0x55
            RegisterItem TEMP_CONT_H = Parent.RegMgr.GetRegisterItem("TEMP_TRIM[5]");      // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
            TEMP_CONT_H.Read();
            TEMP_CONT_L.Read();

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, 1, 63, 31);
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();
            Set_TestInOut_For_VTEMP(false);
#if false // cal
            Parent.xlMgr.Sheet.Select("Temp_Sen");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
                u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                TEMP_CONT_L.Value = (u_adc_val & 0x1f);
                TEMP_CONT_L.Write();
                TEMP_CONT_H.Value = (u_adc_val >> 5);
                TEMP_CONT_H.Write();
            }
#else // fix
            u_adc_val = 0;
            TEMP_CONT_L.Value = (u_adc_val & 0x1f);
            TEMP_CONT_L.Write();
            TEMP_CONT_H.Value = (u_adc_val >> 5);
            TEMP_CONT_H.Write();
            lsl = 0;
            usl = 255;
#endif
            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = Read_ADC_Result(false, 1, 63, 31) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            Set_TestInOut_For_VTEMP(true);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

            Set_TestInOut_For_VTEMP(false);
#if false
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_Voltage_Scaler(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_adc_code;
            int diff_val, target_val = 45; // 2.0V
            int lsl = 40, usl = 50;
            uint u_adc_val, u_adc_val_1;
            double d_volt_mv;

            RegisterItem VOLSCAL_CON_L = Parent.RegMgr.GetRegisterItem("O_VOLSCAL_CON[3:0]");  // 0x56
            RegisterItem VOLTAGE_CON_H = Parent.RegMgr.GetRegisterItem("VOLTAGE_CON[5:4]");    // 0x5B

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
                return true;
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.0,(@3)");
#else
            PowerSupply0.Write("VOLT 2.0");
#endif
            VOLTAGE_CON_H.Read();
            VOLSCAL_CON_L.Read();

            Set_TestInOut_For_VS(true);

            System.Threading.Thread.Sleep(1);
            u_adc_code = Read_ADC_Result(false, 1, 63, 31);
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
            Disable_ADC();

            Parent.xlMgr.Sheet.Select("VS");
            Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());

            u_adc_val = 32;
            u_adc_val_1 = 32;
            diff_val = 5000;
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());

            for (int val = 4; val >= 0; val--)
            {
                if (u_adc_code < target_val)
                {
                    u_adc_val += (uint)(1 << val);
                }
                else
                {
                    u_adc_val -= (uint)(1 << val);
                }
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
                u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
                Disable_ADC();
                Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
                if (val == 0)
                {
                    u_adc_val_1 = u_adc_val;
                    diff_val = Math.Abs((int)u_adc_code - target_val);
                }
            }
            if (u_adc_code < target_val)
            {
                if (u_adc_val < 63)
                    u_adc_val += 1;
                else
                    u_adc_val = 0;
            }
            else
            {
                if (u_adc_val > 0)
                    u_adc_val -= 1;
                else
                    u_adc_val = 63;
            }
            VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
            VOLSCAL_CON_L.Write();
            VOLTAGE_CON_H.Value = (u_adc_val >> 4);
            VOLTAGE_CON_H.Write();
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((2 + (int)u_adc_val), (1 + cnt), u_adc_code.ToString());
            if (Math.Abs((int)u_adc_code - target_val) > diff_val)
            {
                u_adc_val = u_adc_val_1;
                VOLSCAL_CON_L.Value = (u_adc_val & 0x0f);
                VOLSCAL_CON_L.Write();
                VOLTAGE_CON_H.Value = (u_adc_val >> 4);
                VOLTAGE_CON_H.Write();
            }

            Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
            Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), u_adc_val.ToString());
            u_adc_code = (Read_ADC_Result(false, 1, 63, 31) >> 8) & 0xff;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), u_adc_code.ToString());
            d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
            Disable_ADC();
            Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), d_volt_mv.ToString("F3"));

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif

            Set_TestInOut_For_VS(false);
#if true
            if ((u_adc_code < lsl) || u_adc_code > usl)
                return false;
            else
                return true;
#else
            return true;
#endif
        }

        private bool Run_Cal_32M_XTAL_Load_Cap(bool start, int cnt, int x_pos, int y_pos)
        {
            double d_freq_mhz;
            double d_diff_mhz, d_target_mhz = 2402;
            double d_lsl = 2401.9952, d_usl = 2402.0048;
            uint osc_val, osc_val_1;

            RegisterItem XTAL_LOAD_CONT = Parent.RegMgr.GetRegisterItem("O_XTAL_LOAD_CONT[4:0]");  // 0x4F

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                return true;
            }

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            XTAL_LOAD_CONT.Read();
            if (x_pos != 0)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
                Parent.xlMgr.Sheet.Select("XTAL_Cap");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
            }

            osc_val = 15;
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            for (int val = 2; val >= 0; val--)
            {
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

                if (x_pos != 0)
                {
                    Parent.xlMgr.Cell.Write((2 + (int)val), (1 + cnt), d_freq_mhz.ToString("F4"));
                }
                if (d_freq_mhz > d_target_mhz)
                {
                    osc_val += (uint)(1 << val);
                }
                else
                {
                    osc_val -= (uint)(1 << val);
                }
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            osc_val_1 = osc_val;
            d_diff_mhz = Math.Abs(d_freq_mhz - d_target_mhz);

            if (d_freq_mhz > d_target_mhz)
            {
                if (osc_val != 31) osc_val += 1;
            }
            else
            {
                if (osc_val != 0) osc_val -= 1;
            }
            XTAL_LOAD_CONT.Value = osc_val;
            XTAL_LOAD_CONT.Write();
            System.Threading.Thread.Sleep(1);

            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            if (x_pos != 0)
            {
                Parent.xlMgr.Cell.Write((2 + (int)osc_val), (1 + cnt), d_freq_mhz.ToString("F4"));
            }
            if (Math.Abs(d_freq_mhz - d_target_mhz) > d_diff_mhz)
            {
                osc_val = osc_val_1;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
            }

            System.Threading.Thread.Sleep(1);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;

            if (x_pos != 0)
            {
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            }

            if ((d_freq_mhz < d_lsl) || (d_freq_mhz > d_usl))
            {
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return false;
            }
            else
            {
#if true // offset
                osc_val += 3;
                if (osc_val > 31) osc_val = 31;
                XTAL_LOAD_CONT.Value = osc_val;
                XTAL_LOAD_CONT.Write();
                System.Threading.Thread.Sleep(1);
                SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), osc_val.ToString());
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));
#endif
                Write_Register_Fractional_Calc_Ch(0);
                Write_Register_Tx_Tone_Send(false);
                return true;
            }

        }

        private bool Run_Cal_PLL_2PM(bool start, int cnt, int x_pos, int y_pos)
        {
            uint u_fsm;
            uint crc_flag;

            RegisterItem PM_IN = Parent.RegMgr.GetRegisterItem("O_PM_IN[7:0]");                    // 0x58

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            SendCommand("ook 70");
            System.Threading.Thread.Sleep(100);

            u_fsm = Run_Read_FSM_Status() & 0x000f;
            crc_flag = CRC_FLAG_READ();

            if (u_fsm != 7)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "FSM_Error");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) > 1)) // WUR_PKT_END = 1
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CRC_ERROR");
                return false;
            }
            else if ((u_fsm == 7) && ((crc_flag >> 16) == 1))
            {
                PM_IN.Read();
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), (PM_IN.Value).ToString());
            }
            else
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "CHECK_LOG");
                Console.WriteLine("u_fsm : {0}", u_fsm);
                Console.WriteLine("crc_flag : {0}", crc_flag);
                return false;
            }
            return true;
        }

        private void Set_AON_REG_For_BLE(byte[] aon_ble)
        {
            RegisterItem B04_BA_0 = Parent.RegMgr.GetRegisterItem("B04_BA_0[7:0]");            // 0x04
            RegisterItem B05_BA_1 = Parent.RegMgr.GetRegisterItem("B05_BA_1[7:0]");            // 0x05
            RegisterItem B06_BA_2 = Parent.RegMgr.GetRegisterItem("B06_BA_2[7:0]");            // 0x06
            RegisterItem B07_BA_3 = Parent.RegMgr.GetRegisterItem("B07_BA_3[7:0]");            // 0x07
            RegisterItem B08_BA_4 = Parent.RegMgr.GetRegisterItem("B08_BA_4[7:0]");            // 0x08
            RegisterItem B09_BA_5 = Parent.RegMgr.GetRegisterItem("B09_BA_5[7:0]");            // 0x09
            RegisterItem B10_LEN = Parent.RegMgr.GetRegisterItem("B10_LEN[7:0]");              // 0x0A
            RegisterItem B11_MSDF = Parent.RegMgr.GetRegisterItem("B11_MSDF[7:0]");            // 0x0B
            RegisterItem B12_CID_0 = Parent.RegMgr.GetRegisterItem("B12_CID_0[7:0]");          // 0x0C
            RegisterItem B13_CID_1 = Parent.RegMgr.GetRegisterItem("B13_CID_1[7:0]");          // 0x0D
            RegisterItem B14_BA_0 = Parent.RegMgr.GetRegisterItem("B14_BA_0[7:0]");            // 0x0E
            RegisterItem B15_BA_1 = Parent.RegMgr.GetRegisterItem("B15_BA_1[7:0]");            // 0x0F
            RegisterItem B16_BA_2 = Parent.RegMgr.GetRegisterItem("B16_BA_2[7:0]");            // 0x10
            RegisterItem B17_BA_3 = Parent.RegMgr.GetRegisterItem("B17_BA_3[7:0]");            // 0x11
            RegisterItem B18_BA_4 = Parent.RegMgr.GetRegisterItem("B18_BA_4[7:0]");            // 0x12
            RegisterItem B19_BA_5 = Parent.RegMgr.GetRegisterItem("B19_BA_5[7:0]");            // 0x13
            RegisterItem B20_DC_0 = Parent.RegMgr.GetRegisterItem("B20_DC_0[7:0]");            // 0x14
            RegisterItem B21_DC_1 = Parent.RegMgr.GetRegisterItem("B21_DC_1[7:0]");            // 0x15
            RegisterItem B22_WP_0 = Parent.RegMgr.GetRegisterItem("B22_WP_0[7:0]");            // 0x16
            RegisterItem B23_WP_1 = Parent.RegMgr.GetRegisterItem("B23_WP_1[7:0]");            // 0x17
            RegisterItem B24_AI_0 = Parent.RegMgr.GetRegisterItem("B24_AI_0[7:0]");            // 0x18
            RegisterItem B25_AI_1 = Parent.RegMgr.GetRegisterItem("B25_AI_1[7:0]");            // 0x19
            RegisterItem B26_AI_2 = Parent.RegMgr.GetRegisterItem("B26_AI_2[7:0]");            // 0x1A
            RegisterItem B27_AI_3 = Parent.RegMgr.GetRegisterItem("B27_AI_3[7:0]");            // 0x1B
            RegisterItem B28_RSV_0 = Parent.RegMgr.GetRegisterItem("B28_PUTC_0[7:0]");         // 0x1C
            RegisterItem B29_RSV_1 = Parent.RegMgr.GetRegisterItem("B29_PUTC_1[7:0]");         // 0x1D
            RegisterItem B30_AD_0 = Parent.RegMgr.GetRegisterItem("B30_AD_0[7:0]");            // 0x1E
            RegisterItem B31_AD_1 = Parent.RegMgr.GetRegisterItem("B31_AD_1[7:0]");            // 0x1F
            RegisterItem B32_PUTC_0 = Parent.RegMgr.GetRegisterItem("B32_CC_0[7:0]");          // 0x20
            RegisterItem B33_PUTC_1 = Parent.RegMgr.GetRegisterItem("B33_CC_1[7:0]");          // 0x21
            RegisterItem B34_PUTC_2 = Parent.RegMgr.GetRegisterItem("B34_CC_2[7:0]");          // 0x22
            RegisterItem B35_PUTC_3 = Parent.RegMgr.GetRegisterItem("B35_CC_3[7:0]");          // 0x23
            RegisterItem B36_BL = Parent.RegMgr.GetRegisterItem("B36_BL[7:0]");                // 0x24
            RegisterItem B37_TP = Parent.RegMgr.GetRegisterItem("B37_TP[7:0]");                // 0x25
            RegisterItem B38_SN_0 = Parent.RegMgr.GetRegisterItem("B38_SN_0[7:0]");            // 0x26
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27
            RegisterItem B40_MIC = Parent.RegMgr.GetRegisterItem("B40_MIC[7:0]");              // 0x28
            RegisterItem B41_advDelay = Parent.RegMgr.GetRegisterItem("B41_ADD_0[7:0]");       // 0x29
            RegisterItem B42_advDelay = Parent.RegMgr.GetRegisterItem("B42_ADD_1[7:0]");       // 0x2A
            RegisterItem B43_advDelay = Parent.RegMgr.GetRegisterItem("B43_ADD_2[7:0]");       // 0x2B

            B04_BA_0.Value = aon_ble[0];
            B04_BA_0.Write();

            B05_BA_1.Value = aon_ble[1];
            B05_BA_1.Write();

            B06_BA_2.Value = aon_ble[2];
            B06_BA_2.Write();

            B07_BA_3.Value = aon_ble[3];
            B07_BA_3.Write();

            B08_BA_4.Value = aon_ble[4];
            B08_BA_4.Write();

            B09_BA_5.Value = aon_ble[5];
            B09_BA_5.Write();

            B10_LEN.Value = aon_ble[6];
            B10_LEN.Write();

            B11_MSDF.Value = aon_ble[7];
            B11_MSDF.Write();

            B12_CID_0.Value = aon_ble[8];
            B12_CID_0.Write();

            B13_CID_1.Value = aon_ble[9];
            B13_CID_1.Write();

            B14_BA_0.Value = aon_ble[10];
            B14_BA_0.Write();

            B15_BA_1.Value = aon_ble[11];
            B15_BA_1.Write();

            B16_BA_2.Value = aon_ble[12];
            B16_BA_2.Write();

            B17_BA_3.Value = aon_ble[13];
            B17_BA_3.Write();

            B18_BA_4.Value = aon_ble[14];
            B18_BA_4.Write();

            B19_BA_5.Value = aon_ble[15];
            B19_BA_5.Write();

            B20_DC_0.Value = aon_ble[16];
            B20_DC_0.Write();

            B21_DC_1.Value = aon_ble[17];
            B21_DC_1.Write();

            B22_WP_0.Value = aon_ble[18];
            B22_WP_0.Write();

            B23_WP_1.Value = aon_ble[19];
            B23_WP_1.Write();

            B24_AI_0.Value = aon_ble[20];
            B24_AI_0.Write();

            B25_AI_1.Value = aon_ble[21];
            B25_AI_1.Write();

            B26_AI_2.Value = aon_ble[22];
            B26_AI_2.Write();

            B27_AI_3.Value = aon_ble[23];
            B27_AI_3.Write();

            B28_RSV_0.Value = aon_ble[24];
            B28_RSV_0.Write();

            B29_RSV_1.Value = aon_ble[25];
            B29_RSV_1.Write();

            B30_AD_0.Value = aon_ble[26];
            B30_AD_0.Write();

            B31_AD_1.Value = aon_ble[27];
            B31_AD_1.Write();

            B32_PUTC_0.Value = aon_ble[28];
            B32_PUTC_0.Write();

            B33_PUTC_1.Value = aon_ble[29];
            B33_PUTC_1.Write();

            B34_PUTC_2.Value = aon_ble[30];
            B34_PUTC_2.Write();

            B35_PUTC_3.Value = aon_ble[31];
            B35_PUTC_3.Write();

            B36_BL.Value = aon_ble[32];
            B36_BL.Write();

            B37_TP.Value = aon_ble[33];
            B37_TP.Write();

            B38_SN_0.Value = aon_ble[34];
            B38_SN_0.Write();

            B39_SN_1.Value = aon_ble[35];
            B39_SN_1.Write();

            B40_MIC.Value = aon_ble[36];
            B40_MIC.Write();

            B41_advDelay.Value = aon_ble[37];
            B41_advDelay.Write();

            B42_advDelay.Value = aon_ble[38];
            B42_advDelay.Write();

            B43_advDelay.Value = aon_ble[39];
            B43_advDelay.Write();
        }

        private void Run_Write_OTP_With_OOK(bool start, uint page)
        {
            string cmd;

            if (start == false)
            {
                return;
            }

            cmd = "ook 41." + page.ToString("X2");
            SendCommand(cmd);

            System.Threading.Thread.Sleep(500);
        }

        private uint Run_Verify_OTP_With_BIST(bool start, uint page)
        {
            byte[] SendData = new byte[8];
            byte[] RcvData = new byte[4];

            if (start == false)
            {
                return 65535;
            }
            Enable_NVM_BIST();
            Power_On_NVM();

            // set page_flag
            SendData[0] = 0x0C;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = (byte)(page & 0xff);
            SendData[5] = (byte)((page >> 8) & 0xff);
            SendData[6] = (byte)(((page >> 16) & 0x3f) | 0x80);
            SendData[7] = 0xBB;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=1
            SendData[0] = 0x00;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            SendData[4] = 0x1F;
            SendData[5] = 0x00;
            SendData[6] = 0x00;
            SendData[7] = 0x00;
            I2C.WriteBytes(SendData, SendData.Length, true);

            // bist_sel=1, bist_type=7, bist_cmd=0
            SendData[4] = 0x1D;
            I2C.WriteBytes(SendData, SendData.Length, true);

            System.Threading.Thread.Sleep(100);

            // read read_ff_fail
            SendData[0] = 0x10;
            SendData[1] = 0x00;
            SendData[2] = 0x06;
            SendData[3] = 0x00;
            I2C.WriteBytes(SendData, 4, false);
            RcvData = I2C.ReadBytes(RcvData.Length);

            Power_Off_NVM();
            Disable_NVM_BIST();

            if (RcvData[0] == 0x0b)
            {
                return 0x400;
            }
            else
            {
                return RcvData[0];
            }
        }

        private bool Run_Measure_Initial(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }
            // BGR, ALLDO, MLDO
            Set_TestInOut_For_BGR(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                // BGR
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 295) || (d_val > 305))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_12");
                        result = false;
                    }
                }
                // ALLDO
                d_val = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 3), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 800) || (d_val > 820))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_13");
                        result = false;
                    }
                }

                // MLDO
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i + 6), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 985) || (d_val > 1015))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_14");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_BGR(false);

            // RCOSC
            Set_TestInOut_For_RCOSC(true);
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@2)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        break;
                    default:
                        break;
                }
                d_val = double.Parse(DigitalMultimeter3.WriteAndReadString("MEAS:FREQ?")) / 1000.0;
                Parent.xlMgr.Cell.Write((x_pos + i + 9), (y_pos + cnt), d_val.ToString("F3"));
                if ((d_val < 32.113) || (d_val > 33.423))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_16");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_RCOSC(false);

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_ADC(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            uint u_adc_code, u_temp_code, u_vbat_code;
            uint u_lsl_temp = 130, u_usl_temp = 150, u_lsl_vbat = 255, u_usl_vbat = 0;
            double d_volt_mv;

            if (start == false)
            {
                for (int i = 0; i < 8; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return true;
            }

            Set_TestInOut_For_VTEMP(true);

            System.Threading.Thread.Sleep(1);
            for (int i = 0; i < 4; i++)
            {
                switch (i)
                {
                    case 0:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@3)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                        u_lsl_vbat = 186;
                        u_usl_vbat = 196;
                        break;
                    case 1:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        u_lsl_vbat = 96;
                        u_usl_vbat = 106;
                        break;
                    case 2:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.0,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.0");
#endif
                        u_lsl_vbat = 40;
                        u_usl_vbat = 50;
                        break;
                    case 3:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@3)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                        u_lsl_vbat = 6;
                        u_usl_vbat = 16;
                        break;
                    default:
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                        PowerSupply0.Write("VOLT 2.5");
#endif
                        break;
                }
                System.Threading.Thread.Sleep(100);
                u_adc_code = Read_ADC_Result(false, 1, 63, 31);
                u_temp_code = u_adc_code & 0xff;
                u_vbat_code = (u_adc_code >> 8) & 0xff;
                d_volt_mv = double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 0), (y_pos + cnt), u_temp_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 1), (y_pos + cnt), u_vbat_code.ToString("F3"));
                Parent.xlMgr.Cell.Write((x_pos + i * 3 + 2), (y_pos + cnt), d_volt_mv.ToString("F3"));
                Disable_ADC();
#if false
                if ((u_temp_code < u_lsl_temp) || (u_temp_code > u_usl_temp))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_17");
                        result = false;
                    }
                }
#endif
                if ((u_vbat_code < u_lsl_vbat) || (u_vbat_code > u_usl_vbat))
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_18");
                        result = false;
                    }
                }
            }
            Set_TestInOut_For_VTEMP(false);
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@3)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            return result;
        }

        private bool Run_Measure_VCO_Range(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_freq_mhz;
            double d_sl = 2402;

            RegisterItem VCO_TEST = Parent.RegMgr.GetRegisterItem("O_VCO_TEST");           // 0x41
            RegisterItem VCO_CBANK_L = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[7:0]");  // 0x42
            RegisterItem VCO_CBANK_H = Parent.RegMgr.GetRegisterItem("O_VCO_CBANK[9:8]");  // 0x43

            if (start == false)
            {
                Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "Skip");
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(true);

            VCO_TEST.Read();
            VCO_CBANK_L.Read();
            VCO_CBANK_H.Read();

            VCO_TEST.Value = 1;
            VCO_TEST.Write();

            // VCO Low = 0
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 0;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:SPAN 500 MHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.2 GHZ");
            System.Threading.Thread.Sleep(400);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 0), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz > d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO High = 1023
            VCO_CBANK_L.Value = 255;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 3;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.7 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_freq_mhz.ToString("F4"));
            if (d_freq_mhz < d_sl)
            {
                if (result == true)
                {
                    Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_19");
                    result = false;
                }
            }
            // VCO Mid = 512
            VCO_CBANK_L.Value = 0;
            VCO_CBANK_L.Write();
            VCO_CBANK_H.Value = 2;
            VCO_CBANK_H.Write();
            SpectrumAnalyzer.Write("FREQ:CENT 2.5 GHZ");
            System.Threading.Thread.Sleep(50);
            SpectrumAnalyzer.Write("CALC:MARK1:MAX");
            d_freq_mhz = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:X?")) / 1000000.0;
            Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), d_freq_mhz.ToString("F4"));

            Write_Register_Fractional_Calc_Ch(0);
            Write_Register_Tx_Tone_Send(false);
            VCO_TEST.Value = 0;
            VCO_TEST.Write();

            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");

            return result;
        }

        private bool Run_Measure_Tx_Power_Harmonic(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_power_dbm, d_freq_MHz, d_cur_mA;
            double d_lsl = -2, d_usl = 5;
            uint ch;

            if (start == false)
            {
                for (int i = 0; i < 12; i++)
                {
                    Parent.xlMgr.Cell.Write((x_pos + i), (y_pos + cnt), "Skip");
                }
                return result;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.High;
            PowerSupply0.Write("SENS:CURR:RANG 0.01,(@3)");

            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                    case 1:
                        ch = 12;
                        d_freq_MHz = 2426;
                        break;
                    case 2:
                        ch = 39;
                        d_freq_MHz = 2480;
                        break;
                    default:
                        ch = 0;
                        d_freq_MHz = 2402;
                        break;
                }
                Write_Register_Fractional_Calc_Ch(ch);
                for (int j = 0; j < 2; j++)
                {
                    if (j == 0)
                    {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 3.3,(@2)");
#else
                        PowerSupply0.Write("VOLT 3.3");
#endif
                    }
                    else
                    {
#if (POWER_SUPPLY_NEW)
                        PowerSupply0.Write("VOLT 1.7,(@2)");
#else
                        PowerSupply0.Write("VOLT 1.7");
#endif
                    }
                    Write_Register_Tx_Tone_Send(true);
                    SpectrumAnalyzer.Write("FREQ:CENT " + d_freq_MHz + " MHZ");
                    System.Threading.Thread.Sleep(400);
                    d_cur_mA = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000.0;
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 2), (y_pos + cnt), d_cur_mA.ToString("F2"));
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < d_lsl) || (d_power_dbm > d_usl))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_21");
                            result = false;
                        }
                    }
                    if ((d_cur_mA < 8) || (d_cur_mA > 13))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_28");
                            result = false;
                        }
                    }

                    SpectrumAnalyzer.Write("FREQ:CENT " + (d_freq_MHz * 2) + " MHZ");
                    System.Threading.Thread.Sleep(50);
                    SpectrumAnalyzer.Write("CALC:MARK1:MAX");
                    d_power_dbm = double.Parse(SpectrumAnalyzer.WriteAndReadString("CALC:MARK:Y?"));
                    Parent.xlMgr.Cell.Write((x_pos + i * 6 + j * 3 + 1), (y_pos + cnt), d_power_dbm.ToString("F4"));
                    if ((d_power_dbm < (d_lsl - 68)) || (d_power_dbm > (d_usl - 30)))
                    {
                        if (result == true)
                        {
                            Parent.xlMgr.Cell.Write(3, (y_pos + cnt), "FAIL_22");
                            result = false;
                        }
                    }
                }
                Write_Register_Tx_Tone_Send(false);
            }
#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 2.5,(@2)");
#else
            PowerSupply0.Write("VOLT 2.5");
#endif
            SpectrumAnalyzer.Write("FREQ:SPAN 500 KHZ");
            SpectrumAnalyzer.Write("FREQ:CENT 2.402 GHZ");
            Write_Register_Fractional_Calc_Ch(0);
            PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");

            return result;
        }

        private bool Run_Measure_INTB(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            double d_val_H;
            double d_val_L;

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            d_val_H = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
            I2C.GPIOs[3].Direction = GPIO_Direction.Output; // AD7
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].State = GPIO_State.Low;
            System.Threading.Thread.Sleep(115);
            I2C.GPIOs[3].State = GPIO_State.High;
            System.Threading.Thread.Sleep(100);
            I2C.GPIOs[3].Direction = GPIO_Direction.Input;
            System.Threading.Thread.Sleep(100);
            d_val_L = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));

            if (d_val_H > 0.8 && d_val_L < 0.5)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");
            }
            else
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL");
                return false;
            }

            return result;
        }

        private bool Run_Measure_TAMPER(bool start, int cnt, int x_pos, int y_pos, bool result)
        {
            RegisterItem B39_SN_1 = Parent.RegMgr.GetRegisterItem("B39_SN_1[7:0]");            // 0x27 TAMPER dectec : pass = 0 fail = 1

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

#if !(POWER_SUPPLY_NEW)
            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
            return true;
#endif

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.9,(@2)");    // TAMPER 0.9 set
            PowerSupply0.Write("OUTP ON,(@2)");     // TAMPER 0.9 set            
#else
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 0.9");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_H");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.3,(@2)");   //TAMPER 0.3 set
#else
            PowerSupply0.Write("VOLT 0.3");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x80)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_L");
                result = false;
                return result;
            }

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("VOLT 0.6,(@2)");   //TAMPER 0.6 set
#else
            PowerSupply0.Write("VOLT 0.6");
#endif
            System.Threading.Thread.Sleep(100);

            SendCommand("ook c0");
            System.Threading.Thread.Sleep(100);
            if (B39_SN_1.Read() != 0x00)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "FAIL_M");
                result = false;
                return result;
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "PASS");

#if (POWER_SUPPLY_NEW)
            PowerSupply0.Write("OUTP OFF,(@2)");
#else
            PowerSupply0.Write("VOLT 0.0");
            PowerSupply0.Write("INST:NSEL 2");
#endif
            return result;
        }

        private bool Run_Measure_BLE_Packet(bool start, int cnt, int x_pos, int y_pos, bool result, byte[] aon_ble)
        {
            int len;
            byte b;
            byte[] r_data = new byte[37];

            if (start == false)
            {
                Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "Skip");
                return true;
            }

            // For DUT
            I2C.GPIOs[5].Direction = GPIO_Direction.Output; // AC1
            I2C.GPIOs[5].State = GPIO_State.Low;

            Serial.RcvQueue.Clear();
            SendCommand("ook 30");
            System.Threading.Thread.Sleep(150);

            len = Serial.RcvQueue.Count;
            Console.WriteLine("UART len : " + len);
            SendCommand("mode ook");
            for (int i = 0; i < 11; i++)
            {
                b = Serial.RcvQueue.Get();
            }

            for (int i = 0; i < 36; i++)
            {
                b = Serial.RcvQueue.Get();

                if (b > 0x60)
                {
                    r_data[i] = (byte)((b - 87) << 4); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] = (byte)((b - 55) << 4); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] = (byte)((b - 48) << 4); // '0' = 0x30 = 48
                }

                b = Serial.RcvQueue.Get();
                if (b > 0x60)
                {
                    r_data[i] += (byte)((b - 87)); // 'a' = 0x61 = 97
                }
                else if (b > 0x40)
                {
                    r_data[i] += (byte)((b - 55)); // 'A' = 0x41 = 65
                }
                else
                {
                    r_data[i] += (byte)((b - 48)); // '0' = 0x30 = 48
                }

                //b = UartRcvQueue.GetByte(); // space

                Console.Write("{0:X}-", r_data[i]);
                if (r_data[i] == aon_ble[i])
                {
                    continue;
                }
                else if ((i == 32) || (i == 33)) // BL, TP
                {
                    if (r_data[i] > 150)
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if (i == 35) // TAMPER
                {
                    if (!((r_data[i] == 0x00) || (r_data[i] == 0x80)))
                    {
                        Console.Write("\r\nFail!!{0} read : {1:X} \r\n", i, r_data[i]);
                        result = false;
                        return result;
                    }
                }
                else if ((i < 3) || (i > 9) && (i < 12))
                {
                    continue;
                }
                else
                {
                    Console.Write("\r\nFail!!{0} read : {1:X}, write : {2:X}\r\n", i, r_data[i], aon_ble[i]);
                    Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), i.ToString());
                    result = false;
                    return result;
                }

                if ((b == '\r') || (b == '\n'))
                {
                    break;
                }
            }

            Parent.xlMgr.Cell.Write((x_pos), (y_pos + cnt), "P");
            return result;
        }

        private void Test_Good_Chip_Sorting_Rev3(int start_cnt)
        {
            int cnt = 0, pass = 0, fail = 0;
            double d_val;
            int x_pos = 2, y_pos = 12;
            bool result;
            uint mac_code = 0x000000;
            bool OTP_W_Flag = false;

            byte[] aon_ble = { 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0x1E, 0xFF, 0x6D, 0x0B, 0x00, 0x00, 0x00, 0x7D, 0x46, 0x78, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

            Check_Instrument();

            if (I2C.IsOpen == false) // Check I2C connection
            {
                MessageBox.Show("Check I2C");
                return;
            }

            MessageBox.Show("장비 설정을 확인해주세요.\r\n\r\n1.N6705B Power Supply\r\n  - VBAT\r\n  - TEST_IN_OUT(N6705B)\r\n\r\n" +
                "2.SpectrumAnalyzer\r\n  - TX\r\n\r\n3.Digital Multimeter\r\n  - Current mode : VBAT\r\n  - Voltage mode : ALLDO / MLDO / TEST_IN_OUT\r\n\r\n" +
                "4.UM232H\r\n  - AC0_GPIOH0 : I2C EN\r\n  - AD6_GPIOL2 : RX/TX Switch\r\n  - AD7_GPIOL3 : INTB");

            cnt = start_cnt - 1;
            pass = cnt;

            while (true)
            {
                I2C.GPIOs[3].Direction = GPIO_Direction.Input;   // AD7_GPIOL3(INTB)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].Direction = GPIO_Direction.Output;   // AD6_GPIOL2(TRX S/W)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].Direction = GPIO_Direction.Output;   // AC0_GPIOH0(Level Shifter EN)
                System.Threading.Thread.Sleep(10);
                I2C.GPIOs[4].State = GPIO_State.High;
                System.Threading.Thread.Sleep(10);

                // Power off
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 0.0,(@3)");
                PowerSupply0.Write("OUTP ON,(@3)");
                PowerSupply0.Write("SENS:CURR:RANG 1e-6,(@3)");
#else
                PowerSupply0.Write("INST:NSEL 1");
                PowerSupply0.Write("VOLT 0.0");
                PowerSupply0.Write("OUTP ON");
#endif
                SendCommand("mode ook\n");
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n\r\nTest\t: " + cnt + "\r\nPass\t: " + pass + "\r\nFail\t: " + fail
                                                        , Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    cnt++;
                    fail++;
                    Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
                    Parent.xlMgr.Cell.Write(x_pos, (y_pos + cnt), cnt.ToString());
                    result = true;
                }
                else
                {
                    return;
                }

                DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:AC?"); // ALLDO
                // Power on
#if (POWER_SUPPLY_NEW)
                PowerSupply0.Write("VOLT 3.3,(@3)");
                System.Threading.Thread.Sleep(200);
                PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                PowerSupply0.Write("VOLT 3.3");
                System.Threading.Thread.Sleep(500);
                PowerSupply0.Write("VOLT 2.5");
#endif
                // Check MLDO Voltage

                System.Threading.Thread.Sleep(800);
                d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                Parent.xlMgr.Cell.Write((x_pos + 2), (y_pos + cnt), d_val.ToString("F3"));
                if (d_val > 0.2)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                    continue;
                }

                // Seq1 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#else
                d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                d_val /= 5;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                d_val = 600;
                Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), "Skip");
#endif
                if ((d_val < 500) || (d_val > 850))
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                    continue;
                }

                I2C.GPIOs[4].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(10);
                // Wake-up
                WakeUp_I2C();
                System.Threading.Thread.Sleep(5);

                // READ_DIVICE_ID
                if (Check_Revision_Information(0x5F) == false) // N1 = 0x5D, N1B = 0x5E, N1C = 0x5F
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "F");
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_24");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 4), (y_pos + cnt), "P");
                }

                mac_code = Read_Mac_Address();
                aon_ble[0] = (byte)(mac_code & 0xff);
                aon_ble[1] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[2] = (byte)((mac_code >> 16) & 0xff);
                aon_ble[10] = (byte)(mac_code & 0xff);
                aon_ble[11] = (byte)((mac_code >> 8) & 0xff);
                aon_ble[12] = (byte)((mac_code >> 16) & 0xff);
                Parent.xlMgr.Cell.Write((x_pos + 81), (y_pos + cnt), (aon_ble[5].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 82), (y_pos + cnt), (aon_ble[4].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 83), (y_pos + cnt), (aon_ble[3].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 84), (y_pos + cnt), (aon_ble[2].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 85), (y_pos + cnt), (aon_ble[1].ToString("X")));
                Parent.xlMgr.Cell.Write((x_pos + 86), (y_pos + cnt), (aon_ble[0].ToString("X")));

                if (Check_OTP_VALID(true, 21) == true)
                {
                    goto CAL_SKIP;
                }

                // Seq1 5.OTP read
                if (Read_FF_NVM(false) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_10");
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "F");
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 5), (y_pos + cnt), "Skip");
                }

                Write_Register_AON_Fix_Value();

#if true // For Test
                Parent.xlMgr.Sheet.Select("LDO_Default");
                Parent.xlMgr.Cell.Write(1, (1 + cnt), cnt.ToString());
                Parent.xlMgr.Cell.Write(3, (1 + cnt), (double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Cell.Write(4, (1 + cnt), (double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000).ToString("F2"));
                Parent.xlMgr.Sheet.Select("IRIS_Chip_Test");
#endif
                // Seq2 1.BGR Trim
                if (Run_Cal_BGR(true, cnt, (x_pos + 6), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_4");
                    continue;
                }
                // Seq2 2.ALON LDO Trim
                if (Run_Cal_ALLDO(true, cnt, (x_pos + 8), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_5");
                    continue;
                }
                // Seq2 3.MLDO Trim
                if (Run_Cal_MLDO(true, cnt, (x_pos + 11), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_6");
                    continue;
                }
                // Seq2 4.32K RCOSC Trim
                if (Run_Cal_32K_RCOSC(true, cnt, (x_pos + 14), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_7");
                    continue;
                }
                // Seq2 5.Temp Sensor Trim
                if (Run_Cal_Temp_Sensor(true, cnt, (x_pos + 18), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_8");
                    continue;
                }
                // Seq2 5.Voltage Scaler Trim
                if (Run_Cal_Voltage_Scaler(true, cnt, (x_pos + 22), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_25");
                    continue;
                }
                // Seq2 6.32M X-tal Load Cap Trim
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                if (Run_Cal_32M_XTAL_Load_Cap(true, cnt, (x_pos + 26), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_9");
                    continue;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);
                // Seq2 7.PLL 2PM Trim
                if (Run_Cal_PLL_2PM(true, cnt, (x_pos + 29), y_pos) == false)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_20");
                    continue;
                }
                // Seq2 8.OTP Trim data write
                // Set AON_REG(BLE) For OTP
                Set_AON_REG_For_BLE(aon_ble);

                Run_Write_OTP_With_OOK(OTP_W_Flag, 0);
                // Run_Write_OTP_With_OOK(OTP_W_Flag, 15);
                Run_Write_OTP_With_OOK(OTP_W_Flag, 21);
                d_val = Run_Verify_OTP_With_BIST(OTP_W_Flag, 0x208001);
                if (d_val == 65535)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "Skip");
                }
                else if (d_val != 0x400)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_11");
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), d_val.ToString());
                    continue;
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 30), (y_pos + cnt), "P");
                }

            CAL_SKIP:
                if (OTP_W_Flag) // Reset
                {
                    I2C.GPIOs[4].State = GPIO_State.High;
                    System.Threading.Thread.Sleep(10);
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 0.0,(@3)");
#else
                    PowerSupply0.Write("VOLT 0.0");
#endif
                    System.Threading.Thread.Sleep(500);

#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("VOLT 3.3,(@3)");
                    System.Threading.Thread.Sleep(200);
                    PowerSupply0.Write("VOLT 2.5,(@3)");
#else
                    PowerSupply0.Write("VOLT 3.3");
                    System.Threading.Thread.Sleep(500);
                    PowerSupply0.Write("VOLT 2.5");
#endif
                    // Check MLDO Voltage                    
                    System.Threading.Thread.Sleep(800);
                    d_val = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?"));
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), d_val.ToString("F3"));
                    if (d_val > 0.2)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_2");
                        continue;
                    }

                    // Seq2 4.Sleep Current Test
#if true
#if (POWER_SUPPLY_NEW)
                    d_val = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), (d_val).ToString("F3"));
#else
                    d_val = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    for (int i = 0; i < 4; i++) d_val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:CURR:DC?")) * 1000000000.0;
                    d_val /= 5;
                    Parent.xlMgr.Cell.Write((x_pos + 3), (y_pos + cnt), (d_val).ToString("F3"));
#endif
#else
                    d_val = 600;
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
#endif
                    if ((d_val < 500) || (d_val > 850))
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_3");
                        continue;
                    }

                    I2C.GPIOs[4].State = GPIO_State.Low;
                    System.Threading.Thread.Sleep(10);
                    // Wake-up
                    WakeUp_I2C();
                    System.Threading.Thread.Sleep(5);
                }
                else
                {
                    Parent.xlMgr.Cell.Write((x_pos + 31), (y_pos + cnt), "Skip");
                    Parent.xlMgr.Cell.Write((x_pos + 32), (y_pos + cnt), "Skip");
                }

                // Seq3 1.Initial Test
                if (Run_Measure_Initial(true, cnt, (x_pos + 33), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 2.ADC Test
                if (Run_Measure_ADC(true, cnt, (x_pos + 45), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.High;
                System.Threading.Thread.Sleep(1);
                // Seq3 4.VCO Range Test
                if (Run_Measure_VCO_Range(true, cnt, (x_pos + 57), y_pos, result) == false)
                {
                    result = false;
                }
                // Seq3 4. Tx output power and Harmonic
                if (Run_Measure_Tx_Power_Harmonic(true, cnt, (x_pos + 60), y_pos, result) == false)
                {
                    result = false;
                }
                I2C.GPIOs[2].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(1);

                // INTB Test
                if (Run_Measure_INTB(true, cnt, (x_pos + 79), y_pos, result) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_27");
                    }
                    result = false;
                }
                // TAMPER Test
                if (Run_Measure_TAMPER(true, cnt, (x_pos + 78), y_pos, result) == false)
                {
#if (POWER_SUPPLY_NEW)
                    PowerSupply0.Write("OUTP OFF,(@2)");
#endif
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_26");
                    }
                    result = false;
                }

                // BLE Test
                if (Run_Measure_BLE_Packet(true, cnt, (x_pos + 80), y_pos, result, aon_ble) == false)
                {
                    if (result == true)
                    {
                        Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "FAIL_23");
                    }
                    result = false;
                }

                if (result == true)
                {
                    Parent.xlMgr.Cell.Write((x_pos + 1), (y_pos + cnt), "PASS");
                    fail--;
                    pass++;
                }
            }
        }

        private void DUT_TempSensor_Test()
        {
            byte Addr;
            byte[] RcvData = new byte[2];

            Parent.xlMgr.Sheet.Add(DateTime.Now.ToString("MMddHHmmss_") + "DUT_TS");
            Parent.xlMgr.Cell.Write(1, 1, "No.");
            Parent.xlMgr.Cell.Write(2, 1, "ADC");
            Parent.xlMgr.Cell.Write(3, 1, "Temp");

            I2C.GPIOs[2].Direction = GPIO_Direction.Output;
            System.Threading.Thread.Sleep(10);

            Addr = 0x00;
#if false
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(200); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Parent.xlMgr.Cell.Write(1, 2 + i, i.ToString());
                Parent.xlMgr.Cell.Write(2, 2 + i, ((RcvData[0] << 8) | RcvData[1]).ToString());
                Parent.xlMgr.Cell.Write(3, 2 + i, (((RcvData[0] << 8) | RcvData[1]) * 0.0078125).ToString());

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(500);
            }
#else
            for (int i = 0; i < 100; i++)
            {
                I2C.GPIOs[3].State = GPIO_State.High;
                System.Threading.Thread.Sleep(330); // GPIO Low to High
                System.Threading.Thread.Sleep(300); // GPIO High Delay

                RcvData = I2C.ftMPSSE.I2C_WriteAndReadBytes(0x48, new byte[] { Addr }, 1, 2);
                Console.WriteLine("ADC : {0}\tTemp : {1}", (RcvData[0] << 8) | RcvData[1], ((RcvData[0] << 8) | RcvData[1]) * 0.0078125);

                I2C.GPIOs[3].State = GPIO_State.Low;
                System.Threading.Thread.Sleep(100);
            }
#endif
        }
        #endregion

        #region Function for DTM
        private void SetBleDTM_AonControlReg()
        {
            WriteRegister(0x33, 0x17);
            WriteRegister(0x34, 0x3F);       //ADC_EOC_SKIP_TEMP
            WriteRegister(0x35, 0x3F);       //ADC_EOC_SKIP_VOLT
            WriteRegister(0x36, 0xFF);
            WriteRegister(0x37, 0x47);
            WriteRegister(0x38, 0x00);
            WriteRegister(0x39, 0x40);
            WriteRegister(0x3A, 0x4A);
            WriteRegister(0x3B, 0x00);
        }

        private void SetBleDTM_AonControlReg_OTP()
        {
            WriteRegister(0x33, 0x17);
            WriteRegister(0x37, 0x47);

            WriteRegister(0x45, 0x7);   //CT_RES32_FINE = 7 (Bandwidth change)
            WriteRegister(0x64, 0xD2);  //TX_SEL = CoraseLock & PLL_PM_GIAN = 18;
            WriteRegister(0x65, 0x50);  //DA_PEN = CoraseLock & CH_0_GAIN = 16;
        }

        private void SetBleDTM_AonAnalogReg()
        {
            WriteRegister(0x62, 0x66);      //r3 ONLY
            WriteRegister(0x61, 0x33);
            WriteRegister(0x53, 0xC7);      //FC_LGA
            WriteRegister(0x54, 0xBC);
            WriteRegister(0x3C, 0x10);
            WriteRegister(0x3D, 0x12);
            WriteRegister(0x3E, 0x00);      //R3 ONLY
            WriteRegister(0x3F, 0x72);
            WriteRegister(0x40, 0x82);
            WriteRegister(0x41, 0x80);
            WriteRegister(0x42, 0x00);
            WriteRegister(0x43, 0xF2);
            WriteRegister(0x44, 0xFF);
            WriteRegister(0x45, 0x02);
            WriteRegister(0x46, 0xFF);
            WriteRegister(0x47, 0x42);
            WriteRegister(0x48, 0x62);
            WriteRegister(0x49, 0xFF);
            WriteRegister(0x4A, 0x3F);
            WriteRegister(0x4B, 0x16);
            WriteRegister(0x4C, 0x00);
            WriteRegister(0x4D, 0x29);
            WriteRegister(0x4E, 0xB2);
            WriteRegister(0x4F, 0x61);
            WriteRegister(0x50, 0xC5);
            WriteRegister(0x51, 0x2F);      //R3 ONLY
            WriteRegister(0x52, 0x22);

            WriteRegister(0x55, 0xCE);
            WriteRegister(0x56, 0x00);  //R3 ONLY DYNAMIC CONTROL FOR BLE LINK
            WriteRegister(0x57, 0x07);  //TX_BUF, DA_PEN, PRE_DA_PEN, DA_GAIN R3 ONLY DYNIMIC CONTROL FOR BLE LINK
            WriteRegister(0x58, 0x33);  //2PM_CAL R3 ONLY  
            WriteRegister(0x59, 0x07);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
            WriteRegister(0x5A, 0x00);  //EXT_DA_GAIN_LSB_BIT
            WriteRegister(0x5B, 0xE2);  //EXT_DA_GAIN_SET = 1; (BLE_LINK) R3                                                       
            WriteRegister(0x5C, 0x00);  //TX_SEL(W),PRE_DA_PEN(W), TX_BUF_PEN(W), DA_PEN(W) R3 ONLY                                         
            WriteRegister(0x5D, 0x60);  //PM_RESETB(W), PLL_PEN(W)
            WriteRegister(0x5E, 0x20);
            WriteRegister(0x5F, 0x45);  //R3 ONLY
            WriteRegister(0x60, 0x08);

            WriteRegister(0x63, 0x00);
            WriteRegister(0x64, 0xB0);  //TX_SEL = DATA_ST

            //FOR R3 ONLY
            WriteRegister(0x65, 0x34);  //DA_PEN = DATA_ST
            WriteRegister(0x66, 0x12);
            WriteRegister(0x67, 0x40);
        }

        private void Set_PLLRESET(int delay)
        {
            WriteRegister(0x59, 0x1);
            System.Threading.Thread.Sleep(delay);
            WriteRegister(0x59, 0x7);
            System.Threading.Thread.Sleep(delay);
        }

        private void Set_BLE_LINK(uint addr, uint data)
        {
            byte[] Addrs = new byte[8];

            Addrs[0] = (byte)(addr & 0xff);
            Addrs[1] = (byte)(addr >> 8 & 0xff);
            Addrs[2] = (byte)(addr >> 16 & 0xff);
            Addrs[3] = 0x00;
            Addrs[4] = (byte)(data & 0xff);
            Addrs[5] = (byte)(data >> 8 & 0xff);
            Addrs[6] = (byte)(data >> 16 & 0xff);
            Addrs[7] = (byte)(data >> 24 & 0xff);

            I2C.WriteBytes(Addrs, 8, true);
        }

        private void Run_BLE_DTM_MODE(uint PayLoadMode)
        {
            uint PmGain = 0;

            PmGain = ReadRegister(0x58);

            if (PmGain == 0x42)
            {
                SetBleDTM_AonControlReg();
                System.Threading.Thread.Sleep(2);
                SetBleDTM_AonAnalogReg();
                System.Threading.Thread.Sleep(2);
            }
            else
            {
                SetBleDTM_AonControlReg_OTP();
                System.Threading.Thread.Sleep(2);
            }

            Set_BLE_LINK(0x058000, 0x80);   //reset
            Set_BLE_LINK(0x058000, 0x48);   //stop dtm
            Set_BLE_LINK(0x0580d8, 0x78);   //bb_clk_freq_minus_1 
            Set_BLE_LINK(0x0581b4, 0x72);   //RADIO_DATA host datain [15:8] / host dataout[15:8]
            Set_BLE_LINK(0x0581ac, 0x9043); //RADIO_ACCESS/RADIO_CNTRL 
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9400);
            Set_BLE_LINK(0x0581b4, 0x00);
            Set_BLE_LINK(0x0581ac, 0x9401);
            Set_BLE_LINK(0x0580d8, 0x78);
            Set_BLE_LINK(0x058190, 0x9280);
            Set_BLE_LINK(0x058198, 0x5b5b);
            Set_BLE_LINK(0x0581b0, 0x03);
            Set_BLE_LINK(0x0581b8, 0x05);

            uint pmode = ((PayLoadMode * 128) + 0);
            Set_BLE_LINK(0x058170, pmode);
            Set_BLE_LINK(0x05819c, (uint)DTM_PayLoadLength); //packet length
            Set_BLE_LINK(0x040014, (0x280 + (uint)DTM_Channel)); //Ch Sel Mode & CH 0;
            Set_PLLRESET(2);
            Set_BLE_LINK(0x058000, 0x46); //start dtm
        }

        private void Stop_BLE_DTM_MODE()
        {
            Set_BLE_LINK(0x058000, 0x48);
            WriteRegister(0x59, 0x03);  //PLL_PEN, PM_RESETB, PLL_2PM_CAL_HOLD
        }
        #endregion
    }
}
