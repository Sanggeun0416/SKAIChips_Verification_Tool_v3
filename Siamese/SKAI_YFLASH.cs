using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SKAIChips_Verification;
using JLcLib.Chip;
using JLcLib.Comn;

namespace SKAI_YFLASH
{
    public class SLM4781A : ChipControl
    {
        #region Variable and declaration

        public enum TEST_ITEMS
        {
            TEST,
            NUM_TEST_ITEMS,
        }

        private JLcLib.Custom.I2C I2C { get; set; }

        public int SlaveAddress { get; private set; } = 0x29;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI TempChamber = null;
        #endregion Variable and declaration

        public SLM4781A(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            CalibrationData = new byte[256];

            /* Init test items combo box */
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            if (Parent.xlMgr.Sheet.Name == "I2C_Analog")
            {
                SendBytes.Add((byte)((Address >> 8) & 0xFF));
                SendBytes.Add((byte)((Address >> 0) & 0xFF));
                SendBytes.Add((byte)((Data >> 8) & 0xFF));
                SendBytes.Add((byte)((Data >> 0) & 0xFF));
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            }
            else
            {
                //I2C Write: Erase Command
                SendBytes.Add(0x00);
                SendBytes.Add(0x10);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(210);
                System.Threading.Thread.Sleep(70);

                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Data);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);

                //I2C Write: Write Command
                SendBytes.Add(0x00);
                SendBytes.Add(0x12);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(110);
            }
            I2C.Config.SlaveAddress = sa;
        }

        private uint ReadRegister(uint Address)
        {
            if (Parent.xlMgr.Sheet.Name == "I2C_Analog")
            {
                SLM4781A_ReadI2C();
            }
            else
            {
                SLM4781A_ReadFlash();
            }
            return 0;
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
            Parent.ChipCtrlButtons[0].Text = "Erase";
            Parent.ChipCtrlButtons[0].Visible = true;
            Parent.ChipCtrlButtons[0].Click += EraseFlash_Click;

            Parent.ChipCtrlButtons[1].Text = "Write";
            Parent.ChipCtrlButtons[1].Visible = true;
            Parent.ChipCtrlButtons[1].Click += WriteFlash_Click;

            Parent.ChipCtrlButtons[3].Text = "Read";
            Parent.ChipCtrlButtons[3].Visible = true;
            Parent.ChipCtrlButtons[3].Click += ReadFlash_Click;

            Parent.ChipCtrlTextboxes[4].Text = "300"; // Delay Time
            Parent.ChipCtrlTextboxes[4].Visible = true;

            Parent.ChipCtrlButtons[5].Text = "E/W_0";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += Erase_Time_Test_Write_All_zero_Click;

            Parent.ChipCtrlButtons[6].Text = "E/W_1";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Erase_Time_Test_Write_All_one_Click;

            Parent.ChipCtrlButtons[7].Text = "E/W_D";
            Parent.ChipCtrlButtons[7].Visible = true;
            Parent.ChipCtrlButtons[7].Click += Erase_Time_Test_Write_Default_code_Click;
        }

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
                    case JLcLib.Instrument.InstrumentTypes.OscilloScope0:
                        if (OscilloScope0 == null)
                            OscilloScope0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (OscilloScope0.IsOpen == false)
                            OscilloScope0.Open();
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

        private void EraseFlash_Click(object sender, EventArgs e)
        {
            SLM4781A_EraseFlash();
        }

        private void WriteFlash_Click(object sender, EventArgs e)
        {
            SLM4781A_WriteFlash();
        }

        private void ReadFlash_Click(object sender, EventArgs e)
        {
            if (Parent.xlMgr.Sheet.Name == "I2C_Analog")
            {
                SLM4781A_ReadI2C();
            }
            else
            {
                SLM4781A_ReadFlash();
            }
        }

        private void Erase_Time_Test_Write_All_zero_Click(object sender, EventArgs e)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();

            System.Threading.Thread.Sleep(int.Parse(Parent.ChipCtrlTextboxes[4].Text) - 2);

            for (int Address = 0x30; Address <= 0x37; Address++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                if (Address == 0x37)
                {
                    SendBytes.Add(0x80);
                }
                else
                {
                    SendBytes.Add(0x00);
                }
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }
            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(110);

            I2C.Config.SlaveAddress = sa;
        }

        private void Erase_Time_Test_Write_All_one_Click(object sender, EventArgs e)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();

            System.Threading.Thread.Sleep(int.Parse(Parent.ChipCtrlTextboxes[4].Text) - 2);

            for (int Address = 0x30; Address <= 0x37; Address++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                SendBytes.Add(0xFF);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }
            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(110);

            I2C.Config.SlaveAddress = sa;
        }

        private void Erase_Time_Test_Write_Default_code_Click(object sender, EventArgs e)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();

            System.Threading.Thread.Sleep(int.Parse(Parent.ChipCtrlTextboxes[4].Text) - 2);

            for (int Address = 0x30; Address <= 0x37; Address++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                switch (Address)
                {
                    case 0x30:
                        SendBytes.Add(0x9F);
                        break;
                    case 0x31:
                        SendBytes.Add(0xD5);
                        break;
                    case 0x32:
                        SendBytes.Add(0x57);
                        break;
                    case 0x33:
                        SendBytes.Add(0x55);
                        break;
                    case 0x34:
                        SendBytes.Add(0x17);
                        break;
                    case 0x35:
                        SendBytes.Add(0xC2);
                        break;
                    case 0x36:
                        SendBytes.Add(0xD4);
                        break;
                    case 0x37:
                        SendBytes.Add(0xD9);
                        break;
                }
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }
            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(110);

            I2C.Config.SlaveAddress = sa;
        }

        #region Flase control methods
        private void SLM4781A_WriteFlash()
        {
            RegisterItem FS5V = Parent.RegMgr.GetRegisterItem("FS5V[5:0]");                 // 0x00
            RegisterItem PWM_FREQ = Parent.RegMgr.GetRegisterItem("PWM_FREQ[1:0]");         // 0x00
            RegisterItem ADIM1H = Parent.RegMgr.GetRegisterItem("ADIM1H[4:0]");             // 0x01
            RegisterItem DELTA_SEL = Parent.RegMgr.GetRegisterItem("DELTA_SEL");            // 0x01
            RegisterItem OEC = Parent.RegMgr.GetRegisterItem("OEC[1]");                     // 0x01
            RegisterItem ADIM_RNG = Parent.RegMgr.GetRegisterItem("ADIM_RNG(WO)");          // 0x01
            RegisterItem ADIM1L = Parent.RegMgr.GetRegisterItem("ADIM1L[4:0]");             // 0x02
            RegisterItem VCC_DLY_CNT = Parent.RegMgr.GetRegisterItem("VCC_DLY_CNT[2:0]");   // 0x02
            RegisterItem ADIM2H = Parent.RegMgr.GetRegisterItem("ADIM2H[4:0]");             // 0x03
            RegisterItem PERCENT_SEL = Parent.RegMgr.GetRegisterItem("PERCENT_SEL[1:0]");   // 0x03
            RegisterItem GATE_CNT = Parent.RegMgr.GetRegisterItem("GATE_CNT");              // 0x03
            RegisterItem ADIM2L = Parent.RegMgr.GetRegisterItem("ADIM2L[4:0]");             // 0x04
            RegisterItem G_FREQ = Parent.RegMgr.GetRegisterItem("G_FREQ");                  // 0x04
            RegisterItem PWM_OVER = Parent.RegMgr.GetRegisterItem("PWM_OVER");              // 0x04
            RegisterItem JITTER_CON = Parent.RegMgr.GetRegisterItem("JITTER_CON");          // 0x04
            RegisterItem FSF = Parent.RegMgr.GetRegisterItem("FSF[3:0]");                   // 0x05
            RegisterItem D_REF = Parent.RegMgr.GetRegisterItem("2D_REF[3:0]");              // 0x05
            RegisterItem OD2_REF = Parent.RegMgr.GetRegisterItem("OD2_REF[5:0](WO)");       // 0x06
            RegisterItem ODP_EN = Parent.RegMgr.GetRegisterItem("ODP_EN");                  // 0x06
            RegisterItem FUNC_EN = Parent.RegMgr.GetRegisterItem("FUNC_EN");                // 0x06
            RegisterItem ODP_3D = Parent.RegMgr.GetRegisterItem("ODP_3D[5:0]");             // 0x07
            RegisterItem D_FREQ = Parent.RegMgr.GetRegisterItem("D_FREQ");                  // 0x07
            RegisterItem FD = Parent.RegMgr.GetRegisterItem("FD(WO)");                      // 0x07

            uint data = 0x00;
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            for (int i = 0x30; i <= 0x37; i++)
            {
                switch (i)
                {
                    case 0x30:
                        data = (PWM_FREQ.Value << 6) | FS5V.Value;
                        break;
                    case 0x31:
                        data = (ADIM_RNG.Value << 7) | (OEC.Value << 6) | (DELTA_SEL.Value << 5) | ADIM1H.Value;
                        break;
                    case 0x32:
                        data = (VCC_DLY_CNT.Value << 5) | ADIM1L.Value;
                        break;
                    case 0x33:
                        data = (GATE_CNT.Value << 7) | (PERCENT_SEL.Value << 5) | ADIM2H.Value;
                        break;
                    case 0x34:
                        data = (JITTER_CON.Value << 7) | (PWM_OVER.Value << 6) | (G_FREQ.Value << 5) | ADIM2L.Value;
                        break;
                    case 0x35:
                        data = (D_REF.Value << 4) | FSF.Value;
                        break;
                    case 0x36:
                        data = (FUNC_EN.Value << 7) | (ODP_EN.Value << 6) | OD2_REF.Value;
                        break;
                    case 0x37:
                        data = (FD.Value << 7) | (D_FREQ.Value << 6) | ODP_3D.Value;
                        break;
                    default:
                        data = 0x00;
                        break;
                }

                SendBytes.Add(0x00);
                SendBytes.Add((byte)i);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)data);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }

            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            System.Threading.Thread.Sleep(110);
            I2C.Config.SlaveAddress = sa;
        }

        public void SLM4781A_EraseFlash()
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            System.Threading.Thread.Sleep(210);

            I2C.Config.SlaveAddress = sa;
        }

        private void SLM4781A_ReadFlash()
        {
            RegisterItem FS5V = Parent.RegMgr.GetRegisterItem("FS5V[5:0]");                 // 0x00
            RegisterItem PWM_FREQ = Parent.RegMgr.GetRegisterItem("PWM_FREQ[1:0]");         // 0x00
            RegisterItem ADIM1H = Parent.RegMgr.GetRegisterItem("ADIM1H[4:0]");             // 0x01
            RegisterItem DELTA_SEL = Parent.RegMgr.GetRegisterItem("DELTA_SEL");            // 0x01
            RegisterItem OEC = Parent.RegMgr.GetRegisterItem("OEC[1]");                     // 0x01
            //RegisterItem ADIM_RNG = Parent.RegMgr.GetRegisterItem("ADIM_RNG(WO)");        // 0x01
            RegisterItem ADIM1L = Parent.RegMgr.GetRegisterItem("ADIM1L[4:0]");             // 0x02
            RegisterItem VCC_DLY_CNT = Parent.RegMgr.GetRegisterItem("VCC_DLY_CNT[2:0]");   // 0x02
            RegisterItem ADIM2H = Parent.RegMgr.GetRegisterItem("ADIM2H[4:0]");             // 0x03
            RegisterItem PERCENT_SEL = Parent.RegMgr.GetRegisterItem("PERCENT_SEL[1:0]");   // 0x03
            RegisterItem GATE_CNT = Parent.RegMgr.GetRegisterItem("GATE_CNT");              // 0x03
            RegisterItem ADIM2L = Parent.RegMgr.GetRegisterItem("ADIM2L[4:0]");             // 0x04
            RegisterItem G_FREQ = Parent.RegMgr.GetRegisterItem("G_FREQ");                  // 0x04
            RegisterItem PWM_OVER = Parent.RegMgr.GetRegisterItem("PWM_OVER");              // 0x04
            RegisterItem JITTER_CON = Parent.RegMgr.GetRegisterItem("JITTER_CON");          // 0x04
            RegisterItem FSF = Parent.RegMgr.GetRegisterItem("FSF[3:0]");                   // 0x05
            RegisterItem D_REF = Parent.RegMgr.GetRegisterItem("2D_REF[3:0]");              // 0x05
            RegisterItem OD2_REF = Parent.RegMgr.GetRegisterItem("OD2_REF[5:0](WO)");       // 0x06
            RegisterItem ODP_EN = Parent.RegMgr.GetRegisterItem("ODP_EN");                  // 0x06
            RegisterItem FUNC_EN = Parent.RegMgr.GetRegisterItem("FUNC_EN");                // 0x06
            RegisterItem ODP_3D = Parent.RegMgr.GetRegisterItem("ODP_3D[5:0]");             // 0x07
            RegisterItem D_FREQ = Parent.RegMgr.GetRegisterItem("D_FREQ");                  // 0x07
            RegisterItem FD = Parent.RegMgr.GetRegisterItem("FD(WO)");                      // 0x07

            uint data = 0x00;
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[2];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
#if true
            //I2C Write: Read Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x11);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(2);
#endif
            for (int i = 0x01; i <= 0x06; i++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add(0x20);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)i);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

                RcvData = I2C.ReadBytes(RcvData.Length);
                data = (uint)((RcvData[0] << 8) | RcvData[1]);
                SendBytes.Clear();

                // ADIM_RNG, OD2_REF, FD
                switch (i)
                {
                    case 0x01: // ( “0000” FSF[3:0], ”00”, FS5V[5:0] )
                        FSF.Value = (data & 0x0f00) >> 8;
                        FS5V.Value = (data & 0x003f) >> 0;
                        break;
                    case 0x02: //( PERCENT_SEL[1:0], GATE_CNT, ”000”, ADIM1H[4:0], ADIM1L[4:0] )
                        PERCENT_SEL.Value = (data & 0xc000) >> 14;
                        GATE_CNT.Value = (data & 0x2000) >> 13;
                        ADIM1H.Value = (data & 0x03e0) >> 5;
                        ADIM1L.Value = (data & 0x001f) >> 0;
                        break;
                    case 0x03: //( ODP_EN, DELTA_SEL, “0000”, ADIM2H[4:0], ADIM2L[4:0] )
                        ODP_EN.Value = (data & 0x8000) >> 15;
                        DELTA_SEL.Value = (data & 0x4000) >> 14;
                        ADIM2H.Value = (data & 0x03e0) >> 5;
                        ADIM2L.Value = (data & 0x001f) >> 0;
                        break;
                    case 0x04: //( VCC_DLY_CNT[2:0], “0000”, PWM_FREQ[1:0], ”000000”, ODP_EN_CNT(OEC[1]) )
                        VCC_DLY_CNT.Value = (data & 0xe000) >> 13;
                        PWM_FREQ.Value = (data & 0x0180) >> 7;
                        OEC.Value = (data & 0x0001) >> 0;
                        break;
                    case 0x05: //( JIT_CON, “00”, 2D_REF[3:0], “000”, ODP_3D[5:0] )
                        JITTER_CON.Value = (data & 0x8000) >> 15;
                        D_REF.Value = (data & 0x1e00) >> 9;
                        ODP_3D.Value = (data & 0x003f) >> 0;
                        break;
                    case 0x06: // ( “00000000000”, G_FREQ, PWM_OVER, ODP_EN(????), FUNC_EN, D_FREQ )
                        G_FREQ.Value = (data & 0x0010) >> 4;
                        PWM_OVER.Value = (data & 0x0008) >> 3;
                        //ODP_EN.Value = (data & 0x0004) >> 2;
                        FUNC_EN.Value = (data & 0x0002) >> 1;
                        D_FREQ.Value = (data & 0x0001) >> 0;
                        break;
                    default:
                        break;
                }
            }
            I2C.Config.SlaveAddress = sa;
        }

        private void SLM4781A_ReadI2C()
        {
            RegisterItem FS5V = Parent.RegMgr.GetRegisterItem("FS5V[5:0]");                 // 0x01
            RegisterItem FSF = Parent.RegMgr.GetRegisterItem("FSF[3:0]");                   // 0x02
            RegisterItem ADIM1H = Parent.RegMgr.GetRegisterItem("ADIM1H[4:0]");             // 0x03
            RegisterItem ADIM1L = Parent.RegMgr.GetRegisterItem("ADIM1L[4:0]");             // 0x04
            RegisterItem ADIM2H = Parent.RegMgr.GetRegisterItem("ADIM2H[4:0]");             // 0x05
            RegisterItem ADIM2L = Parent.RegMgr.GetRegisterItem("ADIM2L[4:0]");             // 0x06
            RegisterItem D_REF = Parent.RegMgr.GetRegisterItem("2D_REF[3:0]");              // 0x07
            //RegisterItem OD2_REF = Parent.RegMgr.GetRegisterItem("OD2_REF[5:0](WO)");     // 0x08
            RegisterItem ODP_3D = Parent.RegMgr.GetRegisterItem("ODP_3D[5:0]");             // 0x09
            RegisterItem DELTA_SEL = Parent.RegMgr.GetRegisterItem("DELTA_SEL");            // 0x0A
            RegisterItem JITTER_CON = Parent.RegMgr.GetRegisterItem("JITTER_CON");          // 0x0A
            RegisterItem PERCENT_SEL = Parent.RegMgr.GetRegisterItem("PERCENT_SEL[1:0]");   // 0x0A
            RegisterItem GATE_CNT = Parent.RegMgr.GetRegisterItem("GATE_CNT");              // 0x0A
            RegisterItem G_FREQ = Parent.RegMgr.GetRegisterItem("G_FREQ");                  // 0x0A
            RegisterItem FUNC_EN = Parent.RegMgr.GetRegisterItem("FUNC_EN");                // 0x0A
            RegisterItem D_FREQ = Parent.RegMgr.GetRegisterItem("D_FREQ");                  // 0x0A

            uint data = 0x00;
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[2];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            for (int i = 0x01; i <= 0x06; i++)
            {
                if (i == 0x04) continue;
                SendBytes.Add(0x00);
                SendBytes.Add(0x20);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)i);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

                RcvData = I2C.ReadBytes(RcvData.Length);
                data = (uint)((RcvData[0] << 8) | RcvData[1]);
                SendBytes.Clear();

                // ADIM_RNG, OD2_REF, FD
                switch (i)
                {
                    case 0x01: // ( “0000” FSF[3:0], ”00”, FS5V[5:0] )
                        FSF.Value = (data & 0x0f00) >> 8;
                        FS5V.Value = (data & 0x003f) >> 0;
                        break;
                    case 0x02: //( PERCENT_SEL[1:0], GATE_CNT, ”000”, ADIM1H[4:0], ADIM1L[4:0] )
                        PERCENT_SEL.Value = (data & 0xc000) >> 14;
                        GATE_CNT.Value = (data & 0x2000) >> 13;
                        ADIM1H.Value = (data & 0x03e0) >> 5;
                        ADIM1L.Value = (data & 0x001f) >> 0;
                        break;
                    case 0x03: //( ODP_EN, DELTA_SEL, “0000”, ADIM2H[4:0], ADIM2L[4:0] )
                        DELTA_SEL.Value = (data & 0x4000) >> 14;
                        ADIM2H.Value = (data & 0x03e0) >> 5;
                        ADIM2L.Value = (data & 0x001f) >> 0;
                        break;
                    case 0x05: //( JIT_CON, “00”, 2D_REF[3:0], “000”, ODP_3D[5:0] )
                        JITTER_CON.Value = (data & 0x8000) >> 15;
                        D_REF.Value = (data & 0x1e00) >> 9;
                        ODP_3D.Value = (data & 0x003f) >> 0;
                        break;
                    case 0x06: // ( “00000000000”, G_FREQ, PWM_OVER, ODP_EN(????), FUNC_EN, D_FREQ )
                        G_FREQ.Value = (data & 0x0010) >> 4;
                        FUNC_EN.Value = (data & 0x0002) >> 1;
                        D_FREQ.Value = (data & 0x0001) >> 0;
                        break;
                    default:
                        break;
                }
            }
            I2C.Config.SlaveAddress = sa;
        }
        #endregion Flash control methods

        public override bool CheckConnectionForLog()
        {
            return false;
        }

        public override void RunLog()
        {
        }

        public override void SendCommand(string Command)
        {
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal, Result = 0;
            TEST_ITEMS TestItem = (TEST_ITEMS)TestItemIndex;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch ((TEST_ITEMS)TestItemIndex)
            {
                // FW test functions
                default:
                    break;
            }
            Log.WriteLine(TestItem.ToString() + ":" + iVal.ToString() + ":" + Result.ToString());
        }

        public int RunTest(TEST_ITEMS TestItem, int Arg)
        {
            int iVal = 0;
            //RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");
            //RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");
            //RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");
            //RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");

            switch (TestItem)
            {
                default:
                    break;
            }
            return iVal;
        }
    }

    public class STD1402Q : ChipControl
    {
        #region Variable and declaration

        public enum TEST_ITEMS
        {
            VERSION_SELECT,
            CALIBRATION,
            VCC_CURRENT,
            VREF_TEST,
            UVLO_TEST,
            VPWM_TEST,
            SCP_Threshold_TEST,
            CS_SHORT_PROT_TEST,
            MAX_ONOFF_TIME_TEST,
            DS_SHORT_RPOT_TIME_TEST,
            FET_DS_SHORT_PROT_TEST,
            CS_PIN_OPEN_TEST,
            DS_THRESHOLD_TEST,
            VDRVFBOVP_TEST,
            VDRVFBUVP_TEST,
            VCCOVP_TEST,
            STANDBY_MODE_TEST,
            V2DREF_TEST,
            V2DMIN_MAX_TEST,
            VADIMO_TEST,
            SINGLE_TEST,
            NUM_TEST_ITEMS,
        }

        private JLcLib.Custom.I2C I2C { get; set; }

        public int SlaveAddress { get; private set; } = 0x29;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI PowerSupply1 = null;
        JLcLib.Instrument.SCPI PowerSupply2 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI TempChamber = null;
        JLcLib.Instrument.SCPI SignalGenerator0 = null;
        JLcLib.Instrument.SCPI ElectronicLoad = null;


        #endregion Variable and declaration

        public STD1402Q(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            CalibrationData = new byte[256];

            /* Init test items combo box */
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void WriteRegister(uint Address, uint Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            if (Parent.xlMgr.Sheet.Name == "I2C_Analog")
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Data);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            }
            else
            {
                MessageBox.Show("VPP 전압을 8.75V로 바꾸고 확인을 눌러주세요.");
                //I2C Write: Erase Command;
                SendBytes.Add(0x00);
                SendBytes.Add(0x10);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(210);
                System.Threading.Thread.Sleep(70);

                MessageBox.Show("VPP 전압을 5.5V로 바꾸고 확인을 눌러주세요.");
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Data);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);

                //I2C Write: Write Command
                SendBytes.Add(0x00);
                SendBytes.Add(0x12);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(110);
            }
            I2C.Config.SlaveAddress = sa;
        }

        private void STD1402Q_ReadI2C()
        {
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[2];

            RegisterItem FS5V = Parent.RegMgr.GetRegisterItem("FS5V[5:0]");                 // 0x01
            RegisterItem ADIM1H = Parent.RegMgr.GetRegisterItem("ADIM1H[4:0]");             // 0x02
            RegisterItem ADIM1L = Parent.RegMgr.GetRegisterItem("ADIM1L[4:0]");             // 0x03
            RegisterItem FSF = Parent.RegMgr.GetRegisterItem("FSF[3:0]");                   // 0x04
            RegisterItem TONIUP_1 = Parent.RegMgr.GetRegisterItem("TONIUP_1[2:0]");         // 0x05
            RegisterItem TONIUP_2 = Parent.RegMgr.GetRegisterItem("TONIUP_2[2:0]");         // 0x06
            RegisterItem D_REF = Parent.RegMgr.GetRegisterItem("2D_REF[3:0]");              // 0x07
            RegisterItem OD2_REF = Parent.RegMgr.GetRegisterItem("OD2_REF[5:0]");           // 0x08
            RegisterItem ODP_3D = Parent.RegMgr.GetRegisterItem("ODP_3D[5:0]");             // 0x09

            for (int i = 0x01; i <= 0x08; i++)
            {
                if (i == 7) continue;
                SendBytes.Add(0x00);
                SendBytes.Add(0x20);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)i);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

                RcvData = I2C.ReadBytes(RcvData.Length);
                SendBytes.Clear();

                switch (i)
                {
                    case 0x01:
                        FS5V.Value = (uint)(RcvData[1] & 0x3f);
                        break;
                    case 0x02:
                        ADIM1H.Value = (uint)(RcvData[1] & 0x1f);
                        break;
                    case 0x03:
                        ADIM1L.Value = (uint)(RcvData[1] & 0x1f);
                        break;
                    case 0x04:
                        FSF.Value = (uint)(RcvData[1] & 0x0f);
                        D_REF.Value = (uint)((RcvData[1] & 0xf0) >> 4);
                        break;
                    case 0x05:
                        ODP_3D.Value = (uint)(RcvData[1] & 0x3f);
                        break;
                    case 0x06:
                        TONIUP_1.Value = (uint)(RcvData[1] & 0x07);
                        TONIUP_2.Value = (uint)((RcvData[1] & 0x70) >> 4);
                        break;
                    case 0x08:
                        OD2_REF.Value = (uint)(RcvData[1] & 0x3f);
                        break;
                    default:
                        break;
                }
            }

        }

        private uint ReadRegister(uint Address)
        {
            uint data = 0x00;
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[2];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            if (Parent.xlMgr.Sheet.Name != "I2C_Analog")
            {
                //I2C Write: Read Command
                SendBytes.Add(0x00);
                SendBytes.Add(0x11);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);

                // Data read
                SendBytes.Add(0x00);
                SendBytes.Add(0x20);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)(Address - 0x2F));
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

                RcvData = I2C.ReadBytes(RcvData.Length);
                data = (uint)((RcvData[0] << 8) | RcvData[1]);
                SendBytes.Clear();
            }
            else
            {
                STD1402Q_ReadI2C();
            }

            I2C.Config.SlaveAddress = sa;

            return data;
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

            Parent.ChipCtrlButtons[2].Text = "Dump";
            Parent.ChipCtrlButtons[2].Visible = true;
            Parent.ChipCtrlButtons[2].Click += ReadFlashToAnalog_Click;

            Parent.ChipCtrlButtons[3].Text = "Read";
            Parent.ChipCtrlButtons[3].Visible = true;
            Parent.ChipCtrlButtons[3].Click += ReadFlash_Click;

            Parent.ChipCtrlTextboxes[4].Text = "300"; // Delay Time
            Parent.ChipCtrlTextboxes[4].Visible = true;

            Parent.ChipCtrlButtons[5].Text = "E/W_0";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += Erase_Time_Test_Write_All_zero_Click;

            Parent.ChipCtrlButtons[6].Text = "E/W_1";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Erase_Time_Test_Write_All_one_Click;

            Parent.ChipCtrlButtons[7].Text = "E/W_D";
            Parent.ChipCtrlButtons[7].Visible = true;
            Parent.ChipCtrlButtons[7].Click += Erase_Time_Test_Write_Default_code_Click;


        }

        private void Check_Instrument()
        {
            //PowerSupply0 : 6705B
            //PowerSupply0 : E36313A
            //PowerSupply0 : E36313A
            //DigitalMultimeter0 : 34465A
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
                    case JLcLib.Instrument.InstrumentTypes.PowerSupply2:
                        if (PowerSupply2 == null)
                            PowerSupply2 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (PowerSupply2.IsOpen == false)
                            PowerSupply2.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.DigitalMultimeter0:
                        if (DigitalMultimeter0 == null)
                            DigitalMultimeter0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (DigitalMultimeter0.IsOpen == false)
                            DigitalMultimeter0.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.OscilloScope0:
                        if (OscilloScope0 == null)
                            OscilloScope0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (OscilloScope0.IsOpen == false)
                            OscilloScope0.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.SignalGenerator0:
                        if (SignalGenerator0 == null)
                            SignalGenerator0 = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (SignalGenerator0.IsOpen == false)
                            SignalGenerator0.Open();
                        break;
                    case JLcLib.Instrument.InstrumentTypes.ElectronicLoad:
                        if (ElectronicLoad == null)
                            ElectronicLoad = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (ElectronicLoad.IsOpen == false)
                            ElectronicLoad.Open();
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

        private void ReadFlashToAnalog_Click(object sender, EventArgs e)
        {
            List<byte> SendBytes = new List<byte>();

            //I2C Write: Read Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x11);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(2);
        }

        private void ReadFlash_Click(object sender, EventArgs e)
        {
            if (Parent.xlMgr.Sheet.Name == "I2C_Analog")
            {
                STD1402Q_ReadI2C();
            }
            else
            {
                //STD1402Q_ReadFlash();
            }
        }

        private void Erase_Time_Test_Write_All_zero_Click(object sender, EventArgs e)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            MessageBox.Show("VPP 전압을 8.5V로 바꾸고 확인을 눌러주세요.");
            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();

            System.Threading.Thread.Sleep(int.Parse(Parent.ChipCtrlTextboxes[4].Text) - 2);
            MessageBox.Show("VPP 전압을 5.5V로 바꾸고 확인을 눌러주세요.");
            for (int Address = 0x30; Address <= 0x37; Address++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                if (Address == 0x37)
                {
                    SendBytes.Add(0x80);
                }
                else
                {
                    SendBytes.Add(0x00);
                }
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }
            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(110);

            I2C.Config.SlaveAddress = sa;
        }

        private void Erase_Time_Test_Write_All_one_Click(object sender, EventArgs e)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            MessageBox.Show("VPP 전압을 8.5V로 바꾸고 확인을 눌러주세요.");
            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();

            System.Threading.Thread.Sleep(int.Parse(Parent.ChipCtrlTextboxes[4].Text) - 2);
            MessageBox.Show("VPP 전압을 5.5V로 바꾸고 확인을 눌러주세요.");
            for (int Address = 0x30; Address <= 0x37; Address++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                SendBytes.Add(0xFF);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }
            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(110);

            I2C.Config.SlaveAddress = sa;
        }

        private void Erase_Time_Test_Write_Default_code_Click(object sender, EventArgs e)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            MessageBox.Show("VPP 전압을 8.5V로 바꾸고 확인을 눌러주세요.");
            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();

            System.Threading.Thread.Sleep(int.Parse(Parent.ChipCtrlTextboxes[4].Text) - 2);
            MessageBox.Show("VPP 전압을 5.5V로 바꾸고 확인을 눌러주세요.");
            for (int Address = 0x30; Address <= 0x37; Address++)
            {
                SendBytes.Add(0x00);
                SendBytes.Add((byte)Address);
                SendBytes.Add(0x00);
                switch (Address)
                {
                    case 0x30:
                        SendBytes.Add(0x9F);
                        break;
                    case 0x31:
                        SendBytes.Add(0x54);
                        break;
                    case 0x32:
                        SendBytes.Add(0x8D);
                        break;
                    case 0x33:
                        SendBytes.Add(0x82);
                        break;
                    case 0x34:
                        SendBytes.Add(0x59);
                        break;
                    case 0x35:
                        SendBytes.Add(0xC4);
                        break;
                    case 0x36:
                        SendBytes.Add(0x24);
                        break;
                    case 0x37:
                        SendBytes.Add(0xB0);
                        break;
                }
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }
            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(110);

            I2C.Config.SlaveAddress = sa;
        }

        #region Flase control methods
        private void STD1402Q_WriteFlash()
        {
            RegisterItem FS5V = Parent.RegMgr.GetRegisterItem("FS5V[5:0]");                 // 0x00
            RegisterItem PWM_FREQ = Parent.RegMgr.GetRegisterItem("PWM_FREQ[1:0]");         // 0x00
            RegisterItem ADIM1H = Parent.RegMgr.GetRegisterItem("ADIM1H[4:0]");             // 0x01
            RegisterItem DELTA_SEL = Parent.RegMgr.GetRegisterItem("DELTA_SEL");            // 0x01
            RegisterItem OEC = Parent.RegMgr.GetRegisterItem("OEC[1]");                     // 0x01
            RegisterItem ADIM_RNG = Parent.RegMgr.GetRegisterItem("ADIM_RNG(WO)");          // 0x01
            RegisterItem ADIM1L = Parent.RegMgr.GetRegisterItem("ADIM1L[4:0]");             // 0x02
            RegisterItem VCC_DLY_CNT = Parent.RegMgr.GetRegisterItem("VCC_DLY_CNT[2:0]");   // 0x02
            RegisterItem ADIM2H = Parent.RegMgr.GetRegisterItem("ADIM2H[4:0]");             // 0x03
            RegisterItem PERCENT_SEL = Parent.RegMgr.GetRegisterItem("PERCENT_SEL[1:0]");   // 0x03
            RegisterItem GATE_CNT = Parent.RegMgr.GetRegisterItem("GATE_CNT");              // 0x03
            RegisterItem ADIM2L = Parent.RegMgr.GetRegisterItem("ADIM2L[4:0]");             // 0x04
            RegisterItem G_FREQ = Parent.RegMgr.GetRegisterItem("G_FREQ");                  // 0x04
            RegisterItem PWM_OVER = Parent.RegMgr.GetRegisterItem("PWM_OVER");              // 0x04
            RegisterItem JITTER_CON = Parent.RegMgr.GetRegisterItem("JITTER_CON");          // 0x04
            RegisterItem FSF = Parent.RegMgr.GetRegisterItem("FSF[3:0]");                   // 0x05
            RegisterItem D_REF = Parent.RegMgr.GetRegisterItem("2D_REF[3:0]");              // 0x05
            RegisterItem OD2_REF = Parent.RegMgr.GetRegisterItem("OD2_REF[5:0](WO)");       // 0x06
            RegisterItem ODP_EN = Parent.RegMgr.GetRegisterItem("ODP_EN");                  // 0x06
            RegisterItem FUNC_EN = Parent.RegMgr.GetRegisterItem("FUNC_EN");                // 0x06
            RegisterItem ODP_3D = Parent.RegMgr.GetRegisterItem("ODP_3D[5:0]");             // 0x07
            RegisterItem D_FREQ = Parent.RegMgr.GetRegisterItem("D_FREQ");                  // 0x07
            RegisterItem FD = Parent.RegMgr.GetRegisterItem("FD(WO)");                      // 0x07

            uint data = 0x00;
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            for (int i = 0x30; i <= 0x37; i++)
            {
                switch (i)
                {
                    case 0x30:
                        data = (PWM_FREQ.Value << 6) | FS5V.Value;
                        break;
                    case 0x31:
                        data = (ADIM_RNG.Value << 7) | (OEC.Value << 6) | (DELTA_SEL.Value << 5) | ADIM1H.Value;
                        break;
                    case 0x32:
                        data = (VCC_DLY_CNT.Value << 5) | ADIM1L.Value;
                        break;
                    case 0x33:
                        data = (GATE_CNT.Value << 7) | (PERCENT_SEL.Value << 5) | ADIM2H.Value;
                        break;
                    case 0x34:
                        data = (JITTER_CON.Value << 7) | (PWM_OVER.Value << 6) | (G_FREQ.Value << 5) | ADIM2L.Value;
                        break;
                    case 0x35:
                        data = (D_REF.Value << 4) | FSF.Value;
                        break;
                    case 0x36:
                        data = (FUNC_EN.Value << 7) | (ODP_EN.Value << 6) | OD2_REF.Value;
                        break;
                    case 0x37:
                        data = (FD.Value << 7) | (D_FREQ.Value << 6) | ODP_3D.Value;
                        break;
                    default:
                        data = 0x00;
                        break;
                }

                SendBytes.Add(0x00);
                SendBytes.Add((byte)i);
                SendBytes.Add(0x00);
                SendBytes.Add((byte)data);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(2);
            }

            //I2C Write: Write Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x12);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            System.Threading.Thread.Sleep(110);
            I2C.Config.SlaveAddress = sa;
        }

        public void STD1402Q_EraseFlash()
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            //I2C Write: Erase Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x10);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            System.Threading.Thread.Sleep(210);

            I2C.Config.SlaveAddress = sa;
        }

        private uint STD1402Q_ReadFlash(uint Address)
        {
            uint data = 0x00;
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[2];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            //I2C Write: Read Command
            SendBytes.Add(0x00);
            SendBytes.Add(0x11);
            SendBytes.Add(0x00);
            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
            SendBytes.Clear();
            System.Threading.Thread.Sleep(2);

            // Data read
            SendBytes.Add(0x00);
            SendBytes.Add(0x20);
            SendBytes.Add(0x00);
            SendBytes.Add((byte)(Address - 0x29));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            RcvData = I2C.ReadBytes(RcvData.Length);
            data = (uint)((RcvData[0] << 8) | RcvData[1]);
            SendBytes.Clear();

            I2C.Config.SlaveAddress = sa;

            return data;
        }
        #endregion Flash control methods

        public override bool CheckConnectionForLog()
        {
            return false;
        }

        public override void RunLog()
        {
        }

        private void TEST_Run_version_select(int num)
        {
            double VRT_VAL = 0.0;
            int y_pos;

            Check_Instrument();

            //Power off
            PowerSupply0.Write("OUTP OFF,(@1)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@2)"); // VREF CP_EN
            PowerSupply0.Write("OUTP OFF,(@3)"); // VCC
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //Power Set
            PowerSupply0.Write("VOLT 1.8,(@1)"); // VPP
            PowerSupply0.Write("VOLT 5.0,(@2)"); // VREF CP_EN
            PowerSupply0.Write("VOLT 8.0,(@3)"); // VCC
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            // MessageBox.Show("VCC = 8V\nVPP = 1.8V\nVREF = 5V\nCP_EN = 5V\nPWM = 5V", "Okinawa Verison select");
            MessageBox.Show("CH1 : VPP = 1.8V\nCH2 : VCC = 8V\nCH3 : VREF = PWM = 5V\n", "Okinawa Verison select");

            Parent.xlMgr.Sheet.Select("VER_SEL");
            Parent.xlMgr.Cell.Write(1, 1, "No.");
            Parent.xlMgr.Cell.Write(2, 1, "Version");
            y_pos = num;

            I2C.GPIOs[0].Direction = GPIO_Direction.Output;
            I2C.GPIOs[0].State = GPIO_State.Low;

            while (true)
            {
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n", Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    y_pos++;
                }
                else
                {
                    return;
                }

                Parent.xlMgr.Cell.Write(1, y_pos, (y_pos - 1).ToString());

                PowerSupply0.Write("OUTP ON,(@3)"); // VCC
                JLcLib.Delay.Sleep(500); // wait for power supply to stabilize

                PowerSupply0.Write("OUTP ON,(@1)"); // VPP
                JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

                PowerSupply0.Write("OUTP ON,(@2)"); // VREF CP_EN
                JLcLib.Delay.Sleep(500); // wait for power supply to stabilize
                I2C.GPIOs[0].State = GPIO_State.High;
                JLcLib.Delay.Sleep(500); // wait for power supply to stabilize

                VRT_VAL = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
                JLcLib.Delay.Sleep(500); // wait for power supply to stabilize

                //Power off
                PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
                PowerSupply0.Write("OUTP OFF,(@3)"); // VPP
                PowerSupply0.Write("OUTP OFF,(@1)"); // VREF CP_EN
                I2C.GPIOs[0].State = GPIO_State.Low;
                JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

                if (VRT_VAL > 2.8 && VRT_VAL < 2.9)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V1");
                else if (VRT_VAL > 2.65 && VRT_VAL < 2.75)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V2");
                else if (VRT_VAL > 2.5 && VRT_VAL < 2.6)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V3");
                else if (VRT_VAL > 2.35 && VRT_VAL < 2.45)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V4");
                else if (VRT_VAL > 2.2 && VRT_VAL < 2.3)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V5");
                else if (VRT_VAL > 2.05 && VRT_VAL < 2.15)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V6");
                else if (VRT_VAL > 1.9 && VRT_VAL < 2.0)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V7");
                else if (VRT_VAL > 1.75 && VRT_VAL < 1.85)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V8");
                else if (VRT_VAL > 1.6 && VRT_VAL < 1.7)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V9");
                else if (VRT_VAL > 1.45 && VRT_VAL < 1.55)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V10");
                else if (VRT_VAL > 1.3 && VRT_VAL < 1.4)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V11");
                else if (VRT_VAL > 1.15 && VRT_VAL < 1.25)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V12");
                else if (VRT_VAL > 1.0 && VRT_VAL < 1.1)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V13");
                else if (VRT_VAL > 0.85 && VRT_VAL < 0.95)
                    Parent.xlMgr.Cell.Write(2, y_pos, "V14");
                else
                    Parent.xlMgr.Cell.Write(2, y_pos, "Fail");
            }
        }

        private void TEST_Run_CAL_FS5V_ADIM_FSF(int num)
        {
            RegisterItem FS5V = Parent.RegMgr.GetRegisterItem("FS5V[5:0]");     // 0x01
            RegisterItem ADIM1H = Parent.RegMgr.GetRegisterItem("ADIM1H[4:0]"); // 0x02
            RegisterItem ADIM1L = Parent.RegMgr.GetRegisterItem("ADIM1L[4:0]"); // 0x03
            RegisterItem FSF = Parent.RegMgr.GetRegisterItem("FSF[3:0]");       // 0x04

            double VREF_Target = 5; //5V
            double VREF_Delta = 0;

            double ADIMH_Target = 1.32; //1.32V
            double ADIMH_Delta = 0;

            double ADIML_Target = 0.528; //0.528V
            double ADIML_Delta = 0;

            double FSF_Target = 10; //10ms
            double FSF_Delta = 0;

            uint val_fs5v;
            uint val_ADIMH;
            uint val_ADIML;
            uint val_FSF;

            int y_pos;
            double VREF = 0;
            double CS_Voltage = 0;
            double GATE_ON = 0;
            double OFF_TIME = 0;
            List<byte> SendBytes = new List<byte>();

            Check_Instrument();

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2:3)"); // VCC / CS
            PowerSupply1.Write("OUTP OFF,(@2:3)"); // VDRVFB / ADIM
            PowerSupply2.Write("OUTP OFF,(@1:3)"); // VPP / PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //Power Set
            PowerSupply0.Write("VOLT 13.0,(@2)"); // VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); // VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); // CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); // VDRVFB
            PowerSupply1.Write("VOLT 3.3,(@3)"); // ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); // PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize


            I2C.GPIOs[0].Direction = GPIO_Direction.Output;
            I2C.GPIOs[0].State = GPIO_State.Low; // ADIM Floating
            I2C.GPIOs[1].Direction = GPIO_Direction.Output;
            I2C.GPIOs[1].State = GPIO_State.Low; // I2C Floating

            y_pos = num;

            while (true)
            {
                DialogResult dialog = MessageBox.Show("새로운 칩을 넣고 확인을 눌러주세요.\r\n", Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    y_pos++;
                    Parent.xlMgr.Cell.Write(1, y_pos, (y_pos - 1).ToString());
                }
                else
                {
                    return;
                }
                // FS5V
                //Power Set
                PowerSupply0.Write("VOLT 13.0,(@2)"); // VCC
                PowerSupply2.Write("VOLT 1.8,(@3)"); // VPP
                PowerSupply0.Write("VOLT 0.2,(@3)"); // CS
                PowerSupply1.Write("VOLT 1.5,(@2)"); // VDRVFB
                PowerSupply1.Write("VOLT 3.3,(@3)"); // ADIM
                PowerSupply2.Write("VOLT 5.0,(@1)"); // PWM
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                // Power On
                PowerSupply0.Write("OUTP ON,(@2:3)"); // VCC / CS
                PowerSupply1.Write("OUTP ON,(@2:3)"); // VDRVFB / ADIM
                PowerSupply2.Write("OUTP ON,(@1:3)"); // VPP / PWM
                JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

                Parent.xlMgr.Sheet.Select("I2C_Analog");

                I2C.GPIOs[0].State = GPIO_State.Low; // ADIM Floating
                I2C.GPIOs[1].State = GPIO_State.High; // I2C On

                val_fs5v = 31;
                FS5V.Value = val_fs5v;
                FS5V.Write();

                JLcLib.Delay.Sleep(100); // wait for power supply to stabilize
                VREF = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
                VREF_Delta = VREF_Target - VREF;

                for (int i = 5; i >= 0; i--)
                {
                    //Delta_code +-
                    if (VREF_Delta < -0.05)
                        val_fs5v += 5;
                    else if (VREF_Delta < -0.04)
                        val_fs5v += 4;
                    else if (VREF_Delta < -0.03)
                        val_fs5v += 3;
                    else if (VREF_Delta < -0.02)
                        val_fs5v += 2;
                    else if (VREF_Delta < -0.01)
                        val_fs5v += 1;
                    else if (VREF_Delta < 0.01)
                        val_fs5v += 0;
                    else if (VREF_Delta < 0.02)
                        val_fs5v -= 1;
                    else if (VREF_Delta < 0.03)
                        val_fs5v -= 2;
                    else if (VREF_Delta < 0.04)
                        val_fs5v -= 3;
                    else if (VREF_Delta < 0.05)
                        val_fs5v -= 4;
                    else if (VREF_Delta >= 0.05)
                        val_fs5v -= 5;

                    FS5V.Value = val_fs5v;
                    FS5V.Write();

                    JLcLib.Delay.Sleep(100); // wait for power supply to stabilize
                    VREF = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
                    VREF_Delta = VREF_Target - VREF;

                    if (VREF_Delta >= -0.01 && VREF_Delta < 0.01)
                        break;
                }

                Parent.xlMgr.Sheet.Select("CAL_VAL");
                Parent.xlMgr.Cell.Write(2, y_pos, val_fs5v.ToString());
                Parent.xlMgr.Cell.Write(3, y_pos, VREF.ToString("F3"));
                Parent.xlMgr.Sheet.Select("I2C_Analog");

                // ADIM
                //Power on
                // PowerSupply0.Write("VOLT 13.0,(@2)"); // VCC
                // PowerSupply2.Write("VOLT 1.8,(@3)"); // VPP
                // PowerSupply1.Write("VOLT 1.5,(@2)"); // VDRVFB
                PowerSupply1.Write("VOLT 5.0,(@3)"); // ADIM
                // PowerSupply2.Write("VOLT 5.0,(@1)"); // PWM
                PowerSupply0.Write("VOLT 1.35,(@3)"); // CS

                SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
                SignalGenerator0.Write("SOURce1:VOLTage 0.5"); //DS 0-p set 
                SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.25"); //DS 0-p set 
                SignalGenerator0.Write("SOURce1:FREQ 10000"); //DS 10kHz set
                // JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 
                JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

                /////////////////////////ADD jun ////////////////////////////////////
                val_ADIMH = 20;
                val_ADIML = 13;

                ADIM1H.Value = val_ADIMH;
                ADIM1H.Write();

                ADIM1L.Value = val_ADIML;
                ADIM1L.Write();

                I2C.GPIOs[0].State = GPIO_State.High; // ADIM On
                I2C.GPIOs[1].State = GPIO_State.Low; // I2C Floating

                for (CS_Voltage = 1.35; CS_Voltage >= 1.29; CS_Voltage -= 0.002)
                {
                    PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                    JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                    GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                    if (GATE_ON > 15)
                        break;
                }

                ADIMH_Delta = ADIMH_Target - CS_Voltage;

                for (int i = 5; i >= 0; i--)
                {
                    //Delta_code +-
                    if (ADIMH_Delta < -0.0344)
                        val_ADIMH += 5;
                    else if (ADIMH_Delta < -0.0274)
                        val_ADIMH += 4;
                    else if (ADIMH_Delta < -0.0204)
                        val_ADIMH += 3;
                    else if (ADIMH_Delta < -0.0134)
                        val_ADIMH += 2;
                    else if (ADIMH_Delta < -0.0064)
                        val_ADIMH += 1;
                    else if (ADIMH_Delta < 0.0064)
                        val_ADIMH += 0;
                    else if (ADIMH_Delta < 0.0134)
                        val_ADIMH -= 1;
                    else if (ADIMH_Delta < 0.0204)
                        val_ADIMH -= 2;
                    else if (ADIMH_Delta < 0.0274)
                        val_ADIMH -= 3;
                    else if (ADIMH_Delta < 0.0344)
                        val_ADIMH -= 4;
                    else if (ADIMH_Delta >= 0.0344)
                        val_ADIMH -= 5;

                    I2C.GPIOs[0].State = GPIO_State.Low; // ADIM Floating
                    I2C.GPIOs[1].State = GPIO_State.High; // I2C On

                    ADIM1H.Value = val_ADIMH;
                    ADIM1H.Write();

                    I2C.GPIOs[0].State = GPIO_State.High; // ADIM On
                    I2C.GPIOs[1].State = GPIO_State.Low; // I2C Floating

                    for (CS_Voltage = 1.35; CS_Voltage >= 1.29; CS_Voltage -= 0.002)
                    {
                        PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                        JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                        GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                        if (GATE_ON > 15)
                            break;
                    }
                    ADIMH_Delta = ADIMH_Target - CS_Voltage;
                    if (ADIMH_Delta >= -0.0064 && ADIMH_Delta < 0.0064)
                        break;
                }

                Parent.xlMgr.Sheet.Select("CAL_VAL");
                Parent.xlMgr.Cell.Write(4, y_pos, val_ADIMH.ToString());
                Parent.xlMgr.Cell.Write(5, y_pos, CS_Voltage.ToString("F3"));
                Parent.xlMgr.Sheet.Select("I2C_Analog");

                PowerSupply0.Write("VOLT 0.0,(@3)"); // CS
                PowerSupply1.Write("VOLT 0.0,(@3)"); // ADIM
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

                I2C.GPIOs[0].State = GPIO_State.High; // ADIM On
                I2C.GPIOs[1].State = GPIO_State.Low; // I2C Floating

                for (CS_Voltage = 0.6; CS_Voltage >= 0.512; CS_Voltage -= 0.002)
                {
                    PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                    JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                    GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                    if (GATE_ON > 15)
                        break;
                }
                ADIML_Delta = ADIML_Target - CS_Voltage;

                for (int i = 5; i >= 0; i--)
                {
                    if (ADIML_Delta < -0.0154)
                        val_ADIML += 4;
                    else if (ADIML_Delta < -0.0114)
                        val_ADIML += 3;
                    else if (ADIML_Delta < -0.0074)
                        val_ADIML += 2;
                    else if (ADIML_Delta < -0.0034)
                        val_ADIML += 1;
                    else if (ADIML_Delta < 0.0034)
                        val_ADIML += 0;
                    else if (ADIML_Delta < 0.0074)
                        val_ADIML -= 1;
                    else if (ADIML_Delta < 0.0114)
                        val_ADIML -= 2;
                    else if (ADIML_Delta < 0.0154)
                        val_ADIML -= 3;
                    else if (ADIML_Delta >= 0.0154)
                        val_ADIML -= 4;

                    I2C.GPIOs[0].State = GPIO_State.Low; // ADIM Floating
                    I2C.GPIOs[1].State = GPIO_State.High; // I2C On

                    ADIM1L.Value = val_ADIML;
                    ADIM1L.Write();

                    I2C.GPIOs[0].State = GPIO_State.High; // ADIM On
                    I2C.GPIOs[1].State = GPIO_State.Low; // I2C Floating

                    for (CS_Voltage = 1.35; CS_Voltage >= 1.29; CS_Voltage -= 0.002)
                    {
                        PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                        JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                        GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                        if (GATE_ON > 15)
                            break;
                    }
                    ADIML_Delta = ADIML_Target - CS_Voltage;
                    if (ADIML_Delta >= -0.0034 && ADIMH_Delta < 0.0034)
                        break;
                }

                Parent.xlMgr.Sheet.Select("CAL_VAL");
                Parent.xlMgr.Cell.Write(6, y_pos, val_ADIML.ToString());
                Parent.xlMgr.Cell.Write(7, y_pos, CS_Voltage.ToString("F3"));
                Parent.xlMgr.Sheet.Select("I2C_Analog");

                // FSF
                //Power on
                // PowerSupply0.Write("VOLT 13.0,(@2)"); // VCC
                // PowerSupply2.Write("VOLT 1.8,(@3)"); // VPP
                PowerSupply0.Write("VOLT 0.2,(@3)"); // CS
                // PowerSupply1.Write("VOLT 1.5,(@2)"); // VDRVFB
                PowerSupply1.Write("VOLT 3.3,(@3)"); // ADIM
                // PowerSupply2.Write("VOLT 5.0,(@1)"); // PWM
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

                OscilloScope0.WriteAndReadString(":TIM:SCAL 5E-3");
                JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

                I2C.GPIOs[0].State = GPIO_State.Low; // ADIM Floating
                I2C.GPIOs[1].State = GPIO_State.High; // I2C On

                val_FSF = 2;
                FSF.Value = val_FSF;
                FSF.Write();

                OFF_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")) * 1E+3; //CH1 GATE
                FSF_Delta = FSF_Target - OFF_TIME;

                for (int i = 5; i >= 0; i--)
                {
                    if (FSF_Delta < -0.45)
                        val_FSF += 2;
                    else if (FSF_Delta < -0.25)
                        val_FSF += 1;
                    else if (FSF_Delta < 0.25)
                        val_FSF += 0;
                    else if (FSF_Delta < 0.45)
                        val_FSF -= 1;
                    else if (FSF_Delta >= 0.45)
                        val_FSF -= 2;

                    FSF.Value = val_FSF;
                    FSF.Write();

                    OFF_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")) * 1E+3; //CH1 GATE
                    FSF_Delta = FSF_Target - OFF_TIME;

                    if (FSF_Delta >= -0.25 && FSF_Delta < 0.25)
                        break;
                }

                Parent.xlMgr.Sheet.Select("CAL_VAL");
                Parent.xlMgr.Cell.Write(8, y_pos, val_FSF.ToString());
                Parent.xlMgr.Cell.Write(9, y_pos, OFF_TIME.ToString("F3"));
                Parent.xlMgr.Sheet.Select("Flash");

                PowerSupply2.Write("VOLT 8.5,(@3)"); // VPP
                JLcLib.Delay.Sleep(500); // wait for power supply to stabilize
                //I2C Write: Erase Command
                SendBytes.Add(0x00);
                SendBytes.Add(0x10);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                JLcLib.Delay.Sleep(500); // wait

                PowerSupply2.Write("VOLT 5.5,(@3)"); // VPP
                JLcLib.Delay.Sleep(500); // wait for power supply to stabilize
                for (int Address = 0x30; Address <= 0x37; Address++)
                {
                    SendBytes.Add(0x00);
                    SendBytes.Add((byte)Address);
                    SendBytes.Add(0x00);
                    switch (Address)
                    {
                        case 0x30:
                            // SendBytes.Add(0x9F);
                            SendBytes.Add((byte)(0x80 | val_fs5v));
                            break;
                        case 0x31:
                            //SendBytes.Add(0x54);
                            SendBytes.Add((byte)(0x50 | val_ADIMH));
                            break;
                        case 0x32:
                            //SendBytes.Add(0x8D);
                            SendBytes.Add((byte)(0x80 | val_ADIML));
                            break;
                        case 0x33:
                            //SendBytes.Add(0x82);
                            SendBytes.Add((byte)(0x80 | val_FSF));
                            break;
                        case 0x34:
                            SendBytes.Add(0x59);
                            break;
                        case 0x35:
                            SendBytes.Add(0xC4);
                            break;
                        case 0x36:
                            SendBytes.Add(0x24);
                            break;
                        case 0x37:
                            SendBytes.Add(0xB0);
                            break;
                    }
                    I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                    SendBytes.Clear();
                    System.Threading.Thread.Sleep(2);
                }
                //I2C Write: Write Command
                SendBytes.Add(0x00);
                SendBytes.Add(0x12);
                SendBytes.Add(0x00);
                SendBytes.Add(0x00);
                I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);
                SendBytes.Clear();
                System.Threading.Thread.Sleep(110);

                //Power off
                PowerSupply0.Write("OUTP OFF,(@2:3)"); // VCC / CS
                PowerSupply1.Write("OUTP OFF,(@2:3)"); // VDRVFB / ADIM
                PowerSupply2.Write("OUTP OFF,(@1:3)"); // VPP / PWM

                I2C.GPIOs[0].State = GPIO_State.Low; // ADIM Floating
                I2C.GPIOs[1].State = GPIO_State.Low; // I2C Floating
            }
        }

        private void TEST_Run_VCC_Current()
        {
            double IST = 0;
            double IOP11 = 0;
            double IOP12 = 0;
            double IOP13 = 0;
            double IOP14 = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 6.4V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 1.5V\nPWM = 5V", "Okinawa VCC Current");


            //Power off
            PowerSupply0.Write("SENS:CURR:RANG 10E-3,(@2)");
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 6.4,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 1.5,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM


            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

            //Start up current (IST)
            IST = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)")) * 1000000;
            JLcLib.Delay.Sleep(500); // wait for power supply to stabilize


            //Operating current VCC 13V (IOP11)
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply0.Write("SENS:CURR:RANG AUTO,(@2)");
            JLcLib.Delay.Sleep(500); // wait for power supply to stabilize

            IOP11 = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)")) * 1000;

            //Operating current VCC 9V (IOP12)
            PowerSupply0.Write("VOLT 9.0,(@2)"); //VCC
            JLcLib.Delay.Sleep(500); // wait for power supply to stabilize

            IOP12 = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)")) * 1000;

            //Operating current VCC 20V (IOP13)
            PowerSupply0.Write("VOLT 20.0,(@2)"); //VCC
            JLcLib.Delay.Sleep(500); // wait for power supply to stabilize

            IOP13 = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)")) * 1000;

            MessageBox.Show("IST : " + IST.ToString() + "\n" + "IOP11 : " + IOP11.ToString() + "\n" + "IOP12 : " + IOP12.ToString() + "\n" + "IOP13 : " + IOP13.ToString());


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize
        }

        private void TEST_Run_VREF_TEST()
        {
            double VREF = 0;
            double VREF_A = 0;
            double VREF_LOAD = 0;
            double VLINE = 0;
            double VLOAD = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 3.3V\nPWM = 5V", "Okinawa VREF_TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            ElectronicLoad.Write(":OUTP:STAT OFF"); // VREF_Load
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 3.3,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            ElectronicLoad.Write("FUNC CC"); //load mode set cc
            ElectronicLoad.Write("CURR 0.001"); // load curremt set
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

            //VREF
            VREF = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

            // VREF_A 
            PowerSupply0.Write("VOLT 20.0,(@2)"); //VCC
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            VREF_A = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
            ElectronicLoad.Write(":OUTP:STAT ON"); // PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            VREF_LOAD = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize


            VLOAD = VREF - VREF_LOAD;
            VLINE = VREF - VREF_A;



            MessageBox.Show("VREF : " + VREF.ToString("F3") + "\n" + "VREF_A : " + VREF_A.ToString("F3") + "\n" + "VLINE : " + VLINE.ToString("F3") + "\n" + "VLOAD : " + VLOAD.ToString("F3"));


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            ElectronicLoad.Write(":OUTP:STAT OFF"); // VREF_Load

            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize
        }

        private void TEST_Run_UVLO_TEST()
        {
            double VSTH = 0;
            double VSTL = 0;
            double VCC_Voltage = 0;
            double GATE_High = 0;
            double GATE_Low = 0;



            Check_Instrument();
            MessageBox.Show("VCC = 7.3V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5V\nPWM = 5V", "Okinawa UVLO TEST");

            //Power off
            //PowerSupply0.Write("SENS:CURR:RANG AUTO,(@2)");
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 7.3,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize


            //VSTH
            for (VCC_Voltage = 7.3; VCC_Voltage < 9.3; VCC_Voltage += 0.02)
            {

                PowerSupply0.Write("VOLT " + VCC_Voltage.ToString() + ",(@2)"); //VCC
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_High = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1"));
                if (GATE_High > 1)
                    break;

            }
            VSTH = VCC_Voltage;

            //VSTL
            for (VCC_Voltage = VSTH; VCC_Voltage >= 6.5; VCC_Voltage -= 0.02)
            {

                PowerSupply0.Write("VOLT " + VCC_Voltage.ToString() + ",(@2)"); //VCC
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_Low = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1"));
                if (GATE_Low < 0.5)
                    break;

            }
            VSTL = VCC_Voltage;

            MessageBox.Show("VSTH : " + VSTH.ToString("F3") + "\n" + "VSTL : " + VSTL.ToString("F3"));



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize


        }

        private void TEST_Run_VPWM()
        {
            double VPWM = 0;
            double VPWMHY = 0;
            double GATE_High = 0;
            double GATE_Low = 0;
            double VPWM_Low = 0;
            double VGH = 0;
            double VGL = 0;
            double PWM_Voltage = 0;


            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5V\nPWM = sweep", "Okinawa VPWM TEST");



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            for (PWM_Voltage = 1.16; PWM_Voltage < 1.84; PWM_Voltage += 0.02)
            {

                PowerSupply2.Write("VOLT " + PWM_Voltage.ToString() + ",(@1)"); //PWM
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_High = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
                if (GATE_High > 1)
                    break;
            }

            VGH = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VTOP? CHAN1")); //CH1 GATE
            VPWM = PWM_Voltage;

            for (PWM_Voltage = VPWM; PWM_Voltage >= 0.6; PWM_Voltage -= 0.02)
            {

                PowerSupply2.Write("VOLT " + PWM_Voltage.ToString() + ",(@1)"); //PWM
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_Low = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
                if (GATE_Low < 1)
                    break;
            }

            VGL = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
            VPWMHY = VPWM - PWM_Voltage;

            MessageBox.Show("VPWM : " + VPWM.ToString("F3") + "\n" + "VPWMHY : " + VPWMHY.ToString("F3") + "\n" + "VGH : " + VGH.ToString("F3") + "\n" + "VGL : " + VGL.ToString("F3"));

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

        }

        private void TEST_Run_SCP_Threshold_TEST()
        {

            double CS_Voltage = 0;
            double VRST = 0;
            double OFF_TIME = 0;
            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 3.3V\nPWM = 5\nDS = 100kHz, 3V ", "Okinawa VRST TEST");


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF "); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 3.3,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:FREQ 10000"); //DS 10kHz set
            SignalGenerator0.Write("SOURce1:VOLTage 1.5"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.75"); //DS 0-p set 
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            for (CS_Voltage = 1.5; CS_Voltage < 2.16; CS_Voltage += 0.01)
            {
                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                OFF_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")); //CH1 GATE
                if (OFF_TIME > 2)
                    break;
            }

            VRST = CS_Voltage;

            MessageBox.Show("VRST : " + VRST.ToString("F3"));



            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF "); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize



        }

        private void TEST_Run_CS_SHORT_PROT_TEST()
        {
            double CS_Voltage = 0;
            double VCSRST = 0;
            double TCSMNT = 0;
            double TCSRST = 0;
            double ON_TIME = 0;
            double OFF_TIME = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 3.3V\nPWM = 5\n ", "Okinawa CS_SHORT_PROT TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF "); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 3.3,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            for (CS_Voltage = 0.2; CS_Voltage >= 0; CS_Voltage -= 0.005)
            {
                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                ON_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                if (ON_TIME < 15)
                    break;
            }

            VCSRST = CS_Voltage;

            ON_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE

            TCSMNT = ON_TIME;

            OscilloScope0.WriteAndReadString(":TIM:SCAL 5E-3");
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize


            OFF_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")) * 1E+3; //CH1 GATE

            TCSRST = OFF_TIME;


            MessageBox.Show("VCSRST : " + VCSRST.ToString("F3") + "\n" + "TCSMNT : " + TCSMNT.ToString("F3") + "\n" + "TCSRST : " + TCSRST.ToString("F3"));

            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF "); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize
        }

        private void TEST_Run_MAX_ONOFF_TIME_TEST()
        {
            double ON_TIME = 0;
            double OFF_TIME = 0;


            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 3.3V\nPWM = 5\nDS = 0V ", "Okinawa VRST TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 3.3,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            SignalGenerator0.Write("SOURce1:FUNC DC"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:VOLTage 0"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0"); //DS 0-p set 
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 
            OscilloScope0.WriteAndReadString(":TIM:SCAL 2E-5");
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            OFF_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")) * 1E+6; //CH1 GATE
            ON_TIME = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE

            MessageBox.Show("TONMX : " + ON_TIME.ToString() + "\n" + "TOFFMX : " + OFF_TIME.ToString());


            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize


        }

        private void TEST_RUN_DS_SHORT_PROT_TIME()
        {
            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 0V\nPWM = 5V", "Okinawa FET DS PROT TEST");



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            OscilloScope0.Write(":TRIG:SOUR CHAN3");
            OscilloScope0.Write(":TRIG:LEV 3");
            OscilloScope0.Write(":TRIG:EDGE:SLOP POS");
            OscilloScope0.Write(":SING");

            PowerSupply0.Write("VOLT 3.5,(@3)"); //CS




            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

        }

        private void TEST_Run_FET_DS_SHORT_PROT_TEST()
        {
            double CS_Voltage = 0;
            double FAIL_low = 0;
            double VPDS = 0;
            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 0V\nPWM = 5V", "Okinawa FET DS PROT TEST");



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize


            for (CS_Voltage = 0.2; CS_Voltage < 1.0; CS_Voltage += 0.005)
            {

                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                FAIL_low = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN2")); //CH1 GATE
                if (FAIL_low < 10)
                    break;
            }
            VPDS = CS_Voltage;

            MessageBox.Show("VPDS : " + VPDS.ToString());

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize
        }

        private void TEST_Run_CS_PIN_OPEN()
        {
            double VPCSO = 0;
            double ICS = 0;
            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 0V\nPWM = 5V", "Okinawa CS_OPEN_TEST");

            //Power off
            PowerSupply0.Write("SENS:CURR:RANG 10E-3,(@3)");
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-6");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

            VPCSO = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN3")); //CH3 CS

            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            JLcLib.Delay.Sleep(100); // wait for power supply to stabilize

            ICS = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@3)")) * 1000000;

            MessageBox.Show("VPCSO : " + VPCSO.ToString("F3") + "\n" + "ICS : " + ICS.ToString("F3"));



            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize



        }

        private void TEST_Run_DS_THRESHOLD_TEST()
        {
            double GATE_L = 0;
            double GATE_H = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 1.8V\nVDRVFB = 1.5V\nADIM = 3.3V\nPWM = 5\nDS = 0V ", "Okinawa VDSH TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 3.3,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 1.8,(@3)"); //CS

            SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:VOLTage 0.05"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.0025"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:FREQ 100000"); //DS 10kHz set
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            OscilloScope0.WriteAndReadString(":TIM:SCAL 2E-5");

            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            GATE_H = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")) * 1E+6; //CH1 GATE
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize


            SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:VOLTage 0.15"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.075"); //DS 0-p set 
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            GATE_L = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:NWIDth? CHAN1")) * 1E+6; //CH1 GATE

            if (GATE_H > 15 && GATE_L < 15)
                MessageBox.Show("VDSH : PASS");
            else
                MessageBox.Show("VDSH : FAIL");


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize


        }

        private void TEST_Run_VDRVFBOVP_TEST()
        {
            double VDRVFBOVP = 0;
            double VDRVFBOVPR = 0;
            double VDOVPHY = 0;
            double VDRVFB_Voltage = 0;
            double GATE_High = 0;
            double GATE_Low = 0;



            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5V\nPWM = 5V", "Okinawa UVLO TEST");



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 1E-3");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize


            for (VDRVFB_Voltage = 2.0; VDRVFB_Voltage < 2.4; VDRVFB_Voltage += 0.02)
            {

                PowerSupply1.Write("VOLT " + VDRVFB_Voltage.ToString() + ",(@2)"); //VDRVFB
                GATE_Low = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
                if (GATE_Low < 1)
                    break;
            }

            VDRVFBOVP = VDRVFB_Voltage;

            for (VDRVFB_Voltage = VDRVFBOVP; VDRVFB_Voltage >= 1.9; VDRVFB_Voltage -= 0.02)
            {
                PowerSupply1.Write("VOLT " + VDRVFB_Voltage.ToString() + ",(@2)"); //VDRVFB
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_High = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
                if (GATE_High > 1)
                    break;
            }

            VDRVFBOVPR = VDRVFB_Voltage;
            VDOVPHY = VDRVFBOVP - VDRVFBOVPR;

            MessageBox.Show("VDRVFBOVP : " + VDRVFBOVP.ToString("F3") + "\n" + "VDRVFBOVPR : " + VDRVFBOVPR.ToString("F3") + "\n" + "VDOVPHY : " + VDOVPHY.ToString("F3"));

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize



        }

        private void TEST_Run_VDRVFBUVP_TEST()
        {
            double VDRVFBUVP = 0;
            double VDRVFBUVPR = 0;
            double VDUVPHY = 0;
            double VDRVFB_Voltage = 0;
            double GATE_High = 0;
            double GATE_Low = 0;



            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5V\nPWM = 5V", "Okinawa UVLO TEST");



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize


            for (VDRVFB_Voltage = 1.3; VDRVFB_Voltage >= 0.9; VDRVFB_Voltage -= 0.02)
            {

                PowerSupply1.Write("VOLT " + VDRVFB_Voltage.ToString() + ",(@2)"); //VDRVFB
                GATE_Low = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
                if (GATE_Low < 1)
                    break;
            }

            VDRVFBUVP = VDRVFB_Voltage;

            for (VDRVFB_Voltage = VDRVFBUVP; VDRVFB_Voltage < 1.4; VDRVFB_Voltage += 0.02)
            {
                PowerSupply1.Write("VOLT " + VDRVFB_Voltage.ToString() + ",(@2)"); //VDRVFB
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_High = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN1")); //CH1 GATE
                if (GATE_High > 1)
                    break;
            }

            VDRVFBUVPR = VDRVFB_Voltage;
            VDUVPHY = VDRVFBUVP - VDRVFBUVPR;

            MessageBox.Show("VDRVFBUVP : " + VDRVFBUVP.ToString("F3") + "\n" + "VDRVFBUVPR : " + VDRVFBUVPR.ToString("F3") + "\n" + "VDUVPHY : " + VDUVPHY.ToString("F3"));

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize



        }

        private void TEST_Run_VCCOVP_TEST()
        {
            double VCCOVP = 0;
            double VCCOVPR = 0;
            double VCC_Voltage = 0;
            double FAIL_High = 0;
            double FAIL_Low = 0;



            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5V\nPWM = 5V", "Okinawa UVLO TEST");



            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 2E-5");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize


            for (VCC_Voltage = 14.7; VCC_Voltage < 16.5; VCC_Voltage += 0.02)
            {

                PowerSupply0.Write("VOLT " + VCC_Voltage.ToString() + ",(@2)"); //VCC
                FAIL_Low = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN2")); //CH2 FAIL
                if (FAIL_Low < 5)
                    break;
            }

            VCCOVP = VCC_Voltage;

            for (VCC_Voltage = VCCOVP; VCC_Voltage >= 6.1; VCC_Voltage -= 0.02)
            {
                PowerSupply0.Write("VOLT " + VCC_Voltage.ToString() + ",(@2)"); //VCC
                JLcLib.Delay.Sleep(50); // wait for power supply to stabilize
                FAIL_High = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN2")); //CH1 GATE
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

                if (FAIL_High > 5)
                    break;
            }

            VCCOVPR = VCC_Voltage;

            MessageBox.Show("VCCOVP : " + VCCOVP.ToString("F3") + "\n" + "VCCOVPR : " + VCCOVPR.ToString("F3"));

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize



        }

        private void TEST_Run_STANDBY_MODE_TEST()
        {

            double iSTBY = 0;
            double ACTLD = 0;
            Double VREFst = 0;
            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5V\nPWM = 5V", "Okinawa Standby mode TEST");



            //Power off
            PowerSupply0.Write("SENS:CURR:RANG 10E-3,(@2)");
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.2,(@3)"); //CS
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            OscilloScope0.WriteAndReadString(":TIM:SCAL 2E-5");
            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB

            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            iSTBY = double.Parse(PowerSupply0.WriteAndReadString("MEAS:CURR? (@2)")) * 1000000;
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                                    // ACTLD = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:VAV? CHAN4")); //CH4 ACTLD
                                    //JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
            VREFst = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            MessageBox.Show("iSTBY : " + iSTBY.ToString("F3") + /*"\n" + "ACTLD : " + ACTLD.ToString("F3") +*/ "\n" + "VREFst : " + VREFst.ToString("F3"));

            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM



        }

        private void TEST_Run_V2DREF_TEST()
        {
            double CS_Voltage = 0;
            double GATE_ON = 0;
            double V2DREF = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5.0V\nPWM = 5\nDS = pulse ", "Okinawa V2DREF TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 1.35,(@3)"); //CS

            SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:VOLTage 0.5"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.25"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:FREQ 10000"); //DS 10kHz set
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 


            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            for (CS_Voltage = 1.35; CS_Voltage >= 1.29; CS_Voltage -= 0.002)
            {

                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE

                if (GATE_ON > 15)
                    break;

            }
            V2DREF = CS_Voltage;


            MessageBox.Show("V2DREF : " + V2DREF.ToString("F3"));


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize


        }

        private void TEST_Run_V2DMIN_MAX_TEST()
        {
            double CS_Voltage = 0;
            double GATE_ON = 0;
            double V2DMAX = 0;
            double V2DMIN = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5.0V\nPWM = 5.0\nDS = pulse ", "Okinawa V2DMIN_MAX TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 1.35,(@3)"); //CS

            SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:VOLTage 0.5"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.25"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:FREQ 10000"); //DS 10kHz set
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            for (CS_Voltage = 1.35; CS_Voltage >= 1.29; CS_Voltage -= 0.002)
            {

                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                if (GATE_ON > 15)
                    break;

            }
            V2DMAX = CS_Voltage;

            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize


            for (CS_Voltage = 0.6; CS_Voltage >= 0.512; CS_Voltage -= 0.002)
            {

                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                JLcLib.Delay.Sleep(10); // wait for power supply to stabilize
                GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE
                if (GATE_ON > 15)
                    break;

            }
            V2DMIN = CS_Voltage;





            MessageBox.Show("V2DREF : " + V2DMAX.ToString("F3") + "\n" + "V2DMIN : " + V2DMIN.ToString("F3"));


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

        }

        private void TEST_Run_VADIMO_TEST()
        {
            double CS_Voltage = 0;
            double GATE_ON = 0;
            double VADIMO = 0;

            Check_Instrument();
            MessageBox.Show("VCC = 13V\nVPP = 1.8V\nCS = 0.2V\nVDRVFB = 1.5V\nADIM = 5.0V\nPWM = 5\nDS = pulse ", "Okinawa VADIMO TEST");

            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize

            //initial set
            PowerSupply0.Write("VOLT 0.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 0.0,(@3)"); //VPP
            PowerSupply0.Write("VOLT 0.0,(@3)"); //CS
            PowerSupply1.Write("VOLT 0.0,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 0.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 0.0,(@1)"); //PWM
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            //Power on
            PowerSupply0.Write("VOLT 13.0,(@2)"); //VCC
            PowerSupply2.Write("VOLT 1.8,(@3)"); //VPP
            PowerSupply1.Write("VOLT 1.5,(@2)"); //VDRVFB
            PowerSupply1.Write("VOLT 5.0,(@3)"); //ADIM
            PowerSupply2.Write("VOLT 5.0,(@1)"); //PWM
            PowerSupply0.Write("VOLT 1.35,(@3)"); //CS

            SignalGenerator0.Write("SOURce1:FUNC SQU"); //DS 구형파 set
            SignalGenerator0.Write("SOURce1:VOLTage 0.5"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:VOLTage:OFFSet 0.25"); //DS 0-p set 
            SignalGenerator0.Write("SOURce1:FREQ 10000"); //DS 10kHz set
            JLcLib.Delay.Sleep(10); // wait for power supply to stabilize

            PowerSupply0.Write("OUTP ON,(@2)"); // VCC
            PowerSupply2.Write("OUTP ON,(@3)"); // VPP
            PowerSupply0.Write("OUTP ON,(@3)"); // CS
            PowerSupply1.Write("OUTP ON,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP ON,(@3)"); // ADIM
            PowerSupply2.Write("OUTP ON,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT ON"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1000); // wait for power supply to stabilize

            MessageBox.Show("ADIM PIN OPEN 해주세요", "Okinawa VADIMO TEST");


            for (CS_Voltage = 1.372; CS_Voltage >= 1.28; CS_Voltage -= 0.002)
            {

                PowerSupply0.Write("VOLT " + CS_Voltage.ToString() + ",(@3)"); //CS
                GATE_ON = double.Parse(OscilloScope0.WriteAndReadString(":MEAS:PWIDth? CHAN1")) * 1E+6; //CH1 GATE

                if (GATE_ON > 15)
                    break;

            }
            VADIMO = CS_Voltage;

            MessageBox.Show("VADIMO : " + VADIMO.ToString("F3"));


            //Power off
            PowerSupply0.Write("OUTP OFF,(@2)"); // VCC
            PowerSupply2.Write("OUTP OFF,(@3)"); // VPP
            PowerSupply0.Write("OUTP OFF,(@3)"); // CS
            PowerSupply1.Write("OUTP OFF,(@2)"); // VDRVFB
            PowerSupply1.Write("OUTP OFF,(@3)"); // ADIM
            PowerSupply2.Write("OUTP OFF,(@1)"); // PWM
            SignalGenerator0.Write("OUTP:STAT OFF"); //DS p-p 3V 100kHz 
            JLcLib.Delay.Sleep(1); // wait for power supply to stabilize


        }

        private void SINGLE_TEST()
        {
            TEST_Run_VCC_Current();
            TEST_Run_VREF_TEST();
            TEST_Run_UVLO_TEST();
            TEST_Run_VPWM();
            TEST_Run_SCP_Threshold_TEST();
            TEST_Run_CS_SHORT_PROT_TEST();
            TEST_Run_MAX_ONOFF_TIME_TEST();
            TEST_Run_FET_DS_SHORT_PROT_TEST();
            TEST_Run_CS_PIN_OPEN();
            TEST_Run_DS_THRESHOLD_TEST();
            TEST_Run_VDRVFBOVP_TEST();
            TEST_Run_VDRVFBUVP_TEST();
            TEST_Run_VCCOVP_TEST();
            TEST_Run_STANDBY_MODE_TEST();
            TEST_Run_V2DREF_TEST();
            TEST_Run_V2DMIN_MAX_TEST();
            TEST_Run_VADIMO_TEST();


            MessageBox.Show("끝");


        }

        public override void SendCommand(string Command)
        {
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal, Result = 0;
            TEST_ITEMS TestItem = (TEST_ITEMS)TestItemIndex;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch ((TEST_ITEMS)TestItemIndex)
            {
                // FW test functions
                case TEST_ITEMS.VERSION_SELECT:
                    if ((Arg == "") || (iVal == 0))
                        iVal = 1;
                    TEST_Run_version_select(iVal);
                    break;
                case TEST_ITEMS.CALIBRATION:
                    if ((Arg == "") || (iVal == 0))
                        iVal = 1;
                    TEST_Run_CAL_FS5V_ADIM_FSF(iVal);
                    break;
                case TEST_ITEMS.VCC_CURRENT:
                    TEST_Run_VCC_Current();
                    break;
                case TEST_ITEMS.VREF_TEST:
                    TEST_Run_VREF_TEST();
                    break;
                case TEST_ITEMS.UVLO_TEST:
                    TEST_Run_UVLO_TEST();
                    break;
                case TEST_ITEMS.VPWM_TEST:
                    TEST_Run_VPWM();
                    break;
                case TEST_ITEMS.SCP_Threshold_TEST:
                    TEST_Run_SCP_Threshold_TEST();
                    break;
                case TEST_ITEMS.CS_SHORT_PROT_TEST:
                    TEST_Run_CS_SHORT_PROT_TEST();
                    break;

                case TEST_ITEMS.MAX_ONOFF_TIME_TEST:
                    TEST_Run_MAX_ONOFF_TIME_TEST();
                    break;

                case TEST_ITEMS.DS_SHORT_RPOT_TIME_TEST:
                    TEST_RUN_DS_SHORT_PROT_TIME();
                    break;


                case TEST_ITEMS.FET_DS_SHORT_PROT_TEST:
                    TEST_Run_FET_DS_SHORT_PROT_TEST();
                    break;
                case TEST_ITEMS.CS_PIN_OPEN_TEST:
                    TEST_Run_CS_PIN_OPEN();
                    break;


                case TEST_ITEMS.DS_THRESHOLD_TEST:
                    TEST_Run_DS_THRESHOLD_TEST();
                    break;
                case TEST_ITEMS.VDRVFBOVP_TEST:
                    TEST_Run_VDRVFBOVP_TEST();
                    break;
                case TEST_ITEMS.VDRVFBUVP_TEST:
                    TEST_Run_VDRVFBUVP_TEST();
                    break;
                case TEST_ITEMS.VCCOVP_TEST:
                    TEST_Run_VCCOVP_TEST();
                    break;
                case TEST_ITEMS.STANDBY_MODE_TEST:
                    TEST_Run_STANDBY_MODE_TEST();
                    break;

                case TEST_ITEMS.V2DREF_TEST:
                    TEST_Run_V2DREF_TEST();
                    break;
                case TEST_ITEMS.V2DMIN_MAX_TEST:
                    TEST_Run_V2DMIN_MAX_TEST();
                    break;
                case TEST_ITEMS.VADIMO_TEST:
                    TEST_Run_VADIMO_TEST();
                    break;
                case TEST_ITEMS.SINGLE_TEST:
                    SINGLE_TEST();
                    break;

                default:
                    break;
            }
            Log.WriteLine(TestItem.ToString() + ":" + iVal.ToString() + ":" + Result.ToString());
        }

        public int RunTest(TEST_ITEMS TestItem, int Arg)
        {
            int iVal = 0;
            //RegisterItem CommandReg = Parent.RegMgr.GetRegisterItem("TEST_COMMAND[7:0]");
            //RegisterItem StatusReg = Parent.RegMgr.GetRegisterItem("TEST_STATUS[7:0]");
            //RegisterItem Arg0Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG0[7:0]");
            //RegisterItem Arg1Reg = Parent.RegMgr.GetRegisterItem("TEST_ARG1[7:0]");

            switch (TestItem)
            {
                default:
                    break;
            }
            return iVal;
        }
    }
}
