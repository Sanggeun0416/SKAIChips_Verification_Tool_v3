using System;
using System.Text;
using SKAIChips_Verification;
using JLcLib.Comn;
namespace HD_MOBIS
{
    public class Santorini : ChipControl
    {
        #region Variable and declaration
        public enum TEST_ITEMS
        {
            TEST,

            NUM_TEST_ITEMS,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Comn.Serial Serial { get; set; } = new JLcLib.Comn.Serial();
        private bool IsSerialReceivedData = false;
        private bool IsRunCal = false;
        #endregion Variable and declaration

        public Santorini(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "Santorini");
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

        byte[] crcTable = { 0x00 ,0x1D ,0x3A ,0x27 ,0x74 ,0x69 ,0x4E ,0x53 ,0xE8 ,0xF5 ,0xD2 ,0xCF ,0x9C ,0x81 ,0xA6 ,0xBB,
                            0xCD ,0xD0 ,0xF7 ,0xEA ,0xB9 ,0xA4 ,0x83 ,0x9E ,0x25 ,0x38 ,0x1F ,0x02 ,0x51 ,0x4C ,0x6B ,0x76,
                            0x87 ,0x9A ,0xBD ,0xA0 ,0xF3 ,0xEE ,0xC9 ,0xD4 ,0x6F ,0x72 ,0x55 ,0x48 ,0x1B ,0x06 ,0x21 ,0x3C,
                            0x4A ,0x57 ,0x70 ,0x6D ,0x3E ,0x23 ,0x04 ,0x19 ,0xA2 ,0xBF ,0x98 ,0x85 ,0xD6 ,0xCB ,0xEC ,0xF1,
                            0x13 ,0x0E ,0x29 ,0x34 ,0x67 ,0x7A ,0x5D ,0x40 ,0xFB ,0xE6 ,0xC1 ,0xDC ,0x8F ,0x92 ,0xB5 ,0xA8,
                            0xDE ,0xC3 ,0xE4 ,0xF9 ,0xAA ,0xB7 ,0x90 ,0x8D ,0x36 ,0x2B ,0x0C ,0x11 ,0x42 ,0x5F ,0x78 ,0x65,
                            0x94 ,0x89 ,0xAE ,0xB3 ,0xE0 ,0xFD ,0xDA ,0xC7 ,0x7C ,0x61 ,0x46 ,0x5B ,0x08 ,0x15 ,0x32 ,0x2F,
                            0x59 ,0x44 ,0x63 ,0x7E ,0x2D ,0x30 ,0x17 ,0x0A ,0xB1 ,0xAC ,0x8B ,0x96 ,0xC5 ,0xD8 ,0xFF ,0xE2,
                            0x26 ,0x3B ,0x1C ,0x01 ,0x52 ,0x4F ,0x68 ,0x75 ,0xCE ,0xD3 ,0xF4 ,0xE9 ,0xBA ,0xA7 ,0x80 ,0x9D,
                            0xEB ,0xF6 ,0xD1 ,0xCC ,0x9F ,0x82 ,0xA5 ,0xB8 ,0x03 ,0x1E ,0x39 ,0x24 ,0x77 ,0x6A ,0x4D ,0x50,
                            0xA1 ,0xBC ,0x9B ,0x86 ,0xD5 ,0xC8 ,0xEF ,0xF2 ,0x49 ,0x54 ,0x73 ,0x6E ,0x3D ,0x20 ,0x07 ,0x1A,
                            0x6C ,0x71 ,0x56 ,0x4B ,0x18 ,0x05 ,0x22 ,0x3F ,0x84 ,0x99 ,0xBE ,0xA3 ,0xF0 ,0xED ,0xCA ,0xD7,
                            0x35 ,0x28 ,0x0F ,0x12 ,0x41 ,0x5C ,0x7B ,0x66 ,0xDD ,0xC0 ,0xE7 ,0xFA ,0xA9 ,0xB4 ,0x93 ,0x8E,
                            0xF8 ,0xE5 ,0xC2 ,0xDF ,0x8C ,0x91 ,0xB6 ,0xAB ,0x10 ,0x0D ,0x2A ,0x37 ,0x64 ,0x79 ,0x5E ,0x43,
                            0xB2 ,0xAF ,0x88 ,0x95 ,0xC6 ,0xDB ,0xFC ,0xE1 ,0x5A ,0x47 ,0x60 ,0x7D ,0x2E ,0x33 ,0x14 ,0x09,
                            0x7F ,0x62 ,0x45 ,0x58 ,0x0B ,0x16 ,0x31 ,0x2C ,0x97 ,0x8A ,0xAD ,0xB0 ,0xE3 ,0xFE ,0xD9 ,0xC4
                        };

        private void WriteRegister(uint Address, uint Data)
        {
            byte[] SendData = new byte[6];
            byte[] RcvData = new byte[7];
            byte crc = 0xff;
            byte rev_data;

            SendData[0] = 0xAA;                         // Sync
            SendData[1] = 0x80;                         // ID(W)
            SendData[2] = (byte)(Address & 0xFF);       // ADDR
            SendData[3] = (byte)((Data >> 8) & 0xFF);   // DATA_MSB
            SendData[4] = (byte)(Data & 0xFF);          // DATA_LSB
            for (int i = 1; i < 5; i++)
            {
                crc = crcTable[crc ^ SendData[i]];
            }
            crc = (byte)~crc;
            // Log.WriteLine(crc.ToString("X"));
            SendData[5] = (byte)crc;                          // CRC

            // MSB first
            for (int i = 0; i < SendData.Length; i++)
            {
                rev_data = 0x00;
                for (int j = 0; j < 8; j++)
                {
                    rev_data <<= 1;
                    rev_data |= (byte)((SendData[i] >> j) & 0x01);
                }
                SendData[i] = rev_data;
            }

            iComn.QueueClear();
            iComn.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(100);
            RcvData = iComn.ReadBytes(RcvData.Length);

            if (RcvData != null)
            {
                // MSB first
                for (int i = 0; i < RcvData.Length; i++)
                {
                    rev_data = 0x00;
                    for (int j = 0; j < 8; j++)
                    {
                        rev_data <<= 1;
                        rev_data |= (byte)((RcvData[i] >> j) & 0x01);
                    }
                    RcvData[i] = rev_data;
                }
                if (RcvData[6] != 0xA6)
                {
                    Log.WriteLine("Abnormal writing!!");
                }
                /*
                for (int i = 0; i < RcvData.Length; i++)
                {
                    Log.WriteLine(i.ToString() + " : " + RcvData[i].ToString("X"));
                }
                */
            }
            else
            {
                Log.WriteLine("Abnormal writing!!");
            }
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[4];
            byte[] RcvData = new byte[9];
            byte crc = 0xff;
            byte rev_data;

            SendData[0] = 0xAA;                     // Sync
            SendData[1] = 0x00;                     // ID(R)
            SendData[2] = (byte)(Address & 0xFF);   // ADDR
            for (int i = 1; i < 3; i++)
            {
                crc = crcTable[crc ^ SendData[i]];
            }
            crc = (byte)~crc;
            // Log.WriteLine(crc.ToString("X"));
            SendData[3] = crc;                      // CRC

            // MSB first
            for (int i = 0; i < SendData.Length; i++)
            {
                rev_data = 0x00;
                for (int j = 0; j < 8; j++)
                {
                    rev_data <<= 1;
                    rev_data |= (byte)((SendData[i] >> j) & 0x01);
                }
                SendData[i] = rev_data;
            }

            iComn.QueueClear();
            iComn.WriteBytes(SendData, SendData.Length, true);
            System.Threading.Thread.Sleep(100);
            RcvData = iComn.ReadBytes(RcvData.Length);

            if (RcvData != null)
            {
                // MSB first
                for (int i = 0; i < RcvData.Length; i++)
                {
                    rev_data = 0x00;
                    for (int j = 0; j < 8; j++)
                    {
                        rev_data <<= 1;
                        rev_data |= (byte)((RcvData[i] >> j) & 0x01);
                    }
                    RcvData[i] = rev_data;
                }
                /*
                for (int i = 0; i < RcvData.Length;i++)
                {
                    Log.WriteLine(i.ToString() + " : " + RcvData[i].ToString("X"));
                }
                */
            }
            else
            {
                Log.WriteLine("RevData is Null!!");
                return 0xffff;
            }

            return (uint)((RcvData[6] << 8) | RcvData[7]);
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
            Parent.ChipCtrlButtons[0].Text = "WakeUp";
            Parent.ChipCtrlButtons[0].Visible = true;
            Parent.ChipCtrlButtons[0].Click += Santorini_WakeUp_Click;
        }

        private void Santorini_WakeUp_Click(object sender, EventArgs e)
        {
            byte[] SendData = new byte[2];

            Serial.Config.BaudRate = 5000;


            SendData[0] = 0x55;
            SendData[1] = 0xD5;

            iComn.WriteBytes(SendData, SendData.Length, true);
        }

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
    }
}
