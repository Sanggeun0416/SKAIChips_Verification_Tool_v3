using System;
using System.Text;
using SKAIChips_Verification;
using JLcLib.Chip;

namespace TI
{
    public class TMP117 : ChipControl
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

        public TMP117(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            Serial.ReadSettingFile(form.IniFile, "TMP117");
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
            byte[] SendData = new byte[3];

            SendData[0] = (byte)(Address & 0xFF);
            SendData[1] = (byte)((Data >> 8) & 0xFF);
            SendData[2] = (byte)(Data & 0xFF);

            iComn.WriteBytes(SendData, SendData.Length, true);
        }

        private uint ReadRegister(uint Address)
        {
            byte[] SendData = new byte[1];
            byte[] RcvData = new byte[2];

            SendData[0] = (byte)(Address & 0xFF);

            iComn.WriteBytes(SendData, SendData.Length, true);
            RcvData = iComn.ReadBytes(RcvData.Length);

            return (uint)((RcvData[0] << 8) | RcvData[1]);
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
            Parent.ChipCtrlButtons[0].Text = "TEMP";
            Parent.ChipCtrlButtons[0].Visible = true;
            Parent.ChipCtrlButtons[0].Click += Read_Temperature_Click;
        }

        private void Read_Temperature_Click(object sender, EventArgs e)
        {
            RegisterItem TEMP = Parent.RegMgr.GetRegisterItem("TEMP[15:0]");
            double temp;

            TEMP.Read();
            temp = (Int16)TEMP.Value * 0.0078125;
            Log.WriteLine("TEMP[15:0] : " + TEMP.Value.ToString() + "  Temp : " + temp.ToString());
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
