using JLcLib.Chip;
using JLcLib.Comn;
using SKAIChips_Verification;

namespace GCT
{
    public class GRF7255 : ChipControl
    {
        private JLcLib.Custom.SPI SPI = null;

        public GRF7255(RegContForm form) : base(form)
        {
            SPI = Parent.SPI;
        }

        #region Chip register control methods
        private void GRF7255RF_WriteRegister(uint Address, uint Data)
        {
            byte Cmd, Addr;
            byte[] Bytes = new byte[3];

            if (iComn.ComnType == WireComnTypes.SPI)

                if (SPI == null || SPI.ftMPSSE == null || SPI.IsOpen == false) // Check SPI connection
                    return;

            // Prepare data
            Cmd = (byte)((0x2 << 2) | ((Address >> 8) & 0x3)); // command 3bit, address msb 2bit
            Cmd <<= 3; // MSB first, 5 bits sending
            Addr = (byte)(Address & 0xFF); // address lsb 8bit
            Bytes[0] = (byte)((Data >> 16) & 0xff);
            Bytes[1] = (byte)((Data >> 8) & 0xff);
            Bytes[2] = (byte)((Data >> 0) & 0xff);

            // Send data
            SPI.GPIOs[1].State = GPIO_State.Low; // SPI_MODE
            SPI.GPIOs[0].State = GPIO_State.Low; // BG_EN low
            SPI.ftMPSSE.SPI_SetStart(); // SCK, MOSI, MISO, SS are low.
            SPI.ftMPSSE.SSC_SetBits(Cmd, 5, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6], bit[5] .. bit[3].
            SPI.ftMPSSE.SSC_SetBits(Addr, 8, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6], bit[5] .. bit[0].
            SPI.ftMPSSE.SSC_SetAllPins(JLcLib.Custom.MPSSE.Command.PinConfig.Write, 1, false, 0x01, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising);
            for (int i = 0; i < 4; i++) // Data msb 4bit [27:24]
                SPI.ftMPSSE.SSC_SetAllPins(JLcLib.Custom.MPSSE.Command.PinConfig.Write, 1, ((Data >> (27 - i)) & 1) == 1, 0x01, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising);
            SPI.ftMPSSE.SSC_SetBytes(Bytes, 3, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // Data lsb 24 bits [23:0]
            SPI.ftMPSSE.GPIOL_SetPins(0xFB, 0xE8); // SS is high
            SPI.ftMPSSE.SSC_SetBits(0x00, 2, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // 2 clock dummy
            SPI.ftMPSSE.GPIOL_SetPins(0xFB, 0xF8); // BG_EN is high
            SPI.ftMPSSE.SendCommand();
        }

        private uint GRF7255RF_ReadRegister(uint Address)
        {
            byte Cmd, Addr;
            uint Data = 0xFFFFFFFF;

            if (SPI == null || SPI.ftMPSSE == null || SPI.IsOpen == false) // Check SPI connection
                return Data;

            // Prepare data
            Cmd = (byte)((0x3 << 2) | ((Address >> 8) & 0x3)); // command 3bit, address msb 2bit
            Cmd <<= 3; // MSB first, 5 bits sending
            Addr = (byte)(Address & 0xFF); // address lsb 8bit

            // Send data
            SPI.GPIOs[1].State = GPIO_State.Low; // SPI_MODE
            //SPI.GPIOs[0].State = GPIO_State.Low; // BG_EN low
            SPI.ftMPSSE.SPI_SetStart(); // SCK, MOSI, MISO, SS are low.
            SPI.ftMPSSE.SSC_SetBits(Cmd, 5, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6], bit[5] .. bit[3].
            SPI.ftMPSSE.SSC_SetBits(Addr, 8, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6], bit[5] .. bit[0].
            SPI.ftMPSSE.SSC_SetAllPins(JLcLib.Custom.MPSSE.Command.PinConfig.Write, 2, false, 0x03, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising);
            SPI.ftMPSSE.SSC_SetBits(0x00, 4, JLcLib.Custom.MPSSE.Command.PinConfig.Read, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.InOutRising); // Data msb 4 bit [27:24]
            SPI.ftMPSSE.SSC_SetBytes(null, 3, JLcLib.Custom.MPSSE.Command.PinConfig.Read, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.InOutRising); // Data lsb 24 bits [23:0]
            SPI.ftMPSSE.GPIOL_SetPins(0xFB, 0xF8); // SS is high
            SPI.ftMPSSE.SSC_SetBits(0x00, 1, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // 1 clock dummy
            SPI.ftMPSSE.SendCommand();
            byte[] Bytes = SPI.ftMPSSE.GetReceivedBytes(4);

            if (Bytes != null && Bytes.Length >= 4)
                Data = (uint)(((Bytes[0] & 0xF) << 24) | (Bytes[1] << 16) | (Bytes[2] << 8) | (Bytes[3] << 0));

            return Data;
        }

        private void GRF7255IF_WriteRegister(uint Address, uint Data)
        {
            byte Cmd, Addr;
            byte[] Bytes = new byte[3];
            byte rwFlag = 0;

            if (SPI == null || SPI.ftMPSSE == null || SPI.IsOpen == false) // Check SPI connection
                return;

            // Prepare data
            Cmd = (byte)((rwFlag << 1) | (0 << 0)); // bit[0] is Address[8]
            Cmd <<= 6; // MSB first, 2 bit sending
            Addr = (byte)(Address & 0xFF); // address lsb 8bit
            Bytes[0] = (byte)((Data >> 16) & 0xff);
            Bytes[1] = (byte)((Data >> 8) & 0xff);
            Bytes[2] = (byte)((Data >> 0) & 0xff);

            // Send data
            SPI.ftMPSSE.SPI_SetStart(); // SCK, MOSI, MISO, SS are low.
            SPI.ftMPSSE.SSC_SetBits(Cmd, 2, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6]
            SPI.ftMPSSE.SSC_SetBits(Addr, 8, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6], bit[5] .. bit[0].
            for (int i = 0; i < 4; i++) // Data msb 4bit [27:24]
                SPI.ftMPSSE.SSC_SetAllPins(JLcLib.Custom.MPSSE.Command.PinConfig.Write, 1, ((Data >> (27 - i)) & 1) == 1, 0x00, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising);
            SPI.ftMPSSE.SSC_SetBytes(Bytes, 3, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // Data lsb 24 bits [23:0]
            SPI.ftMPSSE.SPI_SetStop();
            SPI.ftMPSSE.SendCommand();
        }

        private uint GRF7255IF_ReadRegister(uint Address)
        {
            byte Cmd, Addr;
            byte rwFlag = 1;
            uint Data = 0xFFFFFFFF;

            if (SPI == null || SPI.ftMPSSE == null || SPI.IsOpen == false) // Check SPI connection
                return Data;

            // Prepare data
            Cmd = (byte)((rwFlag << 1) | (0 << 0)); // bit[0] is Address[8]
            Cmd <<= 6; // MSB first, 2 bit sending
            Addr = (byte)(Address & 0xFF); // address lsb 8bit

            // Send data
            SPI.ftMPSSE.SPI_SetStart(); // SCK, MOSI, MISO, SS are low.
            SPI.ftMPSSE.SSC_SetBits(Cmd, 2, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6]
            SPI.ftMPSSE.SSC_SetBits(Addr, 8, JLcLib.Custom.MPSSE.Command.PinConfig.Write, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.OutRising); // bit[7], bit[6], bit[5] .. bit[0].
            SPI.ftMPSSE.SSC_SetBits(0x00, 4, JLcLib.Custom.MPSSE.Command.PinConfig.Read, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.InOutRising); // Data msb 4 bit [27:24]
            SPI.ftMPSSE.SSC_SetBytes(null, 3, JLcLib.Custom.MPSSE.Command.PinConfig.Read, JLcLib.Custom.MPSSE.Command.BitFirst.MSB, JLcLib.Custom.MPSSE.Command.ClockEdge.InOutRising); // Data lsb 24 bits [23:0]
            SPI.ftMPSSE.SPI_SetStop();
            SPI.ftMPSSE.SendCommand();
            byte[] Bytes = SPI.ftMPSSE.GetReceivedBytes(4);

            if (Bytes != null && Bytes.Length >= 4)
                Data = (uint)(((Bytes[0] & 0xF) << 24) | (Bytes[1] << 16) | (Bytes[2] << 8) | (Bytes[3] << 0));

            return Data;
        }
        #endregion Chip register control methods

        public override JLcLib.Chip.RegisterGroup MakeRegisterGroup(string GroupName, string[,] RegData)
        {
            RegisterGroup rg = null;

            if (ChipType == RegContForm.ChipTypes.GRF7255RF)
                rg = new RegisterGroup(GroupName, GRF7255RF_WriteRegister, GRF7255RF_ReadRegister);
            //else if (ChipType == RegContForm.ChipTypes.GRF7255IF)
            else
                rg = new RegisterGroup(GroupName, GRF7255IF_WriteRegister, GRF7255IF_ReadRegister);

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
        }

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
        }
    }
}
