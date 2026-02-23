using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using FTD2XX_NET;
using JLcLib.Chip;
using JLcLib.Comn;
using JLcLib.Custom;
using SKAIChips_Verification;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using static SKAIChips_Verification.RegContForm;

namespace SKAI_MCUP
{
    public class Barcelona : ChipControl
    {
        #region Variable and declaration

        public enum FW_TARGET
        {
            NV_MEM = 0,
            RAM = 1,
        }

        public enum TEST_ITEMS
        {
            TEST,
            NUM_TEST_ITEMS,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private FW_TARGET FirmwareTarget { get; set; } = FW_TARGET.NV_MEM;

        public int SlaveAddress { get; private set; } = 0x52;
        public FTDI.FT_DEVICE_INFO_NODE Device { get; private set; }

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI DigitalMultimeter0 = null;
        JLcLib.Instrument.SCPI OscilloScope0 = null;
        JLcLib.Instrument.SCPI TempChamber = null;
        #endregion Variable and declaration

        public Barcelona(RegContForm form) : base(form)
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

            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private uint ReadRegister(uint Address)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[4];
            uint Data;

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            RcvData = iComn.ReadBytes(RcvData.Length);
            Data = (uint)((RcvData[3] << 24) | (RcvData[2] << 16) | (RcvData[1] << 8) | (RcvData[0]));

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

        #region SKAI Barcelona host control methods
        public void HaltMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0x56, 0x78 };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }

        public void ResetMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0xAB, 0xCD };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }

        byte[] ReadMemory(uint Address, int Length)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            byte[] RcvBytes = I2C.ReadBytes(Length);

            I2C.Config.SlaveAddress = sa;

            return RcvBytes;
        }

        void WriteMemory(uint Address, uint Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        void WriteMemory(uint Address, byte[] Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            if (Data.Length < 4)
                return;

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
#if false
            SendBytes.Add(Data[3]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[0]);
#else
            SendBytes.Add(Data[0]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[3]);
#endif
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }
        #endregion SKAI Barcelona host control methods

        public override void SetChipSpecificUI()
        {
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */
            Parent.ChipCtrlButtons[0].Text = "BGR";
            Parent.ChipCtrlButtons[0].Visible = true;
            Parent.ChipCtrlButtons[0].Click += Test_BGR_Measurement;

            Parent.ChipCtrlButtons[7].Text = "DN RAM";
            Parent.ChipCtrlButtons[7].Visible = true;
            Parent.ChipCtrlButtons[7].Click += DownloadFirmwareToRAM_Click;
            Parent.ChipCtrlTextboxes[9].Text = ""; // Firmware size textbox
            Parent.ChipCtrlTextboxes[9].Visible = true;

            Parent.ChipCtrlButtons[10].Text = "Dump FW";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += new EventHandler(delegate { DumpFW(); });
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

        private void Test_BGR_Measurement(object sender, EventArgs e)
        {
            bool bRetrunTemp = true;
            string sheet_name, sTempVal;
            int x_pos = 1, y_pos = 1;
            double bgr_v;
            double[] dVal = new double[4];

            RegisterItem DCDC_RSVD = Parent.RegMgr.GetRegisterItem("DCDC_RSVD");        // 0x50110004[31:24]

            Check_Instrument();

            PowerSupply0.Write("INST:NSEL 1");
            PowerSupply0.Write("VOLT 2.0");
            PowerSupply0.Write("CURR 0.03");
            PowerSupply0.Write("OUTP ON");

            //TempChamber.Write("01,POWER,ON");
            //System.Threading.Thread.Sleep(500);
            //sTempVal = TempChamber.WriteAndReadString("TEMP?");

            DCDC_RSVD.Read();
            DCDC_RSVD.Value = 99;
            DCDC_RSVD.Write();

            sheet_name = DateTime.Now.ToString("MMddHHmmss_") + "BGR";
            Parent.xlMgr.Sheet.Add(sheet_name);

            for (int temp = -40; temp <= 120; temp += 10)
            {
                y_pos++;
                x_pos = 1;

                sTempVal = TempChamber.WriteAndReadString("01,TEMP,S" + temp.ToString());
                System.Threading.Thread.Sleep(1000);
                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                System.Threading.Thread.Sleep(100);
                Log.WriteLine("Run Chamber : " + sTempVal);
                Parent.xlMgr.Cell.Write(x_pos, y_pos, temp.ToString());
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

                    if (dVal[0] + dTHVal == temp || dVal[0] - dTHVal == temp || dVal[0] == temp)
                    {
                        bRetrunTemp = false;
                        Log.WriteLine("Done SetTemp!");
                    }
                    else
                    {
                        bRetrunTemp = true;
                    }
                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                    System.Threading.Thread.Sleep(1000 * 5);
                }
                System.Threading.Thread.Sleep(1000 * 60);
                Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());

                for (double volt = 2.0; volt <= 5.0; volt += 0.1)
                {
                    x_pos++;
                    PowerSupply0.Write("VOLT " + volt.ToString());
                    System.Threading.Thread.Sleep(500);
                    DCDC_RSVD.Write();
                    DCDC_RSVD.Write();
                    System.Threading.Thread.Sleep(100);

                    if (y_pos == 2)
                    {
                        Parent.xlMgr.Cell.Write(x_pos, y_pos - 1, volt.ToString());
                    }

                    bgr_v = double.Parse(DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?"));
                    if (bgr_v < 1)
                    {
                        DCDC_RSVD.Write();
                        DCDC_RSVD.Write();
                    }
                    bgr_v = double.Parse(DigitalMultimeter0.WriteAndReadString("MEAS:VOLT:DC?"));
                    Parent.xlMgr.Cell.Write(x_pos, y_pos, bgr_v.ToString("F3"));
                }
                PowerSupply0.Write("VOLT 2.0");
            }
            TempChamber.Write("01,TEMP,S25");
        }

        private void DownloadFirmwareToRAM_Click(object sender, EventArgs e)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = FW_TARGET.RAM;
                DownloadFirmware(BARCELONA_DownloadFW);
            }
        }

        private void DownloadFirmwareToNV_Click(object sender, EventArgs e)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = FW_TARGET.NV_MEM;
                DownloadFirmware(BARCELONA_DownloadFW);
            }
        }

        private void EraseFW()
        {
            EraseFirmware(BARCELONA_EraseFW);
        }

        private void DumpFW()
        {
            ReadFirmwareSize = GetInt_ChipCtrlTextboxes(9);
            DumpFirmware(BARCELONA_DumpFW);
        }

        #region Firmware control methods
        private void BARCELONA_DownloadFW()
        {
            int PageSize = 256;
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
            if (FirmwareTarget == FW_TARGET.RAM)
            {
                PageSize = 4;
            }

            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (FirmwareData.Length + PageSize - 1) / PageSize;
            })));

            /* Mass erase */
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                BARCELONA_EraseFW();
                if (Status == false)
                    goto EXIT;
            }
            else
                HaltMCU();

            /* Program */
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

                // Write firmware data
                WriteMemory(FlashAddress, SendBytes);

                // Compare memory
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    for (uint i = 0; i < PageSize; i += 4)
                    {
                        uint RcvBytes = ReadRegister(FlashAddress + i);
                        for (int j = 0; j < 4; j++)
                        {
                            if (SendBytes[i * 4 + j] != ((RcvBytes >> (8 * j)) & 0xff))
                            {
                                Log.WriteLine("Faile to match: (0x" + (FlashAddress + i * 4 + j).ToString("X4") + ":" + SendBytes[i * 4 + j].ToString("X2") + " " + ((RcvBytes >> (8 * j)) & 0xff).ToString("X2") + ")",
                                    System.Drawing.Color.Coral, Log.RichTextBox.BackColor);
                                goto EXIT;
                            }
                        }
                    }
                }

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            if (FirmwareTarget == FW_TARGET.RAM)
                ResetMCU();

            Status = true;
            br.Close();
            fs.Close();

        EXIT:
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }

        public void BARCELONA_EraseFW()
        {
            Status = true;
        }

        public void BARCELONA_DumpFW()
        {
            const int PageSize = 4;
            uint RcvBytes;
            List<byte> FirmwareData = new List<byte>();
            string file_name;

            Status = false;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (ReadFirmwareSize + PageSize - 1) / PageSize + 1; // Mass erase
            })));

            HaltMCU();

            for (uint Addr = 0; Addr < ReadFirmwareSize; Addr += PageSize)
            {
                RcvBytes = ReadRegister(Addr);
                FirmwareData.Add((byte)(RcvBytes & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 8) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 16) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 24) & 0xff));

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            Status = true;
            ReadFirmwareData = FirmwareData.ToArray();
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                file_name = "ReadFirmwareBinary_NVM.bin";
            }
            else
            {
                file_name = "ReadFirmwareBinary_RAM.bin";
            }
            System.IO.FileStream fs = new System.IO.FileStream(file_name, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(ReadFirmwareData);
            bw.Close();
            fs.Close();

            // 3. Reset system
            ResetMCU();

            Status = true;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }
        #endregion Firmware control methods

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

    public class Toscana : ChipControl
    {
        #region Variable and declaration

        public enum FW_TARGET
        {
            NV_MEM = 0,
            RAM = 1,
            SPI,
        }

        public enum FLASH_CMD
        {
            WRSR = 0x01,    // write stat. reg
            PP = 0x02,      // page program
            RDCMD = 0x03,   // read data
            WRDI = 0x04,    // write disable
            RDSR = 0x05,    // read status reg
            WREN = 0x06,    // write enable
            F_RD = 0x0B,    // fast read datwjda
            SE = 0x20,      // 4KB sector erase
            BE32 = 0x52,    // 32KB block erase
            RSTEN = 0x66,   // reset enable
            REMS = 0x90,    // read manufacture &
            RST = 0x99,     // reset
            RDID = 0x9F,    // read identificatio
            RES = 0xAB,     // read signature
            ENSO = 0xB1,    // enter secured OTP
            DP = 0xB9,      // deep power down
            EXSO = 0xC1,    // exit secured OTP
            CE = 0xC7,      // chip(bulk) erase
            BE64 = 0xD8,	// 64KB block erase
        }

        public enum TEST_ITEMS
        {
            ENTER_TEST_PT, // 0x01 (FW command)
            EXIT_TEST_PT, // 0x02
            SET_TEST_VRECT, // 0x03

            LDO_TURNON, // 0x08 (FW command)
            LDO_TURNOFF, // 0x09
            LDO_SET, // 0x0A

            FIRM_ON_CLEAR, // 0x0F

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

            NUM_TEST_ITEMS,
        }

        public enum AUTO_TEST_ITEMS
        {
            TEST_IOUT,
            TEST_ISEN,
            TEST_VOUT,
            TEST_VOUT_WPC,
            TEST_VRECT_WPC,
            TEST_ISEN_REG,
            TEST_ACTIVE_LOAD,
            TEST_LOAD_SWEEP,
            TEST_ADC_CODE,

            NUM_TEST_ITEMS,
        }

        public enum FW_DN_ITEMS
        {
            FLASH_ERASE,
            FLASH_WRITE,
            FLASH_READ,
            RAM_WRITE,
            RAM_READ,

            FIRM_ON_CLEAR, // 0x0F

            NUM_TEST_ITEMS,
        }

        public enum CAL_ITEMS
        {
            CAL_RUN,
            CAL_WRITE,
            CAL_READ,
            CAL_APPLY,

            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            TEST,
            AUTO,
            CAL,
            FW,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Custom.SPI SPI { get; set; }
        private FW_TARGET FirmwareTarget { get; set; } = FW_TARGET.NV_MEM;

        public int SlaveAddress { get; private set; } = 0x52;
        private bool IsRunCal = false;
        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.TEST;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI PowerSupply1 = null;
        JLcLib.Instrument.SCPI Oscilloscope = null;
        JLcLib.Instrument.SCPI ElectronicLoad = null;
        #endregion Variable and declaration

        public Toscana(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            SPI = form.SPI;
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
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private void WriteRegister(uint Address, byte[] Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add(Data[0]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[3]);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private uint ReadRegister(uint Address)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[4];
            uint Data;

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            RcvData = iComn.ReadBytes(RcvData.Length);
            Data = (uint)((RcvData[3] << 24) | (RcvData[2] << 16) | (RcvData[1] << 8) | (RcvData[0]));

            I2C.Config.SlaveAddress = sa;
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

        #region SKAI Toscana host control methods
        public void HaltMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0x56, 0x78 };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }

        public void ResetMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0xAB, 0xCD };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }
        #endregion SKAI Toscana host control methods

        public override void SetChipSpecificUI()
        {
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */

            Parent.ChipCtrlTextboxes[0].Text = "140"; // Freq(kHz)
            Parent.ChipCtrlTextboxes[0].Visible = true;

            Parent.ChipCtrlTextboxes[1].Text = "50"; // Duty(%)
            Parent.ChipCtrlTextboxes[1].Visible = true;

            Parent.ChipCtrlButtons[2].Text = "SET";
            Parent.ChipCtrlButtons[2].Visible = true;
            Parent.ChipCtrlButtons[2].Click += Run_PWM_SET;

            Parent.ChipCtrlButtons[4].Text = "GPIO";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Run_PWM_GPIO_SET;

            Parent.ChipCtrlButtons[5].Text = "Initial";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += Run_PWM_Initial;

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

        private void Test_Iout_ADC_CODE()
        {
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>");      // 0x50100030[13:10]
            string sheet_name;
            int xPos = 0, yPos = 0; // cell offset
            double Vrect, Vout;

            Log.WriteLine("Start TEST_ADC_CODE");

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
                    case JLcLib.Instrument.InstrumentTypes.ElectronicLoad:
                        if (ElectronicLoad == null)
                            ElectronicLoad = new JLcLib.Instrument.SCPI(Ins.Type);
                        if (ElectronicLoad.IsOpen == false)
                            ElectronicLoad.Open();
                        break;
                }
            }

            if (Oscilloscope == null || Oscilloscope.IsOpen == false || PowerSupply0 == null || PowerSupply0.IsOpen == false
                || PowerSupply1 == null || PowerSupply1.IsOpen == false || ElectronicLoad == null || ElectronicLoad.IsOpen == false)
            {
                MessageBox.Show("Check Instrument");
                return;
            }

            if (Parent.xlMgr == null)
                return;

            MessageBox.Show("Power Supply0\n- CH1 : AP1P8\n- CH2 : VRECT\n- CH3 : GPIO\n\nPower Supply1\n- CH1 : AP1P8\n\nOscilloscope\n- CH1 : VRECT\n- CH2 : VOUT\n- CH3 : N/C\n- CH4 : N/C");
            ElectronicLoad.Write("SYST:REM");
            ElectronicLoad.Write("FUNC CURR");
            ElectronicLoad.Write("INP 0");      // output off
            ElectronicLoad.Write("CURR 0");     // load current
            System.Threading.Thread.Sleep(10);
            ElectronicLoad.Write("INP 1");      // output on

            PowerSupply0.Write("OUTP OFF,(@1:3)");
            PowerSupply0.Write("INST:NSEL 1");  // sel port 1
            PowerSupply0.Write("VOLT 0.0");     // output voltage
            PowerSupply0.Write("CURR 0.3");     // current limit
            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT 5");
            PowerSupply0.Write("CURR 1");
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT 0");
            PowerSupply0.Write("CURR 0.5");
            PowerSupply0.Write("OUTP ON,(@1:3)");
            PowerSupply1.Write("OUTP OFF");
            PowerSupply1.Write("VOLT 0.0");     // output voltage
            PowerSupply1.Write("CURR 0.3");     // current limit
            PowerSupply1.Write("OUTP ON");

            Oscilloscope.Write(":TIM:SCAL 1E-4"); // 100usec/div, Scale / 1div
            Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
            Oscilloscope.Write(":TIM:DEL 0"); // delay
            // Cannel.1 use for VRECT
            Oscilloscope.Write(":CHAN1:DISP 1");
            Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN1:OFFS 0"); // offset 0V
            // Cannel.2 use for VOUT
            Oscilloscope.Write(":CHAN2:DISP 1");
            Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // 200mV/div
            Oscilloscope.Write(":CHAN2:OFFS 5"); // offset 5V

            JLcLib.Delay.Sleep(200);

            sheet_name = DateTime.Now.ToString("MMddHHmmss_") + "ADC_CODE";
            Parent.xlMgr.Sheet.Add(sheet_name);

            PowerSupply0.Write("INST:NSEL 2"); // Vrect
            for (double ld = 0; ld < 1.25; ld += 0.05)
            {
                yPos++;
                ElectronicLoad.Write("CURR " + ld.ToString());
                System.Threading.Thread.Sleep(10);
                Parent.xlMgr.Cell.Write(0, yPos, ((ld * 1000).ToString("F2")));
                for (double Vr = 3.9; Vr <= 5.5; Vr += 0.1) // Vrect Power Supply (V)
                {
                    xPos++;
                    Oscilloscope.Write(":CHAN1:OFFS " + Vr.ToString("F2"));
                    PowerSupply0.Write("VOLT " + Vr.ToString("F2"));
                    JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

                    Parent.xlMgr.Cell.Write(xPos, 0, (Vr.ToString("F2")));
                    Vrect = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
                    Parent.xlMgr.Cell.Write(xPos, yPos, (Vrect * 1000).ToString("F3"));
                    Vout = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
                    Parent.xlMgr.Cell.Write(xPos, yPos + 7, (Vout * 1000).ToString("F3"));
                }
                xPos = 0;
            }



            Log.WriteLine("End TEST_ADC_CODE");
        }

        #region Calibration methods
        private void RunCalibration()
        {
            Log.WriteLine("Run Calibratoin!!");
            if (!CAL_CheckInstrument()) return;
            CAL_InitInstrument();
            //JLcLib.Delay.Sleep(2000);
            PowerSupply0.Write("OUTP OFF,(@1:3)");
            MessageBox.Show("Power Supply\n- CH1 : AP1P8\n- CH2 : GPIO\n- CH3 : VRECT\n\nOscilloscope\n- CH1 : LDO5P0\n- CH2 : LDO1P5\n- CH3 : VBAT_SENS\n- CH4 : TEST_DIG");

            if (!IsRunCal) return;
            PowerSupply0.Write("OUTP ON,(@1:3)");
            System.Threading.Thread.Sleep(500);
            CAL_Reg_Init();
            if (!IsRunCal) return;
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
            MessageBox.Show("Change Oscilloscope\n- CH1 : LDO5P0 -> LDO1P8\n- CH2 : LDO1P5 -> VOUT");
            if (!IsRunCal) return;
            CAL_RunLDO1P8();
            if (!IsRunCal) return;
            CAL_RunOSC();
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
            CAL_RunIOUT();
            Log.WriteLine("Done Calibratoin!!");
        }

        private void StopCalibration()
        {
            IsRunCal = false;
        }

        public int Get_ADC_Value(uint ch, int rep)
        {
            uint result;

            RegisterItem ADC_CHSEL = Parent.RegMgr.GetRegisterItem("ADC_CHSEL[3:0]");   // 0x5011800C[27:24,19:16,11:8,3:0]
            RegisterItem ADC_RUN = Parent.RegMgr.GetRegisterItem("ADC_RUN");            // 0x50118014[24,16,8,0]
            RegisterItem ADC_EOC = Parent.RegMgr.GetRegisterItem("ADC_EOC");            // 0x50118018[29,21,13,5]
            RegisterItem ADC_DATA = Parent.RegMgr.GetRegisterItem("ADC_DATA[11:0]");    // 0x5011801C[11:0]

            ADC_CHSEL.Value = (ch << 24) | (ch << 16) | (ch << 8) | (ch << 0);
            ADC_CHSEL.Write();

            ADC_EOC.Value = 0;
            ADC_EOC.Write();

            result = 0;

            rep += 1;
            for (int i = 0; i < rep; i++)
            {
                ADC_RUN.Value = 0x01010101;
                ADC_RUN.Write();
                while (true)
                {
                    ADC_EOC.Read();
                    if (ADC_EOC.Value == 0x21)
                    {
                        break;
                    }
                }
                ADC_EOC.Value = 0;
                ADC_EOC.Write();
                ADC_DATA.Read();
                result += ADC_DATA.Value;
            }

            return (int)(result / rep);
        }

        private bool CAL_CheckInstrument()
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
            if ((PowerSupply0 == null) || (PowerSupply0.IsOpen == false) ||
                (Oscilloscope == null) || (Oscilloscope.IsOpen == false) || (ElectronicLoad == null) || (ElectronicLoad.IsOpen == false))
            {
                MessageBox.Show("Check Instruments");
                return false;
            }
            else
            {
                return true;
            }
        }

        private void CAL_InitInstrument()
        {
            if (PowerSupply0 != null && PowerSupply0.IsOpen)
            {
                PowerSupply0.Write("OUTP OFF,(@1:3)");
                PowerSupply0.Write("INST:NSEL 1");// AP1P8
                PowerSupply0.Write("VOLT 1.8");
                PowerSupply0.Write("CURR 0.3");
                PowerSupply0.Write("INST:NSEL 2"); // GPIO
                PowerSupply0.Write("VOLT 0.0");
                PowerSupply0.Write("CURR 0.3");
                PowerSupply0.Write("INST:NSEL 3"); // VRECT
                PowerSupply0.Write("VOLT 6.0");
                PowerSupply0.Write("CURR 2.0");
            }
            if (Oscilloscope != null && Oscilloscope.IsOpen)
            {
                //Oscilloscope.Write(":TIM:SCAL 2E-4"); // 200usec/div, Scale / 1div
                Oscilloscope.Write(":TIM:SCAL 1E-7"); // 100nsec/div, Scale / 1div
                Oscilloscope.Write(":TIM:REF CENT"); // LEFT, CENT, RIGHt
                Oscilloscope.Write(":TIM:DEL 0"); // delay
                // Cannel.1 use for LDO 5p0 -> 1p8
                Oscilloscope.Write(":CHAN1:DISP 1");
                Oscilloscope.Write(":CHAN1:SCAL 2E-1"); // div
                Oscilloscope.Write(":CHAN1:OFFS 5");    // offset
                // Cannel.2 use for LDO 1p5 -> vout
                Oscilloscope.Write(":CHAN2:DISP 1");
                Oscilloscope.Write(":CHAN2:SCAL 2E-1"); // div
                Oscilloscope.Write(":CHAN2:OFFS 1.5");  // offset
                // Cannel.3 use for VBAT_SENS
                Oscilloscope.Write(":CHAN3:DISP 1");
                Oscilloscope.Write(":CHAN3:SCAL 2E-1"); // div
                Oscilloscope.Write(":CHAN3:OFFS 1.2");  // offset
                // Cannel.4 use for TEST_DIG
                Oscilloscope.Write(":CHAN4:DISP 1");
                Oscilloscope.Write(":CHAN4:SCAL 1");    // div
                Oscilloscope.Write(":CHAN4:OFFS 3.3");  // offset

                Oscilloscope.Write(":TRIG:MODE EDGE"); // trigger mode edge
                Oscilloscope.Write(":TRIG:EDGE:SOUR CHAN4"); // triger source 4
                Oscilloscope.Write(":TRIG:EDGE:LEV 9E-1"); // trigger level 900mV
            }
            if (ElectronicLoad != null && ElectronicLoad.IsOpen)
            {
                ElectronicLoad.Write("FUNC CC"); // CC mode (FUNC CR/CC/CV/CP/CCCV/CRCV
                ElectronicLoad.Write("INP 0");
                ElectronicLoad.Write("CURR 0");
                System.Threading.Thread.Sleep(10);
                ElectronicLoad.Write("INP 1");
            }
        }

        private void CAL_Reg_Init()
        {
            RegisterItem DIV_MCU = Parent.RegMgr.GetRegisterItem("DIV_MCU<3:0>");               // 0x5006000C[3:0]
            RegisterItem PAD_EN = Parent.RegMgr.GetRegisterItem("PAD_EN<13:0>");                // 0x50100034[18:31]
            RegisterItem EN_CLK_8M = Parent.RegMgr.GetRegisterItem("EN_CLK_8M");                // 0x5010003C[19]
            RegisterItem DIV_CLK_8M = Parent.RegMgr.GetRegisterItem("DIV_CLK_8M<3:0>");         // 0x5010003C[23:20]
            RegisterItem EN_ML_CP_CLK = Parent.RegMgr.GetRegisterItem("EN_ML_CP_CLK");          // 0x50100044[28]
            RegisterItem DIV_ML_CP_CLK = Parent.RegMgr.GetRegisterItem("DIV_ML_CP_CLK<3:0>");   // 0x50100044[27:24]
            RegisterItem ADC_TRIM_VREF = Parent.RegMgr.GetRegisterItem("ADC_TRIM_VREF<3:0>");   // 0x50100020[6:3]
            RegisterItem ADC_ENABLE = Parent.RegMgr.GetRegisterItem("ADC_ENABLE");              // 0x50118000[24,16,8,0]
            RegisterItem ADC_RESET = Parent.RegMgr.GetRegisterItem("ADC_RESET");                // 0x50118004[24,16,8,0]
            RegisterItem ISEN_RVAR = Parent.RegMgr.GetRegisterItem("ML_ISEN_RVAR<1:0>");        // 0x50100018[6:5]
            RegisterItem FC_RTRIM = Parent.RegMgr.GetRegisterItem("ML_FC_RTRIM<3:0>");          // 0x5010001C[3:0]
            RegisterItem VREF_TRIM = Parent.RegMgr.GetRegisterItem("ML_VREF_TRIM<3:0>");        // 0x50100018[27:24]


            DIV_MCU.Read();
            DIV_MCU.Value = 1;
            DIV_MCU.Write();
            System.Threading.Thread.Sleep(10);

            PAD_EN.Read();
            PAD_EN.Value = 0;
            PAD_EN.Write();

            EN_CLK_8M.Read();
            EN_CLK_8M.Value = 1;
            DIV_CLK_8M.Value = 15;
            EN_CLK_8M.Write();

            EN_ML_CP_CLK.Read();
            EN_ML_CP_CLK.Value = 1;
            DIV_ML_CP_CLK.Value = 15;
            EN_ML_CP_CLK.Write();

            ADC_TRIM_VREF.Read();
            ADC_TRIM_VREF.Value = 8;
            ADC_TRIM_VREF.Write();

            ADC_ENABLE.Read();
            ADC_ENABLE.Value = 0x01010101;
            ADC_ENABLE.Write();

            ADC_RESET.Read();
            ADC_RESET.Value = 0x01010101;
            ADC_RESET.Write();

            ISEN_RVAR.Read();
            ISEN_RVAR.Value = 2;
            ISEN_RVAR.Write();

            FC_RTRIM.Read();
            FC_RTRIM.Value = 15;
            FC_RTRIM.Write();

            VREF_TRIM.Read();
            VREF_TRIM.Value = 15;
            VREF_TRIM.Write();
        }

        private void CAL_RunBGR()
        {
            double Result = 0.0;
            double Target = 1.22, Margin = 0.005;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "BGR";
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>");      // 0x50100030[13:10]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MU0_EN");    // 0x50100030[8]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");           // 0x50100030[7]
            RegisterItem TRIM_BGR = Parent.RegMgr.GetRegisterItem("PU_TRIM_BGR<4:0>");      // 0x5010000C[4:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
            ANA_SEL.Read();     // Read 0x50100030
            ANA_SEL.Value = 0;  // BGR
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_SEL.Write();    // Write 0x50100030         
            TRIM_BGR.Read();
            for (uint BitPos = 11; BitPos < 16; BitPos++)
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
                CalibrationData[0x00] = (byte)TRIM_BGR.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_BGR.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x00] = (byte)0xFF;
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
            RegisterItem TRIM_LDO5P0 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO5P0<3:0>"); // 0x5010000C[19:16]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            TRIM_LDO5P0.Read();
            TRIM_LDO5P0.Value = 0;
            TRIM_LDO5P0.Write();
            for (uint BitPos = 7; BitPos < 13; BitPos++)
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
                CalibrationData[0x01] = (byte)TRIM_LDO5P0.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO5P0.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x01] = (byte)0xFF;
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
            RegisterItem TRIM_LDO1P5 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO1P5<3:0>"); // 0x50100014[14:11]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN2:OFFS 1.5"); // offset 1.5V

            TRIM_LDO1P5.Read();
            for (uint BitPos = 4; BitPos < 9; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO1P5.Value = BitPos;
                TRIM_LDO1P5.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
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
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN2"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x02] = (byte)TRIM_LDO1P5.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO1P5.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x02] = (byte)0xFF;
            }
        }

        private void CAL_RunADC_LDO3P3()
        {
            double Result = 0.0;
            double Target = 3.3, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "ADC_LDO3P3";
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>");      // 0x50100030[13:10]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MU0_EN");    // 0x50100030[8]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");           // 0x50100030[7]
            RegisterItem TRIM_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_ADC_LDO3P3<3:0>"); // 0x50100010[3:0]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 3.3"); // offset 3.3V
            ANA_SEL.Read();
            ANA_SEL.Value = 6; // ADC LDO 3.3            
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_SEL.Write();    // Write 0x50100030

            TRIM_LDO3P3.Read();
            for (uint BitPos = 10; BitPos < 15; BitPos++)
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
                CalibrationData[0x03] = (byte)TRIM_LDO3P3.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x03] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_SEL.Write(); // Write 0x50100030
        }

        private void CAL_RunFSK_LDO3P3()
        {
            double Result = 0.0;
            double Target = 3.3, Margin = 0.02;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.01; // pass/fail limit
            string Name = "FSK_LDO3P3";
            RegisterItem ANA_SEL = Parent.RegMgr.GetRegisterItem("TEST_ANA_SEL<3:0>");      // 0x50100030[13:10]
            RegisterItem ANA_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_MU0_EN");    // 0x50100030[8]
            RegisterItem ANA_PEN = Parent.RegMgr.GetRegisterItem("TEST_ANA_PEN");           // 0x50100030[7]
            RegisterItem TRIM_LDO3P3 = Parent.RegMgr.GetRegisterItem("PU_TRIM_FSK_LDO3P3<3:0>"); // 0x50100010[15:12]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN3:OFFS 3.3"); // offset 3.3V
            ANA_SEL.Read();
            ANA_SEL.Value = 5; // FSK LDO 3.3
            ANA_MUX_PEN.Value = 1;
            ANA_PEN.Value = 1;
            ANA_SEL.Write(); // Write 0x50100030

            TRIM_LDO3P3.Read();
            for (uint BitPos = 8; BitPos < 13; BitPos++)
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
                CalibrationData[0x04] = (byte)TRIM_LDO3P3.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO3P3.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x04] = (byte)0xFF;
            }
            ANA_SEL.Value = 0;
            ANA_MUX_PEN.Value = 0;
            ANA_PEN.Value = 0;
            ANA_PEN.Write(); // Write 0x50100030
        }

        private void CAL_RunLDO1P8()
        {
            double Result = 0.0;
            double Target = 1.8, Margin = 0.04;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.03; // pass/fail limit
            string Name = "LDO1P8";
            RegisterItem TRIM_LDO1P8 = Parent.RegMgr.GetRegisterItem("PU_TRIM_LDO18<3:0>"); // 0x50100010[31:28]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            Oscilloscope.Write(":CHAN1:OFFS 1.8"); // offset 1.8V
            TRIM_LDO1P8.Read();
            for (uint BitPos = 5; BitPos < 10; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_LDO1P8.Value = BitPos;
                TRIM_LDO1P8.Write();
                JLcLib.Delay.Sleep(50);
                Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
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
            Result = double.Parse(Oscilloscope.WriteAndReadString("MEAS:VAV? DISP,CHAN1"));
            if (Result >= (Target - Limit) && Result <= (Target + Limit))
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x05] = (byte)TRIM_LDO1P8.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_LDO1P8.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x05] = (byte)0xFF;
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
            RegisterItem DIG_SEL = Parent.RegMgr.GetRegisterItem("TEST_DIG_SEL<3:0>");      // 0x50100030[19:16]
            RegisterItem DIG_MUX_PEN = Parent.RegMgr.GetRegisterItem("TEST_DIG_MU0_EN");    // 0x50100030[9]
            RegisterItem TRIM_OSC = Parent.RegMgr.GetRegisterItem("PU_TRIM_FOSC_16M<7:0>"); // 0x5010000C[15:8]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
            DIG_SEL.Read();
            DIG_SEL.Value = 0; // OSC
            DIG_MUX_PEN.Value = 1;
            DIG_SEL.Write(); // Write 0x50100030

            TRIM_OSC.Read();
            for (uint BitPos = 0; BitPos < 11; BitPos++)
            {
                if (!IsRunCal)
                    return;
                if (BitPos > 4)
                {
                    TRIM_OSC.Value = BitPos - 5;
                }
                else
                {
                    TRIM_OSC.Value = 251 + BitPos;
                }
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
                CalibrationData[0x06] = (byte)TRIM_OSC.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_OSC.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x06] = 0xFF;
            }
            DIG_MUX_PEN.Value = 0;
            DIG_SEL.Write(); // Write 0x50100030
        }

        private void CAL_RunISENVREF()
        {
            double Vrect = 4.9;
            double Result = 0.0;
            double Target = 300, Margin = 10;
            double min = 4095.0;
            uint reg_value = 0;
            double Limit = Target * 0.05; // pass/fail limit
            string Name = "ML_ISEN_FB";
            RegisterItem TRIM_ISEN = Parent.RegMgr.GetRegisterItem("ML_ISEN_VREF_TRIM<3:0>");   // 0x5010001C[7:4]
            RegisterItem TRIM_ISEN_FB = Parent.RegMgr.GetRegisterItem("ML_ISEN_FB_TRIM<3:0>");  // 0x5010001C[31:28]
            RegisterItem TRIM_VILIM = Parent.RegMgr.GetRegisterItem("ML_VILIM_TRIM<3:0>");      // 0x50100018[31:28]

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);
#if POWER_SUPPLY_E36313A
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));
#else
            PowerSupply1.Write("INST:NSEL 2");
            PowerSupply1.Write("VOLT " + Vrect.ToString("F1"));
#endif

            ElectronicLoad.Write("CURR 0.1");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            TRIM_VILIM.Read();
            TRIM_VILIM.Value = 0;
            TRIM_VILIM.Write();

            TRIM_ISEN.Read();
            TRIM_ISEN.Value = 15;
            TRIM_ISEN.Write();

            TRIM_ISEN_FB.Read();
            for (uint BitPos = 0; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_ISEN_FB.Value = BitPos;
                TRIM_ISEN_FB.Write();
                JLcLib.Delay.Sleep(50);
                Result = Get_ADC_Value(2, 1);
                if (Math.Abs(Target - Result) < min)
                {
                    min = Math.Abs(Target - Result);
                    reg_value = TRIM_ISEN_FB.Value;
                }
                Log.WriteLine(Name + ":" + BitPos.ToString("D2") + ":" + TRIM_ISEN_FB.Value.ToString("D3") + ":" + Result.ToString("F3"));
            }
            TRIM_ISEN_FB.Value = reg_value;
            TRIM_ISEN_FB.Write();
            JLcLib.Delay.Sleep(50);
            Result = Get_ADC_Value(2, 1);

            if (Result >= 0 && Result <= 800)
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_ISEN_FB.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x07] = (byte)TRIM_ISEN_FB.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_ISEN_FB.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x07] = (byte)0xFF;
            }

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
#if POWER_SUPPLY_E36313A
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));
#else
            PowerSupply1.Write("INST:NSEL 2");
            PowerSupply1.Write("VOLT " + Vrect.ToString("F1"));
#endif

            ElectronicLoad.Write("CURR 0.6");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize

            TRIM_VILIM.Read();
            for (uint BitPos = 0; BitPos < 16; BitPos++)
            {
                if (!IsRunCal)
                    return;
                TRIM_VILIM.Value = BitPos;
                TRIM_VILIM.Write();
                JLcLib.Delay.Sleep(50);
                Result = Get_ADC_Value(2, 1);
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
            Result = Get_ADC_Value(2, 1);

            if (Result >= 700 && Result <= 1300)
            {
                Log.WriteLine(Name + ":PASS:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.ForestGreen, Log.RichTextBox.BackColor);
                CalibrationData[0x08] = (byte)TRIM_VILIM.Value;
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + TRIM_VILIM.Value.ToString("D3") + ":" + Result.ToString("F3"), Color.Coral, Log.RichTextBox.BackColor);
                CalibrationData[0x08] = (byte)0xFF;
            }

            ElectronicLoad.Write("CURR 0");
        }

        private void CAL_RunVRECT()
        {
            double LowVrect = 5, HighVrect = 8;
            int LowVrectCode, HighVrectCode;
            short Slope, Const;
            string Name = "VRECT";

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            // 5V low VRECT calibration
#if POWER_SUPPLY_E36313A
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT " + LowVrect.ToString("F1"));
#else
            PowerSupply1.Write("INST:NSEL 2");
            PowerSupply1.Write("VOLT " + LowVrect.ToString("F1"));
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            LowVrectCode = Get_ADC_Value(0, 1);
            Log.WriteLine(Name + ":LOW:" + LowVrect.ToString("F1") + ":" + LowVrectCode.ToString("F3"));
            if (!IsRunCal)
                return;
            // 8V high VRECT calibration
#if POWER_SUPPLY_E36313A
            PowerSupply0.Write("VOLT " + HighVrect.ToString("F1"));
#else
            PowerSupply1.Write("VOLT " + HighVrect.ToString("F1"));
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            HighVrectCode = Get_ADC_Value(0, 1);
            Log.WriteLine(Name + ":HIGH:" + HighVrect.ToString("F1") + ":" + HighVrectCode.ToString("F3"));

            Slope = (short)((HighVrect - LowVrect) * 1000 * 1024 / (HighVrectCode - LowVrectCode));
            Const = (short)(LowVrect * 1000 - (Slope * LowVrectCode) / 1024);

            if ((Slope >= 5250 && Slope <= 6750) && (Const >= -512 && Const <= 512))
            {
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            }
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 5000;
                Const = 0;
            }
            CalibrationData[0x09] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x0A] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x0B] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x0C] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunVOUT()
        {
            double Result, Vrect = 8.0;
            double[] TargetVoutValues = new double[2] { 5.0, 6.0 };
            double[] VoutValues = new double[2];
            double[] VoutSetValues = new double[2];
            double[] Limits = new double[2] { TargetVoutValues[0] * 0.03, TargetVoutValues[1] * 0.03 };
            int[] VoutCodes = new int[2];
            double Margin = 0.015;
            double min = 4095.0;
            uint reg_value = 0;
            short Slope, Const;
            string Name = "VOUT";
            RegisterItem OCL_EN = Parent.RegMgr.GetRegisterItem("ML_OCL_EN");               // 0x50100014[23]
            //RegisterItem ML_PEN = Parent.RegMgr.GetRegisterItem("ML_PEN");                  // 0x50100014[20]
            RegisterItem TRIM_VOUT_SET = Parent.RegMgr.GetRegisterItem("ML_VOUT_SET<7:0>"); // 0x50100018[23:16]

            // Sets 8V VRECT
#if POWER_SUPPLY_E36313A
            PowerSupply0.Write("INST:NSEL 3");
            PowerSupply0.Write("VOLT " + Vrect.ToString("F1"));
#else
            PowerSupply1.Write("INST:NSEL 2");
            PowerSupply1.Write("VOLT " + Vrect.ToString("F1"));
#endif
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            OCL_EN.Read();
            OCL_EN.Value = 0;
            OCL_EN.Write();

            //ML_PEN.Read();
            //ML_PEN.Value = 1;
            //ML_PEN.Write();
            //ElectronicLoad.Write("CURR 0.1");
            //ElectronicLoad.Write("INP 1");

            TRIM_VOUT_SET.Read();
            for (uint i = 0; i < 2; i++)
            {
                min = 4095.0;

                Oscilloscope.Write(":CHAN2:SCAL 2"); // 2V/div
                Oscilloscope.Write(":CHAN2:OFFS 8"); // offset 8V

                for (uint BitPos = 17; BitPos < 24; BitPos++)
                {
                    if (!IsRunCal)
                        return;
                    TRIM_VOUT_SET.Value = BitPos + (i * 10);
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
                VoutCodes[i] = Get_ADC_Value(1, 1);

                if (VoutValues[i] >= TargetVoutValues[i] - Limits[i] && VoutValues[i] <= TargetVoutValues[i] + Limits[i])
                    Log.WriteLine(Name + "_LVL:PASS:" + reg_value.ToString() + ":" + VoutValues[i].ToString() + ":" + VoutCodes[i].ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
                else
                    Log.WriteLine(Name + "_LVL:FAIL:" + reg_value.ToString() + ":" + VoutValues[i].ToString() + ":" + VoutCodes[i].ToString(), Color.Coral, Log.RichTextBox.BackColor);
            }

            TRIM_VOUT_SET.Value = (uint)VoutSetValues[0];
            TRIM_VOUT_SET.Write();
            //RunTest(TEST_ITEMS.LDO_TURNOFF, 0);
            //CalibrationData[0xB0] = (byte)TRIM_VOUT_SET.Value;
            //ElectronicLoad.Write("CURR 0");

            // Vout voltage calibration
            Slope = (short)((VoutValues[1] - VoutValues[0]) * 1000 * 1024 / (VoutCodes[1] - VoutCodes[0]));
            Const = (short)(VoutValues[0] * 1000 - (Slope * VoutCodes[0]) / 1024);
            if ((Slope >= 4500 && Slope <= 5500) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + ":PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + ":FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 5000;
                Const = 0;
            }
            CalibrationData[0x0D] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x0E] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x0F] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x10] = (byte)((Const >> 8) & 0xFF);

            // Vout setting register calibration
            Slope = (short)((VoutSetValues[1] - VoutSetValues[0]) * 8192 / ((TargetVoutValues[1] - TargetVoutValues[0]) * 1000));
            Const = (short)(VoutSetValues[0] - (Slope * TargetVoutValues[0] * 1000) / 8192);
            if ((Slope >= 50 && Slope <= 150) && (Const >= -512 && Const <= 512))
                Log.WriteLine(Name + "_SET:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_SET:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 341;
                Const = -82;
            }
            CalibrationData[0x11] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x12] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x13] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x14] = (byte)((Const >> 8) & 0xFF);
        }

        private void CAL_RunADC()
        {
            double LowVolt = 0.3, HighVolt = 1.5;
            int LowVoltCode, HighVoltCode;
            short Slope, Const;
            string Name = "ADC";

            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            PowerSupply0.Write("INST:NSEL 2");
            PowerSupply0.Write("VOLT " + LowVolt.ToString("F1"));
            PowerSupply0.Write("OUTP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            LowVoltCode = Get_ADC_Value(5, 1);
            Log.WriteLine(Name + ":LOW:" + LowVolt.ToString("F1") + ":" + LowVoltCode.ToString("F3"));
            if (!IsRunCal)
                return;

            PowerSupply0.Write("VOLT " + HighVolt.ToString("F1"));
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            HighVoltCode = Get_ADC_Value(5, 1);
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
            CalibrationData[0x15] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x16] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x17] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x18] = (byte)((Const >> 8) & 0xFF);

            PowerSupply0.Write("VOLT 0.0");
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

            PowerSupply0.Write("INST:NSEL 3");

            ElectronicLoad.Write("CURR 0");
            ElectronicLoad.Write("INP 1");
            JLcLib.Delay.Sleep(200); // wait for power supply to stabilize
            Log.WriteLine(Name + " Caliration!!", Color.Blue, Log.RichTextBox.BackColor);

            for (int i = 0; i < IoutValues.Length; i++) // Zero(0)/Low(1)/High(2) current
            {
                ElectronicLoad.Write("CURR " + IoutValues[i].ToString("F2"));
                for (int j = 0; j < VrectValues.Length; j++) // Low(0)/High(1) VRECT
                {
#if POWER_SUPPLY_E36313A
                    PowerSupply0.Write("VOLT " + VrectValues[j].ToString("F1")); // Set VRECT voltage
#else
                    PowerSupply1.Write("VOLT " + VrectValues[j].ToString("F1")); // Set VRECT voltage
#endif
                    JLcLib.Delay.Sleep(500);
                    ISenCodes[i, j] = Get_ADC_Value(2, 1);
                    Log.WriteLine(Name + ":" + IoutValues[i].ToString("F2") + ":" + VrectValues[j].ToString("F1") + ":" + ISenCodes[i, j].ToString());
                }
            }
            ElectronicLoad.Write("CURR 0");

            Const = (short)IoutValues[0]; // Low load current
            CalibrationData[0x19] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x1A] = (byte)((Const >> 8) & 0xFF);
            Log.WriteLine(Name + "_Low:PASS:" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            Const = (short)(ISenCodes[2, 0] - ISenCodes[2, 1]);
            CalibrationData[0x1B] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x1C] = (byte)((Const >> 8) & 0xFF);
            Log.WriteLine(Name + "_Diff:PASS:" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);

            // Calculate Isen0
            Slope = (short)((IoutValues[2] - IoutValues[0]) * 1000 * 1024 / (ISenCodes[2, 0] - ISenCodes[0, 0]));
            Const = (short)(IoutValues[0] * 1000 - (Slope * ISenCodes[0, 0]) / 1024);
            Log.WriteLine(Name + "_0:" + Slope.ToString() + ":" + Const.ToString(), Color.DarkGray, Log.RichTextBox.BackColor);

            if ((Slope >= 500 && Slope <= 2500) && (Const >= -1512 && Const <= 512))
                Log.WriteLine(Name + "_0:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_0:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x1D] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x1E] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x1F] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x20] = (byte)((Const >> 8) & 0xFF);

            // Calculate Isen1
            Slope = (short)((IoutValues[2] - IoutValues[1]) * 1000 * 1024 / (ISenCodes[2, 1] - ISenCodes[1, 1]));
            Const = (short)(IoutValues[1] * 1000 - (Slope * ISenCodes[1, 1]) / 1024);
            Log.WriteLine(Name + "_1:" + Slope.ToString() + ":" + Const.ToString(), Color.DarkGray, Log.RichTextBox.BackColor);
            if ((Slope >= 500 && Slope <= 2500) && (Const >= -1512 && Const <= 512))
                Log.WriteLine(Name + "_1:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_1:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x21] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x22] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x23] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x24] = (byte)((Const >> 8) & 0xFF);

            // Calculate Isen2
            Slope = (short)((IoutValues[1] - IoutValues[0]) * 1000 * 1024 / (ISenCodes[1, 1] - ISenCodes[0, 1]));
            Const = (short)(IoutValues[0] * 1000 - (Slope * ISenCodes[0, 1]) / 1024);
            if ((Slope >= 500 && Slope <= 2500) && (Const >= -1512 && Const <= 512))
                Log.WriteLine(Name + "_2:PASS:" + Slope.ToString() + ":" + Const.ToString(), Color.ForestGreen, Log.RichTextBox.BackColor);
            else
            {
                Log.WriteLine(Name + "_2:FAIL:" + Slope.ToString() + ":" + Const.ToString(), Color.Coral, Log.RichTextBox.BackColor);
                Slope = 1000;
                Const = 0;
            }
            CalibrationData[0x25] = (byte)((Slope >> 0) & 0xFF);
            CalibrationData[0x26] = (byte)((Slope >> 8) & 0xFF);
            CalibrationData[0x27] = (byte)((Const >> 0) & 0xFF);
            CalibrationData[0x28] = (byte)((Const >> 8) & 0xFF);
        }

        private void RunCalibration_Click()
        {
            Parent.StopLogThread();
            if (IsRunCal == false)
            {
                IsRunCal = true;
                RunBackgroudFunc(RunCalibration, StopCalibration);
            }
            else
                StopCalibration();
        }

        private void Write_Calibration_Click()
        {
            List<byte> CalData = new List<byte>();
            byte[] WriteData;
            System.IO.FileStream fs = new System.IO.FileStream("Toscana_Cal_Data.bin", System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);

            CalData.Add(0xE0);
            for (int i = 0; i < 0x29; i++)
            {
                CalData.Add(CalibrationData[i]);
            }

            WriteData = CalData.ToArray();
            bw.Write(WriteData);
            bw.Close();
            fs.Close();
            Log.WriteLine("Save Done! Toscana_Cal_Data.bin");
        }

        private void Read_Calibration_Click()
        {
            byte[] ReadeData;
            System.IO.FileStream fs = new System.IO.FileStream("Toscana_Cal_Data.bin", System.IO.FileMode.Open, System.IO.FileAccess.Read);
            System.IO.BinaryReader br = new System.IO.BinaryReader(fs);

            ReadeData = br.ReadBytes(0x2A);
            for (int i = 0; i < 0x29; i++)
            {
                CalibrationData[i] = ReadeData[i + 1];
                Log.Write(CalibrationData[i].ToString() + " ");
            }
            br.Close();
            fs.Close();
            Log.WriteLine("\nRead Done!");
        }

        private void Apply_Calibration_Click()
        {
            List<byte> SendBytes = new List<byte>();
            byte[] RcvBytes = new byte[1];
            int sa;

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(0xE0);
            for (int i = 0; i < 0x29; i++)
            {
                SendBytes.Add(CalibrationData[i]);
            }
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            for (int i = 0; i < 10; i++)
            {
                System.Threading.Thread.Sleep(10);
                RcvBytes = I2C.ReadBytes(1);
                if (RcvBytes[0] == 0)
                    break;
            }

            I2C.Config.SlaveAddress = sa;
            Log.WriteLine("Send Done!");
        }
        #endregion Calibration methods

        #region Firmware control methods
        private void DownloadFirmware(FW_TARGET target)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = target;
                DownloadFirmware(TOSCANA_DownloadFW);
            }
        }

        private void EraseFW(FW_TARGET target)
        {
            if (target == FW_TARGET.RAM)
                return;
            FirmwareTarget = target;
            HaltMCU();
            EraseFirmware(TOSCANA_EraseFW);
        }

        private void DumpFW(FW_TARGET target)
        {
            try
            {
                ReadFirmwareSize = int.Parse(TextBox_TestArgument.Text, System.Globalization.NumberStyles.Number);
            }
            catch
            {
                MessageBox.Show("FW size를 입력하고 다시 시도해주세요.");
                return;
            }

            FirmwareTarget = target;
            DumpFirmware(TOSCANA_DumpFW);
        }

        void WriteMemory_NVM(uint Address, byte[] Data)
        {
            int cnt = 0;
            uint data;

            if (Data.Length < 256)
                return;

            // Fill FIFO for page program
            for (uint i = 0; i < 64; i++)
            {
                data = (uint)((Data[i * 4 + 3] << 24) | (Data[i * 4 + 2] << 16) | (Data[i * 4 + 1] << 8) | Data[i * 4 + 0]);
                WriteRegister((0x50091000 + (i * 4)), data);
            }

            // Page Program Command
            data = ((uint)(FLASH_CMD.PP) << 24) | (Address & 0xFFFFFF);
            WriteRegister(0x50090008, data);
            System.Threading.Thread.Sleep(2);

            // Check Status Register
            while (TOSCANA_Check_Staus_Register() != 0x00) // typ 667.5us, max 3110us
            {
                cnt++;
                System.Threading.Thread.Sleep(1);
                if (cnt > 20) break;
            }
            if (cnt > 20) Status = false;
        }

        private void TOSCANA_DownloadFW()
        {
            int PageSize = 256;
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

            if (FirmwareTarget == FW_TARGET.RAM)
            {
                PageSize = 4;
            }

            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (FirmwareData.Length + PageSize - 1) / PageSize;
            })));

            HaltMCU();

            /* Mass erase */
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                TOSCANA_EraseFW();
                if (Status == false)
                    goto EXIT;
            }

            /* Program */
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

                // Write firmware data
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    WriteRegister(FlashAddress, SendBytes);
                }
                else
                {
                    WriteMemory_NVM(FlashAddress, SendBytes);
                }

                // Compare memory
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    for (uint i = 0; i < PageSize; i += 4)
                    {
                        uint RcvBytes = ReadRegister(FlashAddress + i);
                        for (int j = 0; j < 4; j++)
                        {
                            if (SendBytes[i * 4 + j] != ((RcvBytes >> (8 * j)) & 0xff))
                            {
                                Log.WriteLine("Faile to match: (0x" + (FlashAddress + i * 4 + j).ToString("X4") + ":" + SendBytes[i * 4 + j].ToString("X2") + " " + ((RcvBytes >> (8 * j)) & 0xff).ToString("X2") + ")",
                                    System.Drawing.Color.Coral, Log.RichTextBox.BackColor);
                                goto EXIT;
                            }
                        }
                    }
                }
                else
                {

                }

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                WriteRegister(0x50090008, ((uint)(FLASH_CMD.WRDI) << 24));
                System.Threading.Thread.Sleep(1);
            }

            ResetMCU();

            Status = true;
            fs.Close();
            br.Close();

        EXIT:
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }

        public uint TOSCANA_Check_Staus_Register()
        {
            Status = true;

            return ReadRegister(0x50090020) & 0x01;
        }

        public void TOSCANA_EraseFW()
        {
            Status = true;
            int cnt = 0;
            List<byte> SendBytes = new List<byte>();

            WriteRegister(0x50090008, ((uint)(FLASH_CMD.CE) << 24));
            System.Threading.Thread.Sleep(100);
            while (TOSCANA_Check_Staus_Register() != 0x00) // typ 1s, max 4s
            {
                cnt++;
                System.Threading.Thread.Sleep(200);
                if (cnt > 20) break;
            }
            if (cnt > 20) Status = false;
        }

        public void TOSCANA_DumpFW()
        {
            const int PageSize = 4;
            uint RcvBytes;
            List<byte> FirmwareData = new List<byte>();
            string file_name;

            Status = false;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (ReadFirmwareSize + PageSize - 1) / PageSize + 1; // Mass erase
            })));

            HaltMCU();

            for (uint Addr = 0; Addr < ReadFirmwareSize; Addr += PageSize)
            {
                RcvBytes = ReadRegister(Addr);
                FirmwareData.Add((byte)(RcvBytes & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 8) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 16) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 24) & 0xff));

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            Status = true;
            ReadFirmwareData = FirmwareData.ToArray();
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                file_name = "ReadFirmwareBinary_NVM.bin";
            }
            else
            {
                file_name = "ReadFirmwareBinary_RAM.bin";
            }
            System.IO.FileStream fs = new System.IO.FileStream(file_name, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(ReadFirmwareData);
            bw.Close();
            fs.Close();
            // 3. Reset system
            ResetMCU();

            Status = true;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }
        #endregion Firmware control methods

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

        private void Run_PWM_GPIO_SET(object sender, EventArgs e)
        {
            RegisterItem ENABLE = Parent.RegMgr.GetRegisterItem("ENABLE<11:0>");
            RegisterItem GPIO_03_SEL = Parent.RegMgr.GetRegisterItem("GPIO_03_SEL<7:0>");
            RegisterItem GPIO_02_SEL = Parent.RegMgr.GetRegisterItem("GPIO_02_SEL<7:0>");
            RegisterItem GPIO_01_SEL = Parent.RegMgr.GetRegisterItem("GPIO_01_SEL<7:0>");
            RegisterItem GPIO_00_SEL = Parent.RegMgr.GetRegisterItem("GPIO_00_SEL<7:0>");

            ENABLE.Read();
            GPIO_00_SEL.Read();

            ENABLE.Value = 15;
            GPIO_03_SEL.Value = 177;
            GPIO_02_SEL.Value = 176;
            GPIO_01_SEL.Value = 175;
            GPIO_00_SEL.Value = 174;

            ENABLE.Write();
            GPIO_00_SEL.Write();
        }

        private void Run_PWM_Initial(object sender, EventArgs e)
        {
            RegisterItem DIV_MCU = Parent.RegMgr.GetRegisterItem("DIV_MCU<3:0>");
            RegisterItem EN_CLK_8M = Parent.RegMgr.GetRegisterItem("EN_CLK_8M");
            RegisterItem DIV_CLK_8M = Parent.RegMgr.GetRegisterItem("DIV_CLK_8M<3:0>");
            RegisterItem R_SR_MD_LS = Parent.RegMgr.GetRegisterItem("R_SR_MD_LS     ");
            RegisterItem R_SR_MD_HS = Parent.RegMgr.GetRegisterItem("R_SR_MD_HS");

            RegisterItem PWM0_SYNCMODE = Parent.RegMgr.GetRegisterItem("PWM0_SYNCMODE");
            RegisterItem PWM0_INTEN = Parent.RegMgr.GetRegisterItem("PWM0_INTEN");
            RegisterItem PWM0_CLKSEL = Parent.RegMgr.GetRegisterItem("PWM0_CLKSEL<2:0>");
            RegisterItem PWM0_PWMEN = Parent.RegMgr.GetRegisterItem("PWM0_PWMEN");

            RegisterItem PWM1_SYNCMODE = Parent.RegMgr.GetRegisterItem("PWM1_SYNCMODE");
            RegisterItem PWM1_INTEN = Parent.RegMgr.GetRegisterItem("PWM1_INTEN");
            RegisterItem PWM1_CLKSEL = Parent.RegMgr.GetRegisterItem("PWM1_CLKSEL<2:0>");
            RegisterItem PWM1_PWMEN = Parent.RegMgr.GetRegisterItem("PWM1_PWMEN");

            RegisterItem PWM2_SYNCMODE = Parent.RegMgr.GetRegisterItem("PWM2_SYNCMODE");
            RegisterItem PWM2_INTEN = Parent.RegMgr.GetRegisterItem("PWM2_INTEN");
            RegisterItem PWM2_CLKSEL = Parent.RegMgr.GetRegisterItem("PWM2_CLKSEL<2:0>");
            RegisterItem PWM2_PWMEN = Parent.RegMgr.GetRegisterItem("PWM2_PWMEN");

            RegisterItem PWM3_SYNCMODE = Parent.RegMgr.GetRegisterItem("PWM3_SYNCMODE");
            RegisterItem PWM3_INTEN = Parent.RegMgr.GetRegisterItem("PWM3_INTEN");
            RegisterItem PWM3_CLKSEL = Parent.RegMgr.GetRegisterItem("PWM3_CLKSEL<2:0>");
            RegisterItem PWM3_PWMEN = Parent.RegMgr.GetRegisterItem("PWM3_PWMEN");

            DIV_MCU.Read();
            DIV_MCU.Value = 15; // Set MCU Clk 16MHz
            DIV_MCU.Write();

            EN_CLK_8M.Read();
            EN_CLK_8M.Value = 1;
            DIV_CLK_8M.Value = 1;
            EN_CLK_8M.Write();

            R_SR_MD_LS.Read();
            R_SR_MD_LS.Value = 1;
            R_SR_MD_HS.Value = 1;
            R_SR_MD_LS.Write();


            PWM0_SYNCMODE.Read();
            PWM1_SYNCMODE.Read();
            PWM2_SYNCMODE.Read();
            PWM3_SYNCMODE.Read();

            PWM0_SYNCMODE.Value = 1;
            PWM0_INTEN.Value = 1;
            PWM0_CLKSEL.Value = 0;
            PWM0_PWMEN.Value = 1;

            PWM1_SYNCMODE.Value = 1;
            PWM1_INTEN.Value = 1;
            PWM1_CLKSEL.Value = 0;
            PWM1_PWMEN.Value = 1;

            PWM2_SYNCMODE.Value = 1;
            PWM2_INTEN.Value = 1;
            PWM2_CLKSEL.Value = 0;
            PWM2_PWMEN.Value = 1;

            PWM3_SYNCMODE.Value = 1;
            PWM3_INTEN.Value = 1;
            PWM3_CLKSEL.Value = 0;
            PWM3_PWMEN.Value = 1;

            PWM0_SYNCMODE.Write();
            PWM1_SYNCMODE.Write();
            PWM2_SYNCMODE.Write();
            PWM3_SYNCMODE.Write();

        }

        private void Run_PWM_SET(object sender, EventArgs e)
        {
            uint Freq = uint.Parse(Parent.ChipCtrlTextboxes[0].Text);
            uint Duty = uint.Parse(Parent.ChipCtrlTextboxes[1].Text);
            uint Period = (uint)Math.Round((double)(16000 / Freq)); ;
            uint COUNT_AC1 = (uint)Math.Ceiling((decimal)(Period * Duty / 100));
            uint COUNT_AC2 = (uint)Math.Ceiling((decimal)(Period * (100 - Duty) / 100));


            RegisterItem START = Parent.RegMgr.GetRegisterItem("PWM0_START");

            RegisterItem PWM0_PERIOD = Parent.RegMgr.GetRegisterItem("PWM0_PERIOD<15:0>");
            RegisterItem PWM0_DCOUNT = Parent.RegMgr.GetRegisterItem("PWM0_DCOUNT<15:0>");
            RegisterItem PWM0_UCOUNT = Parent.RegMgr.GetRegisterItem("PWM0_UCOUNT<15:0>");

            RegisterItem PWM1_PERIOD = Parent.RegMgr.GetRegisterItem("PWM1_PERIOD<15:0>");
            RegisterItem PWM1_DCOUNT = Parent.RegMgr.GetRegisterItem("PWM1_DCOUNT<15:0>");
            RegisterItem PWM1_UCOUNT = Parent.RegMgr.GetRegisterItem("PWM1_UCOUNT<15:0>");

            RegisterItem PWM2_PERIOD = Parent.RegMgr.GetRegisterItem("PWM2_PERIOD<15:0>");
            RegisterItem PWM2_DCOUNT = Parent.RegMgr.GetRegisterItem("PWM2_DCOUNT<15:0>");
            RegisterItem PWM2_UCOUNT = Parent.RegMgr.GetRegisterItem("PWM2_UCOUNT<15:0>");

            RegisterItem PWM3_PERIOD = Parent.RegMgr.GetRegisterItem("PWM3_PERIOD<15:0>");
            RegisterItem PWM3_DCOUNT = Parent.RegMgr.GetRegisterItem("PWM3_DCOUNT<15:0>");
            RegisterItem PWM3_UCOUNT = Parent.RegMgr.GetRegisterItem("PWM3_UCOUNT<15:0>");

            PWM0_PERIOD.Read();
            PWM0_DCOUNT.Read();
            PWM0_UCOUNT.Read();

            PWM1_PERIOD.Read();
            PWM1_DCOUNT.Read();
            PWM1_UCOUNT.Read();

            PWM2_PERIOD.Read();
            PWM2_DCOUNT.Read();
            PWM2_UCOUNT.Read();

            PWM3_PERIOD.Read();
            PWM3_DCOUNT.Read();
            PWM3_UCOUNT.Read();

            PWM0_PERIOD.Value = Period;
            PWM0_DCOUNT.Value = COUNT_AC1;
            PWM0_UCOUNT.Value = 1;

            PWM1_PERIOD.Value = Period;
            PWM1_DCOUNT.Value = Period;
            PWM1_UCOUNT.Value = COUNT_AC2 + 1;

            PWM2_PERIOD.Value = Period;
            PWM2_DCOUNT.Value = COUNT_AC2;
            PWM2_UCOUNT.Value = 1;

            PWM3_PERIOD.Value = Period;
            PWM3_DCOUNT.Value = Period;
            PWM3_UCOUNT.Value = COUNT_AC1 + 1;

            PWM0_PERIOD.Write();
            PWM0_DCOUNT.Write();
            PWM0_UCOUNT.Write();

            PWM1_PERIOD.Write();
            PWM1_DCOUNT.Write();
            PWM1_UCOUNT.Write();

            PWM2_PERIOD.Write();
            PWM2_DCOUNT.Write();
            PWM2_UCOUNT.Write();

            PWM3_PERIOD.Write();
            PWM3_DCOUNT.Write();
            PWM3_UCOUNT.Write();

            START.Value = 1;
            START.Write();
        }

        private void StartCommandProcessing(byte CommandReg, byte lsb, byte msb, int Timeout = 1000)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvBytes = new byte[1];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(CommandReg);
            SendBytes.Add(lsb);
            SendBytes.Add(msb);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            if (CommandReg != 0x0F) // TEST_ITEMS.FIRM_ON_CLEAR
            {
                for (int i = 0; i < (Timeout / 10); i++)
                {
                    System.Threading.Thread.Sleep(10);
                    RcvBytes = I2C.ReadBytes(1);
                    if (RcvBytes[0] == 0)
                        break;
                }
                //Log.WriteLine(RcvBytes[0].ToString());
            }

            I2C.Config.SlaveAddress = sa;
        }

        public int StartReadResult()
        {
            int sa;
            int result;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            byte[] RcvBytes = new byte[3];

            SendBytes.Add(0x00);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            RcvBytes = I2C.ReadBytes(3);
            result = (RcvBytes[2] << 8) + RcvBytes[1];

            //Log.WriteLine(RcvBytes[0].ToString());
            //Log.WriteLine(RcvBytes[1].ToString());
            //Log.WriteLine(RcvBytes[2].ToString());

            I2C.Config.SlaveAddress = sa;

            return result;
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal, Result = 0;
            TEST_ITEMS TestItem = (TEST_ITEMS)TestItemIndex;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch (CombBox_Item)
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

    public class Columbus : ChipControl
    {
        #region Variable and declaration

        public enum FW_TARGET
        {
            NV_MEM = 0,
            RAM = 1,
            SPI,
        }

        public enum FLASH_CMD
        {
            WRSR = 0x01,    // write stat. reg
            PP = 0x02,      // page program
            RDCMD = 0x03,   // read data
            WRDI = 0x04,    // write disable
            RDSR = 0x05,    // read status reg
            WREN = 0x06,    // write enable
            F_RD = 0x0B,    // fast read datwjda
            SE = 0x20,      // 4KB sector erase
            BE32 = 0x52,    // 32KB block erase
            RSTEN = 0x66,   // reset enable
            REMS = 0x90,    // read manufacture &
            RST = 0x99,     // reset
            RDID = 0x9F,    // read identificatio
            RES = 0xAB,     // read signature
            ENSO = 0xB1,    // enter secured OTP
            DP = 0xB9,      // deep power down
            EXSO = 0xC1,    // exit secured OTP
            CE = 0xC7,      // chip(bulk) erase
            BE64 = 0xD8,	// 64KB block erase
        }

        public enum TEST_ITEMS
        {
            TEST,

            NUM_TEST_ITEMS,
        }

        public enum AUTO_TEST_ITEMS
        {
            TEST,

            NUM_TEST_ITEMS,
        }

        public enum FW_DN_ITEMS
        {
            FLASH_ERASE,
            FLASH_WRITE,
            FLASH_READ,
            RAM_WRITE,
            RAM_READ,

            FIRM_ON_CLEAR, // 0x0F

            NUM_TEST_ITEMS,
        }

        public enum CAL_ITEMS
        {
            CAL_RUN,
            CAL_WRITE,
            CAL_READ,
            CAL_APPLY,

            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            TEST,
            AUTO,
            CAL,
            FW,
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Custom.SPI SPI { get; set; }
        private FW_TARGET FirmwareTarget { get; set; } = FW_TARGET.RAM;

        public int SlaveAddress { get; private set; } = 0x52;
        private bool IsRunCal = false;
        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.TEST;

        /* Intrument */
        JLcLib.Instrument.SCPI PowerSupply0 = null;
        JLcLib.Instrument.SCPI PowerSupply1 = null;
        JLcLib.Instrument.SCPI Oscilloscope = null;
        JLcLib.Instrument.SCPI ElectronicLoad = null;
        #endregion Variable and declaration

        public Columbus(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            SPI = form.SPI;
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
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private void WriteRegister(uint Address, byte[] Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add(Data[0]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[3]);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private uint ReadRegister(uint Address)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[4];
            uint Data;

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            RcvData = iComn.ReadBytes(RcvData.Length);
            Data = (uint)((RcvData[3] << 24) | (RcvData[2] << 16) | (RcvData[1] << 8) | (RcvData[0]));

            I2C.Config.SlaveAddress = sa;
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

        #region SKAI Columbus host control methods
        public void HaltMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0x56, 0x78 };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }

        public void ResetMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0xAB, 0xCD };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }
        #endregion SKAI Columbus host control methods

        public override void SetChipSpecificUI()
        {
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */

            Parent.ChipCtrlButtons[4].Text = "GH0_H";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Toogle_GPIO_GH0;

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

        #region Firmware control methods
        private void DownloadFirmware(FW_TARGET target)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = target;
                DownloadFirmware(Columbus_DownloadFW);
            }
        }

        private void EraseFW(FW_TARGET target)
        {
            if (target == FW_TARGET.RAM)
                return;
            FirmwareTarget = target;
            HaltMCU();
            EraseFirmware(Columbus_EraseFW);
        }

        private void DumpFW(FW_TARGET target)
        {
            try
            {
                ReadFirmwareSize = int.Parse(TextBox_TestArgument.Text, System.Globalization.NumberStyles.Number);
            }
            catch
            {
                MessageBox.Show("FW size를 입력하고 다시 시도해주세요.");
                return;
            }

            FirmwareTarget = target;
            DumpFirmware(Columbus_DumpFW);
        }

        void WriteMemory_NVM(uint Address, byte[] Data)
        {
            int cnt = 0;
            uint data;

            if (Data.Length < 256)
                return;

            // Fill FIFO for page program
            for (uint i = 0; i < 64; i++)
            {
                data = (uint)((Data[i * 4 + 3] << 24) | (Data[i * 4 + 2] << 16) | (Data[i * 4 + 1] << 8) | Data[i * 4 + 0]);
                WriteRegister((0x50091000 + (i * 4)), data);
            }

            // Page Program Command
            data = ((uint)(FLASH_CMD.PP) << 24) | (Address & 0xFFFFFF);
            WriteRegister(0x50090008, data);
            System.Threading.Thread.Sleep(2);

            // Check Status Register
            while (Columbus_Check_Staus_Register() != 0x00) // typ 667.5us, max 3110us
            {
                cnt++;
                System.Threading.Thread.Sleep(1);
                if (cnt > 20) break;
            }
            if (cnt > 20) Status = false;
        }

        private void Columbus_DownloadFW()
        {
            int PageSize = 256;
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
            if (FirmwareTarget == FW_TARGET.RAM)
            {
                PageSize = 4;
            }

            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (FirmwareData.Length + PageSize - 1) / PageSize;
            })));
            HaltMCU();
            /* Mass erase */
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                Columbus_EraseFW();
                if (Status == false)
                    goto EXIT;
            }
            else if (FirmwareTarget == FW_TARGET.RAM)
            {
                WriteRegister(0x50060030, 0x01000000);
            }

            /* Program */
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

                // Write firmware data
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    WriteRegister(FlashAddress + 0x80000000, SendBytes);
                }
                else
                {
                    WriteMemory_NVM(FlashAddress, SendBytes);
                }

                // Compare memory
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    for (uint i = 0; i < PageSize; i += 4)
                    {
                        uint RcvBytes = ReadRegister(FlashAddress + i + 0x80000000);
                        for (int j = 0; j < 4; j++)
                        {
                            if (SendBytes[i * 4 + j] != ((RcvBytes >> (8 * j)) & 0xff))
                            {
                                Log.WriteLine("Faile to match: (0x" + (FlashAddress + i * 4 + j).ToString("X4") + ":" + SendBytes[i * 4 + j].ToString("X2") + " " + ((RcvBytes >> (8 * j)) & 0xff).ToString("X2") + ")",
                                    System.Drawing.Color.Coral, Log.RichTextBox.BackColor);
                                goto EXIT;
                            }
                        }
                    }
                }

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                WriteRegister(0x50090008, ((uint)(FLASH_CMD.WRDI) << 24));
                System.Threading.Thread.Sleep(1);
            }

            ResetMCU();

            Status = true;
            fs.Close();
            br.Close();

        EXIT:
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }

        public uint Columbus_Check_Staus_Register()
        {
            Status = true;

            return ReadRegister(0x50090020) & 0x01;
        }

        public void Columbus_EraseFW()
        {
            Status = true;
            int cnt = 0;
            List<byte> SendBytes = new List<byte>();

            WriteRegister(0x50090008, ((uint)(FLASH_CMD.CE) << 24));
            System.Threading.Thread.Sleep(100);
            while (Columbus_Check_Staus_Register() != 0x00) // typ 1s, max 4s
            {
                cnt++;
                System.Threading.Thread.Sleep(200);
                if (cnt > 20) break;
            }
            if (cnt > 20) Status = false;
        }

        public void Columbus_DumpFW()
        {
            const int PageSize = 4;
            uint RcvBytes;
            List<byte> FirmwareData = new List<byte>();
            string file_name;

            Status = false;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = 0;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = (ReadFirmwareSize + PageSize - 1) / PageSize + 1; // Mass erase
            })));

            HaltMCU();

            for (uint Addr = 0; Addr < ReadFirmwareSize; Addr += PageSize)
            {
                RcvBytes = ReadRegister(Addr + 0x80000000);
                FirmwareData.Add((byte)(RcvBytes & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 8) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 16) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 24) & 0xff));

                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate ()
                {
                    ProgressBar.Value++;
                })));
            }
            Status = true;
            ReadFirmwareData = FirmwareData.ToArray();
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                file_name = "ReadFirmwareBinary_NVM.bin";
            }
            else
            {
                file_name = "ReadFirmwareBinary_RAM.bin";
            }
            System.IO.FileStream fs = new System.IO.FileStream(file_name, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(ReadFirmwareData);
            bw.Close();
            fs.Close();

            // 3. Reset system
            ResetMCU();

            Status = true;
            ProgressBar?.Invoke((new MethodInvoker(delegate ()
            {
                ProgressBar.Value = ProgressBar.Maximum;
            })));
        }
        #endregion Firmware control methods

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

        private void StartCommandProcessing(byte CommandReg, byte lsb, byte msb, int Timeout = 1000)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvBytes = new byte[1];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(CommandReg);
            SendBytes.Add(lsb);
            SendBytes.Add(msb);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            if (CommandReg != 0x0F) // TEST_ITEMS.FIRM_ON_CLEAR
            {
                for (int i = 0; i < (Timeout / 10); i++)
                {
                    System.Threading.Thread.Sleep(10);
                    RcvBytes = I2C.ReadBytes(1);
                    if (RcvBytes[0] == 0)
                        break;
                }
                //Log.WriteLine(RcvBytes[0].ToString());
            }

            I2C.Config.SlaveAddress = sa;
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal;
            DialogResult dialog;
            TEST_ITEMS TestItem = (TEST_ITEMS)TestItemIndex;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.TEST:
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    break;
                case COMBOBOX_ITEMS.CAL:
                    break;
                case COMBOBOX_ITEMS.FW:
                    switch ((FW_DN_ITEMS)TestItemIndex)
                    {
                        case FW_DN_ITEMS.FLASH_ERASE:
                            EraseFW(FW_TARGET.NV_MEM);
                            break;
                        case FW_DN_ITEMS.FLASH_WRITE:
                            DownloadFirmware(FW_TARGET.NV_MEM);
                            break;
                        case FW_DN_ITEMS.FLASH_READ:
                            dialog = MessageBox.Show("BMODE pin에 1.8V를 인가하였나요?", Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            if (dialog != DialogResult.OK)
                            {
                                break;
                            }
                            DumpFW(FW_TARGET.NV_MEM);
                            break;
                        case FW_DN_ITEMS.RAM_WRITE:
                            DownloadFirmware(FW_TARGET.RAM);
                            break;
                        case FW_DN_ITEMS.RAM_READ:
                            dialog = MessageBox.Show("BMODE pin에 0V를 인가하거나 floating 하였나요?", Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            if (dialog != DialogResult.OK)
                            {
                                break;
                            }
                            DumpFW(FW_TARGET.RAM);
                            break;
                        case FW_DN_ITEMS.FIRM_ON_CLEAR:
                            StartCommandProcessing(0x0F, 0, 0);
                            break;
                        default:
                            break;
                    }
                    break;
            }
        }
    }

    public class Oasis : ChipControl
    {
        #region Variable and declaration
        public enum FW_TARGET
        {
            NV_MEM = 0,
            RAM = 1,
            SPI,
        }

        public enum FLASH_CMD
        {
            WRSR = 0x01,    // write start. reg
            PP = 0x02,      // page program
            RDCMD = 0x03,   // read data
            WRDI = 0x04,    // write disable
            RDSR = 0x05,    // read status reg
            WREN = 0x06,    // write enable
            F_RD = 0x0B,    // fast read datwjda
            SE = 0x20,      // 4KB sector erase
            BE32 = 0x52,    // 32KB block erase
            RSTEN = 0x66,   // reset enable
            REMS = 0x90,    // read manufacture &
            RST = 0x99,     // reset
            RDID = 0x9F,    // read identification
            RES = 0xAB,     // read signature
            ENSO = 0xB1,    // enter secured OTP
            DP = 0xB9,      // deep power down
            EXSO = 0xC1,    // exit secured OTP
            CE = 0xC7,      // chip(bulk) erase
            BE64 = 0xD8,	// 64KB sector erase
        }

        private JLcLib.Custom.I2C I2C { get; set; }
        private JLcLib.Custom.SPI SPI { get; set; }
        private JLcLib.Custom.MPSSE mPSSE { get; set; }
        private Serial Serial { get; set; } = new Serial();
        private SerialPort SerialPort { get; set; } = new SerialPort();

        private FW_TARGET FirmwareTarget { get; set; } = FW_TARGET.RAM;
        public int SlaveAddress { get; private set; } = 0x52;
        private bool IsRunCal = false;
        private COMBOBOX_ITEMS CombBox_Item = COMBOBOX_ITEMS.TEST;

        private BackgroundWorker autotestWorker = new BackgroundWorker();

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
        JLcLib.Instrument.SCPI SignalGenerator0 = null;
        JLcLib.Instrument.SCPI SignalGenerator1 = null;
        #endregion Variable and declaration

        public Oasis(RegContForm form) : base(form)
        {
            I2C = form.I2C;
            SPI = form.SPI;
            mPSSE = form.MPSSE;
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
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add((byte)((Data >> 0) & 0xFF));
            SendBytes.Add((byte)((Data >> 8) & 0xFF));
            SendBytes.Add((byte)((Data >> 16) & 0xFF));
            SendBytes.Add((byte)((Data >> 24) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private void WriteRegister(uint Address, byte[] Data)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            SendBytes.Clear();
            SendBytes.Add(Data[0]);
            SendBytes.Add(Data[1]);
            SendBytes.Add(Data[2]);
            SendBytes.Add(Data[3]);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            I2C.Config.SlaveAddress = sa;
        }

        private uint ReadRegister(uint Address)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvData = new byte[4];
            uint Data;

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;
            SendBytes.Add(0xA1);
            SendBytes.Add(0x2C);
            SendBytes.Add(0x12);
            SendBytes.Add(0x34);
            SendBytes.Add((byte)((Address >> 24) & 0xFF));
            SendBytes.Add((byte)((Address >> 16) & 0xFF));
            SendBytes.Add((byte)((Address >> 8) & 0xFF));
            SendBytes.Add((byte)((Address >> 0) & 0xFF));
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            RcvData = I2C.ReadBytes(RcvData.Length);
            Data = (uint)((RcvData[3] << 24) | (RcvData[2] << 16) | (RcvData[1] << 8) | (RcvData[0]));

            I2C.Config.SlaveAddress = sa;
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

        #region Run_TEST
        public enum TEST_ITEMS
        {
            GPIO_DISABLE,
            GPIO_04_ABGR,
            GPIO_04_RETLDO,
            GPIO_04_MBGR,
            GPIO_04_DALDO,
            GPIO_16_32KOSC,
            NUM_TEST_ITEMS,
        }

        public enum AUTO_TEST_ITEMS
        {
            //GPIO_SWEEP,
            //DMM_SWEEP,
            //GPADC_SWEEP,
            //GPADC_REPEAT,
            //SNIFFER_SWEEP,
            //SNIFFER_REPEAT,
            ABGR_SWEEP,
            //ALDO_SWEEP,
            //RTCOSC_SWEEP,
            //DCDC_SWEEP,
            //RETLDO_SWEEP,
            //FLDO_SWEEP,
            //MLDO_SWEEP,
            //MBGR_SWEEP,
            //DALDO_SWEEP,
            ABGR_TC_SWEEP,
            //DOUT_read_N,
            //DOUT_Sweep_N,
            PMU_TEMP_TEST,
            NUM_TEST_ITEMS,
        }

        public enum FW_DN_ITEMS
        {
            FLASH_ERASE,
            FLASH_WRITE,
            FLASH_READ,
            FLASH_VERIFY,
            FIRM_ON_CLEAR, // 0x0F
            RAM_WRITE,
            RAM_READ,
            RESET,
            NUM_TEST_ITEMS,
        }

        public enum CAL_ITEMS
        {
            TEST_ITEM,
            //ABGR_TRIM,
            //MLDO_TRIM,
            //ALDO_TRIM,
            NUM_TEST_ITEMS,
        }

        public enum COMBOBOX_ITEMS
        {
            TEST,
            AUTO,
            CAL,
            FW,
        }

        public override void RunTest(int TestItemIndex, string Arg)
        {
            int iVal;
            DialogResult dialog;
            TEST_ITEMS TestItem = (TEST_ITEMS)TestItemIndex;

            try { iVal = int.Parse(Arg, System.Globalization.NumberStyles.Number); }
            catch { iVal = 0; }

            switch (CombBox_Item)
            {
                case COMBOBOX_ITEMS.TEST:
                    switch ((TEST_ITEMS)TestItemIndex)
                    {
                        case TEST_ITEMS.GPIO_DISABLE:
                            Set_GPIO_Disable();
                            break;
                        case TEST_ITEMS.GPIO_04_ABGR:
                            Set_GPIO4_ABGR(true);
                            break;
                        case TEST_ITEMS.GPIO_04_RETLDO:
                            Set_GPIO4_RETLDO(true);
                            break;
                        case TEST_ITEMS.GPIO_04_MBGR:
                            Set_GPIO4_MBGR(true);
                            break;
                        case TEST_ITEMS.GPIO_04_DALDO:
                            Set_GPIO4_DALDO(true);
                            break;
                        case TEST_ITEMS.GPIO_16_32KOSC:
                            Set_GPIO16_32KOSC(true);
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.AUTO:
                    switch ((AUTO_TEST_ITEMS)TestItemIndex)
                    {
                        //case AUTO_TEST_ITEMS.GPIO_SWEEP:
                        //    Run_ANA_TEST_BUFF_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.DMM_SWEEP:
                        //    Run_DMM_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.GPADC_SWEEP:
                        //    Run_GPADC_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.GPADC_REPEAT:
                        //    Run_GPADC_REPEAT();
                        //    break;
                        //case AUTO_TEST_ITEMS.SNIFFER_SWEEP:
                        //    Run_SNIFFER_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.SNIFFER_REPEAT:
                        //    Run_SNIFFER_REPEAT();
                        //    break;
                        case AUTO_TEST_ITEMS.ABGR_SWEEP:
                            Run_ABGR_SWEEP(iVal);
                            break;
                        //case AUTO_TEST_ITEMS.ALDO_SWEEP:
                        //    Run_ALDO_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.RTCOSC_SWEEP:
                        //    Run_32KOSC_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.DCDC_SWEEP:
                        //    Run_DCDC_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.RETLDO_SWEEP:
                        //    Run_RETLDO_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.FLDO_SWEEP:
                        //    Run_FLDO_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.MLDO_SWEEP:
                        //    Run_MLDO_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.MBGR_SWEEP:
                        //    Run_MBGR_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.DALDO_SWEEP:
                        //    Run_DALDO_SWEEP();
                        //    break;
                        //case AUTO_TEST_ITEMS.SET_TEMPCHAMBER:
                        //    Set_TempChamber(false, iVal);
                        //    break;
                        case AUTO_TEST_ITEMS.ABGR_TC_SWEEP:
                            Start_ABGR_TC_Sweep();
                            break;
                        //case AUTO_TEST_ITEMS.DOUT_read_N:
                        //    if (Arg == "")
                        //    {
                        //        dialog = MessageBox.Show("측정 횟수를 기입해주세요.");
                        //        break;
                        //    }
                        //    Read_DOUT_N(iVal);
                        //    break;
                        //case AUTO_TEST_ITEMS.DOUT_Sweep_N:
                        //    if (Arg == "")
                        //    {
                        //        dialog = MessageBox.Show("측정 횟수를 기입해주세요.");
                        //        break;
                        //    }
                        //    Read_DOUT_Swwep_N(iVal);
                        //    break;
                        case AUTO_TEST_ITEMS.PMU_TEMP_TEST:
                            Start_PMU_TEMP_TEST();
                            break;
                        default:
                            break;
                    }
                    break;
                case COMBOBOX_ITEMS.CAL:
                    switch ((CAL_ITEMS)TestItemIndex)
                    {
                        //case CAL_ITEMS.ABGR_TRIM:
                        //    Start_ABGR_Trim(iVal);
                        //    break;
                        //case CAL_ITEMS.MLDO_TRIM:
                        //    Start_MLDO_Trim(iVal);
                        //    break;
                        //case CAL_ITEMS.ALDO_TRIM:
                        //    Start_ALDO_Trim(iVal);
                        //    break;
                        default:
                            break;
                    }

                    break;
                case COMBOBOX_ITEMS.FW:
                    switch ((FW_DN_ITEMS)TestItemIndex)
                    {
                        case FW_DN_ITEMS.FLASH_ERASE:
                            EraseFW(FW_TARGET.NV_MEM);
                            break;
                        case FW_DN_ITEMS.FLASH_WRITE:
                            DownloadFirmware(FW_TARGET.NV_MEM);
                            break;
                        case FW_DN_ITEMS.FLASH_READ:
                            DumpFW(FW_TARGET.NV_MEM);
                            break;
                        case FW_DN_ITEMS.RAM_WRITE:
                            DownloadFirmware(FW_TARGET.RAM);
                            break;
                        case FW_DN_ITEMS.RAM_READ:
                            dialog = MessageBox.Show("BMODE pin에 0V를 인가하거나 floating 하였나요?", Application.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            if (dialog != DialogResult.OK)
                            {
                                break;
                            }
                            DumpFW(FW_TARGET.RAM);
                            break;
                        case FW_DN_ITEMS.FLASH_VERIFY:
                            FlashVerifyTestCore();
                            break;
                        case FW_DN_ITEMS.RESET:
                            ResetOasis();
                            break;
                        case FW_DN_ITEMS.FIRM_ON_CLEAR:
                            StartCommandProcessing(0x0F, 0, 0);
                            break;
                        default:
                            break;
                    }
                    break;
            }
        }
        #endregion Run_TEST

        #region SKAI Oasis host control methods
        public void HaltMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0x56, 0x78 };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }

        public void ResetMCU()
        {
            int sa = I2C.Config.SlaveAddress;
            byte[] SendBytes = new byte[4] { 0xA1, 0x2C, 0xAB, 0xCD };

            I2C.Config.SlaveAddress = SlaveAddress;

            I2C.WriteBytes(SendBytes, SendBytes.Length, true);
            System.Threading.Thread.Sleep(50); // wait for system reset

            I2C.Config.SlaveAddress = sa;
        }
        #endregion SKAI Oasis host control methods

        public override void SetChipSpecificUI()
        {
            /*
             * Buttons and TextBoxes index
             * [Chip test combo box] [argument text box] [test start button]
             * [ 0] [ 1] [ 2] [ 3]
             * [ 4] [ 5] [ 6] [ 7]
             * [ 8] [ 9] [10] [11]
             */

            Parent.ChipCtrlButtons[2].Text = "MEAS";
            Parent.ChipCtrlButtons[2].Visible = true;
            Parent.ChipCtrlButtons[2].Click += Click_Button_Measure;

            Parent.ChipCtrlButtons[3].Text = "IO2_H";
            Parent.ChipCtrlButtons[3].Visible = true;
            Parent.ChipCtrlButtons[3].Click += Toogle_FT4222H_GPIO2;

            Parent.ChipCtrlButtons[4].Text = "STOP";
            Parent.ChipCtrlButtons[4].Visible = true;
            Parent.ChipCtrlButtons[4].Click += Stop_ALL_AUTOTESTWORKER;

            Parent.ChipCtrlButtons[5].Text = "TEST";
            Parent.ChipCtrlButtons[5].Visible = true;
            Parent.ChipCtrlButtons[5].Click += Run_TEST_FUNCTION;

            Parent.ChipCtrlButtons[6].Text = "PLL_RS";
            Parent.ChipCtrlButtons[6].Visible = true;
            Parent.ChipCtrlButtons[6].Click += Reset_PLL;

            Parent.ChipCtrlButtons[7].Text = "DOUT";
            Parent.ChipCtrlButtons[7].Visible = true;
            Parent.ChipCtrlButtons[7].Click += Read_DOUT;

            Parent.ChipCtrlButtons[8].Text = "TEST";
            Parent.ChipCtrlButtons[8].Visible = true;
            Parent.ChipCtrlButtons[8].Click += Set_TESTITEMS_MANUAL;

            Parent.ChipCtrlButtons[9].Text = "AUTO";
            Parent.ChipCtrlButtons[9].Visible = true;
            Parent.ChipCtrlButtons[9].Click += Set_TESTITEMS_AUTO;

            Parent.ChipCtrlButtons[10].Text = "TEMP";
            Parent.ChipCtrlButtons[10].Visible = true;
            Parent.ChipCtrlButtons[10].Click += Set_TESTITEMS_TEMP;

            Parent.ChipCtrlButtons[11].Text = "FW";
            Parent.ChipCtrlButtons[11].Visible = true;
            Parent.ChipCtrlButtons[11].Click += Set_TESTITEMS_FW;
        }

        private void Toogle_GPIO_GH0(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                string errorMsg = $"Error in Toggle GPIO_GH0: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Click_Button_Measure(object sender, EventArgs e)
            => Start_TS_Sweep();

        private void Start_TS_Sweep()
        {
            if (autotestWorker != null && autotestWorker.IsBusy)
            {
                MessageBox.Show("다른 작업이 이미 실행 중입니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            autotestWorker = new BackgroundWorker
            {
                WorkerSupportsCancellation = true
            };

            autotestWorker.DoWork += (sender, e) =>
            {
                Excute_TS_Sweep(sender as BackgroundWorker, e);
            };

            autotestWorker.RunWorkerCompleted += Generic_RunWorkerCompleted;

            autotestWorker.RunWorkerAsync();
        }

        private void Excute_TS_Sweep(object sender, EventArgs e)
        {
            try
            {
                Check_Instrument();

                UpdateRegisterBits(0xDC34_040C, 0xC0FF, (24u << 8));

                double[] abgrTrim = Start_ABGR_Trim(300);
                Log.WriteLine($"O_ABGR_CONT[3:0] = {abgrTrim[0]}, ABGR = {abgrTrim[1]:F5}");

                Set_GPIO4_TEMPSENSOR(true);
                for (uint j = 0; j < 16; j++)
                {
                    uint reg_DC34_009C = ReadRegister(0xDC34_009C);
                    WriteRegister(0xDC34_009C, (reg_DC34_009C & 0xFFE1) | (j << 1));
                    Thread.Sleep(100);

                    double dmm_volt = new double();
                    uint avgTime = 5;
                    for (int i = 0; i < avgTime; i++)
                        dmm_volt += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    Log.WriteLine($"O_TS_TRIM[3:0] = {j}\t{Math.Round(dmm_volt /= avgTime, 5)}");
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Start_TS_Sweep: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO4_TEMPSENSOR(false);
            }
        }

        private void Toogle_FT4222H_GPIO2(object sender, EventArgs e)
        {
            try
            {
                if (I2C.ft4222H != null)
                {
                    I2C.ft4222H.GPIO_SetDirection(2, GPIO_Direction.Output);

                    if (Parent.ChipCtrlButtons[3].Text == "IO2_H")
                    {
                        I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);
                        Parent.ChipCtrlButtons[3].Text = "IO2_L";
                    }
                    else
                    {
                        I2C.ft4222H.GPIO_SetState(2, GPIO_State.High);
                        Parent.ChipCtrlButtons[3].Text = "IO2_H";
                    }
                }
                else
                    throw new Exception("FT4222H is Null!");
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Toogle_FT4222H_GPIO2: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Reset_PLL(object sender, EventArgs e)
        {
            try
            {
                RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("w_PLL_PEN");

                PLL_PEN.Read();

                PLL_PEN.Value = 0;
                PLL_PEN.Write();

                PLL_PEN.Value = 1;
                PLL_PEN.Write();
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Toggle w_PLL_PEN: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PLL_TEST(object sender, EventArgs e)
        {
            Check_Instrument();

            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("w_PLL_PEN");
            RegisterItem VCO_TEST = Parent.RegMgr.GetRegisterItem("w_VCO_TEST");
            RegisterItem EXT_CAPS = Parent.RegMgr.GetRegisterItem("w_EXT_CAPS[9:0]");
            RegisterItem ABGR_CONT = Parent.RegMgr.GetRegisterItem("O_ABGR_CONT[3:0]");
            RegisterItem ALDO_CONT = Parent.RegMgr.GetRegisterItem("O_ALDO_CONT[5:0]");
            RegisterItem MLDO_CONT = Parent.RegMgr.GetRegisterItem("O_MLDO_CONT[5:0]");
            RegisterItem PLL_EN = Parent.RegMgr.GetRegisterItem("w_PLL_EN");
            RegisterItem PLL_OUT_EN = Parent.RegMgr.GetRegisterItem("w_PLL_OUT_EN");
            RegisterItem PLL_DIV2_OUT_EN = Parent.RegMgr.GetRegisterItem("w_PLL_DIV2_OUT_EN");
            RegisterItem TX_EN = Parent.RegMgr.GetRegisterItem("w_TX_EN");
            RegisterItem TX_BUF_EN = Parent.RegMgr.GetRegisterItem("w_TX_BUF_EN");
            RegisterItem TX_PRE_EN = Parent.RegMgr.GetRegisterItem("w_TX_PRE_EN");
            RegisterItem TX_DA_EN = Parent.RegMgr.GetRegisterItem("w_TX_DA_EN");
            RegisterItem DA_LDO_EN = Parent.RegMgr.GetRegisterItem("O_DA_LDO_EN");
            RegisterItem TX_PRE_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_TX_PRE_PEN_MODE");
            RegisterItem TX_DA_PEN_MODE = Parent.RegMgr.GetRegisterItem("w_TX_DA_PEN_MODE");
            RegisterItem TX_BUF_PEN = Parent.RegMgr.GetRegisterItem("r_TX_BUF_PEN");
            RegisterItem TX_PRE_PEN = Parent.RegMgr.GetRegisterItem("r_TX_PRE_PEN");
            RegisterItem TX_DA_PEN = Parent.RegMgr.GetRegisterItem("r_TX_DA_PEN");
            RegisterItem TRX_SEL = Parent.RegMgr.GetRegisterItem("r_TRX_SEL");
            RegisterItem SPI_RSV_CORE = Parent.RegMgr.GetRegisterItem("O_SPI_RSV_CORE[7:0]");
            RegisterItem AN_TEST_EN = Parent.RegMgr.GetRegisterItem("O_AN_TEST_EN");
            RegisterItem AN_TEST_MUX = Parent.RegMgr.GetRegisterItem("O_AN_TEST_MUX[2:0]");
            RegisterItem GPIO4_AN_EN = Parent.RegMgr.GetRegisterItem("O_GPIO4_AN_EN");
            RegisterItem GPIO_TEST_BUF_EN = Parent.RegMgr.GetRegisterItem("O_GPIO_TEST_BUF_EN");
            RegisterItem DA_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_DA_LDO_CONT[5:0]");

            VCO_TEST.Value = 1;
            VCO_TEST.Write();

            PLL_PEN.Value = 0;
            PLL_PEN.Write();

            PLL_PEN.Value = 1;
            PLL_PEN.Write();

            ABGR_CONT.Value = 7;
            ABGR_CONT.Write();

            ALDO_CONT.Value = 32;
            ALDO_CONT.Write();

            MLDO_CONT.Value = 16;
            MLDO_CONT.Write();

            AN_TEST_EN.Value = 1;
            AN_TEST_EN.Write();

            GPIO4_AN_EN.Value = 1;
            GPIO4_AN_EN.Write();

            AN_TEST_MUX.Value = 7;
            AN_TEST_MUX.Write();

            GPIO_TEST_BUF_EN.Value = 1;
            GPIO_TEST_BUF_EN.Write();

            DA_LDO_CONT.Value = 0;
            DA_LDO_CONT.Write();

            PLL_EN.Value = 1;
            PLL_EN.Write();

            PLL_OUT_EN.Value = 1;
            PLL_OUT_EN.Write();

            PLL_DIV2_OUT_EN.Value = 1;
            PLL_DIV2_OUT_EN.Write();

            TX_EN.Value = 1;
            TX_EN.Write();

            TX_BUF_EN.Value = 1;
            TX_BUF_EN.Write();

            TX_PRE_EN.Value = 1;
            TX_PRE_EN.Write();

            TX_DA_EN.Value = 1;
            TX_DA_EN.Write();

            DA_LDO_EN.Value = 1;
            DA_LDO_EN.Write();

            TX_PRE_PEN_MODE.Value = 1;
            TX_PRE_PEN_MODE.Write();

            TX_BUF_PEN.Value = 1;
            TX_BUF_PEN.Write();

            TX_PRE_PEN.Value = 1;
            TX_PRE_PEN.Write();

            TX_DA_PEN.Value = 1;
            TX_DA_PEN.Write();

            TRX_SEL.Value = 1;
            TRX_SEL.Write();

            SPI_RSV_CORE.Value = 1;
            SPI_RSV_CORE.Write();

            SpectrumAnalyzer.Write("DISP:TRAC:Y:RLEV 10dBm");
            SpectrumAnalyzer.Write("FREQ:CENT 2.5 GHz");
            SpectrumAnalyzer.Write("FREQ:SPAN 0.5 GHz");

            string time = DateTime.Now.ToString("HHmmss");
            Parent.xlMgr.Sheet.Add($"{time}_PLL_TEST");
            Parent.xlMgr.Cell.Write(1, 1, $"w_EXT_CAPS[9:0]");
            Parent.xlMgr.Cell.Write(2, 1, $"Freq[MHz]");

            for (uint i = 0; i < 1024; i++)
            {
                EXT_CAPS.Value = i;
                EXT_CAPS.Write();

                PLL_PEN.Value = 0;
                PLL_PEN.Write();
                System.Threading.Thread.Sleep(125);

                PLL_PEN.Value = 1;
                PLL_PEN.Write();
                System.Threading.Thread.Sleep(125);

                SpectrumAnalyzer.Write("CALC:MARK:MAX");
                double freqMHz = Math.Round(double.Parse(SpectrumAnalyzer.WriteAndReadString($"CALC:MARK1:X?")) / 1_000_000, 3);

                Parent.xlMgr.Cell.Write(1, 2 + (int)i, $"{EXT_CAPS.Value}");
                Parent.xlMgr.Cell.Write(2, 2 + (int)i, $"{freqMHz}");
            }
        }

        private void Read_DOUT(object sender, EventArgs e)
        {
            try
            {
                RegisterItem DOUT = Parent.RegMgr.GetRegisterItem("w_DOUT[15:0]");
                RegisterItem AVG_DOUT = Parent.RegMgr.GetRegisterItem("w_AVG_DOUT[15:0]");

                WriteRegister(0xDC34_0000, 0x0000_0001);
                Thread.Sleep(10);
                if (ReadRegister(0xDC34_0000) == 0x0000_0000)
                {
                    DOUT.Read();
                    AVG_DOUT.Read();
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Read DOUT: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Run_TEST_FUNCTION(object sender, EventArgs e)
        {

        }

        private void Set_PS0_OUTPUT(object sender, EventArgs e)
        {
            string userInput = RegContForm.Prompt.ShowDialog("Set to PowerSupply [V].", "Set_POWERSUPPLY");

            double voltageValue;
            if (double.TryParse(userInput, out voltageValue))
            {
                try
                {
                    Check_Instrument();

                    string scpiCommand = $"VOLT {voltageValue}";

                    PowerSupply0.Write(scpiCommand);
                }
                catch (Exception ex)
                {
                    string errorMsg = $"Error in Set_PowerSupply_Voltage_From_Input: {ex.Message}";
                    Log.WriteLine(errorMsg);
                    MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (string.IsNullOrEmpty(userInput))
            {
                
            }
            else
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 숫자만 입력하십시오.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Set_TESTITEMS_MANUAL(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.TEST;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }      

        private void Set_TESTITEMS_AUTO(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.AUTO;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)AUTO_TEST_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((AUTO_TEST_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Set_TESTITEMS_TEMP(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.CAL;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)CAL_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((CAL_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
        }

        private void Set_TESTITEMS_FW(object sender, EventArgs e)
        {
            CombBox_Item = COMBOBOX_ITEMS.FW;
            ComboBox_TestItems.Items.Clear();
            for (int i = 0; i < (int)FW_DN_ITEMS.NUM_TEST_ITEMS; i++)
                ComboBox_TestItems.Items.Add(((FW_DN_ITEMS)i).ToString());
            ComboBox_TestItems.SelectedIndex = 0;
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

        private void StartCommandProcessing(byte CommandReg, byte lsb, byte msb, int Timeout = 1000)
        {
            int sa;
            List<byte> SendBytes = new List<byte>();
            byte[] RcvBytes = new byte[1];

            sa = I2C.Config.SlaveAddress;
            I2C.Config.SlaveAddress = SlaveAddress;

            SendBytes.Add(CommandReg);
            SendBytes.Add(lsb);
            SendBytes.Add(msb);
            I2C.WriteBytes(SendBytes.ToArray(), SendBytes.Count, true);

            if (CommandReg != 0x0F) // TEST_ITEMS.FIRM_ON_CLEAR
            {
                for (int i = 0; i < (Timeout / 10); i++)
                {
                    System.Threading.Thread.Sleep(10);
                    RcvBytes = I2C.ReadBytes(1);
                    if (RcvBytes[0] == 0)
                        break;
                }
                //Log.WriteLine(RcvBytes[0].ToString());
            }

            I2C.Config.SlaveAddress = sa;
        }

        #region TEST_ITEMS
        private void Set_GPIO_Disable()
        {
            try
            {
                Set_GPIO4_ABGR(false);
                Set_GPIO16_32KOSC(false);
                Set_GPIO4_RETLDO(false);
                Set_GPIO4_MBGR(false);
                Set_GPIO4_DALDO(false);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO_Disable: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Set_GPIO4_ABGR(bool enable)
        {
            try
            {
                uint reg_DC34_0050 = ReadRegister(0xDC34_0050);     // O_AN_TEST_EN
                uint reg_DC34_0054 = ReadRegister(0xDC34_0054);     // O_AN_TEST_MUX[2:0] & O_GPIO4_AN_EN
                uint reg_DC34_006C = ReadRegister(0xDC34_006C);     // O_GPIO_TEST_BUF_EN

                if (enable)
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 | 15u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C | 1u << 15);
                    Thread.Sleep(10);
                }
                else
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 & 0xFFF0u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C & 0x7FFFu);
                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO4_ABGR: {ex.Message}";
                Log.WriteLine(errorMsg);
            }
        }

        private void Set_GPIO16_32KOSC(bool enable)
        {
            try
            {
                uint reg_5011_0014 = ReadRegister(0x5011_0014);     // OSEL_16
                uint reg_DC34_00A0 = ReadRegister(0xDC34_00A0);     // w_XO_RFC_CLK_EN

                if (enable)
                {
                    WriteRegister(0x5011_0014, (reg_5011_0014 & 0xFFFF_FF00u) | 69u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_00A0, reg_DC34_00A0 | 1u << 15);
                    Thread.Sleep(10);
                }
                else
                {
                    WriteRegister(0x5011_0014, (reg_5011_0014 & 0xFFFF_FF00u) | 16u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_00A0, reg_DC34_00A0 & 0x7FFFu);
                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO16_32KOSC: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Set_GPIO4_RETLDO(bool enable)
        {
            try
            {
                uint reg_DC34_0050 = ReadRegister(0xDC34_0050);     // O_AN_TEST_EN
                uint reg_DC34_0054 = ReadRegister(0xDC34_0054);     // O_AN_TEST_MUX[2:0] & O_GPIO4_AN_EN
                uint reg_DC34_006C = ReadRegister(0xDC34_006C);     // O_GPIO_TEST_BUF_EN
                uint reg_DC34_041C = ReadRegister(0xDC34_041C);     // w_PWR_SRAM_CRET_Static[4]
                uint reg_DC34_0434 = ReadRegister(0xDC34_0434);     // w_PWR_SRAM_MD[4]

                if (enable)
                {
                    WriteRegister(0xDC34_041C, reg_DC34_041C | 1u << 14);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0434, reg_DC34_0434 | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0050, reg_DC34_0050 | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 | 11u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C | 1u << 15);
                    Thread.Sleep(10);
                }
                else
                {
                    WriteRegister(0xDC34_041C, reg_DC34_041C & 0xBFFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0434, reg_DC34_0434 & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0050, reg_DC34_0050 & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 & 0xFFF0u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C & 0x7FFFu);
                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO4_RETLDO: {ex.Message}";
                Log.WriteLine(errorMsg);
            }

        }

        private void Set_GPIO4_MBGR(bool enable)
        {
            try
            {
                uint reg_DC34_0050 = ReadRegister(0xDC34_0050);     // O_AN_TEST_EN
                uint reg_DC34_0054 = ReadRegister(0xDC34_0054);     // O_AN_TEST_MUX[2:0] & O_GPIO4_AN_EN
                uint reg_DC34_006C = ReadRegister(0xDC34_006C);     // O_GPIO_TEST_BUF_EN
                uint reg_DC34_0080 = ReadRegister(0xDC34_0080);     // O_GPADC_PEN

                if (enable)
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 | 9u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0080, reg_DC34_0080 | 1u << 9);
                    Thread.Sleep(10);
                }
                else
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 & 0xFFF0u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0080, reg_DC34_0080 & 0xFDFFu);
                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO4_MBGR: {ex.Message}";
                Log.WriteLine(errorMsg);
            }
        }

        private void Set_GPIO4_DALDO(bool enable)
        {
            try
            {
                uint reg_DC34_0050 = ReadRegister(0xDC34_0050);     // O_AN_TEST_EN
                uint reg_DC34_0054 = ReadRegister(0xDC34_0054);     // O_AN_TEST_MUX[2:0] & O_GPIO4_AN_EN
                uint reg_DC34_006C = ReadRegister(0xDC34_006C);     // O_GPIO_TEST_BUF_EN
                uint reg_DC34_0080 = ReadRegister(0xDC34_0080);     // O_GPADC_PEN

                if (enable)
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 | 7u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0080, reg_DC34_0080 | 1u << 9);
                    Thread.Sleep(10);
                }
                else
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 & 0xFFF0u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0080, reg_DC34_0080 & 0xFDFFu);
                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO4_DALDO: {ex.Message}";
                Log.WriteLine(errorMsg);
            }
        }

        private void Set_GPIO4_TEMPSENSOR(bool enable)
        {
            try
            {
                uint reg_DC34_0050 = ReadRegister(0xDC34_0050);     // O_AN_TEST_EN
                uint reg_DC34_0054 = ReadRegister(0xDC34_0054);     // O_AN_TEST_MUX[2:0] & O_GPIO4_AN_EN
                uint reg_DC34_006C = ReadRegister(0xDC34_006C);     // O_GPIO_TEST_BUF_EN
                uint reg_DC34_0098 = ReadRegister(0xDC34_0098);     // O_TS_PEN

                if (enable)
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 | 13u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C | 1u << 15);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0098, reg_DC34_0098 | 1u);
                    Thread.Sleep(10);
                }
                else
                {
                    WriteRegister(0xDC34_0050, reg_DC34_0050 & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0054, reg_DC34_0054 & 0xFFF0u);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_006C, reg_DC34_006C & 0x7FFFu);
                    Thread.Sleep(10);

                    WriteRegister(0xDC34_0098, reg_DC34_0098 & 0xFFFEu);
                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Set_GPIO4_TEMPSENSOR: {ex.Message}";
                Log.WriteLine(errorMsg);
            }
        }
        #endregion TEST_ITEMS

        #region AUTO_TEST_ITEMS

        private void UpdateRegisterBits(uint address, uint mask, uint value)
        {
            uint currentValue = ReadRegister(address);
            Thread.Sleep(10);
            WriteRegister(address, (currentValue & mask) | value);
            Thread.Sleep(10);
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

        private void Set_TempChamber(BackgroundWorker worker, DoWorkEventArgs e, bool delay, double temp)
        {
            Check_Instrument();
            if (TempChamber.IsOpen == false)
            {
                throw new InvalidOperationException("Check TempChamber Connection!");
            }

            string tempchamber;
            string sTempVal;

            if (delay)
            {
                double[] dVal = new double[5];
                bool bTempReached = false;
                const double dTHVal = 0.1;

                Log.WriteLine($"Setting chamber to {temp} C and waiting for stabilization...");
                tempchamber = TempChamber.WriteAndReadString("01,TEMP,S" + temp);
                Thread.Sleep(1000); // 초기 명령 후 짧은 대기

                while (!bTempReached)
                {
                    // 루프 시작 시 취소 확인
                    if (worker.CancellationPending) { e.Cancel = true; return; }

                    sTempVal = TempChamber.WriteAndReadString("TEMP?");
                    Thread.Sleep(100);

                    string[] ArrBuf = sTempVal.Split(new char[] { ',' });

                    if (ArrBuf.Length < 4)
                    {
                        Log.WriteLine("Unexpected response from TempChamber: " + sTempVal);
                        ResponsiveSleep(worker, e, 5000); // 5초 대기 (취소 가능)
                        if (worker.CancellationPending) { e.Cancel = true; return; }
                        continue;
                    }

                    for (int split = 0; split < 4; split++)
                        dVal[split] = double.Parse(ArrBuf[split]);

                    Log.WriteLine("RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());

                    if (Math.Round(Math.Abs(dVal[0] - temp), 1) <= dTHVal)
                    {
                        Log.WriteLine("Temperature reached. Soaking for 60 seconds to confirm stability...");
                        ResponsiveSleep(worker, e, 60 * 1000); // 60초 대기 (취소 가능)
                        if (worker.CancellationPending) { e.Cancel = true; return; }

                        // 60초 후 재확인
                        sTempVal = TempChamber.WriteAndReadString("TEMP?");
                        Thread.Sleep(100);
                        ArrBuf = sTempVal.Split(new char[] { ',' });

                        if (ArrBuf.Length < 4) continue;
                        for (int split = 0; split < 4; split++)
                            dVal[split] = double.Parse(ArrBuf[split]);

                        if (Math.Round(Math.Abs(dVal[0] - temp), 1) <= dTHVal)
                        {
                            bTempReached = true;
                            Log.WriteLine("Done SetTemp! Temperature is stable.");
                        }
                        else
                        {
                            Log.WriteLine($"Temperature drifted (Real: {dVal[0]}). Resuming polling...");
                        }
                    }
                    else
                    {
                        // 온도가 아직 도달하지 않았을 때 대기
                        ResponsiveSleep(worker, e, 5000); // 5초 대기 (취소 가능)
                        if (worker.CancellationPending) { e.Cancel = true; return; }
                    }
                }

                Log.WriteLine("Waiting additional 60 seconds for final measurement soak...");
                ResponsiveSleep(worker, e, 60 * 1000); // 60초 최종 대기 (취소 가능)
                if (worker.CancellationPending) { e.Cancel = true; return; }

                sTempVal = TempChamber.WriteAndReadString("TEMP?");
                Thread.Sleep(100);
                string[] finalArrBuf = sTempVal.Split(new char[] { ',' });
                if (finalArrBuf.Length >= 2)
                {
                    dVal[0] = double.Parse(finalArrBuf[0]);
                    dVal[1] = double.Parse(finalArrBuf[1]);
                    Log.WriteLine("Final temp check -> RealTemp : " + dVal[0].ToString() + " | SetTemp : " + dVal[1].ToString());
                }
            }
            else if (!delay)
            {
                tempchamber = TempChamber.WriteAndReadString("01,TEMP,S" + temp);
                Log.WriteLine($"TempChamber : {tempchamber}");
            }
        }

        private void ResponsiveSleep(BackgroundWorker worker, DoWorkEventArgs e, int millisecondsToWait)
        {
            int interval = 1000; // 1초마다 확인
            int elapsed = 0;

            while (elapsed < millisecondsToWait)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                Thread.Sleep(interval);
                elapsed += interval;
            }
        }

        private void Run_ABGR_SWEEP(int chipIndex)
        {
            RegisterItem ABGR_CONT = Parent.RegMgr.GetRegisterItem("O_ABGR_CONT[3:0]");
            Check_Instrument();

            string time = DateTime.Now.ToString("HHmmss");
            Parent.xlMgr.Sheet.Add(time + "_SortResult");

            Parent.xlMgr.Cell.Write(1, 1, $"O_ABGR_CONT[3:0]");
            for (int j = 0; j < 16; j++)
            {
                Parent.xlMgr.Cell.Write(1, 2 + j, j.ToString());
            }

            double specMin = 300.0;
            double guardBand = 3.0;
            double testLimit = specMin + guardBand;

            string fwPath = @"G:\Sysapp\oasis\oasis_r3-sdk-20251118_loopcheck.bin";

            while (true)
            {
                DialogResult result = MessageBox.Show($"[{chipIndex + 1}번째 칩] 테스트를 진행하시겠습니까?",
                                                      "ABGR Sorting", MessageBoxButtons.YesNo);
                if (result == DialogResult.No) break;

                try
                {
                    int colVolt = chipIndex + 2;
                    Parent.xlMgr.Cell.Write(colVolt, 1, $"Chip{chipIndex + 1}_Volt");

                    bool isReady = false;
                    int retryCount = 0;
                    const int maxRetries = 3;

                    // --- [Retry Loop] 전원 On ~ 초기 전압 체크 ---
                    while (retryCount < maxRetries)
                    {
                        if (retryCount > 0)
                            Log.WriteLine($"[Retry {retryCount + 1}/{maxRetries}] Retrying Sequence...");

                        // 1. 계측기 설정 및 전원 인가
                        DigitalMultimeter2.Write("CONF:VOLT:DC 1V");
                        PowerSupply0.Write("VOLT 5.00, (@1)");
                        PowerSupply0.Write("VOLT 0.95, (@2)");
                        PowerSupply0.Write("CURR 0.05, (@1, 2)");

                        PowerSupply0.Write("OUTP ON, (@1, 2)");
                        Thread.Sleep(500);

                        // 2. 펌웨어 다운로드
                        if (!System.IO.File.Exists(fwPath))
                        {
                            MessageBox.Show("펌웨어 파일을 찾을 수 없습니다.\n" + fwPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            goto NEXT_CHIP; // 파일 없으면 해당 칩 스킵 (또는 전체 종료)
                        }

                        FirmwareName = fwPath;
                        FirmwareTarget = FW_TARGET.NV_MEM;
                        Oasis_DownloadFW();

                        if (!Status)
                        {
                            Log.WriteLine($"Chip {chipIndex + 1}: FW Download Failed (Try {retryCount + 1})");
                            // 펌웨어 실패 시에도 전원 끄고 재시도
                            PowerSupply0.Write("OUTP OFF, (@1, 2)");
                            Thread.Sleep(500);
                            retryCount++;
                            continue;
                        }

                        // 3. GPIO 설정
                        Set_GPIO4_ABGR(true);
                        Thread.Sleep(500);

                        // 4. 초기 전압 체크
                        double temp = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString("READ?")) * 1000, 5);

                        // 정상 범위: 200 초과 AND 400 미만 (즉, 200이하 OR 400이상은 Fail)
                        if (temp <= 200 || temp >= 400)
                        {
                            Log.WriteLine($"Init Volt Check Fail: {temp}mV (Try {retryCount + 1})");

                            // 실패 시 전원 끄고 재시도
                            Set_GPIO_Disable();
                            PowerSupply0.Write("OUTP OFF, (@1, 2)");
                            Thread.Sleep(500);
                            retryCount++;
                            continue;
                        }

                        // 성공 시
                        Log.WriteLine($"Init Volt Check Pass: {temp}mV");
                        isReady = true;
                        break; // Retry 루프 탈출 -> Sweep 진행
                    }

                    // 3회 시도 후에도 실패 시
                    if (!isReady)
                    {
                        Log.WriteLine($"Chip {chipIndex + 1}: Init Failed after {maxRetries} retries.");
                        Parent.xlMgr.Cell.Write(colVolt, 18, "INIT_FAIL");
                        MessageBox.Show("초기화(전압 체크) 3회 실패.\n다음 칩을 준비해주세요.", "Fail", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        goto NEXT_CHIP;
                    }

                    // --- Sweep Sequence ---
                    bool isPass = false;

                    for (int j = 0; j < 16; j++)
                    {
                        ABGR_CONT.Read();
                        ABGR_CONT.Value = (uint)j;
                        ABGR_CONT.Write();
                        Thread.Sleep(250);

                        double dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString("READ?")) * 1000, 5);

                        Parent.xlMgr.Cell.Write(colVolt, 2 + j, dmm_volt.ToString("F5"));

                        if (dmm_volt >= testLimit) isPass = true;
                    }

                    string resStr = isPass ? "PASS" : "FAIL";
                    Parent.xlMgr.Cell.Write(colVolt, 18, resStr);

                    if (isPass)
                        Log.WriteLine($"Chip {chipIndex + 1}: PASS");
                    else
                    {
                        Log.WriteLine($"Chip {chipIndex + 1}: FAIL");
                        //MessageBox.Show($"Chip {chipIndex + 1} 불량 (Fail)", "Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                NEXT_CHIP:
                    chipIndex++;
                }
                catch (Exception ex)
                {
                    string errorMsg = $"Error: {ex.Message}";
                    Log.WriteLine(errorMsg);
                    MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    Set_GPIO_Disable();
                    PowerSupply0.Write("OUTP OFF, (@1, 2)");
                    Thread.Sleep(500);
                }
            }
        }

        private void Run_ALDO_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                uint[] aldo_cont = {
                    24, 25, 26, 27, 28, 29, 30, 31, 16, 17, 18, 19, 20, 21, 22, 23,
                    56, 57, 58, 59, 60, 61, 62, 63, 48, 49, 50, 51, 52, 53, 54, 55,
                    8, 9, 10, 11, 12, 13, 14, 15, 0, 1, 2, 3, 4, 5, 6, 7,
                    40, 41, 42, 43, 44, 45, 46, 47, 32, 33, 34, 35, 36, 37, 38, 39
                };
                RegisterItem ALDO_CONT = Parent.RegMgr.GetRegisterItem("O_ALDO_CONT[5:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_ALDO_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_ALDO_CONT[5:0]");
                Parent.xlMgr.Cell.Write(2, 1, "ALDO[㎷]");
                for (int i = 0; i < aldo_cont.Length; i++)
                {
                    ALDO_CONT.Read();
                    ALDO_CONT.Value = aldo_cont[i];
                    ALDO_CONT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(2, 2 + (int)aldo_cont[i], dmm_volt.ToString("F3"));
                    Parent.xlMgr.Cell.Write(1, 2 + (int)aldo_cont[i], aldo_cont[i].ToString(""));
                }
                Log.WriteLine("ALDO Sweep completed successfully.");
                MessageBox.Show("ALDO Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_ALDO_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_32KOSC_SWEEP()
        {
            string userInput = Prompt.ShowDialog("RTC_SCKF[10:6] 값을 입력하십시오 (0-31):", "32K OSC Sweep");
            uint highBits_Raw;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 32KOSC SWEEP 입력을 취소했습니다.");
                return;
            }
            if (!uint.TryParse(userInput, out highBits_Raw) || highBits_Raw > 31)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0에서 31 사이의 숫자만 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                RegisterItem RTC_SCKF = Parent.RegMgr.GetRegisterItem("O_RTC_SCKF[10:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_32KOSC_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "10:6 BIN");
                Parent.xlMgr.Cell.Write(2, 1, "10:6 DEC");
                Parent.xlMgr.Cell.Write(3, 1, "5:0 BIN");
                Parent.xlMgr.Cell.Write(4, 1, "5:0 DEC");
                Parent.xlMgr.Cell.Write(5, 1, "10:0 DEC");
                Parent.xlMgr.Cell.Write(6, 1, "Freq[KHz]");
                Set_GPIO16_32KOSC(true);
                Thread.Sleep(250);

                int numSteps = 64;
                for (int idx = 0; idx < numSteps; idx++)
                {
                    uint lowBits_Raw = (uint)idx;
                    uint regValue = (highBits_Raw << 6) | lowBits_Raw;
                    RTC_SCKF.Read();
                    RTC_SCKF.Value = regValue;
                    RTC_SCKF.Write();
                    Thread.Sleep(250);
                    double dmm_freq = new double();
                    dmm_freq = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:FREQ?")) / 1000, 3);
                    int col = 2 + idx;
                    Parent.xlMgr.Cell.Write(1, col, System.Convert.ToString(highBits_Raw, 2).PadLeft(5, '0'));
                    Parent.xlMgr.Cell.Write(2, col, highBits_Raw.ToString(""));
                    Parent.xlMgr.Cell.Write(3, col, System.Convert.ToString(lowBits_Raw, 2).PadLeft(6, '0'));
                    Parent.xlMgr.Cell.Write(4, col, lowBits_Raw.ToString(""));
                    Parent.xlMgr.Cell.Write(5, col, regValue.ToString(""));
                    Parent.xlMgr.Cell.Write(6, col, dmm_freq.ToString("F3"));
                }
                Log.WriteLine("32KOSC Sweep completed successfully.");
                MessageBox.Show("32KOSC Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_32KOSC_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_DCDC_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                RegisterItem DCDC_SET_VOUT = Parent.RegMgr.GetRegisterItem("O_DCDC_SET_VOUT[5:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_DCDC_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_DCDC_SET_VOUT[5:0]");
                Parent.xlMgr.Cell.Write(2, 1, "DCDC_VOUT[V]");
                int startVal = 0;
                int numSteps = 64;
                for (int idx = 0; idx < numSteps; idx++)
                {
                    uint regValue = (uint)(startVal + idx);
                    DCDC_SET_VOUT.Read();
                    DCDC_SET_VOUT.Value = regValue;
                    DCDC_SET_VOUT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")), 4);
                    Parent.xlMgr.Cell.Write(2, 2 + idx, dmm_volt.ToString("F4"));
                    Parent.xlMgr.Cell.Write(1, 2 + idx, regValue.ToString(""));
                }
                Log.WriteLine("DCDC Sweep completed successfully.");
                MessageBox.Show("DCDC Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_DCDC_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_MLDO_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                uint[] mldo_cont = {
                    8, 9, 10, 11, 12, 13, 14, 15, 0, 1, 2, 3, 4, 5, 6, 7,
                    24, 25, 26, 27, 28, 29, 30, 31, 16, 17, 18, 19, 20, 21,
                    22, 23, 40, 41, 42, 43, 44, 45, 46, 47, 32, 33, 34, 35,
                    36, 37, 38, 39, 56, 57, 58, 59, 60, 61, 62, 63, 48, 49,
                    50, 51, 52, 53, 54, 55
                };
                RegisterItem MLDO_CONT = Parent.RegMgr.GetRegisterItem("O_MLDO_CONT[5:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_MLDO_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_MLDO_CONT[5:0]");
                Parent.xlMgr.Cell.Write(2, 1, "MLDO[㎷]");
                for (int i = 0; i < mldo_cont.Length; i++)
                {
                    MLDO_CONT.Read();
                    MLDO_CONT.Value = mldo_cont[i];
                    MLDO_CONT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(2, 2 + (int)mldo_cont[i], dmm_volt.ToString("F3"));
                    Parent.xlMgr.Cell.Write(1, 2 + (int)mldo_cont[i], mldo_cont[i].ToString(""));
                }
                Log.WriteLine("MLDO Sweep completed successfully.");
                MessageBox.Show("MLDO Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_MLDO_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_DALDO_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                RegisterItem DA_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_DA_LDO_CONT[5:0]");
                Check_Instrument();

                Parent.xlMgr.Sheet.Add(time + "_DA_LDO_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_DA_LDO_CONT[5:0]");
                Parent.xlMgr.Cell.Write(2, 1, "DA_LDO[㎷]");
                Set_GPIO4_DALDO(true);
                Thread.Sleep(250);

                for (int j = 0; j < 64; j++)
                {
                    DA_LDO_CONT.Read();
                    DA_LDO_CONT.Value = (uint)j;
                    DA_LDO_CONT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(1, 2 + j, j.ToString(""));
                    Parent.xlMgr.Cell.Write(2, 2 + j, dmm_volt.ToString("F3"));
                }
                Log.WriteLine("DALDO Sweep completed successfully.");
                MessageBox.Show("DALDO Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_DALDO_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_RETLDO_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                RegisterItem RET_LDO_CONT = Parent.RegMgr.GetRegisterItem("O_RET_LDO_CONT[3:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_RETLDO_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_RET_LDO_CONT[3:0]");
                Parent.xlMgr.Cell.Write(2, 1, "RET_LDO[㎷]");
                Set_GPIO4_RETLDO(true);
                Thread.Sleep(250);
                for (int j = 0; j < 16; j++)
                {
                    RET_LDO_CONT.Read();
                    RET_LDO_CONT.Value = (uint)j;
                    RET_LDO_CONT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(1, 2 + j, j.ToString(""));
                    Parent.xlMgr.Cell.Write(2, 2 + j, dmm_volt.ToString("F3"));
                }
                Log.WriteLine("RETLDO Sweep completed successfully.");
                MessageBox.Show("RETLDO Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_RETLDO_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_MBGR_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                RegisterItem MBGR_CONT = Parent.RegMgr.GetRegisterItem("O_MBGR_OUT_TRIM[5:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_MBGR_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_MBGR_OUT_TRIM[5:0]");
                Parent.xlMgr.Cell.Write(2, 1, "MBGR[㎷]");
                Set_GPIO4_MBGR(true);
                Thread.Sleep(250);
                for (int j = 0; j < 64; j++)
                {
                    MBGR_CONT.Read();
                    MBGR_CONT.Value = (uint)j;
                    MBGR_CONT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(1, 2 + j, j.ToString(""));
                    Parent.xlMgr.Cell.Write(2, 2 + j, dmm_volt.ToString("F3"));
                }
                Log.WriteLine("MBGR Sweep completed successfully.");
                MessageBox.Show("MBGR Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_MBGR_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Set_GPIO_Disable();
            }
        }

        private void Run_FLDO_SWEEP()
        {
            try
            {
                string time = DateTime.Now.ToString("HHmmss");
                RegisterItem FLDO_CONT = Parent.RegMgr.GetRegisterItem("O_FLDO_CONT[3:0]");
                Check_Instrument();
                Parent.xlMgr.Sheet.Add(time + "_FLDO_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "O_FLDO_CONT[3:0]");
                Parent.xlMgr.Cell.Write(2, 1, "FLDO[㎷]");
                for (int j = 0; j < 16; j++)
                {
                    FLDO_CONT.Read();
                    FLDO_CONT.Value = (uint)j;
                    FLDO_CONT.Write();
                    Thread.Sleep(250);
                    double dmm_volt = new double();
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:VOLT:DC?")), 5);
                    Parent.xlMgr.Cell.Write(1, 2 + j, j.ToString(""));
                    Parent.xlMgr.Cell.Write(2, 2 + j, dmm_volt.ToString("F5"));
                }
                Log.WriteLine("FLDO Sweep completed successfully.");
                MessageBox.Show("FLDO Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Run_FLDO_SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //Set_GPIO_Disable();
            }
        }

        private void Run_GPADC_SWEEP()
        {
            string userInput = Prompt.ShowDialog("PowerSupply의 GPIO01 Input Port를 선택해주세요. (1-3):", "GPADC Sweep Test");
            uint psChannel;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!uint.TryParse(userInput, out psChannel) || psChannel < 1 || psChannel > 3)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 1 ~ 3 사이의 숫자만 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            userInput = Prompt.ShowDialog("Max Voltage를 입력해주세요. (0 초과):", "GPADC Sweep Test");
            double maxVoltage;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!double.TryParse(userInput, out maxVoltage) || maxVoltage <= 0)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 큰 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            userInput = Prompt.ShowDialog("Voltage Step을 입력해주세요. (0.001 ~ 1V):", "GPADC Sweep Test");
            double stepVoltage;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!double.TryParse(userInput, out stepVoltage) || stepVoltage < 0.001 || stepVoltage > 1)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0.001 ~ 1 사이의 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            RegisterItem GPIO01_AN_EN = null;
            RegisterItem AN_TEST_MUX = null;
            RegisterItem GPIO_TEST_BUF_EN = null;
            RegisterItem DOUT = null;
            RegisterItem AVG_DOUT = null;

            int sweepSteps = (int)Math.Ceiling(maxVoltage / stepVoltage) + 1;

            try
            {
                Check_Instrument();
                string time = DateTime.Now.ToString("HHmmss");
                Parent.xlMgr.Sheet.Add(time + "_GPADC_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "Input[V]");
                Parent.xlMgr.Cell.Write(2, 1, "Measure[V]");
                Parent.xlMgr.Cell.Write(3, 1, "DOUT");
                Parent.xlMgr.Cell.Write(4, 1, "AVG_DOUT");

                GPIO01_AN_EN = Parent.RegMgr.GetRegisterItem("O_GPIO01_AN_EN");
                AN_TEST_MUX = Parent.RegMgr.GetRegisterItem("O_AN_TEST_MUX[2:0]");
                GPIO_TEST_BUF_EN = Parent.RegMgr.GetRegisterItem("O_GPIO_TEST_BUF_EN");
                DOUT = Parent.RegMgr.GetRegisterItem("w_DOUT[15:0]");
                AVG_DOUT = Parent.RegMgr.GetRegisterItem("w_AVG_DOUT[15:0]");

                GPIO01_AN_EN.Value = 1;
                GPIO01_AN_EN.Write();

                AN_TEST_MUX.Value = 4;
                AN_TEST_MUX.Write();

                GPIO_TEST_BUF_EN.Value = 0;
                GPIO_TEST_BUF_EN.Write();

                PowerSupply0.Write($"INST:NSEL {psChannel}");
                PowerSupply0.Write("OUTP ON");

                for (int i = 0; i < sweepSteps; i++)
                {
                    double inputVoltage = i * stepVoltage;
                    if (inputVoltage > maxVoltage)
                    {
                        inputVoltage = maxVoltage;
                    }

                    uint doutVal = 0;
                    uint avgDoutVal = 0;

                    PowerSupply0.Write($"VOLT {inputVoltage:F3}");
                    Thread.Sleep(50);

                    double measuredVoltage = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));

                    WriteRegister(0xDC340000, 0x00000001);
                    Thread.Sleep(10);
                    if (ReadRegister(0xDC340000) == 0x00000000)
                    {
                        doutVal = DOUT.Read();
                        avgDoutVal = AVG_DOUT.Read();
                    }

                    int col = 2 + i;
                    Parent.xlMgr.Cell.Write(1, col, inputVoltage.ToString("F5"));
                    Parent.xlMgr.Cell.Write(2, col, measuredVoltage.ToString("F5"));
                    Parent.xlMgr.Cell.Write(3, col, doutVal.ToString());
                    Parent.xlMgr.Cell.Write(4, col, avgDoutVal.ToString());

                    if (inputVoltage == maxVoltage)
                    {
                        break;
                    }
                }

                Log.WriteLine("GPADC Sweep completed successfully.");
                MessageBox.Show("GPADC Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in GPADC SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (PowerSupply0 != null)
                {
                    PowerSupply0.Write($"INST:NSEL {psChannel}");
                    PowerSupply0.Write("VOLT 0.0");
                }
                if (GPIO01_AN_EN != null)
                {
                    GPIO01_AN_EN.Value = 0;
                    GPIO01_AN_EN.Write();
                }
                if (AN_TEST_MUX != null)
                {
                    AN_TEST_MUX.Value = 0;
                    AN_TEST_MUX.Write();
                }
            }
        }

        private void Run_GPADC_REPEAT()
        {
            string userInput = Prompt.ShowDialog("PowerSupply의 GPIO01 Input Port를 선택해주세요. (1-3):", "GPADC Repeated Test");
            uint psChannel;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!uint.TryParse(userInput, out psChannel) || psChannel < 1 || psChannel > 3)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 1 ~ 3 사이의 숫자만 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            userInput = Prompt.ShowDialog("Voltage를 입력해주세요. (0 초과):", "GPADC Repeated Test");
            double setVoltage;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!double.TryParse(userInput, out setVoltage) || setVoltage <= 0)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 큰 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            userInput = Prompt.ShowDialog("반복 횟수를 입력해주세요.:", "GPADC Repeated Test");
            uint numRepetition;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!uint.TryParse(userInput, out numRepetition) || numRepetition <= 0)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 큰 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            RegisterItem GPIO01_AN_EN = null;
            RegisterItem AN_TEST_MUX = null;
            RegisterItem GPIO_TEST_BUF_EN = null;
            RegisterItem DOUT = null;
            RegisterItem AVG_DOUT = null;

            try
            {
                Check_Instrument();
                string time = DateTime.Now.ToString("HHmmss");
                Parent.xlMgr.Sheet.Add(time + "_GPADC_REPEAT");
                Parent.xlMgr.Cell.Write(1, 1, "Input[V]");
                Parent.xlMgr.Cell.Write(2, 1, "Measure[V]");
                Parent.xlMgr.Cell.Write(3, 1, "DOUT");
                Parent.xlMgr.Cell.Write(4, 1, "AVG_DOUT");

                GPIO01_AN_EN = Parent.RegMgr.GetRegisterItem("O_GPIO01_AN_EN");
                AN_TEST_MUX = Parent.RegMgr.GetRegisterItem("O_AN_TEST_MUX[2:0]");
                GPIO_TEST_BUF_EN = Parent.RegMgr.GetRegisterItem("O_GPIO_TEST_BUF_EN");
                DOUT = Parent.RegMgr.GetRegisterItem("w_DOUT[15:0]");
                AVG_DOUT = Parent.RegMgr.GetRegisterItem("w_AVG_DOUT[15:0]");

                GPIO01_AN_EN.Value = 1;
                GPIO01_AN_EN.Write();

                AN_TEST_MUX.Value = 4;
                AN_TEST_MUX.Write();

                GPIO_TEST_BUF_EN.Value = 0;
                GPIO_TEST_BUF_EN.Write();

                PowerSupply0.Write($"INST:NSEL {psChannel}");
                PowerSupply0.Write("OUTP ON");

                for (int i = 0; i < numRepetition; i++)
                {
                    uint doutVal = 0;
                    uint avgDoutVal = 0;

                    PowerSupply0.Write($"VOLT {setVoltage:F3}");
                    Thread.Sleep(50);

                    double measuredVoltage = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));

                    WriteRegister(0xDC340000, 0x00000001);
                    Thread.Sleep(10);
                    if (ReadRegister(0xDC340000) == 0x00000000)
                    {
                        doutVal = DOUT.Read();
                        avgDoutVal = AVG_DOUT.Read();
                    }

                    int col = 2 + i;
                    Parent.xlMgr.Cell.Write(1, col, setVoltage.ToString("F5"));
                    Parent.xlMgr.Cell.Write(2, col, measuredVoltage.ToString("F5"));
                    Parent.xlMgr.Cell.Write(3, col, doutVal.ToString());
                    Parent.xlMgr.Cell.Write(4, col, avgDoutVal.ToString());
                }

                Log.WriteLine("GPADC Repeated Test completed successfully.");
                MessageBox.Show("GPADC Repeated Test completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in GPADC REPEAT: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (PowerSupply0 != null)
                {
                    PowerSupply0.Write($"INST:NSEL {psChannel}");
                    PowerSupply0.Write("VOLT 0.0");
                }
                if (GPIO01_AN_EN != null)
                {
                    GPIO01_AN_EN.Value = 0;
                    GPIO01_AN_EN.Write();
                }
                if (AN_TEST_MUX != null)
                {
                    AN_TEST_MUX.Value = 0;
                    AN_TEST_MUX.Write();
                }
            }
        }

        private void Run_SNIFFER_SWEEP()
        {
            string userInput = Prompt.ShowDialog("최소 전압(V)을 입력해주세요.:", "SNIFFER SWEEP Test");
            double minVoltage;
            if (string.IsNullOrEmpty(userInput) || !double.TryParse(userInput, out minVoltage) || minVoltage < 0)
            {
                Log.WriteLine("사용자가 입력을 취소했거나 잘못된 최소 전압을 입력했습니다.");
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0 이상의 숫자를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            userInput = Prompt.ShowDialog("최대 전압(V)을 입력해주세요.:", "SNIFFER SWEEP Test");
            double maxVoltage;
            if (string.IsNullOrEmpty(userInput) || !double.TryParse(userInput, out maxVoltage) || maxVoltage <= minVoltage)
            {
                Log.WriteLine("사용자가 입력을 취소했거나 잘못된 최대 전압을 입력했습니다.");
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 최소 전압({minVoltage}V)보다 큰 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            userInput = Prompt.ShowDialog("스텝 전압(V)을 입력해주세요.:", "SNIFFER SWEEP Test");
            double stepVoltage;
            if (string.IsNullOrEmpty(userInput) || !double.TryParse(userInput, out stepVoltage) || stepVoltage <= 0)
            {
                Log.WriteLine("사용자가 입력을 취소했거나 잘못된 스텝 전압을 입력했습니다.");
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 큰 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Check_Instrument();
                string time = DateTime.Now.ToString("HHmmss");
                Parent.xlMgr.Sheet.Add(time + "_SNIFFER_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "Input @2[V]");
                Parent.xlMgr.Cell.Write(2, 1, "Input @3[V]");
                Parent.xlMgr.Cell.Write(3, 1, "Measure @2[V]");
                Parent.xlMgr.Cell.Write(4, 1, "Measure @3[V]");
                Parent.xlMgr.Cell.Write(5, 1, "i1 (binary)");
                Parent.xlMgr.Cell.Write(6, 1, "q1 (binary)");
                Parent.xlMgr.Cell.Write(7, 1, "i2 (binary)");
                Parent.xlMgr.Cell.Write(8, 1, "q2 (binary)");

                uint[] rdat = new uint[4];

                PowerSupply0.Write("OUTP ON,(@2,3)");

                decimal dMin = (decimal)minVoltage;
                decimal dMax = (decimal)maxVoltage;
                decimal dStep = (decimal)stepVoltage;
                int col = 2;

                for (decimal v = 0; ; v += dStep)
                {
                    decimal currentV_P2_dec = dMax - v;
                    decimal currentV_P3_dec = dMin + v;

                    if (currentV_P2_dec < dMin) currentV_P2_dec = dMin;
                    if (currentV_P3_dec > dMax) currentV_P3_dec = dMax;

                    double currentV_P2 = (double)currentV_P2_dec;
                    double currentV_P3 = (double)currentV_P3_dec;

                    PowerSupply0.Write($"VOLT {currentV_P2:F3},(@2)");
                    PowerSupply0.Write($"VOLT {currentV_P3:F3},(@3)");
                    Thread.Sleep(100);

                    double measuredVoltage0 = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
                    double measuredVoltage1 = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?"));

                    uint temp = ReadRegister(0xDC34_060C);
                    for (int j = 0; j < 4; j++)
                    {
                        rdat[j] = (temp >> (8 * j)) & 0xff;
                    }
                    Thread.Sleep(50);

                    Parent.xlMgr.Cell.Write(1, col, currentV_P2.ToString("F5"));
                    Parent.xlMgr.Cell.Write(2, col, currentV_P3.ToString("F5"));
                    Parent.xlMgr.Cell.Write(3, col, measuredVoltage0.ToString("F5"));
                    Parent.xlMgr.Cell.Write(4, col, measuredVoltage1.ToString("F5"));

                    Parent.xlMgr.Cell.Write(5, col, System.Convert.ToString(rdat[3], 2).PadLeft(8, '0'));
                    Parent.xlMgr.Cell.Write(6, col, System.Convert.ToString(rdat[2], 2).PadLeft(8, '0'));
                    Parent.xlMgr.Cell.Write(7, col, System.Convert.ToString(rdat[1], 2).PadLeft(8, '0'));
                    Parent.xlMgr.Cell.Write(8, col, System.Convert.ToString(rdat[0], 2).PadLeft(8, '0'));

                    col++;

                    if (currentV_P2_dec == dMin || currentV_P3_dec == dMax)
                        break;
                }

                Log.WriteLine("Sniffer Sweep Test completed successfully.");
                MessageBox.Show("Sniffer Sweep Test completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Sniffer SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (PowerSupply0 != null)
                {
                    PowerSupply0.Write("VOLT 0.0,(@2,3)");
                    PowerSupply0.Write("OUTP OFF,(@2,3)");
                }
            }
        }

        private void Run_SNIFFER_REPEAT()
        {
            string userInput = Prompt.ShowDialog("반복 횟수를 입력해주세요.:", "SNIFFER Repeated Test");
            uint numRepetition;
            if (string.IsNullOrEmpty(userInput))
            {
                Log.WriteLine("사용자가 입력을 취소했습니다.");
                return;
            }
            if (!uint.TryParse(userInput, out numRepetition) || numRepetition <= 0)
            {
                MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 큰 수를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Check_Instrument();
                string time = DateTime.Now.ToString("HHmmss");
                Parent.xlMgr.Sheet.Add(time + "_SNIFFER_REPEAT");

                Parent.xlMgr.Cell.Write(1, 1, "Input @2[V]");
                Parent.xlMgr.Cell.Write(2, 1, "Input @3[V]");
                Parent.xlMgr.Cell.Write(3, 1, "Measure @2[V]");
                Parent.xlMgr.Cell.Write(4, 1, "Measure @3[V]");
                Parent.xlMgr.Cell.Write(5, 1, "i1 (binary)");
                Parent.xlMgr.Cell.Write(6, 1, "q1 (binary)");
                Parent.xlMgr.Cell.Write(7, 1, "i2 (binary)");
                Parent.xlMgr.Cell.Write(8, 1, "q2 (binary)");

                uint[] rdat = new uint[4];

                for (int i = 0; i < numRepetition; i++)
                {
                    double measuredVoltage0_DMM = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
                    double measuredVoltage1_DMM = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?"));

                    double measuredVoltage0_PS = double.Parse(PowerSupply0.WriteAndReadString(":MEAS:VOLT:DC? (@2)"));
                    double measuredVoltage1_PS = double.Parse(PowerSupply0.WriteAndReadString(":MEAS:VOLT:DC? (@3)"));

                    uint temp = ReadRegister(0xDC34_060C);
                    for (int j = 0; j < 4; j++)
                    {
                        rdat[j] = (temp >> (8 * j)) & 0xff;
                    }
                    Thread.Sleep(50);

                    int col = 2 + i;

                    Parent.xlMgr.Cell.Write(1, col, measuredVoltage0_PS.ToString("F5")); // PS0 P2
                    Parent.xlMgr.Cell.Write(2, col, measuredVoltage1_PS.ToString("F5")); // PS0 P3
                    Parent.xlMgr.Cell.Write(3, col, measuredVoltage0_DMM.ToString("F5")); // DMM0
                    Parent.xlMgr.Cell.Write(4, col, measuredVoltage1_DMM.ToString("F5")); // DMM1

                    Parent.xlMgr.Cell.Write(5, col, System.Convert.ToString(rdat[3], 2).PadLeft(8, '0'));
                    Parent.xlMgr.Cell.Write(6, col, System.Convert.ToString(rdat[2], 2).PadLeft(8, '0'));
                    Parent.xlMgr.Cell.Write(7, col, System.Convert.ToString(rdat[1], 2).PadLeft(8, '0'));
                    Parent.xlMgr.Cell.Write(8, col, System.Convert.ToString(rdat[0], 2).PadLeft(8, '0'));
                }

                Log.WriteLine("Sniffer Repeated Test completed successfully.");
                MessageBox.Show("Sniffer Repeated Test completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Sniffer REPEAT: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {

            }
        }

        private void Run_ANA_TEST_BUFF_SWEEP()
        {
            try
            {
                Check_Instrument();
                double dmm_volt = new double();
                string time = DateTime.Now.ToString("HHmmss");
                int x_offset = new int();

                Parent.xlMgr.Sheet.Add(time + "_GPIO4_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "VBAT");
                Parent.xlMgr.Cell.Write(2, 1, "ABGR[㎷]");
                Parent.xlMgr.Cell.Write(3, 1, "MBGR[㎷]");
                Parent.xlMgr.Cell.Write(4, 1, "RET_LDO[㎷]");
                Parent.xlMgr.Cell.Write(5, 1, "DA_LDO[㎷]");

                for (double i = 1.7; i < 3.7; i += 0.1)
                {
                    PowerSupply0.Write($"VOLT {i}, (@1)");
                    Parent.xlMgr.Cell.Write(1, 2 + x_offset, i.ToString("F1"));

                    Thread.Sleep(500);

                    Set_GPIO4_ABGR(true);
                    Thread.Sleep(250);
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(2, 2 + x_offset, dmm_volt.ToString("F3"));
                    Set_GPIO_Disable();
                    Thread.Sleep(250);

                    Set_GPIO4_MBGR(true);
                    Thread.Sleep(250);
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(3, 2 + x_offset, dmm_volt.ToString("F3"));
                    Set_GPIO_Disable();
                    Thread.Sleep(250);

                    Set_GPIO4_RETLDO(true);
                    Thread.Sleep(250);
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(4, 2 + x_offset, dmm_volt.ToString("F3"));
                    Set_GPIO_Disable();
                    Thread.Sleep(250);

                    Set_GPIO4_DALDO(true);
                    Thread.Sleep(250);
                    dmm_volt = Math.Round(double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(5, 2 + x_offset, dmm_volt.ToString("F3"));
                    Set_GPIO_Disable();
                    Thread.Sleep(250);

                    x_offset++;
                }

            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in ANA TEST BUFF SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Run_DMM_SWEEP()
        {
            try
            {
                string userInput = Prompt.ShowDialog("측정할 DigitalMultimeter 번호를 선택해주세요. (0~3):", "Run_DMM_SWEEP");
                uint dmmSel;
                if (string.IsNullOrEmpty(userInput))
                {
                    Log.WriteLine("사용자가 입력을 취소했습니다.");
                    return;
                }
                if (!uint.TryParse(userInput, out dmmSel) || dmmSel < 0 || dmmSel > 3)
                {
                    MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 1 ~ 3 사이의 숫자만 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Check_Instrument();
                double dmm_volt = new double();
                string time = DateTime.Now.ToString("HHmmss");
                int x_offset = new int();
                JLcLib.Instrument.SCPI[] DMM = { DigitalMultimeter0, DigitalMultimeter1, DigitalMultimeter2, DigitalMultimeter3 };

                Parent.xlMgr.Sheet.Add(time + "_DMM_SWEEP");
                Parent.xlMgr.Cell.Write(1, 1, "VBAT");
                Parent.xlMgr.Cell.Write(2, 1, "DMM[㎷]");

                for (double i = 1.7; i < 3.7; i += 0.1)
                {
                    PowerSupply0.Write($"VOLT {i}, (@1)");
                    Parent.xlMgr.Cell.Write(1, 2 + x_offset, i.ToString("F1"));
                    Thread.Sleep(500);

                    dmm_volt = Math.Round(double.Parse(DMM[dmmSel].WriteAndReadString(":MEAS:VOLT:DC?")) * 1000, 3);
                    Parent.xlMgr.Cell.Write(2, 2 + x_offset, dmm_volt.ToString("F3"));

                    x_offset++;
                }

            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in DMM SWEEP: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private double[] Create_TEMP_ARRAY(double start, double stop, double step)
        {
            List<double> tempBuffer = new List<double>();

            decimal dStart = (decimal)start;
            decimal dStop = (decimal)stop;
            decimal dStep = Math.Abs((decimal)step); // Step은 크기로만 취급

            if (dStep == 0) return new double[] { start };

            if (dStart <= dStop)
            {
                for (decimal i = dStart; i <= dStop; i += dStep)
                {
                    tempBuffer.Add((double)i);
                }
            }
            else
            {
                for (decimal i = dStart; i >= dStop; i -= dStep)
                {
                    tempBuffer.Add((double)i);
                }
            }

            return tempBuffer.ToArray();
        }

        private void Generic_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                Log.WriteLine("Sweep process cancelled by user.");
                MessageBox.Show("Sweep process cancelled.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (e.Error != null)
            {
                string errorMsg = e.Error.InnerException != null ? e.Error.InnerException.Message : e.Error.Message;
                string errorTitle = e.Error.InnerException != null ? "Sweep Runtime Error" : "BackgroundWorker Error";

                Log.WriteLine($"BackgroundWorker Error: {errorMsg}\n{e.Error.StackTrace}");
                MessageBox.Show($"Runtime Error: {errorMsg}", errorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Log.WriteLine("Sweep completed successfully.");
                MessageBox.Show("Sweep completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Stop_ALL_AUTOTESTWORKER(object sender, EventArgs e)
        {
            if (autotestWorker != null && autotestWorker.IsBusy)
            {
                autotestWorker.CancelAsync();
            }
            else
            {
                MessageBox.Show("현재 실행 중인 작업이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Start_ABGR_TC_Sweep()
        {
            if (autotestWorker != null && autotestWorker.IsBusy)
            {
                MessageBox.Show("다른 작업이 이미 실행 중입니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            autotestWorker = new BackgroundWorker
            {
                WorkerSupportsCancellation = true
            };

            autotestWorker.DoWork += (sender, e) =>
            {
                Execute_ABGR_TC_Sweep(sender as BackgroundWorker, e);
            };

            autotestWorker.RunWorkerCompleted += Generic_RunWorkerCompleted;

            autotestWorker.RunWorkerAsync();
        }

        private void Execute_ABGR_TC_Sweep(BackgroundWorker worker, DoWorkEventArgs e)
        {
            string time;
            //uint[] tc_val = new uint[16];
            uint[] tc_val =
            {
                2, 3, 4, 4, 5, 6, 7, 7, 9, 10, 11, 12, 14, 15, 0, 1
            };
            double dmm_volt;
            double[] temps = Create_TEMP_ARRAY(10, -40, 5);
            double[] vbats = { 2.9 };

            double[,,] vbatResults = new double[temps.Length, vbats.Length, 16];

            try
            {
                Check_Instrument();
                I2C.ft4222H.GPIO_SetDirection(2, GPIO_Direction.Output);
                I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);

                PowerSupply0.Write("VOLT 2.9,(@1)");
                PowerSupply0.Write("VOLT 0.9,(@2)");
                PowerSupply0.Write("VOLT 1.1,(@3)");
                PowerSupply0.Write("OUTP ON,(@1)");
                I2C.ft4222H.GPIO_SetState(2, GPIO_State.High);
                Thread.Sleep(250);
                PowerSupply0.Write("OUTP ON,(@2)");
                Thread.Sleep(250);
                PowerSupply0.Write("OUTP ON,(@3)");
                Thread.Sleep(1000);

                time = DateTime.Now.ToString("HHmmss");

                Parent.Invoke((MethodInvoker)delegate
                {
                    Parent.xlMgr.Sheet.Add(time + "_ABGR_TC_SWEEP");
                    Parent.xlMgr.Cell.Write(1, 1, "O_ABGR_TC[3:0]");
                    Parent.xlMgr.Cell.Write(2, 1, "Best_CONT[3:0]");
                    Parent.xlMgr.Cell.Write(3, 1, "Trimmed_Volt[mV]");
                });

                Set_TempChamber(worker, e, true, 25);

                for (int i = 0; i < 16; i++)
                {
                    if (worker.CancellationPending) { e.Cancel = true; return; }

                    Ensure_Device_Ready(worker, e);
                    if (worker.CancellationPending) { e.Cancel = true; return; }

                    UpdateRegisterBits(0xDC34_040C, 0xC0FF, (24u << 8));

                    UpdateRegisterBits(0xDC34_0414, 0xF81F, (8u << 5));

                    Set_GPIO4_ABGR(true);

                    UpdateRegisterBits(0xDC34_040C, 0xFF0F, ((uint)i << 4));
                    Thread.Sleep(100);

                    double[] trim_result = Start_ABGR_Trim(300);

                    tc_val[i] = (uint)trim_result[0]; // Best CONT Code
                    dmm_volt = trim_result[1];        // Measured Voltage

                    int col_i = i;
                    Parent.Invoke((MethodInvoker)delegate
                    {
                        Parent.xlMgr.Cell.Write(1, 2 + col_i, col_i.ToString());
                        Parent.xlMgr.Cell.Write(2, 2 + col_i, tc_val[col_i].ToString());
                        Parent.xlMgr.Cell.Write(3, 2 + col_i, dmm_volt.ToString("F3"));
                    });
                }

                for (int vbat = 0; vbat < vbats.Length; vbat++)
                {
                    int vbat_idx = vbat;
                    Parent.Invoke((MethodInvoker)delegate
                    {
                        Parent.xlMgr.Cell.Write(1, 18 + (17 * vbat_idx), vbats[vbat_idx].ToString());
                        Parent.xlMgr.Cell.Write(2, 18 + (17 * vbat_idx), "O_ABGR_TC[3:0]");
                        for (int i = 0; i < 16; i++)
                        {
                            Parent.xlMgr.Cell.Write(2, 19 + i + (17 * vbat_idx), i.ToString());
                        }
                        for (int i = 0; i < temps.Length; i++)
                        {
                            Parent.xlMgr.Cell.Write(3 + i, 18 + (17 * vbat_idx), temps[i].ToString());
                        }
                    });
                }

                for (int j = 0; j < temps.Length; j++)
                {
                    if (worker.CancellationPending) { e.Cancel = true; return; }

                    Set_TempChamber(worker, e, true, temps[j]);

                    for (int vbat = 0; vbat < vbats.Length; vbat++)
                    {
                        if (worker.CancellationPending) { e.Cancel = true; return; }

                        for (int i = 0; i < 16; i++)
                        {
                            if (worker.CancellationPending) { e.Cancel = true; return; }

                            Ensure_Device_Ready(worker, e);
                            if (worker.CancellationPending) { e.Cancel = true; return; }

                            UpdateRegisterBits(0xDC34_040C, 0xC0FF, (24u << 8));

                            UpdateRegisterBits(0xDC34_0414, 0xF81F, (8u << 5));

                            Set_GPIO4_ABGR(true);

                            UpdateRegisterBits(0xDC34_040C, 0xFF0F, ((uint)i << 4));

                            UpdateRegisterBits(0xDC34_040C, 0xFFF0, tc_val[i]);

                            Thread.Sleep(500);

                            dmm_volt = new double();
                            for(int k = 0; k < 3; k++)
                            {
                                dmm_volt += double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                            }
                            dmm_volt = Math.Round((dmm_volt / 3), 3);

                            Parent.xlMgr.Cell.Write(3 + j, 19 + i + (17 * vbat), dmm_volt.ToString("F3"));
                        }
                    }
                }

                time = DateTime.Now.ToString("MM/dd/HH:mm");
                Parent.Invoke((MethodInvoker)delegate
                {
                    Log.WriteLine($"Complete time : {time}");
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in Start_ABGR_TC_Sweep: {ex.Message}", ex);
            }
            finally
            {
                Set_TempChamber(worker, e, false, 25);
                Set_GPIO4_ABGR(false);

                if (PowerSupply0 != null && PowerSupply0.IsOpen)
                {
                    PowerSupply0.Write("OUTP OFF,(@1, 2, 3)");
                }
                I2C.ft4222H.GPIO_SetDirection(2, GPIO_Direction.Input);
            }
        }

        //private void Execute_ABGR_TC_Sweep(BackgroundWorker worker, DoWorkEventArgs e)
        //{
        //    string time;
        //    uint[] tc_val = new uint[16]; // Trim 결과 저장용 (실 측정에는 미사용)
        //    double dmm_volt;
        //    double[] temps =
        //    {
        //        85, 80, 75, 70,
        //        65, 60, 55, 50, 45,
        //        40, 35, 30, 25, 20,
        //        15, 10, 5, 0, -5,
        //        -10, -15, -20, -25, -30,
        //        -35, -40
        //    };
        //    double[] vbats = { 2.9 };

        //    int[] tc_sweep_values = { 12, 13, 14 };

        //    try
        //    {
        //        Check_Instrument();

        //        I2C.ft4222H.GPIO_SetDirection(2, GPIO_Direction.Output);
        //        I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);

        //        PowerSupply0.Write("VOLT 2.9,(@1)");
        //        PowerSupply0.Write("VOLT 0.9,(@2)");
        //        PowerSupply0.Write("VOLT 1.1,(@3)");
        //        PowerSupply0.Write("OUTP ON,(@1)");
        //        I2C.ft4222H.GPIO_SetState(2, GPIO_State.High);
        //        Thread.Sleep(250);
        //        PowerSupply0.Write("OUTP ON,(@2)");
        //        Thread.Sleep(250);
        //        PowerSupply0.Write("OUTP ON,(@3)");
        //        Thread.Sleep(1000);

        //        time = DateTime.Now.ToString("HHmmss");

        //        Parent.Invoke((MethodInvoker)delegate
        //        {
        //            Parent.xlMgr.Sheet.Add(time + "_ABGR_TC_SWEEP");
        //            Parent.xlMgr.Cell.Write(1, 1, "O_ABGR_TC[3:0]");
        //            Parent.xlMgr.Cell.Write(2, 1, "Best_CONT[3:0]");
        //            Parent.xlMgr.Cell.Write(3, 1, "Trimmed_Volt[mV]");
        //        });

        //        Set_TempChamber(worker, e, true, 25); // 25도 설정

        //        for (int i = 0; i < 16; i++)
        //        {
        //            if (worker.CancellationPending) { e.Cancel = true; return; }

        //            Ensure_Device_Ready(worker, e);
        //            if (worker.CancellationPending) { e.Cancel = true; return; }

        //            UpdateRegisterBits(0xDC34_040C, 0xC0FF, (24u << 8));

        //            UpdateRegisterBits(0xDC34_0414, 0xF81F, (8u << 5));

        //            Set_GPIO4_ABGR(true);

        //            UpdateRegisterBits(0xDC34_040C, 0xFF0F, ((uint)i << 4));
        //            Thread.Sleep(100);

        //            double[] trim_result = Start_ABGR_Trim(300);

        //            tc_val[i] = (uint)trim_result[0]; // Best CONT Code
        //            dmm_volt = trim_result[1];        // Measured Voltage

        //            int col_i = i;
        //            Parent.Invoke((MethodInvoker)delegate
        //            {
        //                Parent.xlMgr.Cell.Write(1, 2 + col_i, col_i.ToString());
        //                Parent.xlMgr.Cell.Write(2, 2 + col_i, tc_val[col_i].ToString());
        //                Parent.xlMgr.Cell.Write(3, 2 + col_i, dmm_volt.ToString("F3"));
        //            });
        //        }

        //        for (int vbat = 0; vbat < vbats.Length; vbat++)
        //        {
        //            int vbat_idx = vbat;
        //            Parent.Invoke((MethodInvoker)delegate
        //            {
        //                Parent.xlMgr.Cell.Write(1, 18 + (17 * vbat_idx), vbats[vbat_idx].ToString());
        //                Parent.xlMgr.Cell.Write(2, 18 + (17 * vbat_idx), "O_ABGR_TC[3:0]");

        //                for (int k = 0; k < tc_sweep_values.Length; k++)
        //                {
        //                    Parent.xlMgr.Cell.Write(2, 19 + k + (17 * vbat_idx), tc_sweep_values[k].ToString());
        //                }

        //                for (int k = 0; k < temps.Length; k++)
        //                {
        //                    Parent.xlMgr.Cell.Write(3 + k, 18 + (17 * vbat_idx), temps[k].ToString());
        //                }
        //            });
        //        }

        //        for (int j = 0; j < temps.Length; j++)
        //        {
        //            if (worker.CancellationPending) { e.Cancel = true; return; }

        //            Set_TempChamber(worker, e, true, temps[j]);

        //            for (int vbat = 0; vbat < vbats.Length; vbat++)
        //            {
        //                if (worker.CancellationPending) { e.Cancel = true; return; }

        //                for (int i_idx = 0; i_idx < tc_sweep_values.Length; i_idx++)
        //                {
        //                    if (worker.CancellationPending) { e.Cancel = true; return; }

        //                    Ensure_Device_Ready(worker, e);
        //                    if (worker.CancellationPending) { e.Cancel = true; return; }

        //                    uint fixed_tc_value = (uint)tc_sweep_values[i_idx];

        //                    UpdateRegisterBits(0xDC34_040C, 0xC0FF, (24u << 8));

        //                    UpdateRegisterBits(0xDC34_0414, 0xF81F, (8u << 5));

        //                    Set_GPIO4_ABGR(true);

        //                    UpdateRegisterBits(0xDC34_040C, 0xFF0F, (fixed_tc_value << 4));

        //                    UpdateRegisterBits(0xDC34_040C, 0xFFF0, 2u);

        //                    Thread.Sleep(100);

        //                    dmm_volt = double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;

        //                    Parent.xlMgr.Cell.Write(3 + j, 19 + i_idx + (17 * vbat), dmm_volt.ToString("F3"));
        //                }
        //            }
        //        }

        //        time = DateTime.Now.ToString("MM/dd/HH:mm");
        //        Parent.Invoke((MethodInvoker)delegate
        //        {
        //            Log.WriteLine($"Complete time : {time}");
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception($"Error in Execute_ABGR_TC_Sweep: {ex.Message}", ex);
        //    }
        //    finally
        //    {
        //        Set_TempChamber(worker, e, false, 25);
        //        Set_GPIO4_ABGR(false);

        //        if (PowerSupply0 != null && PowerSupply0.IsOpen)
        //        {
        //            PowerSupply0.Write("OUTP OFF,(@1, 2, 3)");
        //        }
        //        I2C.ft4222H.GPIO_SetDirection(2, GPIO_Direction.Input);
        //    }
        //}

        private bool Ensure_Device_Ready(BackgroundWorker worker, DoWorkEventArgs e)
        {
            uint i2cFailCount = 0, fwFailCount = 0;
            bool isReset = true;
            while (true)
            {
                if (worker.CancellationPending) { e.Cancel = true; return true; }

                if (ReadRegister(0x5000_0000) == 0x0202_1010)
                {
                    WriteRegister(0xDC34_0000, 0x000F);
                    Thread.Sleep(100);
                    if (ReadRegister(0xDC34_0000) == 0x0000_0000)
                    {
                        UpdateRegisterBits(0xDC34_040C, 0xC0FF, (24u << 8));
                        UpdateRegisterBits(0xDC34_0414, 0xF81F, (8u << 5));

                        break;
                    }
                    else
                    {
                        fwFailCount++;
                        Log.WriteLine($"FW Looping Failed. - Try {fwFailCount}");
                        isReset = false;

                        PowerSupply0.Write("OUTP OFF,(@1, 2, 3)");
                        Thread.Sleep(500);

                        PowerSupply0.Write("VOLT 3.6,(@1)");
                        PowerSupply0.Write("VOLT 0.9,(@2)");
                        PowerSupply0.Write("VOLT 0.95,(@3)");
                        PowerSupply0.Write("OUTP ON,(@1, 2, 3)");
                        Thread.Sleep(250);

                        if(I2C.ft4222H != null)
                        {
                            I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);
                            Thread.Sleep(1500);
                            I2C.ft4222H.GPIO_SetState(2, GPIO_State.High);
                            Thread.Sleep(1500);
                        }
                    }
                }
                else
                {
                    i2cFailCount++;
                    Log.WriteLine($"I2C Communication Failed. - Try {i2cFailCount}");
                    isReset = false;

                    PowerSupply0.Write("OUTP OFF,(@1, 2, 3)");
                    Thread.Sleep(500);

                    PowerSupply0.Write("VOLT 3.6,(@1)");
                    PowerSupply0.Write("VOLT 0.9,(@2)");
                    PowerSupply0.Write("VOLT 0.95,(@3)");
                    PowerSupply0.Write("OUTP ON,(@1, 2, 3)");
                    Thread.Sleep(250);

                    if (I2C.ft4222H != null)
                    {
                        I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);
                        Thread.Sleep(1500);
                        I2C.ft4222H.GPIO_SetState(2, GPIO_State.High);
                        Thread.Sleep(1500);
                    }
                }
            }
            return isReset;
        }

        private void Start_PMU_TEMP_TEST()
        {
            if (autotestWorker != null && autotestWorker.IsBusy)
            {
                MessageBox.Show("다른 작업이 이미 실행 중입니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            autotestWorker = new BackgroundWorker
            {
                WorkerSupportsCancellation = true
            };

            autotestWorker.DoWork += (sender, e) =>
            {
                Execute_PMU_TEMP_TEST(sender as BackgroundWorker, e);
            };

            autotestWorker.RunWorkerCompleted += Generic_RunWorkerCompleted;

            autotestWorker.RunWorkerAsync();
        }

        private class TestItem
        {
            public string Name { get; set; }
            public uint RegAddr { get; set; }      
            public uint Mask { get; set; }         
            public int Shift { get; set; }         
            public double TargetVal { get; set; }  
            public int MeasureType { get; set; }   
            public uint BestCode { get; set; }     
            public int MaxCode { get; set; }       
            public Action EnableAction { get; set; }
        }

        private void Execute_PMU_TEMP_TEST(BackgroundWorker worker, DoWorkEventArgs e)
        {
            // ABGR, DCDC, FLDO, RET, MBGR, DA_LDO, 32KOSC.
            // DMM 0 = DCDC_VOUT
            // DMM 1 = FLDO
            // DMM 2 = RET, MBGR, DALDO
            // DMM 3 = 32KOSC
            string time;
            double[] temps = //Create_TEMP_ARRAY(85, -40, 5);
            {
                85, 55, 25, -20, -40
            };
            double[] vbats = { 1.7, 2.9, 3.6 };
            double trimVBAT = 2.9;

            List<TestItem> items = new List<TestItem>
            {
                new TestItem {
                    Name = "ABGR", RegAddr = 0xDC34_040C, Mask = 0xFFF0, Shift = 0, TargetVal = 300, 
                    MeasureType = 2, MaxCode = 15, EnableAction = () => { Set_GPIO4_ABGR(true); }
                },
                new TestItem {
                    Name = "DCDC", RegAddr = 0xDC34_041C, Mask = 0xFF81, Shift = 1, TargetVal = 1.3, 
                    MeasureType = 0, MaxCode = 63, EnableAction = () => {  }
                },
                new TestItem {
                    Name = "FLDO", RegAddr = 0xDC34_0414, Mask = 0xFFF0, Shift = 0, TargetVal = 1.8, 
                    MeasureType = 1, MaxCode = 15, EnableAction = () => {  }
                },
                new TestItem {
                    Name = "RET", RegAddr = 0xDC34_0418, Mask = 0x0FFF, Shift = 12, TargetVal = 600,
                    MeasureType = 2, MaxCode = 15, EnableAction = () => { Set_GPIO4_RETLDO(true); }
                },
                new TestItem {
                    Name = "MBGR", RegAddr = 0xDC34_0438, Mask = 0xFC0F,  Shift = 4, TargetVal = 300, 
                    MeasureType = 2, MaxCode = 63, EnableAction = () => { Set_GPIO4_MBGR(true); }
                },
                new TestItem {
                    Name = "DA_LDO", RegAddr = 0xDC34_00C0, Mask = 0xFF03, Shift = 2, TargetVal = 450, 
                    MeasureType = 2, MaxCode = 63, EnableAction = () => { Set_GPIO4_DALDO(true); }
                },
                new TestItem {
                    Name = "32KOSC", RegAddr = 0xDC34_0410, Mask = 0xFE03, Shift = 2, TargetVal = 32.768, 
                    MeasureType = 3, MaxCode = 63, EnableAction = () => { Set_GPIO16_32KOSC(true); }
                }
            };

            try
            {
                Check_Instrument();

                PowerSupply0.Write("OUTP OFF,(@1, 2, 3)");
                Thread.Sleep(500);

                PowerSupply0.Write("VOLT 3.6,(@1)");
                PowerSupply0.Write("VOLT 0.9,(@2)");
                PowerSupply0.Write("VOLT 0.95,(@3)");
                PowerSupply0.Write("OUTP ON,(@1, 2, 3)");
                Thread.Sleep(250);

                if(I2C.ft4222H != null)
                {
                    I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);
                    Thread.Sleep(1500);
                    I2C.ft4222H.GPIO_SetState(2, GPIO_State.High);
                    Thread.Sleep(1500);
                }

                Ensure_Device_Ready(worker, e);

                time = DateTime.Now.ToString("HHmmss");

                Parent.Invoke((MethodInvoker)delegate
                {
                    Parent.xlMgr.Sheet.Add(time + "_PMU_TEMP_TEST");

                    Parent.xlMgr.Cell.Write(1, 1, "O_ABGR_TC[3:0]");
                    Parent.xlMgr.Cell.Write(2, 1, "O_ABGR_CONT[3:0]");
                    Parent.xlMgr.Cell.Write(3, 1, "DMM[mV]");

                    Parent.xlMgr.Cell.Write(4, 1, "O_DCDC_SET_VOUT[5:0]");
                    Parent.xlMgr.Cell.Write(5, 1, "DMM[V]");

                    Parent.xlMgr.Cell.Write(6, 1, "O_FLDO_CONT[3:0]");
                    Parent.xlMgr.Cell.Write(7, 1, "DMM[V]");

                    Parent.xlMgr.Cell.Write(8, 1, "O_RET_LDO_CONT[3:0]");
                    Parent.xlMgr.Cell.Write(9, 1, "DMM[mV]");

                    Parent.xlMgr.Cell.Write(10, 1, "O_MBGR_OUT_TRIM[5:0]");
                    Parent.xlMgr.Cell.Write(11, 1, "DMM[mV]");

                    Parent.xlMgr.Cell.Write(12, 1, "O_DA_LDO_CONT[5:0]");
                    Parent.xlMgr.Cell.Write(13, 1, "DMM[mV]");

                    Parent.xlMgr.Cell.Write(14, 1, "RTC_SCKF[10:0]");
                    Parent.xlMgr.Cell.Write(15, 1, "Freq[KHz]");
                });

                Set_TempChamber(worker, e, true, 25); // 25도 설정

                for (int i = 0; i < items.Count; i++)
                {
                    if (worker.CancellationPending) { e.Cancel = true; return; }

                    PowerSupply0.Write($"VOLT {trimVBAT},(@1)");
                    Thread.Sleep(500);

                    var item = items[i];

                    if (item.EnableAction != null)
                    {
                        Set_GPIO_Disable();
                        item.EnableAction();
                        Thread.Sleep(100);
                    }

                    double minDiff = double.MaxValue;
                    uint bestCode = 0;
                    double bestMeas = 0;

                    for (uint code = 0; code <= item.MaxCode; code++)
                    {
                        if (worker.CancellationPending) { e.Cancel = true; return; }

                        UpdateRegisterBits(item.RegAddr, (uint)(item.Mask), (code << item.Shift));
                        Thread.Sleep(250);

                        double measuredVal = Measure_Item(item.MeasureType);

                        double diff = Math.Abs(measuredVal - item.TargetVal);
                        if (diff < minDiff)
                        {
                            minDiff = diff;
                            bestCode = code;
                            bestMeas = measuredVal;
                        }
                    }

                    item.BestCode = bestCode;
                    UpdateRegisterBits(item.RegAddr, (uint)(item.Mask), (bestCode << item.Shift));
                    Thread.Sleep(250);
                    bestMeas = Measure_Item(item.MeasureType);

                    int col_idx = 2 + (i * 2);
                    Parent.Invoke((MethodInvoker)delegate
                    {
                        Parent.xlMgr.Cell.Write(col_idx, 2, bestCode.ToString());
                        Parent.xlMgr.Cell.Write(col_idx + 1, 2, bestMeas.ToString("F5"));
                    });
                }

                for (int vbat = 0; vbat < vbats.Length; vbat++)
                {
                    int vbat_idx = vbat;
                    Parent.Invoke((MethodInvoker)delegate
                    {
                        Parent.xlMgr.Cell.Write(1, 4 + ((items.Count + 2) * vbat_idx), vbats[vbat_idx].ToString());
                        Parent.xlMgr.Cell.Write(2, 4 + ((items.Count + 2) * vbat_idx), "ITEM");

                        for (int k = 0; k < items.Count; k++)
                        {
                            var item = items[k];
                            Parent.xlMgr.Cell.Write(2, 5 + k + ((items.Count + 2) * vbat_idx), item.Name.ToString());
                        }

                        for (int k = 0; k < temps.Length; k++)
                        {
                            Parent.xlMgr.Cell.Write(3 + k, 4 + ((items.Count + 2) * vbat_idx), temps[k].ToString());
                        }
                    });
                }

                for (int j = 0; j < temps.Length; j++)
                {
                    if (worker.CancellationPending) { e.Cancel = true; return; }

                    Set_TempChamber(worker, e, true, temps[j]);

                    for (int i = 0; i < items.Count; i++)
                    {
                        bool isReset = Ensure_Device_Ready(worker, e);
                        if (!isReset)
                        {
                            for (int k = 0; k < items.Count; k++)
                            {
                                var trimedItem = items[k];

                                UpdateRegisterBits(trimedItem.RegAddr, (uint)(trimedItem.Mask), (trimedItem.BestCode << trimedItem.Shift));
                                Thread.Sleep(250);
                            }
                        }

                        if (worker.CancellationPending) { e.Cancel = true; return; }

                        var item = items[i];

                        if (item.EnableAction != null)
                        {
                            Set_GPIO_Disable();
                            item.EnableAction();
                            Thread.Sleep(100);
                        }

                        for (int vbat = 0; vbat < vbats.Length; vbat++)
                        {
                            if (worker.CancellationPending) { e.Cancel = true; return; }

                            PowerSupply0.Write($"VOLT {vbats[vbat]},(@1)");
                            Thread.Sleep(1000);

                            double measuredVal = Measure_Item(item.MeasureType);

                            int row = 3 + j;
                            int col = 5 + i + ((items.Count + 2) * vbat);

                            Parent.Invoke((MethodInvoker)delegate
                            {
                                Parent.xlMgr.Cell.Write(row, col, measuredVal.ToString("F5"));
                            });
                        }
                    }
                }

                time = DateTime.Now.ToString("MM/dd/HH:mm");
                Parent.Invoke((MethodInvoker)delegate
                {
                    Log.WriteLine($"Complete time : {time}");
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in Execute_PMU_TEMP_TEST: {ex.Message}", ex);
            }
            finally
            {
                Set_TempChamber(worker, e, false, 25);

                if (PowerSupply0 != null && PowerSupply0.IsOpen)
                {
                    PowerSupply0.Write("OUTP OFF,(@1, 2, 3)");
                }
                I2C.ft4222H.GPIO_SetState(2, GPIO_State.Low);
            }
        }

        private double Measure_Item(int type)
        {
            double val = 0;
            uint avgTime = 10;
            switch (type)
            {
                case 0: // DMM0 (DCDC)
                    for (int i = 0; i < avgTime; i++)
                        val += double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?"));
                    val /= avgTime;
                    break;
                case 1: // DMM1 (FLDO)
                    for (int i = 0; i < avgTime; i++)
                        val += double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?"));
                    val /= avgTime;
                    break;
                case 2: // DMM2 (ABGR, RET, MBGR, DALDO)
                    for (int i = 0; i < avgTime; i++)
                        val += double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    val /= avgTime;
                    break;
                case 3: // Frequency (32KOSC)
                    for(int i = 0; i < avgTime; i++)
                        val += double.Parse(DigitalMultimeter3.WriteAndReadString(":MEAS:FREQ?")) / 1000;
                    val /= avgTime;
                    break;
            }
            return val;
        }

        private void Run_SWEEP_POWERSUPPLY0_OUTPUT()
        {
            try
            {
                Check_Instrument();

                string userInput;
                userInput = Prompt.ShowDialog("PowerSupply의 Output Port를 선택해주세요. (1-3):", "Output Sweep Test");
                uint psChannel;
                if (string.IsNullOrEmpty(userInput))
                {
                    Log.WriteLine("사용자가 입력을 취소했습니다.");
                    return;
                }
                if (!uint.TryParse(userInput, out psChannel) || psChannel < 1 || psChannel > 3)
                {
                    MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 1 ~ 3 사이의 숫자만 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                userInput = Prompt.ShowDialog("시작 Voltage를 입력하여 주세요. (0 이상)", "Output Sweep Test");
                double startVoltage;
                if (string.IsNullOrEmpty(userInput))
                {
                    Log.WriteLine("사용자가 입력을 취소했습니다.");
                    return;
                }
                if (!double.TryParse(userInput, out startVoltage) || startVoltage < 0)
                {
                    MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 크거나 같은 수를 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                userInput = Prompt.ShowDialog("종료 Voltage를 입력하여 주세요. (0 이상)", "Output Sweep Test");
                double stopVoltage;
                if (string.IsNullOrEmpty(userInput))
                {
                    Log.WriteLine("사용자가 입력을 취소했습니다.");
                    return;
                }
                if (!double.TryParse(userInput, out stopVoltage) || stopVoltage < 0)
                {
                    MessageBox.Show($"잘못된 값 '{userInput}' 입니다. 0보다 크거나 같은 수를 입력하십시오.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                const double stepMag = 0.05;
                const int dwellMs = 50;

                if (Math.Abs(stopVoltage - startVoltage) < 1e-9)
                {
                    PowerSupply0.Write($"INST:NSEL {psChannel}");
                    PowerSupply0.Write("OUTP ON");
                    PowerSupply0.Write($"VOLT {stopVoltage:F3}");
                    Log.WriteLine($"CH{psChannel} VOLT set to {stopVoltage:F3} V");
                    return;
                }

                double dir = Math.Sign(stopVoltage - startVoltage);
                double step = stepMag * dir;

                PowerSupply0.Write($"INST:NSEL {psChannel}");
                PowerSupply0.Write("OUTP ON");
                PowerSupply0.Write($"VOLT {startVoltage:F3}");

                double v = startVoltage;
                while ((dir > 0 && v <= stopVoltage) || (dir < 0 && v >= stopVoltage))
                {
                    PowerSupply0.Write($"VOLT {v:F3}");
                    Thread.Sleep(dwellMs);
                    v += step;
                }

                PowerSupply0.Write($"VOLT {stopVoltage:F3}");
                Log.WriteLine($"CH{psChannel} swept from {startVoltage:F3} V to {stopVoltage:F3} V in 0.1 V steps / {dwellMs} ms dwell.");
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in SWEEP POWERSUPPLY0: {ex.Message}";
                Log.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #region DVPnR RX Test Item
        private bool ShowPromptAndCheckCancel(string message, string caption)
        {
            DialogResult result = MessageBox.Show(message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result != DialogResult.OK)
            {
                SigGen_Reset(true);
                MessageBox.Show("시퀀스가 중단 되었습니다.");
                return false;
            }
            return true;
        }

        private void RX_RF_Gain(uint ch_sel, uint gain)
        {
            Check_Instrument();

            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("w_PLL_PEN");
            RegisterItem SPI_CH_SEL = Parent.RegMgr.GetRegisterItem("m_SPI_CH_SEL[6:0]");
            RegisterItem RX_RF_GC = Parent.RegMgr.GetRegisterItem("r_RX_RF_GC[1:0]");
            RegisterItem ABB_DRV_TPE = Parent.RegMgr.GetRegisterItem("O_ABB_DRV_TPE[3:0]");

            SignalGenerator1.Write(":OUTPut:STATe OFF");

            PLL_PEN.Read();
            PLL_PEN.Value = 0;
            PLL_PEN.Write();

            ABB_DRV_TPE.Read();
            ABB_DRV_TPE.Value = 9;
            ABB_DRV_TPE.Write();

            SPI_CH_SEL.Read();
            SPI_CH_SEL.Value = ch_sel;
            SPI_CH_SEL.Write();

            RX_RF_GC.Read();
            RX_RF_GC.Value = gain;
            RX_RF_GC.Write();

            PLL_PEN.Read();
            PLL_PEN.Value = 1;
            PLL_PEN.Write();

            SignalGenerator1.Write($":FREQ {2400 + ch_sel}E6");
            if (gain != 0)
            {
                SignalGenerator1.Write($":POW {-85}");
            }
            else if (gain == 0)
            {
                SignalGenerator1.Write($":POW {-65}");
            }

            SignalGenerator1.Write(":OUTPut:STATe ON");
        }

        private void RX_Gain(uint ch_sel, uint mode)
        {
            #region Register ITEM
            RegisterItem PLL_PEN = Parent.RegMgr.GetRegisterItem("w_PLL_PEN");
            RegisterItem SPI_CH_SEL = Parent.RegMgr.GetRegisterItem("m_SPI_CH_SEL[6:0]");
            RegisterItem ABB_DRV_TPE = Parent.RegMgr.GetRegisterItem("O_ABB_DRV_TPE[3:0]");
            RegisterItem RX_RF_GC = Parent.RegMgr.GetRegisterItem("r_RX_RF_GC[1:0]");
            RegisterItem ABB_FLT_GC = Parent.RegMgr.GetRegisterItem("w_ABB_FLT_GC[1:0]");
            RegisterItem ABB_TIA_GC = Parent.RegMgr.GetRegisterItem("w_ABB_TIA_GC[2:0]");
            RegisterItem ABB_PGA_GC = Parent.RegMgr.GetRegisterItem("w_ABB_PGA_GC[4:0]");
            #endregion Register ITEM

            if (mode != 1 && mode != 0)
            {
                return;
            }

            PLL_PEN.Read();
            PLL_PEN.Value = 0;
            PLL_PEN.Write();
            System.Threading.Thread.Sleep(10);

            ABB_DRV_TPE.Read();
            ABB_DRV_TPE.Value = 0;
            ABB_DRV_TPE.Write();
            System.Threading.Thread.Sleep(10);

            SPI_CH_SEL.Read();
            SPI_CH_SEL.Value = ch_sel;
            SPI_CH_SEL.Write();
            System.Threading.Thread.Sleep(10);

            if (mode == 1) // Maximum
            {
                RX_RF_GC.Read();
                RX_RF_GC.Value = 3;
                RX_RF_GC.Write();
                System.Threading.Thread.Sleep(10);

                ABB_FLT_GC.Read();
                ABB_FLT_GC.Value = 0;
                ABB_FLT_GC.Write();
                System.Threading.Thread.Sleep(10);

                ABB_TIA_GC.Read();
                ABB_TIA_GC.Value = 6;
                ABB_TIA_GC.Write();
                System.Threading.Thread.Sleep(10);

                ABB_PGA_GC.Read();
                ABB_PGA_GC.Value = 23;
                ABB_PGA_GC.Write();
                System.Threading.Thread.Sleep(10);
            }

            else if (mode == 0) // Minimum
            {
                RX_RF_GC.Read();
                RX_RF_GC.Value = 0;
                RX_RF_GC.Write();
                System.Threading.Thread.Sleep(10);

                ABB_FLT_GC.Read();
                ABB_FLT_GC.Value = 0;
                ABB_FLT_GC.Write();
                System.Threading.Thread.Sleep(10);

                ABB_TIA_GC.Read();
                ABB_TIA_GC.Value = 0;
                ABB_TIA_GC.Write();
                System.Threading.Thread.Sleep(10);

                ABB_PGA_GC.Read();
                ABB_PGA_GC.Value = 0;
                ABB_PGA_GC.Write();
                System.Threading.Thread.Sleep(10);
            }

            PLL_PEN.Read();
            PLL_PEN.Value = 1;
            PLL_PEN.Write();
            System.Threading.Thread.Sleep(10);
        }

        private void SigGen_Reset(bool start)
        {
            if (!start)
            {
                return;
            }

            var devices = new Dictionary<string, dynamic>
            {
                { "I2C", I2C },
                { "SignalGenerator0", SignalGenerator0 },
                { "SignalGenerator1", SignalGenerator1 },
            };

            foreach (var device in devices)
            {
                if (!device.Value.IsOpen)
                {
                    MessageBox.Show($"Check {device.Key} Connection!");
                    return;
                }
            }
            ;

            SignalGenerator0.Write(":OUTPut:STATe OFF");
            SignalGenerator1.Write(":OUTPut:STATe OFF");

            SignalGenerator0.Write($":POW {-85}");
            SignalGenerator1.Write($":POW {-85}");

            SignalGenerator0.Write($":FREQ 2402E6");
            SignalGenerator1.Write($":FREQ 2402E6");
        }

        private void DVPnR_RX_Sequence(int sequence)
        {
            bool Check;
            if (sequence < 13 || sequence > 18)
            {
                MessageBox.Show($"텍스트 상자에 13 ~ 18 의 값을 입력해주세요!");
                return;
            }

            Check_Instrument();

            RegisterItem ABB_DRV_TPE = Parent.RegMgr.GetRegisterItem("O_ABB_DRV_TPE[3:0]");

            var devices = new Dictionary<string, dynamic>
            {
                { "I2C", I2C },
                { "SignalGenerator0", SignalGenerator0 },
                { "SignalGenerator1", SignalGenerator1 },
                { "PowerSupply0", PowerSupply0 }
            };
            uint[] ch_sel =
            {
                2, 26, 40, 80
            };

            foreach (var device in devices)
            {
                if (!device.Value.IsOpen)
                {
                    MessageBox.Show($"Check {device.Key} Connection!");
                    return;
                }
            }
            ;

            switch (sequence)
            {
                // RF13. RX_RF_GAIN.
                case 13:
                    string[] rf_gain = { "Low gain", "", "Mid gain", "High gain" };
                    for (uint i = 0; i < ch_sel.Length; i++)
                    {
                        for (uint j = 0; j < 4; j++)
                        {
                            SignalGenerator1.Write(":OUTPut:STATe OFF");
                            System.Threading.Thread.Sleep(10);

                            RX_RF_Gain(ch_sel[i], j);

                            SignalGenerator1.Write($":FREQ {2400 + ch_sel[i]}E6");
                            if (j == 0)
                            {
                                SignalGenerator1.Write($":POW {-65}");
                            }
                            else
                            {
                                SignalGenerator1.Write($":POW {-85}");
                            }
                            SignalGenerator1.Write(":OUTPut:STATe ON");
                            System.Threading.Thread.Sleep(10);

                            Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tGain_Mode = {rf_gain[j]}\n측정 완료 후 확인, 취소를 선택하세요.", "1. RX_RF_GAIN");
                            if (!Check) return;
                            if (j == 0) j = 1;
                        }
                    }
                    SigGen_Reset(true);
                    break;

                // RF14. RX_GAIN.
                case 14:
                    string[] rx_gain = { "Min gain", "Max gain" };
                    for (uint i = 0; i < ch_sel.Length; i++)
                    {
                        for (uint j = 0; j < 2; j++)
                        {
                            SignalGenerator1.Write(":OUTPut:STATe OFF");
                            System.Threading.Thread.Sleep(10);

                            RX_Gain(ch_sel[i], j);

                            SignalGenerator1.Write($":FREQ {2400 + ch_sel[i]}E6");
                            if (j == 0)
                            {
                                SignalGenerator1.Write($":POW {-35}");
                            }
                            else
                            {
                                SignalGenerator1.Write($":POW {-85}");
                            }
                            SignalGenerator1.Write(":OUTPut:STATe ON");
                            System.Threading.Thread.Sleep(10);

                            Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tGain_Mode = {rx_gain[j]}\n측정 완료 후 확인, 취소를 선택하세요.", "2. RX_GAIN");
                            if (!Check) return;
                        }
                    }
                    SigGen_Reset(true);
                    break;

                // RF15. RX_NOISE_FIGURE.
                case 15:
                    for (uint i = 0; i < ch_sel.Length; i++)
                    {
                        SignalGenerator1.Write(":OUTPut:STATe OFF");
                        System.Threading.Thread.Sleep(10);

                        RX_Gain(ch_sel[i], 1);

                        SignalGenerator1.Write($":FREQ {2400 + ch_sel[i]}E6");
                        SignalGenerator1.Write($":POW {-85}");
                        SignalGenerator1.Write(":OUTPut:STATe ON");
                        System.Threading.Thread.Sleep(10);

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tO_ABB_DRV_TPE[3:0] = 0\nPout 측정 완료 후 확인, 취소를 선택하세요.", "3. RX_NF");
                        if (!Check) return;

                        SignalGenerator1.Write(":OUTPut:STATe OFF");

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tO_ABB_DRV_TPE[3:0] = 0\nNout 측정 완료 후 확인, 취소를 선택하세요.", "3. RX_NF");
                        if (!Check) return;

                        ABB_DRV_TPE.Read();
                        ABB_DRV_TPE.Value = 9;
                        ABB_DRV_TPE.Write();

                        SignalGenerator1.Write(":OUTPut:STATe ON");

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tO_ABB_DRV_TPE[3:0] = 9\nPout 측정 완료 후 확인, 취소를 선택하세요.", "3. RX_NF");
                        if (!Check) return;

                        SignalGenerator1.Write(":OUTPut:STATe OFF");

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tO_ABB_DRV_TPE[3:0] = 9\nNout 측정 완료 후 확인, 취소를 선택하세요.", "3. RX_NF");
                        if (!Check) return;
                    }
                    SigGen_Reset(true);
                    break;

                // RF16. RX_IIP3.
                case 16:
                    for (uint i = 0; i < ch_sel.Length; i++)
                    {
                        SignalGenerator1.Write(":OUTPut:STATe OFF");
                        System.Threading.Thread.Sleep(10);

                        RX_Gain(ch_sel[i], 1);

                        SignalGenerator1.Write($":FREQ {2400 + ch_sel[i]}E6");
                        SignalGenerator1.Write($":POW {-85}");
                        SignalGenerator1.Write(":OUTPut:STATe ON");
                        System.Threading.Thread.Sleep(10);

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tGain_Mode = 1\nPout 측정 완료 후 확인, 취소를 선택하세요.", "4. RX_IIP3");
                        if (!Check) return;

                        SignalGenerator0.Write(":OUTPut:STATe OFF");
                        SignalGenerator1.Write(":OUTPut:STATe OFF");
                        System.Threading.Thread.Sleep(10);

                        SignalGenerator0.Write($":POW {-48}");
                        SignalGenerator1.Write($":POW {-48}");

                        SignalGenerator0.Write($":FREQ {2400 + ch_sel[i] + 4}E6");
                        SignalGenerator1.Write($":FREQ {2400 + ch_sel[i] + 8}E6");

                        SignalGenerator0.Write(":OUTPut:STATe ON");
                        SignalGenerator1.Write(":OUTPut:STATe ON");
                        System.Threading.Thread.Sleep(10);

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tPimd_p 측정 완료 후 확인, 취소를 선택하세요.", "4. RX_IIP3");
                        if (!Check) return;

                        SignalGenerator0.Write(":OUTPut:STATe OFF");
                        SignalGenerator1.Write(":OUTPut:STATe OFF");
                        System.Threading.Thread.Sleep(10);

                        SignalGenerator0.Write($":POW {-48}");
                        SignalGenerator1.Write($":POW {-48}");

                        SignalGenerator0.Write($":FREQ {2400 + ch_sel[i] - 4}E6");
                        SignalGenerator1.Write($":FREQ {2400 + ch_sel[i] - 8}E6");

                        SignalGenerator0.Write(":OUTPut:STATe ON");
                        SignalGenerator1.Write(":OUTPut:STATe ON");
                        System.Threading.Thread.Sleep(10);

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\tPimd_n 측정 완료 후 확인, 취소를 선택하세요.", "4. RX_IIP3");
                        if (!Check) return;
                    }

                    SigGen_Reset(true);
                    break;

                // RF17. RX_OOB.
                case 17:
                    SignalGenerator1.Write(":OUTPut:STATe OFF");
                    System.Threading.Thread.Sleep(10);

                    RX_Gain(40, 1);

                    SignalGenerator1.Write($":FREQ 2440E6");
                    SignalGenerator1.Write($":POW -85");
                    SignalGenerator1.Write(":OUTPut:STATe ON");
                    System.Threading.Thread.Sleep(10);

                    Check = ShowPromptAndCheckCancel($"CH_SEL = 2440\tGain_Mode = 1\nPout 측정 완료 후 확인, 취소를 선택하세요.", "5. RX_OOB");
                    if (!Check) return;

                    SignalGenerator1.Write(":OUTPut:STATe OFF");

                    Check = ShowPromptAndCheckCancel($"CH_SEL = 2440\tGain_Mode = 1\nNout 측정 완료 후 확인, 취소를 선택하세요.", "5. RX_OOB");
                    if (!Check) return;

                    SignalGenerator0.Write($":POW {-15}");

                    SignalGenerator0.Write($":FREQ 2399E6");

                    SignalGenerator0.Write(":OUTPut:STATe ON");
                    SignalGenerator1.Write(":OUTPut:STATe ON");

                    Check = ShowPromptAndCheckCancel($"CH_SEL = 2440\tGain_Mode = 1\n2399MHz, -15dBm\nPout 측정 완료 후 확인, 취소를 선택하세요.", "5. RX_OOB");
                    if (!Check) return;

                    SignalGenerator1.Write(":OUTPut:STATe OFF");

                    Check = ShowPromptAndCheckCancel($"CH_SEL = 2440\tGain_Mode = 1\n2399MHz, -15dBm\nNout 측정 완료 후 확인, 취소를 선택하세요.", "5. RX_OOB");
                    if (!Check) return;

                    SignalGenerator0.Write(":OUTPut:STATe OFF");
                    SignalGenerator1.Write(":OUTPut:STATe OFF");
                    System.Threading.Thread.Sleep(10);

                    SignalGenerator0.Write($":FREQ 2484E6");

                    SignalGenerator0.Write(":OUTPut:STATe ON");
                    SignalGenerator1.Write(":OUTPut:STATe ON");
                    System.Threading.Thread.Sleep(10);

                    Check = ShowPromptAndCheckCancel($"CH_SEL = 2440\tGain_Mode = 1\n2484MHz, -15dBm\nPout 측정 완료 후 확인, 취소를 선택하세요.", "5. RX_OOB");
                    if (!Check) return;

                    SignalGenerator1.Write(":OUTPut:STATe OFF");

                    Check = ShowPromptAndCheckCancel($"CH_SEL = 2440\tGain_Mode = 1\n2484MHz, -15dBm\nPout 측정 완료 후 확인, 취소를 선택하세요.", "5. RX_OOB");
                    if (!Check) return;

                    SigGen_Reset(true);
                    break;

                // RF18. RX_FILTER_RESPONSE.
                case 18:
                    double[] offset =
                    {
                        0.75, 2, 3, 4
                    };

                    for (uint i = 0; i < ch_sel.Length; i++)
                    {
                        SignalGenerator1.Write(":OUTPut:STATe OFF");
                        System.Threading.Thread.Sleep(10);

                        RX_Gain(ch_sel[i], 1);

                        SignalGenerator1.Write($":FREQ {2400 + ch_sel[i]}E6");
                        SignalGenerator1.Write($":POW {-85}");
                        SignalGenerator1.Write(":OUTPut:STATe ON");
                        System.Threading.Thread.Sleep(10);

                        Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i]}\nPout 측정 완료 후 확인, 취소를 선택하세요.", "6. RX_FILTER");
                        if (!Check) return;

                        for (int j = 0; j < offset.Length; j++)
                        {
                            SignalGenerator1.Write(":OUTPut:STATe OFF");
                            System.Threading.Thread.Sleep(10);

                            SignalGenerator1.Write($":FREQ {2400 + ch_sel[i] + offset[j]}E6");

                            SignalGenerator1.Write(":OUTPut:STATe ON");
                            System.Threading.Thread.Sleep(10);

                            Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i] + offset[j]}\nPout 측정 완료 후 확인, 취소를 선택하세요.", "6. RX_FILTER");
                            if (!Check) return;
                        }
                        for (int j = 0; j < offset.Length; j++)
                        {
                            SignalGenerator1.Write(":OUTPut:STATe OFF");
                            System.Threading.Thread.Sleep(10);

                            SignalGenerator1.Write($":FREQ {2400 + ch_sel[i] - offset[j]}E6");

                            SignalGenerator1.Write(":OUTPut:STATe ON");
                            System.Threading.Thread.Sleep(10);

                            Check = ShowPromptAndCheckCancel($"CH_SEL = {2400 + ch_sel[i] - offset[j]}\nPout 측정 완료 후 확인, 취소를 선택하세요.", "6. RX_FILTER");
                            if (!Check) return;
                        }
                    }
                    SigGen_Reset(true);
                    break;
            }
        }
        #endregion DVPnR RX Test Item
        #endregion AUTO_TEST_ITEMS

        #region Firmware control methods
        private void SetStatusLabel(string function)
        {
            StatusBar?.GetCurrentParent()?.Invoke(new MethodInvoker(delegate
            {
                StatusBar.Text = $"{function}";
            }));
        }

        private void InitProgressBar(int value, int min, int max)
        {
            ProgressBar?.Invoke(new MethodInvoker(delegate
            {
                ProgressBar.Value = value;
                ProgressBar.Minimum = min;
                ProgressBar.Maximum = max;
            }));
        }

        private bool CheckI2C_ID()
        {
            uint id = ReadRegister(0x50000000);
            if ((id >> 12) != 0x02021)
            {
                Log.WriteLine($"Fail to Check I2C IP ID.\nR = 0x{(id >> 12):X5}", Color.Coral, Log.RichTextBox.BackColor);
                Status = false;
                return false;
            }
            return true;
        }

        private void DownloadFirmware(FW_TARGET target)
        {
            if (GetFirmwareFileName())
            {
                FirmwareTarget = target;
                DownloadFirmware(Oasis_DownloadFW);
            }
        }

        private void EraseFW(FW_TARGET target)
        {
            if (target == FW_TARGET.RAM)
                return;
            FirmwareTarget = target;
            HaltMCU();
            EraseFirmware(Oasis_EraseFW);
        }

        private void DumpFW(FW_TARGET target)
        {
            try
            {
                ReadFirmwareSize = int.Parse(TextBox_TestArgument.Text, System.Globalization.NumberStyles.Number);
            }
            catch
            {
                MessageBox.Show("FW size를 입력하고 다시 시도해주세요.");
                return;
            }

            FirmwareTarget = target;
            DumpFirmware(Oasis_DumpFW);
        }

        public uint Oasis_Check_Staus_Register()
        {
            Status = true;
            uint data = ReadRegister(0x50090020);

            return data & (uint)0x01;
        }

        void WriteMemory_NVM(uint Address, byte[] Data)
        {
            int cnt = 0;
            uint data;

            if (Data.Length == 0)
                return;

            //Fill FIFO for page program
            for (uint i = 0; i < 64; i++)
            {
                if (i * 4 < Data.Length)
                {
                    data = (uint)((Data[i * 4 + 3] << 24) | (Data[i * 4 + 2] << 16) | (Data[i * 4 + 1] << 8) | Data[i * 4 + 0]);
                }
                else
                {
                    data = 0xFFFFFFFF;
                }
                WriteRegister((0x50091000 + (i * 4)), data);
            }

            // Page Program Command
            data = ((uint)(FLASH_CMD.PP) << 24) | (Address & 0xFFFFFF);
            WriteRegister(0x50090008, data);

            // Check Status Register
            while (Oasis_Check_Staus_Register() != 0) // typ 667.5us, max 3110us
            {
                cnt++;
                System.Threading.Thread.Sleep(1);
                if (cnt > 30) break;
            }
            if (cnt > 30)
            {
                Status = false;
                Log.WriteLine($"Serial Flash Controller Status is False.\nStop WriteMemory_NVM Process.", System.Drawing.Color.Coral, Log.RichTextBox.BackColor);
            }
        }

        private void Oasis_DownloadFW()
        {
            Stopwatch flash_write = new Stopwatch();
            flash_write.Start();

            int PageSize = 256;
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
            if (FirmwareTarget == FW_TARGET.RAM)
            {
                PageSize = 4;
            }

            if (!CheckI2C_ID()) goto EXIT;

            HaltMCU();

            /* Mass erase */
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                Oasis_EraseFW();
                if (Status == false)
                    goto EXIT;
            }

            else if (FirmwareTarget == FW_TARGET.RAM)
            {
                WriteRegister(0x50060030, 0x01000000);
            }

            InitProgressBar(0, 0, (FirmwareData.Length + PageSize - 1) / PageSize);

            for (FlashAddress = 0; FlashAddress < FirmwareData.Length; FlashAddress += (uint)PageSize)
            {
                if ((FlashAddress % 0x100) == 0)
                {
                    SetStatusLabel($"Running - Write verify @ 0x{FlashAddress:X8}");
                }

                // Copy FW data
                for (int i = 0; i < PageSize; i++)
                {
                    if (FlashAddress + i < FirmwareData.Length)
                        SendBytes[i] = FirmwareData[FlashAddress + i];
                    else
                        SendBytes[i] = 0xFF;
                }

                // Write firmware data
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    WriteRegister(FlashAddress + 0x80000000, SendBytes);
                }

                else
                {
                    WriteMemory_NVM(FlashAddress, SendBytes);
                    if (Status != true)
                    {
                        goto EXIT;
                    }
                }

                // Compare memory
                if (FirmwareTarget == FW_TARGET.RAM)
                {
                    for (uint i = 0; i < PageSize; i += 4)
                    {
                        uint RcvBytes = ReadRegister(FlashAddress + i + 0x40000000);
                        for (int j = 0; j < 4; j++)
                        {
                            /*
                            if (SendBytes[i * 4 + j] != ((RcvBytes >> (8 * j)) & 0xff))
                            {
                                Log.WriteLine("Faile to match: (0x" + (FlashAddress + i * 4 + j).ToString("X4") + ":" + SendBytes[i * 4 + j].ToString("X2") + " " + ((RcvBytes >> (8 * j)) & 0xff).ToString("X2") + ")",
                                    System.Drawing.Color.Coral, Log.RichTextBox.BackColor);
                                goto EXIT;
                            }
                            */
                        }
                    }
                }

                else if (FirmwareTarget == FW_TARGET.NV_MEM)
                {
                    for (uint i = 0; i < PageSize; i += 4)
                    {
                        uint RcvBytes = ReadRegister(FlashAddress + i);

                        for (int j = 0; j < 4; j++)
                        {
                            //Log.WriteLine($"0x{(FlashAddress + i + j).ToString("X4")} | Write = 0x{SendBytes[i + j].ToString("X2")} | Read = 0x{((RcvBytes >> (8 * j)) & 0xff).ToString("X2")}");

                            if (SendBytes[i + j] != ((RcvBytes >> (8 * j)) & 0xff))
                            {
                                Status = false;
                                Log.WriteLine($"Fail to match : 0x{(((FlashAddress + i + j) >> 16) & 0xFFFF).ToString("X4")}_{((FlashAddress + i + j) & 0xFFFF).ToString("X4")}" +
                                    $"\nWrite = {SendBytes[i + j].ToString("X2")} != Read = {((RcvBytes >> (8 * j)) & 0xff).ToString("X2")}",
                                    System.Drawing.Color.Coral, Log.RichTextBox.BackColor);
                                goto EXIT;
                            }
                        }
                    }

                }
                // Increase progress bar
                ProgressBar?.Invoke((new MethodInvoker(delegate () { ProgressBar.Value++; })));
            }

            Status = true;

        EXIT:

            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                WriteRegister(0x50090008, ((uint)(FLASH_CMD.WRDI) << 24));
                System.Threading.Thread.Sleep(10);
            }

            fs.Close();
            br.Close();

            ResetMCU();

            ProgressBar?.Invoke((new MethodInvoker(delegate () { ProgressBar.Value = ProgressBar.Maximum; })));

            flash_write.Stop();
            Log.WriteLine($"Run Time : {flash_write.Elapsed.TotalSeconds:F2} sec");
            StatusBar.Text = $"Complete";
        }

        public void Oasis_EraseFW()
        {
            Status = true;
            int cnt = 0;
            List<byte> SendBytes = new List<byte>();
            uint sec_addr = 0;

            if (!CheckI2C_ID()) return;

            InitProgressBar(0, 0, 8);

            // 8 sector erase
            for (uint num = 0; num < 8; num++)
            {
                sec_addr = num * 0x10000;
                SetStatusLabel($"Running - Erase @ {sec_addr:X8}");
                sec_addr = ((uint)(FLASH_CMD.BE64) << 24) | (sec_addr & 0xFFFFFF);
                WriteRegister(0x50090008, sec_addr);
                while (Oasis_Check_Staus_Register() != 0x00) // typ 1s, max 4ss
                {
                    cnt++;
                    System.Threading.Thread.Sleep(200);
                    if (cnt > 20) break;
                }
                if (cnt > 20)
                {
                    Status = false;
                    break;
                }
                ProgressBar?.Invoke(new MethodInvoker(delegate { ProgressBar.Value++; }));
            }
            SetStatusLabel("Complete");
        }

        public void Oasis_DumpFW()
        {
            Stopwatch flash_dump = new Stopwatch();
            flash_dump.Start();

            const int PageSize = 4;
            uint RcvBytes;
            List<byte> FirmwareData = new List<byte>();
            string file_name;
            string time;

            Status = false;
            InitProgressBar(0, 0, (ReadFirmwareSize + PageSize - 1) / PageSize + 1);

            if (!CheckI2C_ID()) goto EXIT;

            HaltMCU();

            for (uint Addr = 0; Addr < ReadFirmwareSize; Addr += PageSize)
            {
                if ((Addr % 0x100) == 0)
                {
                    SetStatusLabel($"Running - Read Flash @ 0x{Addr:X8}");
                }

                //RcvBytes = ReadRegister(Addr + 0x80000);  // ROM code exist after remap.
                RcvBytes = ReadRegister(Addr + 0x0);
                FirmwareData.Add((byte)(RcvBytes & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 8) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 16) & 0xff));
                FirmwareData.Add((byte)((RcvBytes >> 24) & 0xff));

                ProgressBar?.Invoke(new MethodInvoker(delegate () { ProgressBar.Value++; }));
            }
            Status = true;
            ReadFirmwareData = FirmwareData.ToArray();
            time = DateTime.Now.ToString("MMddHHmm");
            if (FirmwareTarget == FW_TARGET.NV_MEM)
            {
                file_name = ($"ReadFW_NVM_{time}.bin");
            }
            else
            {
                file_name = ($"ReadFW_RAM_{time}.bin");
            }
            System.IO.FileStream fs = new System.IO.FileStream(file_name, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(ReadFirmwareData);
            bw.Close();
            fs.Close();

            Status = true;

        EXIT:
            ResetMCU();

            ProgressBar?.Invoke(new MethodInvoker(delegate () { ProgressBar.Value = ProgressBar.Maximum; }));

            flash_dump.Stop();
            Log.WriteLine($"Dump Run Time : {flash_dump.Elapsed.TotalSeconds:F2} sec");
            SetStatusLabel("Complete");
        }

        public void FlashVerifyTestCore() => VerifyPatternFlash(RunFlashVerifyTest);

        public void RunFlashVerifyTest()
        {
            uint flashSize;
            int pageSize = 256;
            byte[] pageBuffer = new byte[pageSize];
            byte[] readBuffer = new byte[pageSize];
            byte[] fixedPatterns = new byte[] { 0xAA, 0x55 };

            try
            {
                flashSize = uint.Parse(TextBox_TestArgument.Text);
            }
            catch
            {
                MessageBox.Show("Flash 크기[Byte]를 입력하고 다시 시도해주세요.");
                return;
            }

            string sizeStr = $"{flashSize / 1024:F2} KB";

            DialogResult dialog = MessageBox.Show(
                $"Start to Flash Verify Sequence?\nFLASH Size: {flashSize} bytes ({sizeStr})",
                "FLASH_VERIFY",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dialog != DialogResult.Yes)
            {
                Log.WriteLine("FLASH test has been cancelled by the user.", Color.Orange, Log.RichTextBox.BackColor);
                Status = false;
                return;
            }

            if (!CheckI2C_ID()) return;

            int loopCount = 0;
            Check_Instrument();

            while (true)
            {
                loopCount++;
                Log.WriteLine($"\n=== LOOP #{loopCount} START ===");

                Stopwatch totalTimer = new Stopwatch();
                totalTimer.Start();

                for (int p = 0; p < fixedPatterns.Length; p++)
                {
                    // Power Toggle
                    PowerSupply0.Write("OUTP 0");
                    Thread.Sleep(1000);
                    PowerSupply0.Write("OUTP 1");
                    Thread.Sleep(1000);

                    HaltMCU();

                    byte pattern = fixedPatterns[p];
                    Stopwatch stepTimer = new Stopwatch();
                    stepTimer.Start();

                    uint flashAddress = 0;
                    Log.WriteLine($"\nPattern {p + 1}: 0x{pattern:X2}");

                    #region Erase Flash
                    Log.WriteLine("Start to erase NV memory...");
                    Oasis_EraseFW();
                    if (!Status)
                    {
                        Log.WriteLine("Fail to Erase.", Color.Coral, Log.RichTextBox.BackColor);
                        goto EXIT;
                    }

                    InitProgressBar(0, 0, (int)((flashSize + pageSize - 1) / pageSize));

                    while (flashAddress < flashSize)
                    {
                        if ((flashAddress % 0x100) == 0)
                            SetStatusLabel($"Loop {loopCount} - Erase verify @ 0x{flashAddress:X8}");

                        // Verify
                        for (uint i = 0; i < pageSize; i += 4)
                        {
                            uint data = ReadRegister(flashAddress + i);
                            for (int j = 0; j < 4; j++)
                            {
                                int idx = (int)(i + j);
                                if (idx < pageSize)
                                    readBuffer[idx] = (byte)((data >> (8 * j)) & 0xFF);
                            }
                        }
                        for (int i = 0; i < pageSize; i++)
                        {
                            if (readBuffer[i] != 0xFF)
                            {
                                Log.WriteLine($"Fail to Erase: Addr=0x{(flashAddress + (uint)i):X8}, Read=0x{readBuffer[i]:X2}",
                                    Color.Coral, Log.RichTextBox.BackColor);
                                Status = false;
                                goto EXIT;
                            }
                        }

                        flashAddress += (uint)pageSize;
                        ProgressBar?.Invoke(new MethodInvoker(delegate { ProgressBar.Value++; }));
                    }

                    Log.WriteLine("Succeed to erase.", Color.ForestGreen, Log.RichTextBox.BackColor);
                    #endregion Erase Flash

                    #region Write Flash
                    flashAddress = 0;
                    InitProgressBar(0, 0, (int)((flashSize + pageSize - 1) / pageSize));

                    string binary = System.Convert.ToString(pattern, 2).PadLeft(8, '0');
                    Log.WriteLine("Start to download Firmware...");
                    Log.WriteLine($"→ 0x{pattern:X2}(0b{binary})");

                    while (flashAddress < flashSize)
                    {
                        for (int i = 0; i < pageSize; i++)
                            pageBuffer[i] = pattern;

                        if ((flashAddress % 0x100) == 0)
                            SetStatusLabel($"Loop {loopCount} - Write verify @ 0x{flashAddress:X8}");

                        // Write Flash
                        WriteMemory_NVM(flashAddress, pageBuffer);
                        if (!Status)
                        {
                            Log.WriteLine($"Fail to Write: Step={p + 1}, Addr=0x{flashAddress:X8}", Color.Coral, Log.RichTextBox.BackColor);
                            goto EXIT;
                        }

                        // Verify
                        for (uint i = 0; i < pageSize; i += 4)
                        {
                            uint data = ReadRegister(flashAddress + i);
                            for (int j = 0; j < 4; j++)
                            {
                                int idx = (int)(i + j);
                                if (idx < pageSize)
                                    readBuffer[idx] = (byte)((data >> (8 * j)) & 0xFF);
                            }
                        }
                        for (int i = 0; i < pageSize; i++)
                        {
                            if (readBuffer[i] != pattern)
                            {
                                Log.WriteLine($"Fail to Verify: Step={p + 1}, Addr=0x{(flashAddress + (uint)i):X8}, " +
                                    $"\nW=0x{pattern:X2}, R=0x{readBuffer[i]:X2}", Color.Coral, Log.RichTextBox.BackColor);
                                Status = false;
                                goto EXIT;
                            }
                        }

                        flashAddress += (uint)pageSize;
                        ProgressBar?.Invoke(new MethodInvoker(delegate { ProgressBar.Value++; }));
                    }

                    Log.WriteLine("Succeed to download.", Color.ForestGreen, Log.RichTextBox.BackColor);
                    #endregion Write Flash

                    #region Retention Test
                    for (int d = 5; d <= 100; d += 95)
                    {
                        Log.WriteLine($"Start Power Retention Verify @ {d}ms...");

                        // Power Toggle
                        PowerSupply0.Write("OUTP 0");
                        Thread.Sleep(d);
                        PowerSupply0.Write("OUTP 1");
                        Thread.Sleep(100);

                        HaltMCU();

                        flashAddress = 0;
                        InitProgressBar(0, 0, (int)((flashSize + pageSize - 1) / pageSize));

                        while (flashAddress < flashSize)
                        {
                            if ((flashAddress % 0x100) == 0)
                                SetStatusLabel($"Power Retention Verify @ 0x{flashAddress:X8}");

                            // Verify
                            for (uint i = 0; i < pageSize; i += 4)
                            {
                                uint data = ReadRegister(flashAddress + i);
                                for (int j = 0; j < 4; j++)
                                {
                                    int idx = (int)(i + j);
                                    if (idx < pageSize)
                                        readBuffer[idx] = (byte)((data >> (8 * j)) & 0xFF);
                                }
                            }
                            for (int i = 0; i < pageSize; i++)
                            {
                                if (readBuffer[i] != pattern)
                                {
                                    Log.WriteLine($"Fail to Verify: Addr=0x{(flashAddress + (uint)i):X8}, W=0x{pattern:X2}, R=0x{readBuffer[i]:X2}",
                                        Color.Coral, Log.RichTextBox.BackColor);
                                    Status = false;
                                    goto EXIT;
                                }
                            }

                            flashAddress += (uint)pageSize;
                            ProgressBar?.Invoke(new MethodInvoker(delegate { ProgressBar.Value++; }));
                        }

                        Log.WriteLine($"Power Retention Passed @ {d}ms.", Color.ForestGreen, Log.RichTextBox.BackColor);
                    }
                    #endregion Retention Test

                    stepTimer.Stop();
                    Log.WriteLine($"Run Time: {stepTimer.Elapsed.TotalSeconds:F2} sec");
                }

                totalTimer.Stop();
                Log.WriteLine($"\nLoop #{loopCount} Complete. Total Run Time: {totalTimer.Elapsed.TotalSeconds:F2} sec",
                    Color.ForestGreen, Log.RichTextBox.BackColor);

                ResetMCU();
                SetStatusLabel($"Loop {loopCount} Complete");

                DialogResult result = TimedMessageBox.Show(
                    "Continue to next loop?\n(Automatically proceeds in 5 sec)",
                    "FLASH_VERIFY_PATTERN",
                    5000);

                if (result == DialogResult.No)
                    break;
            }

        EXIT:
            ProgressBar?.Invoke((new MethodInvoker(delegate () { ProgressBar.Value = ProgressBar.Maximum; })));

            ResetMCU();

            SetStatusLabel("Stopped");
        }

        private void ResetOasis()
        {
            Check_Instrument();

            // Power Toggle
            PowerSupply0.Write("OUTP 0");
            Thread.Sleep(1000);
            PowerSupply0.Write("OUTP 1");
            Thread.Sleep(1000);

            HaltMCU();
        }
        #endregion Firmware control methods

        #region CAL_ITEMS
        private double[] Start_ABGR_Trim(double target)
        {
            uint[] bgr_cont = { 8, 9, 10, 11, 12, 13, 14, 15, 0, 1, 2, 3, 4, 5, 6, 7 };
            double bgr_target_mv = target;
            double[] trim_value = { 0, 0 };

            double minDifference = double.MaxValue;
            uint bestTrimCode = 0;
            double bestVoltage = 0;

            try
            {
                Check_Instrument();
                Set_GPIO4_ABGR(true);

                uint avgTime = 5;

                for (int j = 0; j < bgr_cont.Length; j++)
                {
                    uint currentTrimCode = bgr_cont[j];

                    UpdateRegisterBits(0xDC34_040C, 0xFFF0, currentTrimCode);
                    System.Threading.Thread.Sleep(100);

                    double dmm_volt_mv = 0;
                    for(int i = 0; i < avgTime; i++)
                    {
                        dmm_volt_mv += double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                    }
                    dmm_volt_mv /= avgTime;
                    double currentDifference = Math.Abs(dmm_volt_mv - bgr_target_mv);

                    if (currentDifference < minDifference)
                    {
                        minDifference = currentDifference;
                        bestTrimCode = currentTrimCode;
                        bestVoltage = dmm_volt_mv;
                    }
                }

                UpdateRegisterBits(0xDC34_040C, 0xFFF0, bestTrimCode);
                System.Threading.Thread.Sleep(100);
                for (int i = 0; i < avgTime; i++)
                {
                    bestVoltage += double.Parse(DigitalMultimeter2.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;
                }
                bestVoltage /= avgTime;

                trim_value[0] = bestTrimCode;
                trim_value[1] = bestVoltage;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error in Start_ABGR_Trim: {ex.Message}";
                Log.WriteLine(errorMsg);
            }
            finally
            {
                Set_GPIO4_ABGR(false);
            }

            return trim_value;
        }

        private double[] Start_MLDO_Trim(double target)   // DigitalMultimeter 1.
        {
            RegisterItem MLDO_CONT = Parent.RegMgr.GetRegisterItem("O_MLDO_CONT[5:0]");

            uint[] mldo_cont = {
                8, 9, 10, 11, 12, 13, 14, 15, 0, 1, 2, 3, 4, 5, 6, 7,
                24, 25, 26, 27, 28, 29, 30, 31, 16, 17, 18, 19, 20, 21,
                22, 23, 40, 41, 42, 43, 44, 45, 46, 47, 32, 33, 34, 35,
                36, 37, 38, 39, 56, 57, 58, 59, 60, 61, 62, 63, 48, 49,
                50, 51, 52, 53, 54, 55
            };
            double dmm_volt_mv = 0, mldo_target_mv = target;
            int left = 0, mid = 0, right = mldo_cont.Length - 1;
            double[] trim_value = { 0, 0 };
            Check_Instrument();

            while (left <= right)
            {
                mid = (left + right) / 2;
                MLDO_CONT.Read();
                MLDO_CONT.Value = mldo_cont[mid];
                MLDO_CONT.Write();
                System.Threading.Thread.Sleep(10);

                dmm_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;

                if (Math.Abs(dmm_volt_mv - mldo_target_mv) <= 4.5)
                {
                    break;
                }
                if (dmm_volt_mv >= mldo_target_mv)
                {
                    right = mid - 1;
                }
                else if (dmm_volt_mv <= mldo_target_mv)
                {
                    left = mid + 1;
                }
            }

            MLDO_CONT.Value = mldo_cont[mid];
            MLDO_CONT.Write();
            System.Threading.Thread.Sleep(10);

            dmm_volt_mv = double.Parse(DigitalMultimeter1.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;

            trim_value[0] = mldo_cont[mid];
            trim_value[1] = dmm_volt_mv;

            return trim_value;
        }

        private double[] Start_ALDO_Trim(double target)   // DigitalMultimeter 0.
        {
            RegisterItem ALDO_CONT = Parent.RegMgr.GetRegisterItem("O_ALDO_CONT[5:0]");

            uint[] aldo_cont = {
                24, 25, 26, 27, 28, 29, 30, 31,
                16, 17, 18, 19, 20, 21, 22, 23,
                56, 57, 58, 59, 60, 61, 62, 63,
                48, 49, 50, 51, 52, 53, 54, 55,
                8, 9, 10, 11, 12, 13, 14, 15,
                0, 1, 2, 3, 4, 5, 6, 7,
                40, 41, 42, 43, 44, 45, 46, 47,
                32, 33, 34, 35, 36, 37, 38, 39
            };
            double dmm_volt_mv = 0, aldo_target_mv = target;
            int left = 0, mid = 0, right = aldo_cont.Length - 1;
            double[] trim_value = { 0, 0 };

            Check_Instrument();

            while (left <= right)
            {
                mid = (left + right) / 2;
                ALDO_CONT.Read();
                ALDO_CONT.Value = aldo_cont[mid];
                ALDO_CONT.Write();
                System.Threading.Thread.Sleep(10);

                dmm_volt_mv = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;

                if (Math.Abs(dmm_volt_mv - aldo_target_mv) <= 4.5)
                {
                    break;
                }
                if (dmm_volt_mv >= aldo_target_mv)
                {
                    right = mid - 1;
                }
                else if (dmm_volt_mv <= aldo_target_mv)
                {
                    left = mid + 1;
                }
            }

            ALDO_CONT.Value = aldo_cont[mid];
            ALDO_CONT.Write();
            System.Threading.Thread.Sleep(10);

            dmm_volt_mv = double.Parse(DigitalMultimeter0.WriteAndReadString(":MEAS:VOLT:DC?")) * 1000;

            trim_value[0] = aldo_cont[mid];
            trim_value[1] = dmm_volt_mv;

            return trim_value;
        }
        #endregion CAL_ITEMS
    }
}