using FTD2XX_NET;
using JLcLib.Comn;
using System.Collections.Generic;
using System.Threading;
using static SKAIChips_Verification.RegContForm;

namespace JLcLib.Custom
{
    public class SPI : IComn
    {
        public ChipTypes CurrentChipType { get; private set; } = ChipTypes.Unknown;

        public void SetChipType(ChipTypes type) => CurrentChipType = type;

        public bool IsChicago => CurrentChipType == ChipTypes.Chicago;

        public class Setting
        {
            public string DevName;

            public int DevIndex;

            public int Clock_kHz;

            public int Mode;

            public int BitFirst;
        }

        public Setting Config = new Setting
        {
            DevName = "FT232H",
            DevIndex = 0,
            Mode = 0,
            Clock_kHz = 1000
        };

        public FT4222H ft4222H { get; set; }

        public MPSSE ftMPSSE { get; set; }

        public bool IsOpen { get; private set; }

        public WireComnTypes ComnType { get; }

        public string DeviceName => Config.DevName;

        public List<string> DevicesNames { get; private set; } = new List<string>();

        public string StatusMessage { get; private set; }

        public GPIO_Collection GPIOs { get; set; }

        public ManualResetEvent BusyEvent { get; set; }

        public SPI()
        {
            SetAvailableDevices();
            ComnType = WireComnTypes.SPI;
        }

        private void SetAvailableDevices()
        {
            DevicesNames.Clear();
            foreach (string devName in FT_Device.DevNames)
            {
                DevicesNames.Add(devName);
            }
        }

        public List<string> Search()
        {
            FT_Device.FindAvailableDevices();
            SetAvailableDevices();
            return DevicesNames;
        }

        public bool IsAvailable()
        {
            foreach (string devicesName in DevicesNames)
            {
                if (devicesName == DeviceName)
                {
                    return true;
                }
            }

            return false;
        }

        public bool Open()
        {
            if (IsOpen)
            {
                Close();
            }

            if (StringCtrl.CompareIn("FT4222 A", Config.DevName) >= 0 || StringCtrl.CompareIn("FT4222 B", Config.DevName) >= 0)
            {
                FTDI.FT_DEVICE_INFO_NODE device = FT_Device.GetDevice(Config.DevName);
                ft4222H = new FT4222H();
                if (ft4222H.Open(device))
                {
                    ft4222H.SPI_InitMaster(Config.Clock_kHz, Config.Mode);
                    GPIOs = ft4222H.GPIOs;
                    IsOpen = ft4222H.IsOpen;
                    BusyEvent = ft4222H.BusyEvent;
                }
                else
                {
                    ft4222H = null;
                }
            }
            else
            {
                FTDI.FT_DEVICE_INFO_NODE device2 = FT_Device.GetDevice(Config.DevName);
                ftMPSSE = new MPSSE();
                if (ftMPSSE.Open(device2))
                {
                    ftMPSSE.SPI_Init(Config.Mode, Config.Clock_kHz, (Config.BitFirst != 0) ? MPSSE.Command.BitFirst.LSB : MPSSE.Command.BitFirst.MSB);
                    GPIOs = ftMPSSE.GPIOs;
                    IsOpen = ftMPSSE.IsOpen;
                    BusyEvent = ftMPSSE.BusyEvent;
                    if (IsChicago)
                    {
                        SetGPIOLHigh();
                    }

                }
            }

            if (IsOpen)
            {
                StatusMessage = Config.DevName + "-SPI-" + Config.Clock_kHz + "kHz";
            }
            else
            {
                Close();
                StatusMessage = Config.DevName + " was denied";
            }

            return IsOpen;
        }

        public bool Close()
        {
            if (ft4222H != null)
            {
                ft4222H.Close();
            }

            if (ftMPSSE != null)
            {
                ftMPSSE.Close();
            }

            ft4222H = null;
            ftMPSSE = null;
            IsOpen = false;
            return IsOpen;
        }

        public void QueueClear()
        {
        }

        public void WriteSettingFile(IniFile iniFile, string Section)
        {
            iniFile.Write(Section + "_SPI_Setting", "DevName", Config.DevName);
            iniFile.Write(Section + "_SPI_Setting", "Clock_kHz", Config.Clock_kHz.ToString());
            iniFile.Write(Section + "_SPI_Setting", "Mode", Config.Mode.ToString());
            iniFile.Write(Section + "_SPI_Setting", "BitFirst", Config.BitFirst.ToString());
            iniFile.Write(Section + "_SPI_Setting", "DevIndex", Config.DevIndex.ToString());
        }

        public Setting ReadSettingFile(IniFile iniFile, string Section)
        {
            int Value = 0;
            string text = iniFile.Read(Section + "_SPI_Setting", "DevName");
            if (text != null && text != "")
            {
                Config.DevName = text;
            }

            if (iniFile.Read(Section + "_SPI_Setting", "Clock_kHz", ref Value))
            {
                Config.Clock_kHz = Value;
            }

            if (iniFile.Read(Section + "_SPI_Setting", "Mode", ref Value))
            {
                Config.Mode = Value;
            }

            if (iniFile.Read(Section + "_SPI_Setting", "BitFirst", ref Value))
            {
                Config.BitFirst = Value;
            }

            if (iniFile.Read(Section + "_SPI_Setting", "DevIndex", ref Value))
            {
                Config.DevIndex = Value;
            }

            return Config;
        }

        public byte[] ReadBytes(int NumBytesToRead)
        {
            byte[] ReadBuf = null;
            if (ft4222H != null)
            {
                ReadBuf = new byte[NumBytesToRead];
                ft4222H.SPI_ReadBytes(ref ReadBuf, (ushort)NumBytesToRead, SSEndTrans: true);
            }
            else if (ftMPSSE != null)
            {
                ftMPSSE.SPI_SetStart();
                ftMPSSE.SPI_SetBytes(null, (ushort)NumBytesToRead, MPSSE.Command.PinConfig.Read);
                ftMPSSE.SPI_SetStop();
                ftMPSSE.SendCommand(SendAnswerBackImmediately: true);
                ReadBuf = ftMPSSE.GetReceivedBytes(NumBytesToRead);
            }

            return ReadBuf;
        }

        public void WriteBytes(byte[] WriteBytes, int NumBytesToWrite, bool stop)
        {
            if (ft4222H != null)
            {
                ft4222H.SPI_WriteBytes(WriteBytes, (ushort)NumBytesToWrite, SSEndTrans: true);
            }
            else if (ftMPSSE != null)
            {
                ftMPSSE.SPI_SetStart();
                ftMPSSE.SPI_SetBytes(WriteBytes, (ushort)NumBytesToWrite, MPSSE.Command.PinConfig.Write);
                ftMPSSE.SPI_SetStop();
                ftMPSSE.SendCommand();
            }
        }

        public byte[] ReadWriteBytes(byte[] WriteBytes, int Length)
        {
            byte[] ReadBuf = null;
            if (ft4222H != null)
            {
                ReadBuf = new byte[Length];
                ft4222H.SPI_ReadWriteBytes(ref ReadBuf, WriteBytes, (ushort)Length, SSEndTrans: true);
            }
            else if (ftMPSSE != null)
            {
                ftMPSSE.SPI_SetStart();
                ftMPSSE.SPI_SetBytes(WriteBytes, (ushort)Length, MPSSE.Command.PinConfig.ReadWrite);
                ftMPSSE.SPI_SetStop();
                ftMPSSE.SendCommand(SendAnswerBackImmediately: true);
                ReadBuf = ftMPSSE.GetReceivedBytes(Length);
            }

            return ReadBuf;
        }

        public byte[] WriteAndReadBytes(byte[] WriteBytes, int NumBytesToWrite, int NumBytesToRead)
        {
            byte[] ReadBuf = null;
            if (ft4222H != null)
            {
                ReadBuf = new byte[NumBytesToRead];
                ft4222H.SPI_WriteBytes(WriteBytes, (ushort)NumBytesToWrite, SSEndTrans: false);
                ft4222H.SPI_ReadBytes(ref ReadBuf, (ushort)NumBytesToRead, SSEndTrans: true);
            }
            else if (ftMPSSE != null)
            {
                ftMPSSE.SPI_SetStart();
                ftMPSSE.SPI_SetBytes(WriteBytes, (ushort)NumBytesToWrite, MPSSE.Command.PinConfig.Write);
                ftMPSSE.SPI_SetBytes(null, (ushort)NumBytesToRead, MPSSE.Command.PinConfig.Read);
                ftMPSSE.SPI_SetStop();
                ftMPSSE.SendCommand(SendAnswerBackImmediately: true);
                ReadBuf = ftMPSSE.GetReceivedBytes(NumBytesToRead);
            }

            return ReadBuf;
        }

        public void WriteBytesForChicago(byte[] WriteBytes, int NumBytesToWrite, bool stop)
        {
            if (ftMPSSE != null)
            {
                ftMPSSE.SPI_SetStart();
                ftMPSSE.SPI_SetBytes(WriteBytes, (ushort)NumBytesToWrite, MPSSE.Command.PinConfig.Write);
                ftMPSSE.SPI_SetStop();
                ftMPSSE.SendCommand();
                SetGPIOLHigh();
            }
        }

        public byte[] WriteAndReadBytesForChicago(byte[] WriteBytes, int NumBytesToWrite, int NumBytesToRead)
        {
            byte[] ReadBuf = null;
            if (ftMPSSE != null)
            {
                ftMPSSE.SPI_SetStart();
                ftMPSSE.SPI_SetBytes(WriteBytes, (ushort)NumBytesToWrite, MPSSE.Command.PinConfig.Write);
                ftMPSSE.SPI_HalfDuplex();
                ftMPSSE.SPI_SetBytes(null, (ushort)NumBytesToRead, MPSSE.Command.PinConfig.Read);
                ftMPSSE.SPI_SetStop();
                ftMPSSE.SendCommand(SendAnswerBackImmediately: false);
                ReadBuf = ftMPSSE.GetReceivedBytes(NumBytesToRead);
                SetGPIOLHigh();
            }

            return ReadBuf;
        }

        public void SetGPIOLHigh()
        {
            if (ftMPSSE != null)
            {
                ftMPSSE.GPIOL_SetPins(251, 255, Protected: false);
                ftMPSSE.SendCommand();
            }
        }

        public void SetSPIStop()
        {
            ftMPSSE.SPI_SetStop();
            ftMPSSE.SendCommand(SendAnswerBackImmediately: false);
        }

        public void Toggle_GPIO3()
        {
            if (ft4222H != null)
            {
                ft4222H.GPIO_SetDirection(3, GPIO_Direction.Output);
                ft4222H.GPIO_SetState(3, GPIO_State.Low);
                Thread.Sleep(1000);

                ft4222H.GPIO_SetDirection(3, GPIO_Direction.Input);
                Thread.Sleep(1000);
            }
        }
    }
}
