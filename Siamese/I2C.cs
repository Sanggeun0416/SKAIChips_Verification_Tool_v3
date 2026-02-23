#region 어셈블리 JLcLib, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// G:\Verification_Tools\SKAIChips_TestGUI_V3.1.1_Source\Siamese\bin\Debug\JLcLib.dll
// Decompiled with ICSharpCode.Decompiler 8.2.0.7535
#endregion

using FTD2XX_NET;
using JLcLib.Comn;
using System.Collections.Generic;
using System.Threading;

namespace JLcLib.Custom
{
    public class I2C : IComn
    {
        public class Setting
        {
            public string DevName;

            public int DevIndex;

            public int Clock_kHz;

            public int SlaveAddress;
        }

        public Setting Config = new Setting
        {
            DevName = "FT4222H A",
            DevIndex = 0,
            Clock_kHz = 400,
            SlaveAddress = 36
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

        public I2C()
        {
            SetAvailableDevices();
            ComnType = WireComnTypes.I2C;
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
                    ft4222H.I2C_InitMaster(Config.Clock_kHz);
                    GPIOs = ft4222H.GPIOs;
                    IsOpen = ft4222H.IsOpen;
                    BusyEvent = ft4222H.BusyEvent;

                    ft4222H.GPIO_SetDirection(2, GPIO_Direction.Output);
                    ft4222H.GPIO_SetState(2, GPIO_State.High);
                }
            }
            else
            {
                FTDI.FT_DEVICE_INFO_NODE device2 = FT_Device.GetDevice(Config.DevName);
                ftMPSSE = new MPSSE();
                if (ftMPSSE.Open(device2))
                {
                    ftMPSSE.I2C_Init(Config.Clock_kHz);
                    GPIOs = ftMPSSE.GPIOs;
                    IsOpen = ftMPSSE.IsOpen;
                    BusyEvent = ftMPSSE.BusyEvent;
                }
            }

            if (IsOpen)
            {
                StatusMessage = Config.DevName + "-I2C-" + Config.SlaveAddress.ToString("X2") + "-" + Config.Clock_kHz + "kHz";
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
            iniFile.Write(Section + "_I2C_Setting", "DevName", Config.DevName);
            iniFile.Write(Section + "_I2C_Setting", "Clock_kHz", Config.Clock_kHz.ToString());
            iniFile.Write(Section + "_I2C_Setting", "SlaveAddress", Config.SlaveAddress.ToString());
            iniFile.Write(Section + "_I2C_Setting", "DevIndex", Config.DevIndex.ToString());
        }

        public Setting ReadSettingFile(IniFile iniFile, string Section)
        {
            int Value = 0;
            string text = iniFile.Read(Section + "_I2C_Setting", "DevName");
            if (text != null && text != "")
            {
                Config.DevName = text;
            }

            if (iniFile.Read(Section + "_I2C_Setting", "Clock_kHz", ref Value))
            {
                Config.Clock_kHz = Value;
            }

            if (iniFile.Read(Section + "_I2C_Setting", "SlaveAddress", ref Value))
            {
                Config.SlaveAddress = Value;
            }

            if (iniFile.Read(Section + "_I2C_Setting", "DevIndex", ref Value))
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
                ft4222H.I2C_ReadBytes((byte)Config.SlaveAddress, ref ReadBuf, (ushort)NumBytesToRead);
            }
            else if (ftMPSSE != null)
            {
                ReadBuf = ftMPSSE.I2C_ReadBytes((byte)Config.SlaveAddress, NumBytesToRead);
            }

            return ReadBuf;
        }

        public void WriteBytes(byte[] WriteBytes, int NumBytesToWrite, bool stop)
        {
            if (ft4222H != null)
            {
                ft4222H.I2C_WriteBytes((byte)Config.SlaveAddress, WriteBytes, (ushort)NumBytesToWrite);
            }
            else if (ftMPSSE != null)
            {
                ftMPSSE.I2C_WriteBytes((byte)Config.SlaveAddress, WriteBytes, NumBytesToWrite, stop);
            }
        }

        public byte[] ReadWriteBytes(byte[] WriteBuffer, int Length)
        {
            return null;
        }

        public byte[] WriteAndReadBytes(byte[] WriteBytes, int NumBytesToWrite, int NumBytesToRead)
        {
            byte[] ReadBuf = null;
            ReadBuf = new byte[NumBytesToRead];
            if (ft4222H != null)
            {
                ft4222H.I2C_WriteBytes_NoStop((byte)Config.SlaveAddress, WriteBytes, (ushort)NumBytesToWrite);
                ft4222H.I2C_ReadBytes((byte)Config.SlaveAddress, ref ReadBuf, (ushort)NumBytesToRead);
            }
            else if (ftMPSSE != null)
            {
                ftMPSSE.I2C_WriteBytes((byte)Config.SlaveAddress, WriteBytes, NumBytesToWrite, false);
                ReadBuf = ftMPSSE.I2C_ReadBytes((byte)Config.SlaveAddress, NumBytesToRead);
            }

            return ReadBuf;
        }
    }

}
