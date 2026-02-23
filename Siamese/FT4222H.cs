using FTD2XX_NET;
using JLcLib.Comn;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace JLcLib.Custom
{
    public class FT4222H
    {
        public enum ConnectionType
        {
            SPI_I2C,
            GPIO
        }

        private enum STATUS
        {
            FT4222_OK = 0,
            FT4222_INVALID_HANDLE = 1,
            FT4222_DEVICE_NOT_FOUND = 2,
            FT4222_DEVICE_NOT_OPENED = 3,
            FT4222_IO_ERROR = 4,
            FT4222_INSUFFICIENT_RESOURCES = 5,
            FT4222_INVALID_PARAMETER = 6,
            FT4222_INVALID_BAUD_RATE = 7,
            FT4222_DEVICE_NOT_OPENED_FOR_ERASE = 8,
            FT4222_DEVICE_NOT_OPENED_FOR_WRITE = 9,
            FT4222_FAILED_TO_WRITE_DEVICE = 10,
            FT4222_EEPROM_READ_FAILED = 11,
            FT4222_EEPROM_WRITE_FAILED = 12,
            FT4222_EEPROM_ERASE_FAILED = 13,
            FT4222_EEPROM_NOT_PRESENT = 14,
            FT4222_EEPROM_NOT_PROGRAMMED = 15,
            FT4222_INVALID_ARGS = 16,
            FT4222_NOT_SUPPORTED = 17,
            FT4222_OTHER_ERROR = 18,
            FT4222_DEVICE_LIST_NOT_READY = 19,
            FT4222_DEVICE_NOT_SUPPORTED = 1000,
            FT4222_CLK_NOT_SUPPORTED = 1001,
            FT4222_VENDER_CMD_NOT_SUPPORTED = 1002,
            FT4222_IS_NOT_SPI_MODE = 1003,
            FT4222_IS_NOT_I2C_MODE = 1004,
            FT4222_IS_NOT_SPI_SINGLE_MODE = 1005,
            FT4222_IS_NOT_SPI_MULTI_MODE = 1006,
            FT4222_WRONG_I2C_ADDR = 1007,
            FT4222_INVAILD_FUNCTION = 1008,
            FT4222_INVALID_POINTER = 1009,
            FT4222_EXCEEDED_MAX_TRANSFER_SIZE = 1010,
            FT4222_FAILED_TO_READ_DEVICE = 1011,
            FT4222_I2C_NOT_SUPPORTED_IN_THIS_MODE = 1012,
            FT4222_GPIO_NOT_SUPPORTED_IN_THIS_MODE = 1013,
            FT4222_GPIO_EXCEEDED_MAX_PORTNUM = 1014,
            FT4222_GPIO_WRITE_NOT_SUPPORTED = 1015,
            FT4222_GPIO_PULLUP_INVALID_IN_INPUTMODE = 1016,
            FT4222_GPIO_PULLDOWN_INVALID_IN_INPUTMODE = 1017,
            FT4222_GPIO_OPENDRAIN_INVALID_IN_OUTPUTMODE = 1018,
            FT4222_INTERRUPT_NOT_SUPPORTED = 1019,
            FT4222_GPIO_INPUT_NOT_SUPPORTED = 1020,
            FT4222_EVENT_NOT_SUPPORTED = 1021
        }

        public enum OperationMode
        {
            SPI_I2C_GPIO,
            SPI_3_SLV_GPIO,
            SPI_4_SLV,
            SPI_I2C
        }

        public enum SystemClock
        {
            SYS_CLK_60,
            SYS_CLK_24,
            SYS_CLK_48,
            SYS_CLK_80
        }

        public enum SPI_Mode
        {
            SPI_IO_NONE = 0,
            SPI_IO_SINGLE = 1,
            SPI_IO_DUAL = 2,
            SPI_IO_QUAD = 4
        }

        public enum SPI_ClockDivider
        {
            CLK_NONE,
            CLK_DIV_2,
            CLK_DIV_4,
            CLK_DIV_8,
            CLK_DIV_16,
            CLK_DIV_32,
            CLK_DIV_64,
            CLK_DIV_128,
            CLK_DIV_256,
            CLK_DIV_512
        }

        public enum SPI_SCKRate
        {
            SCK_40M = 40000,
            SCK_30M = 30000,
            SCK_24M = 24000,
            SCK_20M = 20000,
            SCK_15M = 15000,
            SCK_12M = 12000,
            SCK_10M = 10000,
            SCK_7_5M = 7500,
            SCK_6M = 6000,
            SCK_5M = 5000,
            SCK_3_75M = 3750,
            SCK_3M = 3000,
            SCK_2_5M = 2500,
            SCK_1_875M = 1875,
            SCK_1_5M = 1500,
            SCK_1_25M = 1250,
            SCK_937K = 937,
            SCK_750K = 750,
            SCK_625K = 625,
            SCK_469K = 469
        }

        public enum SPI_ClkPolarity
        {
            ACTIVE_LOW,
            ACTIVE_HIGH
        }

        public enum SPI_ClkPhase
        {
            CLK_LEADING,
            CLK_TRAILING
        }

        public enum SPI_DrivingStrength
        {
            DS_4MA,
            DS_8MA,
            DS_12MA,
            DS_16MA
        }

        public enum GPIO_Port
        {
            PORT0,
            PORT1,
            PORT2,
            PORT3
        }

        public enum GPIO_Direction
        {
            OUTPUT,
            INPUT
        }

        public enum GPIO_Trigger
        {
            RISING = 1,
            FALLING = 2,
            LEVEL_HIGH = 4,
            LEVEL_LOW = 8
        }

        public enum GPIO_Output
        {
            LOW,
            HIGH
        }

        public enum I2C_MasterFlag
        {
            NONE = 128,
            START = 2,
            Repeated_START = 3,
            STOP = 4,
            START_AND_STOP = 6
        }

        private const byte FT_OPEN_BY_SERIAL_NUMBER = 1;

        private const byte FT_OPEN_BY_DESCRIPTION = 2;

        private const byte FT_OPEN_BY_LOCATION = 4;

        private IntPtr ComnHandle = default(IntPtr);

        private IntPtr GpioHandle = default(IntPtr);

        private FTDI.FT_DEVICE_INFO_NODE ComnDevice;

        private FTDI.FT_DEVICE_INFO_NODE GpioDevice;

        public ManualResetEvent BusyEvent { get; set; } = new ManualResetEvent(initialState: true);


        public bool IsOpen { get; private set; } = false;


        public GPIO_Direction[] GpioDirection { get; set; } = new GPIO_Direction[4]
        {
        GPIO_Direction.INPUT,
        GPIO_Direction.INPUT,
        GPIO_Direction.OUTPUT,
        GPIO_Direction.INPUT
        };

        public int SPI_Clock { get; private set; } = 10000;

        public int I2C_Clock { get; private set; } = 400;

        public GPIO_Collection GPIOs { get; set; }

        [DllImport("ftd2xx.dll")]
        private static extern FTDI.FT_STATUS FT_OpenEx(uint pvArg1, int dwFlags, ref IntPtr ftHandle);

        [DllImport("ftd2xx.dll")]
        private static extern FTDI.FT_STATUS FT_Close(IntPtr ftHandle);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_UnInitialize(IntPtr ftHandle);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SetClock(IntPtr ftHandle, SystemClock clk);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_GetClock(IntPtr ftHandle, ref SystemClock clk);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SetSuspendOut(IntPtr ftHandle, bool enable);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SetWakeUpInterrupt(IntPtr ftHandle, bool enable);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SetInterruptTrigger(IntPtr ftHandle, GPIO_Trigger trigger);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_GetMaxTransferSize(IntPtr ftHandle, ref ushort MaxSize);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPIMaster_Init(IntPtr ftHandle, SPI_Mode ioLine, SPI_ClockDivider clock, SPI_ClkPolarity cpol, SPI_ClkPhase cpha, byte ssoMap);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPIMaster_SetLines(IntPtr ftHandle, SPI_Mode Mode);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPIMaster_SingleRead(IntPtr ftHandle, ref byte buffer, ushort bytesToRead, ref ushort sizeOfRead, bool isEndTransaction);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPIMaster_SingleWrite(IntPtr ftHandle, ref byte buffer, ushort bytesToWrite, ref ushort sizeTransferred, bool isEndTransaction);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPIMaster_SingleReadWrite(IntPtr ftHandle, ref byte readBuffer, ref byte writeBuffer, ushort bufferSize, ref ushort sizeTransferred, bool isEndTransaction);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPIMaster_MultiReadWrite(IntPtr ftHandle, ref byte readBuffer, ref byte writeBuffer, byte singleWriteBytes, ushort multiWriteBytes, ushort multiReadBytes, ref uint sizeOfRead);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPISlave_Init(IntPtr ftHandle);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPISlave_GetRxStatus(IntPtr ftHandle, ref ushort RxSize);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPISlave_Read(IntPtr ftHandle, ref byte buffer, ushort bytesToRead, ref ushort sizeOfRead);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPISlave_Write(IntPtr ftHandle, ref byte buffer, ushort bytesToWrite, ref ushort sizeTransferred);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPI_ResetTransaction(IntPtr ftHandle, byte spiIdx);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPI_Reset(IntPtr ftHandle);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_SPI_SetDrivingStrength(IntPtr ftHandle, SPI_DrivingStrength clkStrength, SPI_DrivingStrength ioStrength, SPI_DrivingStrength ssoStregth);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_I2CMaster_Init(IntPtr ftHandle, uint kbps);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_I2CMaster_Read(IntPtr ftHandle, ushort slaveAddress, ref byte buffer, ushort bytesToRead, ref ushort sizeOfRead);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_I2CMaster_Write(IntPtr ftHandle, ushort slaveAddress, ref byte buffer, ushort bytesToWrite, ref ushort sizeTransferred);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_I2CMaster_WriteEx(IntPtr ftHandle, ushort slaveAddress, uint flag, ref byte buffer, ushort bytesToWrite, ref ushort sizeTransferred);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_I2CMaster_Reset(IntPtr ftHandle);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_GPIO_Init(IntPtr ftHandle, GPIO_Direction[] GpioDir);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_GPIO_SetInputTrigger(IntPtr ftHandle, GPIO_Port PortNum, GPIO_Trigger Trigger);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_GPIO_Read(IntPtr ftHandle, GPIO_Port PortNum, ref bool Value);

        [DllImport("LibFT4222-64.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern STATUS FT4222_GPIO_Write(IntPtr ftHandle, GPIO_Port PortNum, bool Value);

        public FT4222H()
        {
            GPIOs = new GPIO_Collection(GPIO_SetDirection, GPIO_SetState, GPIO_GetState)
        {
            {
                0,
                "GPIO0",
                JLcLib.Comn.GPIO_Direction.Input,
                GPIO_State.Low
            },
            {
                1,
                "GPIO1",
                JLcLib.Comn.GPIO_Direction.Input,
                GPIO_State.Low
            },
            {
                2,
                "GPIO2",
                JLcLib.Comn.GPIO_Direction.Output,
                GPIO_State.Low
            },
            {
                3,
                "GPIO3",
                JLcLib.Comn.GPIO_Direction.Input,
                GPIO_State.Low
            }
        };
        }

        ~FT4222H()
        {
            Close();
        }

        private bool CheckAvailableDevices(FTDI.FT_DEVICE_INFO_NODE Device)
        {
            ComnDevice = null;
            GpioDevice = null;
            foreach (FTDI.FT_DEVICE_INFO_NODE fT4222H in FT_Device.FT4222Hs)
            {
                if (fT4222H.Description == "FT4222 A" && (fT4222H.LocId & 0xFFF0) == (Device.LocId & 0xFFF0))
                {
                    ComnDevice = fT4222H;
                }
                else if (fT4222H.Description == "FT4222 B" && (fT4222H.LocId & 0xFFF0) == (Device.LocId & 0xFFF0))
                {
                    GpioDevice = fT4222H;
                }
            }

            if (ComnDevice != null && GpioDevice != null)
            {
                return true;
            }

            return false;
        }

        private bool CheckAvailableDevices()
        {
            ComnDevice = null;
            GpioDevice = null;
            foreach (FTDI.FT_DEVICE_INFO_NODE fT4222H in FT_Device.FT4222Hs)
            {
                if (fT4222H.Description == "FT4222 A")
                {
                    ComnDevice = fT4222H;
                }
                else if (fT4222H.Description == "FT4222 B")
                {
                    GpioDevice = fT4222H;
                }
            }

            if (ComnDevice != null && GpioDevice != null)
            {
                return true;
            }

            return false;
        }

        private bool WaitBusyEvent(int Timeout_msec, string DebugMessage)
        {
            if (BusyEvent.WaitOne(Timeout_msec))
            {
                BusyEvent.Reset();
                return true;
            }

            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff") + "> FT4222:BDF:" + DebugMessage);
            return false;
        }

        public bool Open(FTDI.FT_DEVICE_INFO_NODE Device)
        {
            if (Device == null)
            {
                return IsOpen;
            }

            if (IsOpen)
            {
                Close();
            }

            if (CheckAvailableDevices(Device) && FT_OpenEx(ComnDevice.LocId, 4, ref ComnHandle) == FTDI.FT_STATUS.FT_OK)
            {
                FT4222_SetClock(ComnHandle, SystemClock.SYS_CLK_60);
                if (FT_OpenEx(GpioDevice.LocId, 4, ref GpioHandle) == FTDI.FT_STATUS.FT_OK)
                {
                    FT4222_SetClock(GpioHandle, SystemClock.SYS_CLK_60);
                    if (FT4222_GPIO_Init(GpioHandle, GpioDirection) == STATUS.FT4222_OK)
                    {
                        IsOpen = true;
                        return IsOpen;
                    }
                }
            }

            Close();
            return IsOpen;
        }

        public bool Open()
        {
            if (IsOpen)
            {
                Close();
            }

            if (CheckAvailableDevices() && FT_OpenEx(ComnDevice.LocId, 4, ref ComnHandle) == FTDI.FT_STATUS.FT_OK)
            {
                FT4222_SetClock(ComnHandle, SystemClock.SYS_CLK_60);
                if (FT_OpenEx(GpioDevice.LocId, 4, ref GpioHandle) == FTDI.FT_STATUS.FT_OK)
                {
                    FT4222_SetClock(GpioHandle, SystemClock.SYS_CLK_60);
                    if (FT4222_GPIO_Init(GpioHandle, GpioDirection) == STATUS.FT4222_OK)
                    {
                        IsOpen = true;
                        return IsOpen;
                    }
                }
            }

            Close();
            return IsOpen;
        }

        public bool Close()
        {
            if (ComnDevice != null && ComnHandle != IntPtr.Zero)
            {
                FT4222_UnInitialize(ComnHandle);
                FT_Close(ComnHandle);
            }

            if (GpioDevice != null && GpioHandle != IntPtr.Zero)
            {
                FT4222_UnInitialize(GpioHandle);
                FT_Close(GpioHandle);
            }

            ComnDevice = null;
            ComnHandle = IntPtr.Zero;
            GpioDevice = null;
            GpioHandle = IntPtr.Zero;
            IsOpen = false;
            return IsOpen;
        }

        private void SPI_GetClockParamter(int SckRate, ref SystemClock SysClk, ref SPI_ClockDivider ClkDiv)
        {
            SysClk = SystemClock.SYS_CLK_80;
            ClkDiv = SPI_ClockDivider.CLK_DIV_8;
            for (int i = 1; i < 8; i++)
            {
                if (SckRate / (80000 / (1 << i)) > 0)
                {
                    SysClk = SystemClock.SYS_CLK_80;
                    ClkDiv = (SPI_ClockDivider)i;
                    break;
                }

                if (SckRate / (60000 / (1 << i)) > 0)
                {
                    SysClk = SystemClock.SYS_CLK_60;
                    ClkDiv = (SPI_ClockDivider)i;
                    break;
                }

                if (SckRate / (48000 / (1 << i)) > 0)
                {
                    SysClk = SystemClock.SYS_CLK_48;
                    ClkDiv = (SPI_ClockDivider)i;
                    break;
                }
            }
        }

        public bool SPI_InitMaster(int SckRate, int Mode)
        {
            SystemClock SysClk = SystemClock.SYS_CLK_80;
            SPI_ClockDivider ClkDiv = SPI_ClockDivider.CLK_DIV_8;
            SPI_ClkPolarity cpol = SPI_ClkPolarity.ACTIVE_LOW;
            SPI_ClkPhase cpha = SPI_ClkPhase.CLK_LEADING;
            bool result = false;
            SPI_GetClockParamter(SckRate, ref SysClk, ref ClkDiv);
            if (Mode == 2 || Mode == 3)
            {
                cpol = SPI_ClkPolarity.ACTIVE_HIGH;
            }

            if (Mode == 1 || Mode == 3)
            {
                cpha = SPI_ClkPhase.CLK_TRAILING;
            }

            WaitBusyEvent(200, "SPI_InitMaster");
            if (FT4222_SetClock(ComnHandle, SysClk) == STATUS.FT4222_OK && FT4222_SPIMaster_Init(ComnHandle, SPI_Mode.SPI_IO_SINGLE, ClkDiv, cpol, cpha, 1) == STATUS.FT4222_OK)
            {
                FT4222_SPI_SetDrivingStrength(ComnHandle, SPI_DrivingStrength.DS_4MA, SPI_DrivingStrength.DS_4MA, SPI_DrivingStrength.DS_4MA);
                FT4222_SetSuspendOut(ComnHandle, enable: false);
                FT4222_SetWakeUpInterrupt(ComnHandle, enable: false);
                switch (SysClk)
                {
                    case SystemClock.SYS_CLK_80:
                        SPI_Clock = 80000 / (1 << (int)ClkDiv);
                        break;
                    case SystemClock.SYS_CLK_60:
                        SPI_Clock = 60000 / (1 << (int)ClkDiv);
                        break;
                    case SystemClock.SYS_CLK_48:
                        SPI_Clock = 48000 / (1 << (int)ClkDiv);
                        break;
                }

                result = true;
            }

            BusyEvent.Set();
            return result;
        }

        public bool SPI_ResetTransaction(OperationMode Mode)
        {
            bool result = false;
            WaitBusyEvent(200, "SPI_ResetTransaction");
            if (FT4222_SPI_ResetTransaction(ComnHandle, (byte)Mode) == STATUS.FT4222_OK)
            {
                result = true;
            }

            BusyEvent.Set();
            return result;
        }

        public bool SPI_Reset()
        {
            bool result = false;
            WaitBusyEvent(200, "SPI_Reset");
            if (FT4222_SPI_Reset(ComnHandle) == STATUS.FT4222_OK)
            {
                result = true;
            }

            BusyEvent.Set();
            return result;
        }

        public ushort SPI_WriteBytes(byte[] WriteBuf, ushort Length, bool SSEndTrans)
        {
            ushort sizeTransferred = 0;
            WaitBusyEvent(200, "SPI_WriteBytes");
            FT4222_SPIMaster_SingleWrite(ComnHandle, ref WriteBuf[0], Length, ref sizeTransferred, SSEndTrans);
            BusyEvent.Set();
            return sizeTransferred;
        }

        public ushort SPI_ReadBytes(ref byte[] ReadBuf, ushort Length, bool SSEndTrans)
        {
            ushort sizeOfRead = 0;
            WaitBusyEvent(200, "SPI_ReadBytes");
            FT4222_SPIMaster_SingleRead(ComnHandle, ref ReadBuf[0], Length, ref sizeOfRead, SSEndTrans);
            BusyEvent.Set();
            return sizeOfRead;
        }

        public ushort SPI_ReadWriteBytes(ref byte[] ReadBuf, byte[] WriteBuf, ushort Length, bool SSEndTrans)
        {
            ushort sizeTransferred = 0;
            WaitBusyEvent(200, "SPI_ReadWriteBytes");
            try
            {
                FT4222_SPIMaster_SingleReadWrite(ComnHandle, ref ReadBuf[0], ref WriteBuf[0], Length, ref sizeTransferred, SSEndTrans);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            BusyEvent.Set();
            return sizeTransferred;
        }

        public bool I2C_InitMaster(int SclRate)
        {
            bool result = false;
            WaitBusyEvent(200, "I2C_InitMaster");
            if (FT4222_I2CMaster_Init(ComnHandle, (uint)SclRate) == STATUS.FT4222_OK)
            {
                FT4222_SetSuspendOut(ComnHandle, enable: false);
                FT4222_SetWakeUpInterrupt(ComnHandle, enable: false);
                I2C_Clock = SclRate;
                result = true;
            }

            BusyEvent.Set();
            return result;
        }

        public bool I2C_Reset()
        {
            bool result = false;
            WaitBusyEvent(200, "I2C_Reset");
            if (FT4222_I2CMaster_Reset(ComnHandle) == STATUS.FT4222_OK)
            {
                result = true;
            }

            BusyEvent.Set();
            return result;
        }

        public ushort I2C_ReadBytes(byte SlaveAddress, ref byte[] ReadBuf, ushort Length)
        {
            ushort sizeOfRead = 0;
            WaitBusyEvent(200, "I2C_ReadBytes");
            FT4222_I2CMaster_Read(ComnHandle, SlaveAddress, ref ReadBuf[0], Length, ref sizeOfRead);
            BusyEvent.Set();
            return sizeOfRead;
        }

        public ushort I2C_WriteBytes(byte SlaveAddress, byte[] WriteBuf, ushort Length)
        {
            ushort sizeTransferred = 0;
            WaitBusyEvent(200, "I2C_WriteBytes");
            FT4222H.STATUS Status = FT4222_I2CMaster_Write(ComnHandle, SlaveAddress, ref WriteBuf[0], Length, ref sizeTransferred);
            BusyEvent.Set();
            return sizeTransferred;
        }

        public ushort I2C_WriteBytes_NoStop(byte SlaveAddress, byte[] WriteBuf, ushort Length)
        {
            ushort sizeTransferred = 0;
            WaitBusyEvent(200, "I2C_WriteBytes_NoStop");
            var result = FT4222_I2CMaster_WriteEx(ComnHandle, SlaveAddress, 0x03, ref WriteBuf[0], Length, ref sizeTransferred);
            BusyEvent.Set();
            return sizeTransferred;
        }

        public bool GPIO_SetDirection(int PortNum, JLcLib.Comn.GPIO_Direction Direction)
        {
            bool result = false;
            if (PortNum < GPIOs.Count)
            {
                if (Direction == JLcLib.Comn.GPIO_Direction.Input)
                {
                    GpioDirection[PortNum] = GPIO_Direction.INPUT;
                }
                else
                {
                    GpioDirection[PortNum] = GPIO_Direction.OUTPUT;
                }

                WaitBusyEvent(200, "GPIO_Init");
                FT4222_UnInitialize(GpioHandle);
                if (FT4222_GPIO_Init(GpioHandle, GpioDirection) == STATUS.FT4222_OK)
                {
                    result = true;
                }

                BusyEvent.Set();
            }

            return result;
        }

        public bool GPIO_SetState(int PortNum, GPIO_State State)
        {
            bool result = false;
            if (PortNum < GPIOs.Count)
            {
                bool value = State == GPIO_State.High;
                WaitBusyEvent(200, "GPIO_SetState");
                if (FT4222_GPIO_Write(GpioHandle, (GPIO_Port)PortNum, value) == STATUS.FT4222_OK)
                {
                    result = true;
                }

                BusyEvent.Set();
            }

            return result;
        }

        public GPIO_State GPIO_GetState(int PortNum)
        {
            GPIO_State result = GPIO_State.Low;
            if (PortNum < GPIOs.Count)
            {
                bool Value = false;
                WaitBusyEvent(200, "GPIO_GetState");
                if (FT4222_GPIO_Read(GpioHandle, (GPIO_Port)PortNum, ref Value) == STATUS.FT4222_OK)
                {
                    result = (Value ? GPIO_State.High : GPIO_State.Low);
                }

                BusyEvent.Set();
            }

            return result;
        }
    }
}
