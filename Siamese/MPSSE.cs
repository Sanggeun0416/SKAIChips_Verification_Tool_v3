using FTD2XX_NET;
using JLcLib.Comn;
using System;
using System.Threading;

namespace JLcLib.Custom
{
    public class MPSSE
    {
        public static class Command
        {
            public static class Clock
            {
                public static readonly byte SetDivisor = 134;

                public static readonly byte Disable5Divisor = 138;

                public static readonly byte Enable5Divisor = 139;
            }

            public static class GPIO
            {
                public static readonly byte SetGPIOL = 128;

                public static readonly byte GetGPIOL = 129;

                public static readonly byte SetGPIOH = 130;

                public static readonly byte GetGPIOH = 131;
            }

            public static class Loopback
            {
                public static readonly byte Enable = 132;

                public static readonly byte Disable = 133;
            }

            public static class ThreePhase
            {
                public static readonly byte Enable = 140;

                public static readonly byte Disable = 141;
            }

            public static class AdaptiveClocking
            {
                public static readonly byte Enable = 150;

                public static readonly byte Disable = 151;
            }

            public enum PinConfig
            {
                Write = 16,
                Read = 32,
                ReadWrite = 48,
                TMS_Write = 64,
                TMS_ReadWrite = 96
            }

            public enum DataUnit
            {
                Byte = 0,
                Bit = 2
            }

            public enum BitFirst
            {
                MSB = 0,
                LSB = 8
            }

            public enum ClockEdge
            {
                OutRising = 1,
                OutFalling = 0,
                InRising = 0,
                InFalling = 4,
                InOutRising = 1,
                InOutFalling = 4,
                InRising_OutFalling = 0,
                InFalling_OutRising = 5
            }

            public static readonly byte BadCommand = 250;

            public static readonly byte SendAnswerBackImmediately = 135;

            public static readonly byte DriveOnlyZero = 158;

            public static byte Get(PinConfig Config, ClockEdge Edge, DataUnit Unit, BitFirst First)
            {
                if ((Config == PinConfig.TMS_Write || Config == PinConfig.TMS_ReadWrite) && Unit == DataUnit.Byte)
                {
                    Unit = DataUnit.Bit;
                }

                return (byte)((byte)Config | (byte)Edge | (byte)Unit | (byte)First);
            }
        }

        public enum TapState
        {
            TestLogicReset,
            RunTestIdle,
            SelectDR,
            SelectIR,
            CaptureDR,
            CaptureIR,
            ShiftDR,
            ShiftIR,
            Exit1DR,
            Exit1IR,
            PauseDR,
            PauseIR,
            Exit2DR,
            Exit2IR,
            UpdateDR,
            UpdateIR,
            Undefined
        }

        private FTDI ftdi = new FTDI();

        private Queue CommandQueue = new Queue(65536);

        private int NumBytesToRead = 0;

        private int SPI_Mode = 0;

        private Command.ClockEdge SPI_Edge;

        private Command.BitFirst SPI_BitFirst;

        private TapState CurTapState = TapState.Undefined;

        private const byte RFFE_Mask = 248;

        private const byte RFFE_Direction = 11;

        public ManualResetEvent BusyEvent { get; set; } = new ManualResetEvent(initialState: true);


        public FTDI.FT_DEVICE_INFO_NODE Device { get; private set; }

        public bool IsOpen => ftdi.IsOpen;

        public GPIO_Collection GPIOs { get; set; } = new GPIO_Collection();


        public ushort ClockDivisor { get; private set; }

        public double Freq
        {
            get
            {
                return CLK_Get();
            }
            set
            {
                CLK_Set(value);
            }
        }

        public byte LowPinsDirections { get; private set; } = 251;


        public byte LowPinsStates { get; private set; } = 251;


        public byte HighPinsDirections { get; private set; } = byte.MaxValue;


        public byte HighPinsStates { get; private set; } = byte.MaxValue;


        ~MPSSE()
        {
            Close();
        }

        public FTDI.FT_DEVICE_INFO_NODE GetDeviceInfo(string DevName)
        {
            foreach (FTDI.FT_DEVICE_INFO_NODE mPSS in FT_Device.MPSSEs)
            {
                if (mPSS.Description == DevName)
                {
                    return mPSS;
                }
            }

            return null;
        }

        private bool WaitBusyEvent(int Timeout_msec, string DebugMessage)
        {
            if (BusyEvent.WaitOne(Timeout_msec))
            {
                BusyEvent.Reset();
                return true;
            }

            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff") + "> MPSSE:BDF:" + DebugMessage);
            return false;
        }

        public bool Open(FTDI.FT_DEVICE_INFO_NODE Device)
        {
            bool flag = false;
            if (Device == null)
            {
                return Close();
            }

            this.Device = Device;
            FTDI.FT_STATUS fT_STATUS = ftdi.OpenByLocation(Device.LocId);
            if (fT_STATUS != 0)
            {
                Console.WriteLine("Can't open FTD2XX device!");
            }
            else
            {
                fT_STATUS |= ftdi.ResetDevice();
                fT_STATUS |= ftdi.SetCharacters(0, EventCharEnable: false, 0, ErrorCharEnable: false);
                fT_STATUS |= ftdi.SetTimeouts(5000u, 5000u);
                fT_STATUS |= ftdi.SetLatency(16);
                fT_STATUS |= ftdi.SetBitMode(0, 0);
                if ((fT_STATUS | ftdi.SetBitMode(0, 2)) != 0)
                {
                    Console.WriteLine("Failed to initialize FTD2XX device!");
                    ftdi.Close();
                }
                else
                {
                    Delay.Sleep(50);
                    byte b = 170;
                    while (true)
                    {
                        if (b <= 171)
                        {
                            CommandQueue.Clear();
                            CommandQueue.Set(b);
                            if (!SendCommand())
                            {
                                Console.WriteLine("Write timed out!");
                                ftdi.Close();
                                break;
                            }

                            byte[] receivedBytes = GetReceivedBytes(2);
                            if (receivedBytes == null)
                            {
                                return ftdi.IsOpen;
                            }

                            if (receivedBytes.Length >= 2)
                            {
                                flag = false;
                                for (int i = 0; i < receivedBytes.Length - 1; i++)
                                {
                                    if (receivedBytes[i] == Command.BadCommand && receivedBytes[i + 1] == b)
                                    {
                                        flag = true;
                                        break;
                                    }
                                }
                            }

                            if (!flag)
                            {
                                Console.WriteLine("fail to synchronize MPSSE with command 0x" + b.ToString("X") + " \n");
                                ftdi.Close();
                                return ftdi.IsOpen;
                            }

                            b++;
                            continue;
                        }

                        CommandQueue.Clear();
                        CommandQueue.Set(Command.Clock.Disable5Divisor);
                        CommandQueue.Set(Command.Loopback.Disable);
                        SendCommand();
                        CommandQueue.Set(Command.GPIO.SetGPIOL);
                        CommandQueue.Set(LowPinsStates);
                        CommandQueue.Set(LowPinsDirections);
                        CommandQueue.Set(Command.GPIO.SetGPIOH);
                        CommandQueue.Set(HighPinsStates);
                        CommandQueue.Set(HighPinsDirections);
                        SendCommand();
                        GPIO_Init();
                        break;
                    }
                }
            }

            return ftdi.IsOpen;
        }

        public bool Close()
        {
            if (ftdi.IsOpen)
            {
                ftdi.Close();
                GPIOs.Clear();
            }

            return ftdi.IsOpen;
        }

        public bool SendCommand(bool SendAnswerBackImmediately = false)
        {
            uint numBytesWritten = 0u;
            if (SendAnswerBackImmediately)
            {
                CommandQueue.Set(Command.SendAnswerBackImmediately);
            }

            byte[] bytes = CommandQueue.GetBytes(CommandQueue.Count);
            if (bytes.Length != 0)
            {
                WaitBusyEvent(200, "SendCommand()");
                FTDI.FT_STATUS fT_STATUS = ftdi.Write(bytes, bytes.Length, ref numBytesWritten);
                BusyEvent.Set();
                if (fT_STATUS == FTDI.FT_STATUS.FT_OK)
                {
                    return true;
                }
            }

            return false;
        }

        public byte[] GetReceivedBytes(int Length)
        {
            uint RxQueue = 0u;
            uint numBytesRead = 0u;
            int num = 0;
            byte[] array = null;
            WaitBusyEvent(200, "GetReceivedBytes()");
            FTDI.FT_STATUS rxBytesAvailable;
            do
            {
                Delay.Sleep(1);
                rxBytesAvailable = ftdi.GetRxBytesAvailable(ref RxQueue);
                num++;
            }
            while (RxQueue < Length && num < 500);
            if (rxBytesAvailable == FTDI.FT_STATUS.FT_OK && Length <= RxQueue)
            {
                array = new byte[RxQueue];
                ftdi.Read(array, RxQueue, ref numBytesRead);
                NumBytesToRead = 0;
            }

            BusyEvent.Set();
            return array;
        }

        private bool CLK_Set(ushort ClockDivisor)
        {
            CommandQueue.Set(Command.Clock.SetDivisor);
            CommandQueue.Set((byte)(ClockDivisor & 0xFFu));
            CommandQueue.Set((byte)((uint)(ClockDivisor >> 8) & 0xFFu));
            if (SendCommand())
            {
                this.ClockDivisor = ClockDivisor;
                return true;
            }

            return false;
        }

        private bool CLK_Set(double Clock_kHz)
        {
            double num = ((Device.Type == FTDI.FT_DEVICE.FT_DEVICE_2232) ? 12000f : 60000f);
            ushort clockDivisor = (ushort)(num / (2.0 * Clock_kHz) - 1.0);
            return CLK_Set(clockDivisor);
        }

        private double CLK_Get()
        {
            if (Device.Type == FTDI.FT_DEVICE.FT_DEVICE_2232)
            {
                return 12000.0 / ((1.0 + (double)(int)ClockDivisor) * 2.0);
            }

            return 60000.0 / ((1.0 + (double)(int)ClockDivisor) * 2.0);
        }

        private void GPIO_Init()
        {
            GPIOs.Clear();
            for (int i = 0; i < 4; i++)
            {
                GPIOs.Add(i, "GPIOL" + i, (GPIO_Direction)((LowPinsDirections >> i + 4) & 1), (GPIO_State)((LowPinsStates >> i + 4) & 1), GPIOL_SetDirection, GPIOL_SetState, GPIOL_GetState);
            }

            switch (Device.Type)
            {
                case FTDI.FT_DEVICE.FT_DEVICE_2232:
                    {
                        for (int k = 0; k < 4; k++)
                        {
                            GPIOs.Add(k, "GPIOH" + k, (GPIO_Direction)((HighPinsDirections >> k) & 1), (GPIO_State)((HighPinsStates >> k) & 1), GPIOH_SetDirection, GPIOH_SetState, GPIOH_GetState);
                        }

                        break;
                    }
                case FTDI.FT_DEVICE.FT_DEVICE_2232H:
                case FTDI.FT_DEVICE.FT_DEVICE_232H:
                    {
                        for (int j = 0; j < 8; j++)
                        {
                            GPIOs.Add(j, "GPIOH" + j, (GPIO_Direction)((HighPinsDirections >> j) & 1), (GPIO_State)((HighPinsStates >> j) & 1), GPIOH_SetDirection, GPIOH_SetState, GPIOH_GetState);
                        }

                        break;
                    }
                case FTDI.FT_DEVICE.FT_DEVICE_232R:
                case FTDI.FT_DEVICE.FT_DEVICE_4232H:
                    break;
            }
        }

        public void GPIOL_SetPins(byte Directions, byte States, bool Protected = true)
        {
            if (Protected)
            {
                byte b = GPIOL_GetPins();
                LowPinsStates = (byte)((States & 0xF0u) | (b & 0xFu));
                LowPinsDirections = (byte)((Directions & 0xF0u) | (LowPinsDirections & 0xFu));
            }
            else
            {
                LowPinsStates = States;
                LowPinsDirections = Directions;
            }

            CommandQueue.Set(Command.GPIO.SetGPIOL);
            CommandQueue.Set(LowPinsStates);
            CommandQueue.Set(LowPinsDirections);
        }

        public byte GPIOL_GetPins()
        {
            byte BitMode = 0;
            WaitBusyEvent(200, "GetLowPins()");
            ftdi.GetPinStates(ref BitMode);
            LowPinsStates = BitMode;
            BusyEvent.Set();
            return LowPinsStates;
        }

        public bool GPIOL_SetDirection(int PortNum, GPIO_Direction Direction)
        {
            int num = PortNum + 4;
            if (Direction == GPIO_Direction.Output)
            {
                LowPinsDirections |= (byte)(1 << num);
            }
            else
            {
                LowPinsDirections &= (byte)(~(1 << num));
            }

            GPIOL_SetPins(LowPinsDirections, LowPinsStates);
            return SendCommand();
        }
        
        public bool GPIOL_SetState(int PortNum, GPIO_State State)
        {
            int num = PortNum + 4;
            if (State == GPIO_State.High)
            {
                LowPinsStates |= (byte)(1 << num);
            }
            else
            {
                LowPinsStates &= (byte)(~(1 << num));
            }

            GPIOL_SetPins(LowPinsDirections, LowPinsStates);
            return SendCommand();
        }

        public GPIO_State GPIOL_GetState(int PortNum)
        {
            int num = PortNum + 4;
            GPIOL_GetPins();
            if (((LowPinsStates >> num) & 1) == 1)
            {
                return GPIO_State.High;
            }

            return GPIO_State.Low;
        }

        public void GPIOH_SetPins(byte Directions, byte States)
        {
            HighPinsStates = States;
            HighPinsDirections = Directions;
            CommandQueue.Set(Command.GPIO.SetGPIOH);
            CommandQueue.Set(HighPinsStates);
            CommandQueue.Set(HighPinsDirections);
        }

        public byte GPOH_GetPins()
        {
            CommandQueue.Set(Command.GPIO.GetGPIOH);
            NumBytesToRead++;
            SendCommand();
            byte[] receivedBytes = GetReceivedBytes(NumBytesToRead);
            if (receivedBytes != null && receivedBytes.Length != 0)
            {
                HighPinsStates = receivedBytes[receivedBytes.Length - 1];
            }

            return HighPinsStates;
        }

        public bool GPIOH_SetDirection(int PortNum, GPIO_Direction Direction)
        {
            if (Direction == GPIO_Direction.Output)
            {
                HighPinsDirections |= (byte)(1 << PortNum);
            }
            else
            {
                HighPinsDirections &= (byte)(~(1 << PortNum));
            }

            GPIOH_SetPins(HighPinsDirections, HighPinsStates);
            return SendCommand();
        }

        public bool GPIOH_SetState(int PortNum, GPIO_State State)
        {
            if (State == GPIO_State.High)
            {
                HighPinsStates |= (byte)(1 << PortNum);
            }
            else
            {
                HighPinsStates &= (byte)(~(1 << PortNum));
            }

            GPIOH_SetPins(HighPinsDirections, HighPinsStates);
            return SendCommand();
        }

        public GPIO_State GPIOH_GetState(int PortNum)
        {
            GPOH_GetPins();
            if (((HighPinsStates >> PortNum) & 1) == 1)
            {
                return GPIO_State.High;
            }

            return GPIO_State.Low;
        }

        public void SSC_SetBytes(byte[] Bytes, ushort Length, Command.PinConfig Config, Command.BitFirst First, Command.ClockEdge Edge)
        {
            if (Config == Command.PinConfig.TMS_Write || Config == Command.PinConfig.TMS_ReadWrite || ((Config == Command.PinConfig.Write || Config == Command.PinConfig.ReadWrite) && (Bytes == null || Bytes.Length < Length)))
            {
                return;
            }

            CommandQueue.Set(Command.Get(Config, Edge, Command.DataUnit.Byte, First));
            CommandQueue.Set((byte)((uint)(Length - 1) & 0xFFu));
            CommandQueue.Set((byte)((uint)(Length - 1 >> 8) & 0xFFu));
            if (Config == Command.PinConfig.Write || Config == Command.PinConfig.ReadWrite)
            {
                for (int i = 0; i < Length; i++)
                {
                    CommandQueue.Set(Bytes[i]);
                }
            }

            if (Config == Command.PinConfig.Read || Config == Command.PinConfig.ReadWrite)
            {
                NumBytesToRead += Length;
            }
        }

        public void SSC_SetBits(byte SendBits, byte Length, Command.PinConfig Config, Command.BitFirst First, Command.ClockEdge Edge)
        {
            if (Config != Command.PinConfig.TMS_Write && Config != Command.PinConfig.TMS_ReadWrite && Length <= 8 && Length >= 1)
            {
                CommandQueue.Set(Command.Get(Config, Edge, Command.DataUnit.Bit, First));
                CommandQueue.Set((byte)(Length - 1));
                if (Config == Command.PinConfig.Write || Config == Command.PinConfig.ReadWrite)
                {
                    CommandQueue.Set(SendBits);
                }

                if (Config == Command.PinConfig.Read || Config == Command.PinConfig.ReadWrite)
                {
                    NumBytesToRead++;
                }
            }
        }

        public void SSC_SetAllPins(Command.PinConfig Config, byte Length, bool TDIBit, byte TMSBits, Command.ClockEdge Edge)
        {
            if (Config != Command.PinConfig.Read && Config != Command.PinConfig.Write && Config != Command.PinConfig.ReadWrite && Length <= 8 && Length >= 1)
            {
                byte value = ((!TDIBit) ? ((byte)(TMSBits & 0x7Fu)) : ((byte)(0x80u | (TMSBits & 0x7Fu))));
                CommandQueue.Set(Command.Get(Config, Edge, Command.DataUnit.Bit, Command.BitFirst.LSB));
                CommandQueue.Set((byte)(Length - 1));
                CommandQueue.Set(value);
                if (Config == Command.PinConfig.TMS_ReadWrite)
                {
                    NumBytesToRead++;
                }
            }
        }

        public void SPI_Init(int Mode, int Clock_kHz, Command.BitFirst First)
        {
            CommandQueue.Clear();
            CommandQueue.Set(Command.AdaptiveClocking.Disable);
            CommandQueue.Set(Command.ThreePhase.Disable);
            SendCommand();
            CLK_Set(Clock_kHz);
            Delay.Sleep(20);
            SPI_Mode = Mode;
            SPI_BitFirst = First;
            if (Mode == 0 || Mode == 3)
            {
                SPI_Edge = Command.ClockEdge.OutRising;
            }
            else
            {
                SPI_Edge = Command.ClockEdge.InFalling;
            }

            if (SPI_Mode < 2)
            {
                GPIOL_SetPins(251, 248, Protected: false);
            }
            else
            {
                GPIOL_SetPins(251, 249, Protected: false);
            }

            GPIOH_SetPins(byte.MaxValue, byte.MaxValue);
            SendCommand();
            Delay.Sleep(30);
        }

        public void SPI_SetStart()
        {
            if (SPI_Mode < 2)
            {
                for (int i = 0; i < 10; i++)
                {
                    GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 8u), Protected: false);
                }

                GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
            }
            else
            {
                for (int j = 0; j < 10; j++)
                {
                    GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 9u), Protected: false);
                }

                GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 1u), Protected: false);
            }
        }

        public void SPI_SetStop()
        {
            if (SPI_Mode < 2)
            {
                for (int i = 0; i < 10; i++)
                {
                    GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
                }

                GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 8u), Protected: false);
            }
            else
            {
                for (int j = 0; j < 10; j++)
                {
                    GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 1u), Protected: false);
                }

                GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 9u), Protected: false);
            }
        }

        public void SPI_SetBytes(byte[] Bytes, ushort Length, Command.PinConfig Config)
        {
            SSC_SetBytes(Bytes, Length, Config, SPI_BitFirst, SPI_Edge);
        }

        public void SPI_SetBits(byte Bits, byte Length, Command.PinConfig Config)
        {
            SSC_SetBits(Bits, Length, Config, SPI_BitFirst, SPI_Edge);
        }

        public void SPI_SetDummyBits(bool MOSI_State, bool SS_State, byte Length)
        {
            SSC_SetAllPins(Command.PinConfig.TMS_Write, Length, MOSI_State, (byte)(SS_State ? 127u : 0u), SPI_Edge);
        }

        public void SPI_HalfDuplex()
        {
            GPIOL_SetPins(249, (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
        }

        public void I2C_Init(int Clock_kHz)
        {
            CommandQueue.Clear();
            CommandQueue.Set(Command.AdaptiveClocking.Disable);
            CommandQueue.Set(Command.ThreePhase.Enable);
            if (Device.Type == FTDI.FT_DEVICE.FT_DEVICE_232H)
            {
                CommandQueue.Set(Command.DriveOnlyZero);
                CommandQueue.Set(7);
                CommandQueue.Set(0);
            }

            SendCommand();
            CLK_Set(Clock_kHz);
            Delay.Sleep(20);
            GPIOL_SetPins(240, 240, Protected: false);
            GPIOH_SetPins(byte.MaxValue, byte.MaxValue);
            SendCommand();
            Delay.Sleep(30);
        }

        private void I2C_SetStart()
        {
            byte directions = (byte)((LowPinsDirections & 0xF0u) | 3u);
            for (int i = 0; i < 20; i++)
            {
                GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 3u), Protected: false);
            }

            for (int j = 0; j < 20; j++)
            {
                GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 1u), Protected: false);
            }

            GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
        }

        private void I2C_SetStop()
        {
            byte directions = (byte)((LowPinsDirections & 0xF0u) | 0u);
            byte directions2 = (byte)((LowPinsDirections & 0xF0u) | 3u);
            for (int i = 0; i < 20; i++)
            {
                GPIOL_SetPins(directions2, (byte)((LowPinsStates & 0xF0u) | 1u), Protected: false);
            }

            for (int j = 0; j < 20; j++)
            {
                GPIOL_SetPins(directions2, (byte)((LowPinsStates & 0xF0u) | 3u), Protected: false);
            }

            GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
        }

        private void I2C_SetWriteByte(byte Data)
        {
            byte directions = (byte)((LowPinsDirections & 0xF0u) | 3u);
            SSC_SetBits(Data, 8, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutRising);
            for (int i = 0; i < 20; i++)
            {
                GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 2u), Protected: false);
            }

            SSC_SetBits(0, 1, Command.PinConfig.Read, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            for (int j = 0; j < 100; j++)
            {
                GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 2u), Protected: false);
            }
        }

        private void I2C_SetReadByte(bool Nak)
        {
            byte directions = (byte)((LowPinsDirections & 0xF0u) | 3u);
            for (int i = 0; i < 50; i++)
            {
                GPIOL_SetPins(directions, (byte)((LowPinsStates & 0xF0u) | 2u), Protected: false);
            }

            SSC_SetBits(0, 8, Command.PinConfig.Read, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SSC_SetBits((byte)(Nak ? 255u : 0u), 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutRising);
        }

        public bool I2C_WriteBytes(byte SlaveAddress, byte[] Bytes, int Length, bool stop)
        {
            I2C_SetStart();
            I2C_SetWriteByte((byte)((uint)(SlaveAddress << 1) | 0u));
            for (int i = 0; i < Length; i++)
            {
                I2C_SetWriteByte(Bytes[i]);
            }

            if (stop)
            {
                I2C_SetStop();
            }

            if (SendCommand(SendAnswerBackImmediately: true))
            {
                byte[] receivedBytes = GetReceivedBytes(NumBytesToRead);
                if (receivedBytes != null && receivedBytes.Length >= Length)
                {
                    return true;
                }
            }

            return false;
        }

        public byte[] I2C_ReadBytes(byte SlaveAddress, int Length)
        {
            byte[] array = null;
            I2C_SetStart();
            I2C_SetWriteByte((byte)((uint)(SlaveAddress << 1) | 1u));
            for (int i = 0; i < Length; i++)
            {
                I2C_SetReadByte(i == Length - 1);
            }

            I2C_SetStop();
            if (SendCommand(SendAnswerBackImmediately: true))
            {
                byte[] receivedBytes = GetReceivedBytes(NumBytesToRead);
                if (receivedBytes != null && receivedBytes.Length >= Length)
                {
                    array = new byte[Length];
                    for (int j = 0; j < Length; j++)
                    {
                        array[Length - 1 - j] = receivedBytes[receivedBytes.Length - 1 - j];
                    }
                }
            }

            return array;
        }

        public byte[] I2C_WriteAndReadBytes(byte SlaveAddress, byte[] WriteBytes, int NumBytesToWrite, int NumBytesToRead)
        {
            byte[] array = null;
            I2C_SetStart();
            I2C_SetWriteByte((byte)((uint)(SlaveAddress << 1) | 0u));
            for (int i = 0; i < NumBytesToWrite; i++)
            {
                I2C_SetWriteByte(WriteBytes[i]);
            }

            I2C_SetStop();
            for (int j = 0; j < 10; j++)
            {
                GPIOL_SetPins(LowPinsDirections, (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
            }

            I2C_SetStart();
            I2C_SetWriteByte((byte)((uint)(SlaveAddress << 1) | 1u));
            for (int k = 0; k < NumBytesToRead; k++)
            {
                I2C_SetReadByte(k == NumBytesToRead - 1);
            }

            I2C_SetStop();
            if (SendCommand(SendAnswerBackImmediately: true))
            {
                byte[] receivedBytes = GetReceivedBytes(this.NumBytesToRead);
                if (receivedBytes != null && receivedBytes.Length >= NumBytesToRead)
                {
                    array = new byte[NumBytesToRead];
                    for (int l = 0; l < NumBytesToRead; l++)
                    {
                        array[NumBytesToRead - 1 - l] = receivedBytes[receivedBytes.Length - 1 - l];
                    }
                }
            }

            return array;
        }

        public void JTAG_Init(int Clock_kHz)
        {
            CommandQueue.Clear();
            CommandQueue.Set(Command.AdaptiveClocking.Enable);
            CommandQueue.Set(Command.ThreePhase.Disable);
            SendCommand();
            CLK_Set(Clock_kHz);
            Delay.Sleep(20);
            GPIOL_SetPins(251, 252, Protected: false);
            GPIOH_SetPins(byte.MaxValue, byte.MaxValue);
            SendCommand();
            Delay.Sleep(30);
            JTAG_Reset();
        }

        private TapState JTAG_GetNextTapState(bool TMSVal)
        {
            TapState tapState = TapState.TestLogicReset;
            switch (CurTapState)
            {
                case TapState.TestLogicReset:
                    tapState = ((!TMSVal) ? TapState.RunTestIdle : TapState.TestLogicReset);
                    break;
                case TapState.RunTestIdle:
                    tapState = ((!TMSVal) ? TapState.RunTestIdle : TapState.SelectDR);
                    break;
                case TapState.SelectDR:
                    tapState = (TMSVal ? TapState.SelectIR : TapState.CaptureDR);
                    break;
                case TapState.SelectIR:
                    tapState = ((!TMSVal) ? TapState.CaptureIR : TapState.TestLogicReset);
                    break;
                case TapState.CaptureDR:
                    tapState = (TMSVal ? TapState.Exit1DR : TapState.ShiftDR);
                    break;
                case TapState.CaptureIR:
                    tapState = (TMSVal ? TapState.Exit1IR : TapState.ShiftIR);
                    break;
                case TapState.ShiftDR:
                    tapState = (TMSVal ? TapState.Exit1DR : TapState.ShiftDR);
                    break;
                case TapState.ShiftIR:
                    tapState = (TMSVal ? TapState.Exit1IR : TapState.ShiftIR);
                    break;
                case TapState.Exit1DR:
                    tapState = (TMSVal ? TapState.UpdateDR : TapState.PauseDR);
                    break;
                case TapState.Exit1IR:
                    tapState = (TMSVal ? TapState.UpdateIR : TapState.PauseIR);
                    break;
                case TapState.PauseDR:
                    tapState = (TMSVal ? TapState.Exit2DR : TapState.PauseDR);
                    break;
                case TapState.PauseIR:
                    tapState = (TMSVal ? TapState.Exit2IR : TapState.PauseIR);
                    break;
                case TapState.Exit2DR:
                    tapState = (TMSVal ? TapState.UpdateDR : TapState.ShiftDR);
                    break;
                case TapState.Exit2IR:
                    tapState = (TMSVal ? TapState.UpdateIR : TapState.ShiftIR);
                    break;
                case TapState.UpdateDR:
                    tapState = ((!TMSVal) ? TapState.RunTestIdle : TapState.SelectDR);
                    break;
                case TapState.UpdateIR:
                    tapState = ((!TMSVal) ? TapState.RunTestIdle : TapState.SelectDR);
                    break;
            }

            CurTapState = tapState;
            return tapState;
        }

        public bool JTAG_Reset()
        {
            SSC_SetAllPins(Command.PinConfig.TMS_Write, 7, TDIBit: false, 63, Command.ClockEdge.OutRising);
            bool flag = SendCommand();
            if (flag)
            {
                CurTapState = TapState.RunTestIdle;
            }

            return flag;
        }

        public bool JTAG_ChangeTAP(TapState NextTapState, bool ReadTDO, bool TDIVal)
        {
            byte b = 1;
            TapState curTapState = CurTapState;
            TapState tapState = TapState.TestLogicReset;
            if (CurTapState == NextTapState)
            {
                return false;
            }

            for (b = 1; b < 8; b++)
            {
                for (byte b2 = 0; b2 < 1 << (int)b; b2++)
                {
                    for (int i = 0; i < b; i++)
                    {
                        tapState = (((b2 & (1 << i)) != 0) ? JTAG_GetNextTapState(TMSVal: true) : JTAG_GetNextTapState(TMSVal: false));
                    }

                    if (tapState == NextTapState)
                    {
                        if (!ReadTDO)
                        {
                            SSC_SetAllPins(Command.PinConfig.TMS_Write, b, TDIVal, b2, Command.ClockEdge.OutRising);
                        }
                        else
                        {
                            SSC_SetAllPins(Command.PinConfig.TMS_ReadWrite, b, TDIVal, b2, Command.ClockEdge.OutRising);
                        }

                        CurTapState = NextTapState;
                        return true;
                    }

                    CurTapState = curTapState;
                }
            }

            return false;
        }

        public void JTAG_SetIR(int Data, int Length, Command.PinConfig TRxMode)
        {
            byte[] bytes = BitConverter.GetBytes(Data);
            int num = (Length - 1) / 8;
            int num2 = (Length - 1) % 8;
            int num3 = (Length + 7) / 8 - 1;
            JTAG_ChangeTAP(TapState.ShiftIR, ReadTDO: false, TDIVal: false);
            if (num > 0)
            {
                SSC_SetBytes(bytes, (ushort)num, TRxMode, Command.BitFirst.LSB, Command.ClockEdge.OutRising);
            }

            if (num2 > 0)
            {
                SSC_SetBits(bytes[num], (byte)(num2 - 1), TRxMode, Command.BitFirst.LSB, Command.ClockEdge.OutRising);
            }

            if ((bytes[num3] & (1 << num2)) == 0)
            {
                if (TRxMode == Command.PinConfig.Write)
                {
                    JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: false, TDIVal: false);
                }
                else
                {
                    JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: true, TDIVal: false);
                }
            }
            else if (TRxMode == Command.PinConfig.Write)
            {
                JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: false, TDIVal: true);
            }
            else
            {
                JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: true, TDIVal: true);
            }
        }

        public void JTAG_SetDR(int Data, int Length, Command.PinConfig TRxMode)
        {
            byte[] bytes = BitConverter.GetBytes(Data);
            int num = (Length - 1) / 8;
            int num2 = (Length - 1) % 8;
            int num3 = (Length + 7) / 8 - 1;
            JTAG_ChangeTAP(TapState.ShiftDR, ReadTDO: false, TDIVal: false);
            if (num > 0)
            {
                SSC_SetBytes(bytes, (ushort)num, TRxMode, Command.BitFirst.LSB, Command.ClockEdge.OutRising);
            }

            if (num2 > 0)
            {
                SSC_SetBits(bytes[num], (byte)(num2 - 1), TRxMode, Command.BitFirst.LSB, Command.ClockEdge.OutRising);
            }

            if ((bytes[num3] & (1 << num2)) == 0)
            {
                if (TRxMode == Command.PinConfig.Write)
                {
                    JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: false, TDIVal: false);
                }
                else
                {
                    JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: true, TDIVal: false);
                }
            }
            else if (TRxMode == Command.PinConfig.Write)
            {
                JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: false, TDIVal: true);
            }
            else
            {
                JTAG_ChangeTAP(TapState.RunTestIdle, ReadTDO: true, TDIVal: true);
            }
        }

        public void RFFE_Init(int Clock_kHz)
        {
            CommandQueue.Clear();
            CommandQueue.Set(Command.AdaptiveClocking.Disable);
            CommandQueue.Set(Command.ThreePhase.Disable);
            SendCommand();
            CLK_Set(Clock_kHz);
            Delay.Sleep(20);
            GPIOL_SetPins(251, 240, Protected: false);
            GPIOH_SetPins(byte.MaxValue, byte.MaxValue);
            SendCommand();
            Delay.Sleep(30);
        }

        private void RFFE_SetSSC()
        {
            int num = (int)(10.0 / CLK_Get());
            if (num <= 0)
            {
                num = 1;
            }

            LowPinsStates = (byte)((LowPinsStates & 0xF8u) | 2u);
            LowPinsDirections = (byte)((LowPinsDirections & 0xF8u) | 0xBu);
            for (int i = 0; i < num; i++)
            {
                GPIOL_SetPins(LowPinsDirections, LowPinsStates, Protected: false);
            }

            LowPinsStates = (byte)((LowPinsStates & 0xF8u) | 2u);
            for (int j = 0; j < num; j++)
            {
                GPIOL_SetPins(LowPinsDirections, LowPinsStates, Protected: false);
            }
        }

        public bool RFFE_WriteBytes(byte SlaveAddress, byte Address, byte[] Data, int Length)
        {
            int num = Length - 1;
            if (Length < 1 || Length > 16)
            {
                return false;
            }

            int num2 = (SlaveAddress << 8) | 0 | num;
            byte b = (byte)Comn.Convert.GetOddParity(num2);
            num2 = (num2 << 1) | b;
            RFFE_SetSSC();
            SSC_SetBits((byte)(num2 >> 5), 5, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SSC_SetBits((byte)((uint)num2 & 0xFFu), 8, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            b = (byte)Comn.Convert.GetOddParity(Address);
            SSC_SetBits(Address, 8, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SSC_SetBits(b, 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            for (int i = 0; i < Length; i++)
            {
                b = (byte)Comn.Convert.GetOddParity(Data[i]);
                SSC_SetBits(Data[i], 8, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
                SSC_SetBits(b, 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            }

            SSC_SetBits(0, 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            return SendCommand();
        }

        public int RFFE_ReadBytes(byte SlaveAddress, byte Address, ref byte[] Data, int Length)
        {
            int num = Length - 1;
            if (Length < 1 || Length > 16)
            {
                return 0;
            }

            int num2 = (SlaveAddress << 8) | 0x20 | num;
            byte b = (byte)Comn.Convert.GetOddParity(num2);
            num2 = (num2 << 1) | b;
            RFFE_SetSSC();
            SSC_SetBits((byte)(num2 >> 5), 5, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SSC_SetBits((byte)((uint)num2 & 0xFFu), 8, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            b = (byte)Comn.Convert.GetOddParity(Address);
            SSC_SetBits(Address, 8, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SSC_SetBits(b, 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SSC_SetBits(0, 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            GPIOL_SetPins((byte)((LowPinsDirections & 0xF0u) | 1u), (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
            for (int i = 0; i < Length; i++)
            {
                SSC_SetBits(0, 8, Command.PinConfig.Read, Command.BitFirst.MSB, Command.ClockEdge.InFalling);
                SSC_SetBits(0, 1, Command.PinConfig.Read, Command.BitFirst.MSB, Command.ClockEdge.InFalling);
            }

            GPIOL_SetPins((byte)((LowPinsDirections & 0xF0u) | 3u), (byte)((LowPinsStates & 0xF0u) | 0u), Protected: false);
            SSC_SetBits(0, 1, Command.PinConfig.Write, Command.BitFirst.MSB, Command.ClockEdge.OutFalling);
            SendCommand();
            byte[] receivedBytes = GetReceivedBytes(NumBytesToRead);
            if (receivedBytes == null || receivedBytes.Length == 0)
            {
                return 0;
            }

            if (receivedBytes.Length >= Length)
            {
                for (int j = 0; j < Length; j++)
                {
                    Data[Length - 1 - j] = receivedBytes[receivedBytes.Length - 1 - j];
                }
            }

            return receivedBytes.Length;
        }
    }
}
