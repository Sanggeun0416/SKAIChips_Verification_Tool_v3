using System;
using System.Runtime.InteropServices;
using Ivi.Visa.Interop;

namespace JLcLib.Custom
{
    public class SCPI : IDisposable
    {
        private readonly object _sync = new object();
        private ResourceManager _rMgr;
        private FormattedIO488 _ins;

        public InsInformation InsInfo { get; set; }
        public bool IsOpen { get; private set; }

        public SCPI(InstrumentTypes type)
        {
            InsInfo = null;
            foreach (var insInfo in InstrumentForm.InsInfoList)
            {
                if (insInfo.Type == type)
                {
                    InsInfo = insInfo;
                    break;
                }
            }

            _rMgr = new ResourceManager();
            _ins = new FormattedIO488();
        }

        ~SCPI()
        {
            Dispose(false);
        }

        public bool Open(int openTimeoutMs = 200, int ioTimeoutMs = 5000)
        {
            if (InsInfo == null || !InsInfo.Valid) return false;

            lock (_sync)
            {
                if (IsOpen) Close();

                try
                {
                    _ins.IO = (IMessage)_rMgr.Open(InsInfo.Address, AccessMode.NO_LOCK, openTimeoutMs, "");
                    if (_ins.IO == null) return false;
                    _ins.IO.Timeout = ioTimeoutMs;
                    IsOpen = true;
                    return true;
                }
                catch
                {
                    IsOpen = false;
                    return false;
                }
            }
        }

        public bool Close()
        {
            lock (_sync)
            {
                try
                {
                    if (_ins != null && _ins.IO != null)
                    {
                        _ins.IO.Close();
                        Marshal.FinalReleaseComObject(_ins.IO);
                        _ins.IO = null;
                    }
                }
                catch { }
                IsOpen = false;
                return true;
            }
        }

        public void Write(string command)
        {
            if (!IsOpen) return;
            lock (_sync)
            {
                _ins.WriteString(command, true);
            }
        }

        public void Write(string command, byte[] sendBytes, bool flushAndEnd = false)
        {
            if (!IsOpen) return;
            lock (_sync)
            {
                _ins.WriteIEEEBlock(command, sendBytes, flushAndEnd);
            }
        }

        public string WriteAndReadString(string command, int timeoutMs = 5000)
        {
            if (!IsOpen) return string.Empty;
            lock (_sync)
            {
                try
                {
                    _ins.IO.Timeout = timeoutMs;
                    _ins.WriteString(command, true);
                    return _ins.ReadString();
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        public byte[] WriteAndReadBytes(string command, int timeoutMs = 5000)
        {
            if (!IsOpen) return null;
            lock (_sync)
            {
                try
                {
                    _ins.IO.Timeout = timeoutMs;
                    _ins.WriteString(command, true);
                    return (byte[])_ins.ReadIEEEBlock(IEEEBinaryType.BinaryType_UI1);
                }
                catch
                {
                    return null;
                }
            }
        }

        public void WaitAllOperationsComplete(int timeoutMs = 10000)
        {
            WriteAndReadString("*OPC?", timeoutMs);
        }

        public string GetInstrumentName()
        {
            return WriteAndReadString("*IDN?", 1000);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            Close();

            if (_ins != null)
            {
                try { Marshal.FinalReleaseComObject(_ins); } catch { }
                _ins = null;
            }

            if (_rMgr != null)
            {
                try { Marshal.FinalReleaseComObject(_rMgr); } catch { }
                _rMgr = null;
            }
        }
    }
}
