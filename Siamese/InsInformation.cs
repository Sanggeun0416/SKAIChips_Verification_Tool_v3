using System;

namespace JLcLib.Custom
{
    [Serializable]
    public class InsInformation
    {
        public InstrumentTypes Type { get; set; }

        public bool Valid { get; set; }

        public string Address { get; set; }

        public string Name { get; set; }

        public InsInformation(InstrumentTypes Type, bool Valid, string Address)
        {
            this.Type = Type;
            this.Valid = Valid;
            this.Address = Address;
            Name = "";
        }
    }
}
