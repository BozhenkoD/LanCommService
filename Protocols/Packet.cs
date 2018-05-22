using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Protocols
{
    [Serializable]
    public class Packet
    {
        public string CardType { get; set; }

        public bool CVV { get; set; }

        public bool MSOffice { get; set; }

        public bool FindNumber { get; set; }

        public List<string> FilePath { get; set; }

        public string FileInfo { get; set; }

        public bool Rar { get; set; }

        public int Progress { get; set; }
    }
}
