using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Prolocol
{
    [Serializable]
    public class Packet
    {
        public string CardType { get; set; }

        public bool CVV { get; set; }

        public bool MSOffice { get; set; }
    }
}
