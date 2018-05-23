using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using Protocols;

namespace WinServices
{
    public class TODO : Packet
    {
        private Packet pack { get; set; }

        private string LogFile { get; set; }

        public TODO(Packet pak)
        {
            pack = pak;
        }

        public void Work()
        {
            Search rea = new Search(pack);

            rea.FindFiles(pack , 0);

            
        }

        public void CountFiles()
        {
            Search rea = new Search(pack);

            rea.FindFiles(pack, 1);
        }

        
    }
}
