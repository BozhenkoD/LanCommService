using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinServices;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            QuoteServer qs = new QuoteServer("127.0.0.1", "127.0.0.1", 4567);//( "192.168.42.175","192.168.42.27", 4567);
            qs.Start();
            qs.StartWork();
            Console.WriteLine("Hit return to exit");
            Console.ReadLine();
            qs.Stop();
        }
    }
}
