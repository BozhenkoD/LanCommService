using System;
using WinServices;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            QuoteServer qs = new QuoteServer("127.0.0.1", 4567);
            qs.StartWork();
            Console.WriteLine("Hit return to exit");
            Console.ReadLine();
            qs.Stop();
        }
    }
}
