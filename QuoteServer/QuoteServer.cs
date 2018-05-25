using Protocols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;


namespace WinServices
{
    public class QuoteServer
    {
        private string MyIPAddr;
        private Thread listenerThread;
        private Thread StartThread;
        private Thread ProgressThread;

        private string MangerIP;

        private Packet Packet;

        public QuoteServer(string MyIP, string ManagerIP)
        {
            this.MyIPAddr = MyIP;
            this.MangerIP = ManagerIP;
        }

        public bool StopTCP()
        {
            try
            {
                Start_bool = false;
                StartWork_bool = false;

                SetQuoteClient.Close();
                SendProgressTcp.Close();
                listenThread.Stop();

                listenerThread.Abort();
                StartThread.Abort();
                ProgressThread.Abort();

                return true;
            }
            catch  { return false; }
        }

        bool Start_bool = false;
        bool StartWork_bool = false;

        public void Start()
        {
            Start_bool = true;
            listenerThread = new Thread(ListenerThread);
            listenerThread.IsBackground = true;
            listenerThread.Name = "Listener";
            listenerThread.Start();
        }

        public void StartWork()
        {
            StartWork_bool = true;
            StartThread = new Thread(SetQuote);
            StartThread.IsBackground = true;
            StartThread.Name = "StartThread";
            StartThread.Start();

            ProgressThread = new Thread(SendProgress);
            ProgressThread.IsBackground = true;
            ProgressThread.Name = "StartThreadProgress";
            ProgressThread.Start();
        }

        private TcpClient SetQuoteClient;

        private void SetQuote()
        {
            while (StartWork_bool)
            {
                SetQuoteClient = new TcpClient();

                NetworkStream stream = null;
                try
                {
                    SetQuoteClient.Connect(MangerIP, 4569);

                    stream = SetQuoteClient.GetStream();

                    byte[] buffer = new byte[] { 0x01 };

                    stream.Write(buffer, 0, buffer.Length);

                    Console.Write("|");
                }
                catch (SocketException ex)
                {
                    Console.WriteLine(ex.Message.ToString());
                }
                finally
                {
                    //
                    if (stream != null)
                    {
                        stream.Close();
                    }
                    if (StartWork_bool)
                    {
                        if (SetQuoteClient.Connected)
                        {
                            SetQuoteClient.Close();
                        }
                    }
                }
                Thread.Sleep(1000);
            }

        }

        private TcpClient SendProgressTcp;

        private void SendProgress()
        {
            while (Start_bool)
            {
                SendProgressTcp = new TcpClient();

                NetworkStream stream = null;
                try
                {
                    if (Packet != null)
                    {
                        if (Packet.Progress != 100)
                        {
                            SendProgressTcp.Connect(MangerIP, 4568);

                            stream = SendProgressTcp.GetStream();

                            byte[] buffer = ToByteArray<Packet>(Packet);

                            stream.Write(buffer, 0, buffer.Length);

                            Console.Write(" [" + buffer.Length + " byte ]");
                        }
                        else
                        {
                            SendProgressTcp.Connect(MangerIP, 4568);

                            stream = SendProgressTcp.GetStream();

                            byte[] buffer = ToByteArray<Packet>(Packet);

                            stream.Write(buffer, 0, buffer.Length);

                            Console.WriteLine("Отправлено: " + buffer.Length + " byte");

                            if (!String.IsNullOrEmpty(Packet.FileInfo))
                            {
                                SendBigPacket(Packet.FileInfo);

                                Packet = null;
                            }
                        }
                    }
                    else
                    {
                        Console.Write(".");
                    }
                }
                catch (SocketException ex)
                {
                    Console.WriteLine(ex.Message.ToString());
                }
                finally
                {
                    //
                    if (stream != null)
                    {
                        stream.Close();
                    }

                    if (SendProgressTcp.Connected)
                    {
                        SendProgressTcp.Close();
                    }
                }
                Thread.Sleep(1000);

            }

        }

        private void SendBigPacket(string FilePath)
        {
            Socket client = new Socket(AddressFamily.InterNetwork,
                    SocketType.Stream, ProtocolType.Tcp);

            // Connect the socket to the remote endpoint.
            client.Connect(MangerIP, 4570);

            // There is a text file test.txt located in the root directory.
            string fileName = FilePath;

            // Send file fileName to remote device
            Console.WriteLine("Sending {0} to the host.", fileName);
            client.SendFile(fileName);

            // Release the socket.
            client.Shutdown(SocketShutdown.Both);

            client.Close();
        }

        private TcpListener listenThread;

        protected void ListenerThread()
        {
            try
            {
                IPAddress ipAddress = IPAddress.Parse(MyIPAddr);

                listenThread = new TcpListener(ipAddress, 4567);

                listenThread.Start();

                while (Start_bool)
                {
                    Socket clientSocket = listenThread.AcceptSocket();

                    byte[] buffer = new byte[65000];

                    int res = clientSocket.Receive(buffer);

                    if (res > 1)
                    {
                        byte[] buf = new byte[res];

                        Array.Copy(buffer, buf, res);

                        Packet = FromByteArray<Packet>(buf);

                        if (Packet.Progress != 100)
                        {
                            TODO dowork = new TODO(Packet);

                            dowork.Work();
                        }
                    }

                    clientSocket.Close();
                }
            }
            catch (SocketException ex)
            {
                Trace.TraceError(String.Format("QuoteServer {0}", ex.Message));
            }

        }



        public byte[] ToByteArray<T>(T obj)
        {
            if (obj == null)
                return null;
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream())
            {
                bf.Serialize(ms, obj);
                return ms.ToArray();
            }
        }

        public T FromByteArray<T>(byte[] data)
        {
            if (data == null)
                return default(T);
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream(data))
            {
                object obj = bf.Deserialize(ms);
                return (T)obj; 
            }
        }
    }
}

