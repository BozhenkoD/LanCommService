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
        private TcpListener listener;
        private int port;
        private string MyIPAddr;
        private Thread listenerThread;
        private Thread StartThread;
        private Thread ProgressThread;

        private string MangerIP;

        private Packet Packet;

        public QuoteServer(string MyIP, string ManagerIP, int port)
        {
            this.MyIPAddr = MyIP;
            this.port = port;
            this.MangerIP = ManagerIP;
        }
        

        public void Start()
        {
            listenerThread = new Thread(ListenerThread);
            listenerThread.IsBackground = true;
            listenerThread.Name = "Listener";
            listenerThread.Start();
        }

        public void StartWork()
        {
            StartThread = new Thread(SetQuote);
            StartThread.IsBackground = true;
            StartThread.Name = "StartThread";
            StartThread.Start();

            ProgressThread = new Thread(SendProgress);
            ProgressThread.IsBackground = true;
            ProgressThread.Name = "StartThread";
            ProgressThread.Start();
        }

        private void SetQuote()
        {

            while (true)
            {
                TcpClient client = new TcpClient();

                NetworkStream stream = null;
                try
                {
                    client.Connect(MangerIP, 4568);

                    stream = client.GetStream();

                    byte[] buffer = new byte[] { 0x01 };

                    stream.Write(buffer, 0, buffer.Length);

                    Console.WriteLine("Проверка соединения по адресу: " + MangerIP);
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

                    if (client.Connected)
                    {
                        client.Close();
                    }
                }
                Thread.Sleep(5000);
            }

        }

        private void SendProgress()
        {
            while (true)
            {
                TcpClient client = new TcpClient();

                NetworkStream stream = null;
                try
                {
                    client.Connect(MangerIP, 4568);

                    stream = client.GetStream();

                    if (Packet != null)
                    {
                        byte[] buffer = ToByteArray<Packet>(Packet);

                        stream.Write(buffer, 0, buffer.Length);

                        Console.WriteLine("Отправлено: " + buffer.Length + " byte");
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

                    if (client.Connected)
                    {
                        client.Close();
                    }
                }
                Thread.Sleep(1000);
            }

        }

        public void SendPacket(Packet pak)
        {
            TcpClient client = new TcpClient();

            NetworkStream stream = null;
            try
            {
                client.Connect(MangerIP, 4568);

                stream = client.GetStream();

                byte[] buffer = ToByteArray<Packet>(pak);

                stream.Write(buffer, 0, buffer.Length);

                Console.WriteLine("Отправлено Результат: " + MangerIP);
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

                if (client.Connected)
                {
                    client.Close();
                }
            }
        }


        protected void ListenerThread()
        {
            try
            {
                IPAddress ipAddress =  IPAddress.Parse(MyIPAddr);

                listener = new TcpListener(ipAddress, port);

                listener.Start();

                while (true)
                {
                    Socket clientSocket = listener.AcceptSocket();

                    byte[] buffer = new byte[65000];

                    int res = clientSocket.Receive(buffer);

                    if(res > 1)
                    {
                        byte[] buf = new byte[res];

                        Array.Copy(buffer, buf, res);

                        if (Packet == null)
                        {
                            Packet = new Packet();

                            Packet = FromByteArray<Packet>(buf);

                            Packet.FileInfo = new List<string>();

                            TODO dowork = new TODO(Packet);

                            //dowork.CountFiles();

                            dowork.Work();

                            SendPacket(Packet);

                        }

                        Packet = null;
                    }

                    clientSocket.Close();
                }
            }
            catch (SocketException ex)
            {
                Trace.TraceError(String.Format("QuoteServer {0}", ex.Message));
            }
        }

        public void Stop()
        {
            listener.Stop();
        }
        public void Suspend()
        {
            listener.Stop();
        }
        public void Resume()
        {
            listener.Start();
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

