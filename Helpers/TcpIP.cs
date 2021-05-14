using System;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace SBO.Hub.Helpers
{
    public class TcpIP
    {
        private static TcpClient Client;

        public static string ReadData(string ip, int port, char breakChar, int breakSize)
        {
            try
            {
                if (Client == null || !Client.Connected || ((IPEndPoint)Client.Client.RemoteEndPoint).Address != IPAddress.Parse(ip) || ((IPEndPoint)Client.Client.RemoteEndPoint).Port != port)
                {
                    Client = new TcpClient();
                    Client.Connect(new IPEndPoint(IPAddress.Parse(ip), port));
                    while (!Client.Connected) { } // Wait for connection
                }

                NetworkStream network = Client.GetStream();
                
                StringBuilder message = new StringBuilder();
                while (true)
                {

                    if (network.DataAvailable)
                    {
                        int read = network.ReadByte();
                        if ((char)read == breakChar )
                        {
                            break;
                        }

                        if (read > 0)
                        {
                            message.Append((char)read);
                            if (message.Length == breakSize)
                            {
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                Client.Close();

                return message.ToString();

            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
