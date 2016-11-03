using System;
using System.Net.Sockets;
using System.Threading;

namespace Loaned_PC_Tracker_Server {
    //Class to handle each client request seperately
    public class Client {
        static int UserCount { get; set; }
        public string UserName { get; set; }
        //public PCPacket SendPack = new PCPacket();

        //private byte[] InStream = new byte[10025];
        private TcpClient ClientSocket;
        //private Thread ClientThread;

        public Client() {
            UserName = "Client #" + UserCount++.ToString();
        }

        public Client(TcpClient inClientSocket) {
            ClientSocket = inClientSocket;
            ClientSocket.NoDelay = true;
            startClient();
        }

        /// <summary>
        ///     this creates and starts the tread for the client object, on the server.
        /// </summary>
        /// <param name="inClientSocket"></param>
        private bool startClient() {
            byte[] inStream = new byte[10025];
            try {
                ClientSocket.GetStream().Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                NamePacket handshake = new NamePacket(inStream);
                if (handshake.Name != string.Empty && handshake.Name != null) {
                    UserName = handshake.Name;
                } else {
                    UserName = "Client #" + UserCount++.ToString();
                }
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
                return false;
            }
            return true;
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public PCPacket GetPCPacket() {
            byte[] inStream = new byte[10025];
            PCPacket receivedPacket = new PCPacket();
            try {
                ClientSocket.GetStream().Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                receivedPacket.GetPacket(inStream);
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
            return receivedPacket;
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public void SendPacketToClient(Packet packet) {
            try {
                ClientSocket.GetStream().Write(packet.CreateDataStream(), 0, packet.PacketLength);
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
        }
    }
}
