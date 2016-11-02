using System;
using System.Net.Sockets;
using System.Threading;

namespace Loaned_PC_Tracker_Server {
    //Class to handle each client request seperately
    public class Client {
        public string UserName { get; set; }
        public TcpClient ClientSocket;
        public Thread ClientThread;
        //public PCPacket SendPack = new PCPacket();

        private byte[] InStream = new byte[10025];

        public Client() { }

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
            NamePacket handshake = new NamePacket();
            try {
                NetworkStream networkStream = ClientSocket.GetStream();
                networkStream.Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                handshake.ParsePacket(InStream);
                UserName = handshake.Name;
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
                return false;
            }
            ClearStream();
            return true;
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public PCPacket GetPCPacket() {
            PCPacket receivedPacket = new PCPacket();
            try {
                NetworkStream networkStream = ClientSocket.GetStream();
                networkStream.Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                receivedPacket.GetPacket(InStream);
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
            ClearStream();
            return receivedPacket;
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public void SendPacketToClient(Packet packet) {
            try {
                NetworkStream networkStream = ClientSocket.GetStream();
                networkStream.Write(packet.CreateDataStream(), 0, packet.PacketLength);
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
        }

        private void ClearStream() {
            InStream = new byte[10025];
        }
    }
}
