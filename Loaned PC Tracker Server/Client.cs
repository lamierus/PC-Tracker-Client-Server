using System;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Loaned_PC_Tracker_Server {
    //Class to handle each client request seperately
    public class Client {
        static int UserCount { get; set; }
        public string UserName { get; set; }
        public string Site { get; set; }
        public bool Hotswaps { get; set; }
        public IPAddress IP {
            get { return ((IPEndPoint)ClientSocket.Client.RemoteEndPoint).Address; }
        }

        //private byte[] InStream;
        private TcpClient ClientSocket;

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
            byte[] InStream = new byte[10025];
            try {
                ClientSocket.GetStream().Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                NamePacket handshake = new NamePacket(InStream);
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
            byte[] InStream = new byte[10025];
            PCPacket receivedPacket = new PCPacket();
            try {
                ClientSocket.GetStream().Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                receivedPacket.GetPacket(InStream);
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
            return receivedPacket;
        }

        public void StreamDataToClient(byte[] dataToSend) {
            try {
                NetworkStream outStream = ClientSocket.GetStream();
                outStream.Write(dataToSend, 0, dataToSend.Length);
                outStream.Flush();
            } catch {

            }
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public void SendPacketToClient(NamePacket packet) {
            try {
                NetworkStream outStream = ClientSocket.GetStream();
                outStream.Write(packet.CreateDataStream(), 0, packet.PacketLength);
                outStream.Flush();
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public void SendPacketToClient(NumberPacket packet) {
            try {
                NetworkStream outStream = ClientSocket.GetStream();
                outStream.Write(packet.CreateDataStream(), 0, packet.PacketLength);
                outStream.Flush();
            } catch (Exception ex) {
                //Console.WriteLine(" >> " + ex.Message.ToString());
            }
        }
    }
}
