using System;
using System.Net;
using System.Net.Sockets;
using System.ComponentModel;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Loaned_PC_Tracker_Server {
    //Class to handle each client request seperately
    public class Client {
        static public int UserCount { get; set; }

        private BackgroundWorker bgwWaitForPCRequests = new BackgroundWorker() {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true,
        };
        public string UserName { get; set; }
        public string Site { get; set; }
        public bool Hotswaps { get; set; }
        public IPAddress IP {
            get { return ((IPEndPoint)ClientSocket.Client.RemoteEndPoint).Address; }
        }

        private TcpClient ClientSocket;

        public Client(TcpClient inClientSocket, PCTrackerServerForm siht) {
            ClientSocket = inClientSocket;
            ClientSocket.NoDelay = true;
            startClient(siht);
            initializeBGW();
        }

        private void initializeBGW() {
            bgwWaitForPCRequests.DoWork += new DoWorkEventHandler(bgwWaitForPCRequests_DoWork);
            bgwWaitForPCRequests.ProgressChanged += new ProgressChangedEventHandler(bgwWaitForPCRequests_ProgressChanged);
            bgwWaitForPCRequests.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwWaitForPCRequests_RunWorkerCompleted);
        }
        /// <summary>
        ///     this creates and starts the tread for the client object, on the server.
        /// </summary>
        private bool startClient(PCTrackerServerForm siht) {
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
                siht.UpdateStatus(ex.Message);
                return false;
            }
            bgwWaitForPCRequests.RunWorkerAsync(siht);
            return true;
        }

        /// <summary>
        ///     this is the function that takes care of receiving the packets from the different users,
        ///     interpreting them and sending out the correct broadcast messages to the other users.
        /// </summary>
        public PCPacket GetPCPacket(PCTrackerServerForm siht) {
            byte[] InStream = new byte[10025];
            PCPacket receivedPacket = new PCPacket();
            try {
                ClientSocket.GetStream().Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                receivedPacket.GetPacket(InStream);
            } catch (Exception ex) {
                siht.UpdateStatus(ex.Message);
            }
            return receivedPacket;
        }

        public void StreamDataToClient(byte[] dataToSend, PCTrackerServerForm siht) {
            try {
                NetworkStream outStream = ClientSocket.GetStream();
                outStream.Write(dataToSend, 0, dataToSend.Length);
                outStream.Flush();
            } catch (Exception ex) {
                siht.UpdateStatus(ex.Message);
            }
        }

        private void bgwWaitForPCRequests_DoWork(object sender, DoWorkEventArgs e) {
            var siht = e.Argument as PCTrackerServerForm;
            siht.UpdateStatus("Awaiting requests from " + UserName);
            byte[] inStream = new byte[10025];
            NetworkStream stream;
            while (true) {
                try {
                    stream = ClientSocket.GetStream();
                    stream.Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                    RequestPCPacket pcRequest = new RequestPCPacket(inStream);
                    Site = pcRequest.SiteName;
                    siht.SendPCsForSite(this, pcRequest.SiteName, pcRequest.Type);
                } catch (Exception ex) {
                    siht.UpdateStatus(ex.Message);
                    if (!ClientSocket.Connected) {
                        break;
                    }
                }
            }
        }

        private void bgwWaitForPCRequests_ProgressChanged(object sender, ProgressChangedEventArgs e) {

        }

        private void bgwWaitForPCRequests_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {

        }
    }
}
