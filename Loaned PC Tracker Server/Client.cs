﻿using System;
using System.Net;
using System.Net.Sockets;
using System.ComponentModel;

namespace Loaned_PC_Tracker_Server {
    //Class to handle each client request seperately
    public class Client {
        static public int UserCount { get; set; }
        
        public string UserName { get; set; }
        public string Site { get; set; }
        public bool Hotswaps { get; set; }
        public IPAddress IP {
            get { return ((IPEndPoint)ClientSocket.Client.RemoteEndPoint).Address; }
        }

        private TcpClient ClientSocket;
        private BackgroundWorker bgwWaitForPCRequests = new BackgroundWorker() {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true,
        };

        public Client(TcpClient inClientSocket, PCTrackerServerForm siht) {
            ClientSocket = inClientSocket;
            ClientSocket.NoDelay = true;
            ClientSocket.ReceiveBufferSize = 10025;
            ClientSocket.SendBufferSize = 10025;
            bgwWaitForPCRequests.DoWork += new DoWorkEventHandler(bgwWaitForPCRequests_DoWork);
            //bgwWaitForPCRequests.ProgressChanged += new ProgressChangedEventHandler(bgwWaitForPCRequests_ProgressChanged);
            //bgwWaitForPCRequests.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwWaitForPCRequests_RunWorkerCompleted);
            startClient(siht);
        }
        
        /// <summary>
        ///     this creates and starts the tread for the client object, on the server.
        /// </summary>
        private bool startClient(PCTrackerServerForm siht) {
            byte[] InStream = new byte[10025];
            try {
                //ClientSocket.GetStream().Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                NetworkStream stream = ClientSocket.GetStream();
                stream.Read(InStream, 0, ClientSocket.ReceiveBufferSize);
                NamePacket handshake = new NamePacket(InStream);
                if (handshake.Name != string.Empty && handshake.Name != null) {
                    UserName = handshake.Name;
                } else {
                    UserName = "Client #" + UserCount++.ToString();
                }
            } catch (Exception ex) {
                siht.UpdateStatus("Error Starting Client Connection: " + ex.Message);
                return false;
            }
            bgwWaitForPCRequests.RunWorkerAsync(siht);
            return true;
        }

        public void StreamDataToClient(byte[] dataToSend, PCTrackerServerForm siht) {
            try {
                //siht.UpdateStatus(" >> Sending data to " + UserName);
                NetworkStream outStream = ClientSocket.GetStream();
                outStream.Write(dataToSend, 0, dataToSend.Length);
                outStream.Flush();
            } catch (Exception ex) {
                siht.UpdateStatus("Error Streaming Data to Client:" + ex.Message);
            }
        }

        private void bgwWaitForPCRequests_DoWork(object sender, DoWorkEventArgs e) {
            var siht = e.Argument as PCTrackerServerForm;
            siht.UpdateStatus("> Awaiting requests from " + UserName);
            byte[] inStream = new byte[10025];
            while (true) {
                try {
                    ClientSocket.GetStream().Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                    var streamIdentifier = (DataIdentifier)BitConverter.ToInt32(inStream, 0);
                    if (streamIdentifier == DataIdentifier.Request) {
                        RequestPCPacket pcRequest = new RequestPCPacket(inStream);
                        if (pcRequest.Type == "Hotswaps") {
                            Hotswaps = true;
                        } else {
                            Hotswaps = false;
                        }
                        Site = pcRequest.SiteName;
                        siht.SendPCsForSite(this, pcRequest.SiteName, pcRequest.Type);
                    } else if (streamIdentifier == DataIdentifier.Update) {
                        PCChange changedPC = new PCChange(inStream);
                        //siht.BroadcastUpdateToSite("> User " + UserName + " is modifying " + changedPC.Serial, this);
                        siht.updatePC(changedPC, this);
                    } else if(streamIdentifier == DataIdentifier.Laptop) {

                    }
                } catch (Exception ex) {
                    siht.UpdateStatus("<<< " + UserName + " disconnected: " + ex.Message);
                    if (!ClientSocket.Connected) {
                        siht.RemoveClient(this);
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
