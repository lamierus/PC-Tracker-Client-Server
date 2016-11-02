using System;
using System.Collections.Generic;

namespace Loaned_PC_Tracker_Client {
    public class Packet {
        public DataIdentifier Identifier { get; set; }
        public int PacketLength {
            get { return CreateDataStream().Length; }
        }

        // Default Constructor
        public Packet() {
            Identifier = DataIdentifier.Null;
        }

        public byte[] CreateDataStream() {
            List<byte> dataStream = new List<byte>();

            // Add the dataIdentifier
            dataStream.AddRange(BitConverter.GetBytes((int)Identifier));

            return dataStream.ToArray();
        }
    }
}