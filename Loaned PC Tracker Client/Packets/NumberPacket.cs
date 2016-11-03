using System;
using System.Collections.Generic;

namespace Loaned_PC_Tracker_Client {
    public class NumberPacket : Packet {
        public int Number { get; set; }
        public new int PacketLength {
            get { return CreateDataStream().Length; }
        }

        // Default Constructor
        public NumberPacket() {
            Identifier = DataIdentifier.Message;
            Number = 0;
        }

        public NumberPacket(int numberToSend) {
            Identifier = DataIdentifier.Message;
            Number = numberToSend;
        }

        public NumberPacket(byte[] dataStream) {
            GetPacket(dataStream);
        }

        public void GetPacket(byte[] dataStream) {
            // Read the data identifier from the beginning of the stream (4 bytes)
            Identifier = DataIdentifier.Message;

            // Read the length of the name (4 bytes)
            int nameLength = BitConverter.ToInt32(dataStream, 4);

            // Read the name field
            if (nameLength > 0)
                Number = BitConverter.ToInt32(dataStream, 8);
            else
                Number = 0;
        }

        // Converts the packet into a byte array for sending/receiving 
        public new byte[] CreateDataStream() {
            List<byte> dataStream = new List<byte>();

            // Add the dataIdentifier
            dataStream.AddRange(BitConverter.GetBytes((int)Identifier));

            // Add the name length
            if (Number != 0)
                dataStream.AddRange(BitConverter.GetBytes(Number));
            else
                dataStream.AddRange(BitConverter.GetBytes(0));

            // Add the name
            if (Number != 0)
                dataStream.AddRange(BitConverter.GetBytes(Number));

            return dataStream.ToArray();
        }
    }
}
