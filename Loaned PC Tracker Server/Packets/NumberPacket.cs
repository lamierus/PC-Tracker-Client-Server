using System;
using System.Collections.Generic;
using System.Text;

namespace Loaned_PC_Tracker_Server {
    public class NumberPacket : Packet {
        public int Number { get; set; }
        public override int PacketLength {
            get { return CreateDataStream().Length; }
        }

        // Default Constructor
        public NumberPacket() {
            Identifier = DataIdentifier.Number;
            Number = 0;
        }

        public NumberPacket(int numberToSend) {
            Identifier = DataIdentifier.Number;
            Number = numberToSend;
        }

        public NumberPacket(byte[] dataStream) {
            GetPacket(dataStream);
        }

        public void GetPacket(byte[] dataStream) {
            // Read the data identifier from the beginning of the stream (4 bytes)
            Identifier = DataIdentifier.Number;

            // Read the Number field (4 bytes)
            Number = BitConverter.ToInt32(dataStream, 4);
        }

        // Converts the packet into a byte array for sending/receiving 
        public override byte[] CreateDataStream() {
            List<byte> dataStream = new List<byte>();

            // Add the dataIdentifier
            dataStream.AddRange(BitConverter.GetBytes((int)Identifier));
            
            // Add the number
            if (Number != 0)
                dataStream.AddRange(BitConverter.GetBytes(Number));

            dataStream.AddRange(Encoding.UTF8.GetBytes(";"));

            return dataStream.ToArray();
        }
    }
}
