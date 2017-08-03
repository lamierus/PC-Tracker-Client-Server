using System;
using System.Collections.Generic;
using System.Text;

namespace Loaned_PC_Tracker_Client {
    public class Packet {
        public DataIdentifier Identifier { get; set; }
        public virtual int PacketLength {
            get { return CreateDataStream().Length; }
        }

        // Default Constructor
        public Packet() {
            Identifier = DataIdentifier.Null;
        }

        public virtual byte[] CreateDataStream() {
            List<byte> dataStream = new List<byte>();

            // Add the dataIdentifier
            dataStream.AddRange(BitConverter.GetBytes((int)Identifier));

            dataStream.AddRange(Encoding.UTF8.GetBytes(";"));

            return dataStream.ToArray();
        }

        public string DataStreamToString() {
            string stream = "";

            foreach (byte b in CreateDataStream()) {
                stream += b.ToString() + " ";
            }

            return stream;
        }

        public string DataStreamToString(byte[] dataStream) {
            string stream = "";

            foreach (byte b in dataStream) {
                stream += b.ToString() + " ";
            }

            return stream;
        }
    }
}