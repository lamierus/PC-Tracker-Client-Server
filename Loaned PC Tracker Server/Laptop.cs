using System;
using System.Text;
using System.Collections.Generic;

namespace Loaned_PC_Tracker_Server {
    public class Laptop : IEquatable<Laptop> {

        public int Number { get; set; }
        public string Serial { get; set; }
        public string Brand { get; set; }
        public string Model { get; set; }
        public string Warranty { get; set; }
        public string Username { get; set; }
        public string UserPCSerial { get; set; }
        public string TicketNumber { get; set; }
        public bool CheckedOut;

        public Laptop() {

        }

        public Laptop(int loanerNum, string brand, string model, string serialNumber, string warranty) {
            Number = loanerNum;
            Serial = serialNumber;
            Brand = brand;
            Model = model;
            Warranty = warranty;
        }

        public Laptop(int loanerNumber, string brand, string model, string serialNumber, string warranty, string username, string userSerialNumber, string ticketNumber, bool checkedOut = true) {
            Number = loanerNumber;
            Serial = serialNumber;
            Brand = brand;
            Model = model;
            Warranty = warranty;
            Username = username;
            UserPCSerial = userSerialNumber;
            TicketNumber = ticketNumber;
            CheckedOut = checkedOut;
        }

        public byte[] SerializeLaptop() {
            byte[] seperator = BitConverter.GetBytes(',');
            List<byte> serializedPC = new List<byte>();
            
            serializedPC.AddRange(Encoding.UTF8.GetBytes(Number.ToString()));

            serializedPC.AddRange(seperator);
            
            // Add the name
            if (Serial != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(Serial));

            serializedPC.AddRange(seperator);

            // Add the name
            if (Brand != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(Brand));

            serializedPC.AddRange(seperator);

            // Add the name
            if (Model != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(Model));

            serializedPC.AddRange(seperator);

            // Add the name
            if (Warranty != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(Warranty));

            serializedPC.AddRange(seperator);

            // Add the name
            if (Username != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(Username));

            serializedPC.AddRange(seperator);

            // Add the name
            if (UserPCSerial != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(UserPCSerial));

            serializedPC.AddRange(seperator);

            // Add the name
            if (TicketNumber != null)
                serializedPC.AddRange(Encoding.UTF8.GetBytes(TicketNumber));

            serializedPC.AddRange(seperator);

            //serializedPC.AddRange(BitConverter.GetBytes(CheckedOut));
            serializedPC.AddRange(Encoding.UTF8.GetBytes(CheckedOut.ToString()));
            
            serializedPC.AddRange(BitConverter.GetBytes(';'));


            return serializedPC.ToArray();
        }

        public Laptop DeserializeLaptop(byte[] serializedPC) {
            char[] seperator = new char[] { ',' };
            Laptop deserializedPC = new Laptop();

            string dataString = Encoding.UTF8.GetString(serializedPC);
            string[] splitString = dataString.Split(seperator, StringSplitOptions.RemoveEmptyEntries);

            int parsedNum;
            if (int.TryParse(splitString[0], out parsedNum)) {
                deserializedPC.Number = parsedNum;
            }

            deserializedPC.Serial = splitString[1];
            deserializedPC.Brand = splitString[2];
            deserializedPC.Model = splitString[3];
            deserializedPC.Warranty = splitString[4];
            deserializedPC.Username = splitString[5];
            deserializedPC.UserPCSerial = splitString[6];
            deserializedPC.TicketNumber = splitString[7];

            if (splitString[8].ToLower() == "true") {
                deserializedPC.CheckedOut = true;
            } else {
                deserializedPC.CheckedOut = false;
            }

            return deserializedPC;
        }

        public void MergeChanges(PCChange changes) {
            Serial = changes.Serial;
            Username = changes.UserName;
            UserPCSerial = changes.UserPCSerial;
            TicketNumber = changes.Ticket;
            CheckedOut = changes.CheckedOut;
        }

        // the logic required to be able to compare CSATs to each other
        public override bool Equals(object obj) {
            if (obj == null) {
                return false;
            }
            Laptop objAsPC = obj as Laptop;
            if (objAsPC == null) {
                return false;
            } else {
                return Equals(objAsPC);
            }
        }

        public override int GetHashCode() {
            return Serial.GetHashCode();
        }

        public bool Equals(Laptop other) {
            if (other == null) {
                return false;
            }
            return (Serial.Equals(other.Serial));
        }

        public static bool operator ==(Laptop lhs, Laptop rhs) {
            if (ReferenceEquals(lhs, null)) {
                return ReferenceEquals(rhs, null);
            }
            return lhs.Equals(rhs);
        }

        public static bool operator !=(Laptop lhs, Laptop rhs) {
            if (ReferenceEquals(lhs, null)) {
                return ReferenceEquals(rhs, null);
            }
            return !(lhs.Equals(rhs));
        }
    }
}
