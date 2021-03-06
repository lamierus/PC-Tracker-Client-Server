﻿using System;
using System.Text;
using System.Collections.Generic;

namespace Loaned_PC_Tracker_Client {
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

        static char[] Seperator = new char[] { ',' };
        static char[] TrimChar = new char[] { '\0' };

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

        public Laptop(byte[] dataStream) {
            DeserializeLaptop(dataStream);
        }

        public byte[] SerializeLaptop() {
            var seperator = BitConverter.GetBytes(',');
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

        private void DeserializeLaptop(byte[] serializedPC) {
            string dataString = Encoding.UTF8.GetString(serializedPC);
            string[] splitString = dataString.Split(Seperator, StringSplitOptions.RemoveEmptyEntries);
            
            int parsedNum;
            if (int.TryParse(splitString[0].Trim(TrimChar), out parsedNum)) {
                Number = parsedNum;
            }

            Serial = splitString[1].Trim(TrimChar);
            Brand = splitString[2].Trim(TrimChar);
            Model = splitString[3].Trim(TrimChar);
            Warranty = splitString[4].Trim(TrimChar);
            Username = splitString[5].Trim(TrimChar);
            UserPCSerial = splitString[6].Trim(TrimChar);
            TicketNumber = splitString[7].Trim(TrimChar);

            if (splitString[8].Trim(TrimChar).ToLower() == "true") {
                CheckedOut = true;
            } else {
                CheckedOut = false;
            }
        }

        public Laptop DeserializeLaptop(string serializedPC) {
            Laptop deserializedPC = new Laptop();
            
            string[] splitString = serializedPC.Split(Seperator, StringSplitOptions.RemoveEmptyEntries);
            
            int parsedNum;
            if (int.TryParse(splitString[0].Trim(TrimChar), out parsedNum)) {
                deserializedPC.Number = parsedNum;
            }

            deserializedPC.Serial = splitString[1].Trim(TrimChar);
            deserializedPC.Brand = splitString[2].Trim(TrimChar);
            deserializedPC.Model = splitString[3].Trim(TrimChar);
            deserializedPC.Warranty = splitString[4].Trim(TrimChar);
            deserializedPC.Username = splitString[5].Trim(TrimChar);
            deserializedPC.UserPCSerial = splitString[6].Trim(TrimChar);
            deserializedPC.TicketNumber = splitString[7].Trim(TrimChar);

            if (splitString[8].Trim(TrimChar).ToLower() == "true") {
                deserializedPC.CheckedOut = true;
            } else {
                deserializedPC.CheckedOut = false;
            }

            return deserializedPC;
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
            if (other == null || other.Serial == null) {
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
