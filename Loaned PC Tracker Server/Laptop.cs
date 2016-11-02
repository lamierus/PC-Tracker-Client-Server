using System;

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
            Brand = brand;
            Model = model;
            Serial = serialNumber;
            Warranty = warranty;
        }

        public Laptop(int loanerNumber, string brand, string model, string serialNumber, string warranty, string username, string userSerialNumber, string ticketNumber, bool checkedOut = true) {
            Number = loanerNumber;
            Brand = brand;
            Model = model;
            Serial = serialNumber;
            Warranty = warranty;
            Username = username;
            UserPCSerial = userSerialNumber;
            TicketNumber = ticketNumber;
            CheckedOut = checkedOut;
        }

        // the logic required to be able to compare CSATs to each other
        public override bool Equals(Object obj) {
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
