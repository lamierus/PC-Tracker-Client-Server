using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Loaned_PC_Tracker_Server {
    class Site {
        public List<Laptop> AvailableHotswaps;
        public List<Laptop> CheckedOutHotswaps;
        public List<Laptop> AvailableLoaners;
        public List<Laptop> CheckedOutLoaners;
        public string Name;

        public Site(string name) {
            Name = name;
            AvailableHotswaps = new List<Laptop>();
            CheckedOutHotswaps = new List<Laptop>();
            AvailableLoaners = new List<Laptop>();
            CheckedOutLoaners = new List<Laptop>();
        }
    }
}
