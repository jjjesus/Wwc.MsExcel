using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wwc.MsExcel
{
    public class Customer
    {
        public double Id { get; set; }
        public string FirstName { get; set;  }
        public string LastName { get; set; }
        public string City { get; set; }
        public string State { get; set; }

        public static List<Customer> CustomerList { get; set; }

        static Customer()
        {
            CustomerList = new List<Customer>();
        }

        public override string ToString()
        {
            string output = string.Format("{0}:{1}:{2}:{3}:{4}",
                this.Id, this.FirstName, this.LastName, this.City, this.State);
            return output;
        }
    }
}
