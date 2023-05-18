using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppWorkerWithExcelDoc.Model
{
    public class Order
    {
        public string IdOrder { get; set; }
        public string IdProduct { get; set; }
        public Product Product { get; set; }
        public string IdClient { get; set; }
        public Client Client { get; set; }
        public string NumberOrder { get; set; }
        public int Quantity { get; set; }
        public DateTime DatePosting { get; set; }
    }
}
