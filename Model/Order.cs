using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppWorkerWithExcelDoc.Model
{
    public class Order
    {
        public int IdOrder { get; set; }
        public int IdProduct { get; set; }
        public int IdClient { get; set; }
        public int NumberOrder { get; set; }
        public int Quantity { get; set; }
        public DateTime DatePosting { get; set; }
    }
}
