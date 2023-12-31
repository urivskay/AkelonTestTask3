using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTestExcel.Repository
{
    internal class Order
    {
        public int Code { get; set; }
        public Product Product { get; set; }
        public Client Client { get; set; }
        public int Number { get; set; }
        public int CountProduct { get; set; }
        public DateTime DateCreated { get; set; }

        public Order(Product product, Client client)
        {
            Product = product;
            Client = client;
        }

    }

    internal struct ClientOrders
    {
        public Client Client { get; set; }
        public int CountOrders { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
    }
}
