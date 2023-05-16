using ConsoleAppWorkerWithExcelDoc.Domain.Interface;
using ConsoleAppWorkerWithExcelDoc.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppWorkerWithExcelDoc.Domain
{
    public class OrderRepository : IOrderRepository
    {
        readonly string _filePach;
        public OrderRepository(string filePach)
        {
            _filePach = filePach;
        }

        public List<int> GetIdGoldenCustomerFromData(DateTime selectData)
        {
            throw new NotImplementedException();
        }

        public List<Order> LoadData()
        {
            throw new NotImplementedException();
        }
        public Task<List<Order>> LoadDataAcync()
        {
            throw new NotImplementedException();
        }

        public bool SaveNewData(Order dataForSave)
        {
            throw new NotImplementedException();
        }
        public Task<bool> SaveNewDataAcync(Order dataForSave)
        {
            throw new NotImplementedException();
        }

        public bool Update(Order oldDataobject, Order dataForUpdate)
        {
            throw new NotImplementedException();
        }
        public Task<bool> UpdateAcync(Order oldDataobject, Order dataForUpdate)
        {
            throw new NotImplementedException();
        }
    }
}
