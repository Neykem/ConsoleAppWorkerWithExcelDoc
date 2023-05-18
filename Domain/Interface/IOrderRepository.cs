using ConsoleAppWorkerWithExcelDoc.Model;

namespace ConsoleAppWorkerWithExcelDoc.Domain.Interface
{
    //Интерфейс для обозначения работы с конкретным типом таблицы в данном случай Order, добавления ему дополнительного метода
    internal interface IOrderRepository : IReposirory<Order>
    {
        //возвращает коллекцию id пользователей с самым большим количеством подписок
        public List<Client> GetIdGoldenCustomerFromData(DateTime dateTime);
        public List<Order> GetOrdersByProductName(string productName);
        public string ChangeClientContact(string organizationName, string newContactName);
    }
}
