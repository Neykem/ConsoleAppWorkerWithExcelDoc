using ConsoleAppWorkerWithExcelDoc.Domain;
using ConsoleAppWorkerWithExcelDoc.Domain.Interface;

namespace ConsoleAppWorkerWithExcelDoc
{
    class Program
    {

        static void Main(string[] args)
        {
            try
            {
                string filePath = "E:\\Desktop\\test\\Test.xlsx";
                IOrderRepository orderRepository = new OrderRepository(filePath);
                var collectionOrder = orderRepository.LoadData();
                Console.WriteLine("Номер Заказа" + "|" + "Номер клиента " + "|" + "Количество" + "|" + "Дата доставки" + "|");
                foreach (var item in collectionOrder)
                {
                    Console.WriteLine("\t" + item.IdOrder + " |\t\t" + item.IdClient + "|\t    " + item.Quantity + "|   " + item.DatePosting.ToString("d") + "|");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Не удалось открыть файл");
            }

        }
    }
}