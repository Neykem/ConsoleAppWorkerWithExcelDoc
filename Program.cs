using ConsoleAppWorkerWithExcelDoc.Domain;
using ConsoleAppWorkerWithExcelDoc.Domain.Interface;

namespace ConsoleAppWorkerWithExcelDoc
{
    class Program
    {

        static void Main(string[] args)
        {
            string _filePath = "";
            string _command = "";

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("═════════════════════════════════════════════════════════ ═ ═ ═ ═ ═ ");
            Console.WriteLine(" Вас приведствует программа для работы с таблицей Excel");
            Console.WriteLine(" Введите путь к файлу для продолжения для продолжения: (Для выхода введите | exit |)");
            Console.WriteLine("═════════════════════════════════════════════════════════ ═ ═ ═ ═ ═ ");
            Console.WriteLine("Введите путь до файла:");
            Console.ForegroundColor = ConsoleColor.DarkRed;
            _filePath = Console.ReadLine();
             _filePath = "E:\\Desktop\\test\\Test.xlsx";
            Console.ForegroundColor = ConsoleColor.Green;
            IOrderRepository orderRepository = new OrderRepository(_filePath);
            Console.WriteLine("═════════════════════════════════════════════════════════ ═ ═ ═ ═ ═ ");
            Console.WriteLine(" Введите команду из списка доступных опираций:");
            Console.WriteLine(" Введите | 1 | для вывода информаций клиентов по товару... ");
            Console.WriteLine(" Введите | 2 | для изменения контактного лица в указанной организаций... ");
            Console.WriteLine(" Введите | 3 | выдачи золотого клиента за указанный месяц в году... ");
            Console.WriteLine("═════════════════════════════════════════════════════════ ═ ═ ═ ═ ═ ");
            while (_command != "exit") 
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(" Введите команду:");
                Console.ForegroundColor = ConsoleColor.DarkRed;

                _command = Console.ReadLine();
                if(_command == "1")
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(" Введите название продукта который покупал клиент:");
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    _command = Console.ReadLine();
                    var a = orderRepository.GetOrdersByProductName(_command);
                    if(a != null)
                    {
                        foreach (var item in a)
                        {
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine("\t" + item.Client.Name + "\t" + item.DatePosting.ToString("d"));
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(" По данному товару нет информаций");
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                    }

                }
                else if (_command == "2") 
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(" Введите название фирмы:");
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    string _command1 = Console.ReadLine();

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(" Введите ФИО:");
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    string _command2 = Console.ReadLine();
                    
                    Console.WriteLine(orderRepository.ChangeClientContact(_command1, _command2));
                }
                else if (_command == "3")
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(" Введите год:");
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    int _command1 = int.Parse(Console.ReadLine());
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(" Введите месяц:");
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    int _command2 = int.Parse(Console.ReadLine());
                    DateTime date = new DateTime(_command1, _command2, 01);
                    var b = orderRepository.GetIdGoldenCustomerFromData(date);
                    foreach (var item in b)
                    {
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine("\t" + item.Name + "\t" + item.Id);
                    }
                }

            }
        }
    }
}