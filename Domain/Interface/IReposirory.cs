namespace ConsoleAppWorkerWithExcelDoc.Domain.Interface
{
    //Итерфейс для универсального репозитория с обозначением методов
    internal interface IReposirory<T>
    {
        //Метод возврашающий коллекцию с загруженными данными из каталога
        public List<T> LoadData();
    }
}
