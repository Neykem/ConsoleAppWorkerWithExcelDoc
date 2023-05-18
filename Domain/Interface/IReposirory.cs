namespace ConsoleAppWorkerWithExcelDoc.Domain.Interface
{
    //Итерфейс для универсального репозитория с обозначением методов и их асинхроного аналога
    internal interface IReposirory<T>
    {
        //Метод возврашающий коллекцию с загруженными данными из каталога
        public List<T> LoadData();
        public Task<List<T>> LoadDataAcync();
        //Метод сохраняющий обьект в каталоге, должен возвращять результат успешности выполнения
        public bool SaveNewData(T dataForSave);
        public Task<bool> SaveNewDataAcync(T dataForSave);
        //Метод обновляющий данные обьекта в каталоге, должен возвращять результат успешности выполнения
        public bool Update(T oldDataobject, T dataForUpdate);
        public Task<bool> UpdateAcync(T oldDataobject, T dataForUpdate);
    }
}
