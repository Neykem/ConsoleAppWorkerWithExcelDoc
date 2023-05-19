using ConsoleAppWorkerWithExcelDoc.Domain.Interface;
using ConsoleAppWorkerWithExcelDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleAppWorkerWithExcelDoc.Domain
{
    public class OrderRepository : IOrderRepository
    {
        readonly string _filePath;
        private List<Order> _ordersBuffer { get; set; }

        public OrderRepository(string filePath)
        {
            _filePath = filePath;
            _ordersBuffer = LoadData();
        }

        public List<Client> GetIdGoldenCustomerFromData(DateTime dateTime)
        {

            List<Order> orders = LoadData();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart clientsWorksheetPart = workbookPart.GetPartById("rId2") as WorksheetPart;
                List<Client> groupedOrders = orders.Where(order => order.DatePosting.Year == dateTime.Year && order.DatePosting.Month == dateTime.Month).GroupBy(order => order.IdClient).Select(group =>
                {
                    string clientId = group.Key;
                    int orderCount = group.Count();
                    Client client = GetClientData(clientsWorksheetPart, clientId, workbookPart);
                    return client;
                }).ToList();

                return groupedOrders;
            }
        }

        //Метод загрузки данныых из файла
        public List<Order> LoadData()
        {
            List<Order> orders = new List<Order>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                // Получение ссылок на листы документа
                WorksheetPart productsWorksheetPart = workbookPart.GetPartById("rId1") as WorksheetPart;
                WorksheetPart clientsWorksheetPart = workbookPart.GetPartById("rId2") as WorksheetPart;
                WorksheetPart ordersWorksheetPart = workbookPart.GetPartById("rId3") as WorksheetPart;

                if (productsWorksheetPart == null || clientsWorksheetPart == null || ordersWorksheetPart == null)
                {
                    return orders;
                }
                //Пробе по строчкам, пропуская строку с названиями
                IEnumerable<Row> rows = ordersWorksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                foreach (Row row in rows.Skip(1))
                {
                    Order order = new Order();
                    string clientId = GetCellValue(GetCell(ordersWorksheetPart, row, "C"), workbookPart);
                    string productId = GetCellValue(GetCell(ordersWorksheetPart, row, "B"), workbookPart);
                    if (productId == null || clientId == null || clientId == "" || productId == "")
                    {
                        return orders;
                    }
                    Client client = GetClientData(clientsWorksheetPart, clientId, workbookPart);
                    order.Client = client;
                    order.IdClient = client.Id;

                    Product product = GetProductData(productsWorksheetPart, productId, workbookPart);
                    order.Product = product;
                    order.IdProduct = productId;

                    order.IdOrder = (GetCellValue(GetCell(ordersWorksheetPart, row, "A"), workbookPart));
                    order.NumberOrder = (GetCellValue(GetCell(ordersWorksheetPart, row, "B"), workbookPart));
                    order.Quantity = int.Parse(GetCellValue(GetCell(ordersWorksheetPart, row, "E"), workbookPart));
                    order.DatePosting = DateTime.FromOADate(double.Parse(GetCellValue(GetCell(ordersWorksheetPart, row, "F"), workbookPart)));

                    orders.Add(order);
                }
            }

            return orders;
        }

        //Изменение данных клиента по номеру
        public string ChangeClientContact(string organizationName, string newContactName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart clientsWorksheetPart = workbookPart.GetPartById("rId2") as WorksheetPart;

                IEnumerable<Row> rows = clientsWorksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                foreach (Row row in rows.Skip(1))
                {
                    string organization = GetCellValue(GetCell(clientsWorksheetPart, row, "B"), workbookPart);

                    if (organization == organizationName)
                    {
                        SetCellValue(GetCell(clientsWorksheetPart, row, "D"), newContactName);

                        document.Save();

                        _ordersBuffer = LoadData();
                        return "Информация успешно сохранена!";
                    }
                }
            }

            return "Организация не найдена!";
        }

        // Метод для получения данных о заявке по кленту
        public List<Order> GetOrdersByProductName(string productName)
        {
            if (_ordersBuffer == null)
            {
                _ordersBuffer = LoadData();
            }
            List<Order> filteredOrders = _ordersBuffer.Where(order => order.Product.Name == productName).ToList();
            return filteredOrders;
        }

        //Метод для установки нового значения в ячейку
        private static void SetCellValue(Cell cell, string newData)
        {
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            cell.CellValue = new CellValue(newData);
        }

        //Метод для получения значений в данной ячейке
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            //Нашел вот такую рекомендацию, на проверку типов, решил добавить ее )
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringId = int.Parse(cell.InnerText);
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringTablePart != null)
                {
                    return sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringId).InnerText;
                }
            }

            return cell.InnerText;
        }

        // Метод для получения ячейки по адресу столбца и строки
        private static Cell GetCell(WorksheetPart worksheetPart, Row row, string columnName)
        {
            string cellReference = columnName + row.RowIndex;
            return worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
        }

        // Метод для получения данных о клиенте по его идентификатору
        private static Client GetClientData(WorksheetPart clientsWorksheetPart, string clientId, WorkbookPart workbookPart)
        {
            IEnumerable<Row> rows = clientsWorksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            foreach (Row row in rows.Skip(1))
            {
                string id = GetCellValue(GetCell(clientsWorksheetPart, row, "A"), workbookPart);

                if (id == clientId)
                {
                    Client client = new Client();
                    client.Id = id;
                    client.Organization = GetCellValue(GetCell(clientsWorksheetPart, row, "B"), workbookPart);
                    client.Adress = GetCellValue(GetCell(clientsWorksheetPart, row, "C"), workbookPart);
                    client.Name = GetCellValue(GetCell(clientsWorksheetPart, row, "D"), workbookPart);

                    return client;
                }
            }

            return null;
        }

        // Метод для получения данных о продукте по его идентификатору
        private static Product GetProductData(WorksheetPart productsWorksheetPart, string productId, WorkbookPart workbookPart)
        {
            IEnumerable<Row> rows = productsWorksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            foreach (Row row in rows.Skip(1))
            {
                string id = GetCellValue(GetCell(productsWorksheetPart, row, "A"), workbookPart);

                if (id == productId)
                {
                    Product product = new Product();
                    product.Id = id;
                    product.Name = GetCellValue(GetCell(productsWorksheetPart, row, "B"), workbookPart);
                    product.UnitMeasure = GetCellValue(GetCell(productsWorksheetPart, row, "C"), workbookPart);
                    product.Cost = (GetCellValue(GetCell(productsWorksheetPart, row, "D"), workbookPart));

                    return product;
                }
            }

            return null;
        }
    }
}
