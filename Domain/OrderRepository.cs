using ConsoleAppWorkerWithExcelDoc.Domain.Interface;
using ConsoleAppWorkerWithExcelDoc.Model;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleAppWorkerWithExcelDoc.Domain
{
    public class OrderRepository : IOrderRepository
    {
        readonly string _filePath;
        public OrderRepository(string filePath)
        {
            _filePath = filePath;
        }

        public List<int> GetIdGoldenCustomerFromData(DateTime selectData)
        {
            throw new NotImplementedException();
        }

        public List<Order> LoadData()
        {
            List<Order> orders = new List<Order>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var rows = sheetData.Descendants<Row>();

                

                foreach (Row row in rows.Skip(1)) // Пропускаем заголовок
                {
                    Order order = new Order();

                    Cell idOrderCell = row.Elements<Cell>().ElementAt(0);
                    Cell idProductCell = row.Elements<Cell>().ElementAt(1);
                    Cell idClientCell = row.Elements<Cell>().ElementAt(2);
                    Cell numberOrderCell = row.Elements<Cell>().ElementAt(3);
                    Cell quantityCell = row.Elements<Cell>().ElementAt(4);
                    Cell datePostingCell = row.Elements<Cell>().ElementAt(5);
                    if (idOrderCell.CellValue != null)
                    {
                        order.IdOrder = Convert.ToInt32(GetCellValue(workbookPart, idOrderCell));
                        order.IdProduct = Convert.ToInt32(GetCellValue(workbookPart, idProductCell));
                        order.IdClient = Convert.ToInt32(GetCellValue(workbookPart, idClientCell));
                        order.NumberOrder = Convert.ToInt32(GetCellValue(workbookPart, numberOrderCell));
                        order.Quantity = Convert.ToInt32(GetCellValue(workbookPart, quantityCell));
                        order.DatePosting = DateTime.FromOADate(Convert.ToDouble(GetCellValue(workbookPart, datePostingCell)));
                    }
                    else
                    {
                        return orders;
                    }
                    orders.Add(order);
                }
            }
            return orders;
        }
        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedStringPart.SharedStringTable.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText;
            }
            else
            {
                return cell.CellValue.InnerText;
            }
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
