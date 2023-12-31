using AkelonTestExcel.Extensions;
using AkelonTestExcel.Repository;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Order = AkelonTestExcel.Repository.Order;

namespace AkelonTestExcel
{
    internal class Program
    {
        static string filePath = string.Empty;
        static XLWorkbook book;
        static List<Product> products;
        static List<Order> orders;
        static List<Client> clients;

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.InputEncoding = System.Text.Encoding.UTF8;

            string command = "";
            while (true)
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine("Для работы с программой укажите путь xlsx файла:");
                    filePath = Console.ReadLine();
                }
                try
                {
                    book = new XLWorkbook(filePath);
                }
                catch
                {
                    Console.WriteLine("Не удалось открыть файл. Проверьте путь к файлу/n и убедитесь что он не открыт в другой программе.");
                    filePath = null;
                    continue;
                }

                if (products == null)
                    products = GetProducts(book);
                if (clients == null)
                    clients = GetClients(book);
                if (orders == null)
                    orders = GetOrders(book);

                Console.WriteLine(@"Выберите действие:
                                1) поиск заявок по товару
                                2) изменение контактного лица
                                3) определение золотого клиента
                                4) завершить работу с программой");
                command = Console.ReadLine();

                switch(command)
                {
                    case "1":
                        string text = "";
                        Console.WriteLine("Введите название товара: ");
                        text = Console.ReadLine();
                        var list = orders.Where(x => x.Product.Name.Contains(text??"", StringComparison.OrdinalIgnoreCase)).ToList();
                        if (list.Count == 0)
                            Console.WriteLine("Не найдено заявок по данному товару");
                        else
                        {
                            Console.WriteLine("Компания   Количество  Цена   Дата");
                            list.ForEach(x => Console.WriteLine("{0}     {1}     {2}  {3:d.M.yyyy}\n", x.Client.CompanyName, x.CountProduct, (x.CountProduct * x.Product.Price), x.DateCreated));
                        }
                        break;
                    case "2":
                        Console.WriteLine("Введите имя организации: ");
                        text = Console.ReadLine();
                        var listc = clients.Where(x => x.CompanyName.Contains(text, StringComparison.OrdinalIgnoreCase)).ToList();

                        if (listc.Count > 0)
                        {
                            if (listc.Count > 1)
                            {
                                Console.WriteLine("Найдено несколько организаций: ");
                                Console.WriteLine("Компания     Номер");
                                listc.ForEach(x => Console.WriteLine("{0}      {1}", x.CompanyName, x.Code));
                                Console.WriteLine("Введите номер нужной организации:");
                                var code = Console.ReadLine();

                                if (!listc.Any(x => x.Code.ToString() == code))
                                {
                                    Console.WriteLine("Неверно введен номер организации");
                                }
                                Console.WriteLine("Введите новое имя контактного лица: ");
                                text = Console.ReadLine();

                                var ws = book.Worksheets.ElementAt(1);
                                var range = ws.RangeUsed();
                                var colCount = range.ColumnCount();
                                var rowCount = range.RowCount();

                                var i = 2;

                                while (i < rowCount + 1)
                                {
                                    var value = ws.Cell(i, 1).Value.ToString();

                                    if (value == code)
                                    {
                                        ws.Cell(i, 4).Value = text;
                                        break;
                                    }
                                    else
                                    {
                                        i++;
                                        continue;
                                    }
                                }
                                book.Save();
                                Console.WriteLine("Изменения сохранены");

                            }
                        }
                        else
                            Console.WriteLine("Не найдено такой организации");
                        break;
                    case "3":
                        Console.WriteLine("Введите месяц и год (например 05.2023) :");
                        text = Console.ReadLine();
                        DateTime date;
                        if (DateTime.TryParseExact(text, "MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                        {
                            List<ClientOrders> listt = new List<ClientOrders>();
                            DateTime end = date.AddMonths(1);
                            var groups = orders.Where(x => x.DateCreated >= date && x.DateCreated <= end)
                                               .GroupBy(x => x.Client)
                                               .Select(s => new ClientOrders { Client = s.Key, CountOrders = s.Count(), Month = date.Month, Year = date.Year })
                                               .ToList();

                            if (groups.Count > 0)
                            {
                                Console.WriteLine("Золотым клиентом является :");
                                Console.WriteLine("         Клиент          Кол-во заказов:");

                                var goldClient = groups.MaxBy(x => x.CountOrders);
                                Console.WriteLine("         {0}        {1}", goldClient.Client.CompanyName, goldClient.CountOrders);
                            }
                            else
                                Console.WriteLine("За указанный месяц,год не найдено заявок");
                        }
                        else
                            Console.WriteLine("Похоже вы ввели некорректную дату, потому что нам не удалось её распознать.");
                        break;
                    case "4":
                        Environment.Exit(0);
                        break;
                }
            }

            
        }

        //метод читающий из xlsx файла лист Товары
        static List<Product> GetProducts(XLWorkbook book)
        {
            List<Product> list = new List<Product>();             

                var ws = book.Worksheets.First();
                var range = ws.RangeUsed();
                var colCount = range.ColumnCount();
                var rowCount = range.RowCount();

                var i = 2;
                List<string> unitsName = Product.GetDescriptionsUnits().ToList();

                while (i < rowCount + 1)
                {
                    Product product = new Product();

                    // получим код товара
                    var value = ws.Cell("A"+i).Value;
                    if (!value.IsNumber)
                    {
                        Console.Write("Sheet 1. Product in row {0} have article wrong", i);
                        continue;
                    }
                    else
                    {
                        product.Article = int.Parse(value.ToString());
                    }

                    // получим единицы измерения
                    string units = ws.Cell("C"+i).Value.ToString();
                    if (!unitsName.Contains(units))
                    {
                        Console.Write("Sheet 1. Product in row {0} have units wrong", i);
                        continue;
                    }
                    else
                        product.Units = EnumExtensions.GetValueFromDescription<Units>(units);

                    // получим цену продукта
                    var priceStr = ws.Cell("D"+i).Value.ToString();

                    if (!string.IsNullOrEmpty(priceStr))
                    {
                        float price;

                        if (float.TryParse(priceStr, out price))
                            product.Price = price;
                        else
                        {
                            Console.Write("Sheet 1. Product in row {0} have price wrong", i);
                            continue;
                        }
                    }
                    else
                        product.Price = 0;

                    // получим название товара
                    product.Name = ws.Cell("B"+i).Value.ToString();
                    list.Add(product);
                    i++;
                }
            return list;
        }

        //метод читающий из xlsx файла лист Заявки
        static List<Order> GetOrders(XLWorkbook book)
        {
            List<Order> list = new List<Order>();
            var ws = book.Worksheets.ElementAt(2);
            var range = ws.RangeUsed();
            var colCount = range.ColumnCount();
            var rowCount = range.RowCount();

            var i = 2;

            while (i < rowCount + 1)
            {
                // получим код товара
                var value = ws.Cell("B"+i).Value;
                if (!value.IsNumber)
                {
                    Console.Write("Sheet 3. Order in row :0 have product article wrong", i);
                    continue;
                }
                else
                {
                    int article = int.Parse(value.ToString());
                    Product product = products.SingleOrDefault(x => x.Article == article);

                    if (product != null)
                    {
                        // получим код клиента
                        value = ws.Cell("C"+i).Value;
                        if (!value.IsNumber)
                        {
                            Console.Write("Sheet 3. Order in row :0 have client code wrong", i);
                            continue;
                        }
                        else
                        {
                            int clientCode = int.Parse(value.ToString());
                            Client client = clients.SingleOrDefault(x => x.Code == clientCode);

                            if (client != null)
                            {
                                Order order = new Order(product,client);

                                // получим код заявки
                                value = ws.Cell("A"+i).Value;
                                if (!value.IsNumber)
                                {
                                    Console.Write("Sheet 3. Order in row :0 have order code wrong", i);
                                    continue;
                                }
                                else
                                    order.Code = int.Parse(value.ToString());

                                // получим номер заявки
                                value = ws.Cell("D"+i).Value;
                                if (!value.IsNumber)
                                {
                                    Console.Write("Sheet 3. Order in row :0 have number code wrong", i);
                                    continue;
                                }
                                else
                                    order.Number = int.Parse(value.ToString());

                                // получим количество заказанного товара
                                value = ws.Cell("E"+i).Value;
                                if (!value.IsNumber)
                                {
                                    Console.Write("Sheet 3. Order in row :0 have count product wrong", i);
                                    continue;
                                }
                                else
                                    order.CountProduct = int.Parse(value.ToString());

                                // получим дату заявки
                                value = ws.Cell("F"+i).Value;
                                if (!value.IsDateTime)
                                {
                                    Console.Write("Sheet 3. Order in row :0 have date wrong", i);
                                    continue;
                                }
                                else
                                    order.DateCreated = DateTime.Parse(value.ToString());

                                list.Add(order);
                            }
                            else
                            {
                                Console.Write("Sheet 3. Not found client with code = {0}", clientCode);
                                continue;
                            }
                        }
                    }
                    else
                    {
                        Console.Write("Sheet 3. Not found product with code = {0}", article);
                        continue;
                    }
                }

                i++;
            }
            return list;
        }

        //метод читающий из xlsx файла лист Клиенты
        static List<Client> GetClients(XLWorkbook book)
        {
            List<Client> list = new List<Client>();

                var ws = book.Worksheets.ElementAt(1);
                var range = ws.RangeUsed();
                var colCount = range.ColumnCount();
                var rowCount = range.RowCount();

                var i = 2;

                while (i < rowCount + 1)
                {
                    Client client = new Client();
                    var value = ws.Cell("A"+i).Value;
                    if (!value.IsNumber)
                    {
                        Console.Write("Sheet 2. Client in row :0 have code wrong", i);
                        continue;
                    }
                    else
                    {
                        client.Code = int.Parse(value.ToString());
                    }

                    client.CompanyName = ws.Cell("B"+i).Value.ToString();
                    client.Address = ws.Cell("C"+i).Value.ToString();
                    client.ContactManager = ws.Cell("D"+i).Value.ToString();

                    list.Add(client);
                    i++;
                }
            return list;
        }
    }
}