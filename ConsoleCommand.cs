using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace ExcelSolution
{
    static class ConsoleCommand
    {
        static string? _filePath;
        static List<Product> products = new List<Product>();
        static List<Client> clients = new List<Client>();
        static List<Application> applications = new List<Application>();
        public static void Help()
        {
            Console.WriteLine("Список команд:\n");
            Console.WriteLine("\t<show> - вывести информацию по файлу\n");
            Console.WriteLine("\t<find> - найти информацию по товару\n");
            Console.WriteLine("\t<vip> - показать золотого клиента\n");
            Console.WriteLine("\t<change> - изменить контактные данные\n");
            Console.WriteLine("\t<clear> - очистить консоль\n");
            Console.WriteLine("\t<exit> - выйти из программы");
        }

        public static void FindProduct()
        {
            Console.Write("Введите наименование товара: ");
            var productName = Console.ReadLine().Trim().ToLower();
            var product = products.Where(x => x.Name.ToLower() == productName).FirstOrDefault();
            if (product is null)
            {
                Console.WriteLine($"Товар с названием {productName} - не найден");
                return;
            }
            var productCode = product?.Code;
            var productPrice = product?.PricePerUnit;
            var foundApplications = applications.Where(x => x.ProductCode == productCode);
            Console.WriteLine($"Для товара {productName} найдено заявок от клиентов: {foundApplications.Count()}");
            foreach (var application in foundApplications)
            {
                var foundClient = clients.Where(x => x.Code == application.ClientCode).FirstOrDefault();
                if (foundClient != null)
                {
                    Console.WriteLine($"Организация: {foundClient.CompanyName}, контактное лицо: {foundClient.ContactPerson}");
                    Console.WriteLine($"\t{application.PostingDate.ToLongDateString()} | Количество товара: {application.RequiredQuantity} " +
                        $"({product?.Unit}), на сумму {application.RequiredQuantity * productPrice} рублей");
                }
            }
        }

        public static void ShowGoldenClient()
        {
            Console.Write("Введите год: ");
            var inputYear = Console.ReadLine().Trim().ToLower();
            Console.Write("Введите месяц (числом): ");
            var inputMonth = Console.ReadLine().Trim().ToLower();
            Regex monthRgx = new Regex(@"^[1-9]$|^[1][0-2]$"); // 1 - 12
            Regex yearRgx = new Regex(@"^[12][0-9]{3}$"); // 1000 - 2999
            if (yearRgx.IsMatch(inputYear) && monthRgx.IsMatch(inputMonth))
            {
                var applicationsForMonth = applications.Where(x => x.PostingDate.Year == int.Parse(inputYear) &&
                                               x.PostingDate.Month == int.Parse(inputMonth)).ToArray(); //находим все заявки за указанный период
                var goldenClientCode = applicationsForMonth.GroupBy(x => x.ClientCode).OrderByDescending(x => x.Count()).FirstOrDefault()?
                                   .Select(x => x.ClientCode).FirstOrDefault(); //находим код клиента с наибольшим количеством заявок
                var client = clients.Where(x => x.Code == goldenClientCode).FirstOrDefault();
                if (client != null) Console.WriteLine($"Золотой клиент за {inputMonth}.{inputYear} - {client.CompanyName} ({client.ContactPerson})");
                else Console.WriteLine("В базе отсутствует клиент за указанный пероид");
            }
            else Console.WriteLine("Некорректный формат, необходимо указать год и ввести число от 1-12 для месяца");
        }

        public static void ChangeContactPerson()
        {
            Console.Write("Введите название организации: ");
            var companyName = Console.ReadLine().Trim().ToLower();
            if (!clients.Any(x => x.CompanyName.ToLower().Trim() == companyName))
            {
                Console.WriteLine("По заданному названию не найдена организация");
                return;
            }
            Console.Write("Укажите ФИО нового контактного лица: ");
            var contactPersonName = Console.ReadLine()?.Trim();
            //считываем таблицу Клиенты
            var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(2);
            var rows = worksheet.RowsUsed();
            foreach (var row in rows)
            {
                if (row.Cell(2).Value.ToString().Trim().ToLower() == companyName) //находим строчки клиента и меняем информацию
                    row.Cell(4).Value = contactPersonName;
            }
            workbook.Save();
            ReadFileData(_filePath); //обновляем информацию
            foreach (var client in clients) Console.WriteLine($"{client.CompanyName}\t|\t{client.ContactPerson}"); //выводим список клиентов
        }

        public static void ShowData()
        {
            Console.WriteLine("Товары:");
            foreach (var product in products) Console.WriteLine($"{product.Name}\t\t|\t{product.PricePerUnit} руб."); //выводим список товаров
            Console.WriteLine("\nКлиенты:");
            foreach (var client in clients) Console.WriteLine($"{client.CompanyName}\t|\t{client.ContactPerson}"); //выводим список клиентов
            Console.WriteLine("\nЗаявки:");
            foreach (var app in applications) Console.WriteLine($"{app.PostingDate.ToLongDateString()}\t\t|\tКод: {app.Code}"); //выводим список заявок
        }

        public static void Clear()
        {
            Console.Clear();
        }

        public static void Exit()
        {
            Environment.Exit(0);
        }

        public static bool SetFilePath()
        {
            Console.Write("Введите путь до Excel файла: ");
            string? path = Console.ReadLine();
            if (!File.Exists(path))
            {
                Console.WriteLine("Файл не найден!");
                return true;
            }
            _filePath = path;
            try { var workbook = new XLWorkbook(path); }
            catch
            {
                Console.WriteLine("Закройте файл перед использованием программы!");
                return true;
            }
            ReadFileData(path);
            return false;
        }

        public static void ReadFileData(string path)
        {
            products.Clear();
            clients.Clear();
            applications.Clear();
            var workbook = new XLWorkbook(path);
            //считываем таблицу Товары
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed();
            foreach (var row in rows)
            {
                if (row == rows.ElementAt(0)) continue; //пропускаем строку заголовков
                var product = new Product(row.Cell(1).Value.ToString(), row.Cell(2).Value.ToString(), 
                                          row.Cell(3).Value.ToString(), row.Cell(4).Value.ToString());
                products.Add(product);
            }
            //считываем таблицу Клиенты
            worksheet = workbook.Worksheet(2);
            rows = worksheet.RowsUsed();
            foreach (var row in rows)
            {
                if (row == rows.ElementAt(0)) continue; //пропускаем строку заголовков
                var client = new Client(row.Cell(1).Value.ToString(), row.Cell(2).Value.ToString(),
                                          row.Cell(3).Value.ToString(), row.Cell(4).Value.ToString());
                clients.Add(client);
            }
            //считываем таблицу Заявки
            worksheet = workbook.Worksheet(3);
            rows = worksheet.RowsUsed();
            foreach (var row in rows)
            {
                if (row == rows.ElementAt(0)) continue; //пропускаем строку заголовков
                var application = new Application(row.Cell(1).Value.ToString(), row.Cell(2).Value.ToString(), row.Cell(3).Value.ToString(), 
                                              row.Cell(4).Value.ToString(), row.Cell(5).Value.ToString(), row.Cell(6).Value.ToString());
                applications.Add(application);
            }
        }
    }
}
