namespace ExcelSolution
{
    internal class Program
    {
        static void Main(string[] args)
        {
            while (ConsoleCommand.SetFilePath()) { } //ввод пути до файла
            string? command = "";
            Console.WriteLine("Чтобы посмотреть список команд введите help или ?");
            while (true)
            {
                Console.Write(">> ");
                command = Console.ReadLine().Trim().ToLower();
                if (command == "help" || command == "?") ConsoleCommand.Help(); //вывод справки
                else if (command == "show") ConsoleCommand.ShowData();
                else if (command == "find") ConsoleCommand.FindProduct();
                else if (command == "vip") ConsoleCommand.ShowGoldenClient();
                else if (command == "change") ConsoleCommand.ChangeContactPerson();
                else if (command == "clear") ConsoleCommand.Clear(); //очистить консоль
                else if (command == "exit") ConsoleCommand.Exit(); //выход
                else if (string.IsNullOrWhiteSpace(command)) continue;
                else Console.WriteLine("Неизвестная команда, введите help или ? для получения справки");
            }
        }
    }
}