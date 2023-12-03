namespace ExcelSolution
{
    internal class Application
    {
        public int Code { get; set; }
        public int ProductCode { get; set; }
        public int ClientCode { get; set; }
        public int ApplicationNumber { get; set; }
        public int RequiredQuantity { get; set; }
        public DateTime PostingDate { get; set; }
        public Application(string code, string productCode, string clientCode, string applicationNum, string quantity, string postingDate)
        {
            Code = int.Parse(code);
            ProductCode = int.Parse(productCode);
            ClientCode = int.Parse(clientCode);
            ApplicationNumber = int.Parse(applicationNum);
            RequiredQuantity = int.Parse(quantity);
            PostingDate = DateTime.Parse(postingDate);
        }
    }
}
