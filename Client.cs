namespace ExcelSolution
{
    internal class Client
    {
        public int Code { get; set; }
        public string CompanyName { get; set; }
        public string Address { get; set;}
        public string ContactPerson { get; set;}
        public Client(string code, string companyName, string address, string contactPerson)
        {
            Code = int.Parse(code);
            CompanyName = companyName;
            Address = address;
            ContactPerson = contactPerson;
        }
    }
}
