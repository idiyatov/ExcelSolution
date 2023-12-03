namespace ExcelSolution
{
    internal class Product
    {
        public int Code { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public int PricePerUnit { get; set; }

        public Product(string code, string name, string unit, string pricePerUnit)
        {
            Code = int.Parse(code);
            Name = name;
            Unit = unit;
            PricePerUnit = int.Parse(pricePerUnit);
        }
    }
}
