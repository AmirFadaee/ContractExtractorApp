namespace ContractExtractorApp
{
    public class ContractInfo
    {
        public string ContractType { get; set; } // تجاری یا اداری
        public string UnitNumber { get; set; }
        public string Floor { get; set; }
        public double Area { get; set; }
        public decimal TotalPrice { get; set; }
        public string InstallmentDescription { get; set; }
        public string PurchaseDate { get; set; }
        public string DeliveryDate { get; set; }
    }
}
