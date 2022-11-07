namespace WBParserAPI
{
    public class ProductCard
    {
        public Data data { get; set; }
    }


    public class Data
    {
        public Product[] products { get; set; }
    }

    public class Product
    {
        public int __sort { get; set; }
        public int ksort { get; set; }
        public int ksale { get; set; }
        public int time1 { get; set; }
        public int time2 { get; set; }
        public int dist { get; set; }
        public long id { get; set; }
        public int root { get; set; }
        public int kindId { get; set; }
        public int subjectId { get; set; }
        public int subjectParentId { get; set; }
        public string name { get; set; }
        public string brand { get; set; }
        public int brandId { get; set; }
        public int siteBrandId { get; set; }
        public int sale { get; set; }
        public int priceU { get; set; }
        public int salePriceU { get; set; }
        public int averagePrice { get; set; }
        public int benefit { get; set; }
        public int pics { get; set; }
        public int rating { get; set; }
        public int feedbacks { get; set; }
        public bool diffPrice { get; set; }
        public int panelPromoId { get; set; }
        public string promoTextCat { get; set; }
        public bool isNew { get; set; }
    }

}