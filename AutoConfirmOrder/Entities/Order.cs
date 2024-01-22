namespace AutoConfirmOrder.Entities
{
    public class Order
    {
        public string Price { get; set; } = null!;
        public string Game { get; set; } = null!;
        public string Type { get; set; } = null!;
        public string Region { get; set; } = null!;
        public string QueueType { get; set; } = null!;
        public string Rank { get; set; } = null!;
        public DateTime AcceptAt { get; set; }
    }
}
