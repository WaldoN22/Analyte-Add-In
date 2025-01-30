namespace BlazorAddIn.Model
{
    public class UnreadCount
    {
        public int? UnCount { get; set; }

        public UnreadCount()
        {

        }

        public UnreadCount(int uncount)
        {
            UnCount = uncount;
        }
    }
}


