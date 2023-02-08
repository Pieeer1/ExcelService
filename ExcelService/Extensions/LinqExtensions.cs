namespace ExcelService.Extensions
{
    public static class LinqExtensions
    {
        public static int IndexOf<T>(this IEnumerable<T> source, T value)
        {
            int index = 0;
            var comparer = EqualityComparer<T>.Default; // or pass in as a parameter
            foreach (T item in source)
            {
                if (comparer.Equals(item, value)) return index;
                index++;
            }
            return -1;
        }
        public static IEnumerable<T> ToNonNullableInside<T>(this IEnumerable<T?> obj)
        {
            List<T> result = new List<T>();

            foreach (var o in obj)
            {
                if (o is not null)
                {
                    result.Add(o);
                }
            }
            return result.AsEnumerable<T>();
        }
    }
}
