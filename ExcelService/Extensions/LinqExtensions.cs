﻿namespace ExcelService.Extensions
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
        public static IEnumerable<T> WhereNotNull<T>(this IEnumerable<T?> o) => o.Where(x => x != null)!;
    }
}
