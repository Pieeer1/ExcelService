using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;

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
        public static IEnumerable<T> WhereNotNull<T>(this IEnumerable<T?> o) => o.Where(x => x != null)!;


        public static string? ResolveArgs<T>(this Expression<Func<T, bool>> expression)
        {
            var body = expression.Body as BinaryExpression;
            if (body != null) 
            {
                var left = body.Left as MemberExpression;
                if (left != null) 
                {
                    return left.Member.Name;
                }
            }
            return null;
        }
        public static bool IsIEnumerableOfT(this Type type)
        {
            return type.GetInterfaces().Any(x => x.IsGenericType
                   && x.GetGenericTypeDefinition() == typeof(IEnumerable<>));
        }
    }
}
