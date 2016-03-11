using System.Collections.Generic;
using System.Linq;

namespace Depends
{
    public static class LINQExtension 
    {
        public static IEnumerable<T> Evaluate<T>(this IEnumerable<T> source)
        {
            foreach (var _ in source) ;
            return source;
        }
    }
}
