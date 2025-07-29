using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestWordMerge.Core;
using TestWordMerge.Models;

namespace TestWordMerge.Extensions
{
    public static class CoupleExtension
    {
        public static bool IsEqualTo<TLeft, TRight>(this ICouple<TLeft, TRight> couple, ICouple<TLeft, TRight> other)
        {
            if (couple == null) throw new ArgumentNullException(nameof(couple));
            if (other == null) throw new ArgumentNullException(nameof(other));
            return couple.Equals(other);
        }

        public static IEnumerable<TLeft> LeftItems<TLeft, TRight>(this List<ICouple<TLeft, TRight>> source)
        {
            return source.Select(l => l.Left);
        }

        public static IEnumerable<TRight> RightItems<TLeft, TRight>(this List<ICouple<TLeft, TRight>> source)
        {
            return source.Select(r => r.Right);
        }

    }
}
