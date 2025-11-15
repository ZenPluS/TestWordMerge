using System;
using System.Collections.Generic;
using System.Linq;
using WordMerge.Core;

namespace WordMerge.Extensions
{
    public static class CoupleExtension
    {
        public static bool IsEqualTo<TLeft, TRight>(this ICouple<TLeft, TRight> couple, ICouple<TLeft, TRight> other)
        {
            if (couple == null)
                throw new ArgumentNullException(nameof(couple));

            return other == null
                ? throw new ArgumentNullException(nameof(other))
                : couple.Equals(other);
        }

        public static IEnumerable<TLeft> LeftItems<TLeft, TRight>(this IEnumerable<ICouple<TLeft, TRight>> source)
            => source?.Select(c => c.Left) ?? Enumerable.Empty<TLeft>();

        public static IEnumerable<TRight> RightItems<TLeft, TRight>(this IEnumerable<ICouple<TLeft, TRight>> source)
            => source?.Select(c => c.Right) ?? Enumerable.Empty<TRight>();
    }
}