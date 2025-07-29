using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestWordMerge.Core;

namespace TestWordMerge.Abstract
{
    public abstract partial class AbstractCouple<TLeft, TRight>
    {
        public override string ToString() => $"{Left} - {Right}";

        public override bool Equals(object obj)
        {
            if (obj is ICouple<TLeft, TRight> other)
            {
                return Equals(other);
            }

            return false;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (EqualityComparer<TLeft>.Default.GetHashCode(Left) * 397) ^
                       EqualityComparer<TRight>.Default.GetHashCode(Right);
            }
        }
    }
}
