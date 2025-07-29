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
        public static bool operator ==(AbstractCouple<TLeft, TRight> left, AbstractCouple<TLeft, TRight> right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (left is null || right is null) return false;
            return left.Equals(right);
        }

        public static bool operator !=(AbstractCouple<TLeft, TRight> left, AbstractCouple<TLeft, TRight> right)
        {
            return !(left == right);
        }

        public static bool operator <(AbstractCouple<TLeft, TRight> left, AbstractCouple<TLeft, TRight> right)
        {
            if (left is null) return right != null;
            return left.CompareTo(right) < 0;
        }

        public static bool operator >(AbstractCouple<TLeft, TRight> left, AbstractCouple<TLeft, TRight> right)
        {
            if (left is null) return false;
            if (right is null) return true;
            return left.CompareTo(right) > 0;
        }

        public static bool operator <=(AbstractCouple<TLeft, TRight> left, AbstractCouple<TLeft, TRight> right)
        {
            return left == right || left < right;
        }

        public static bool operator >=(AbstractCouple<TLeft, TRight> left, AbstractCouple<TLeft, TRight> right)
        {
            return left == right || left > right;
        }

        public static implicit operator Tuple<TLeft, TRight>(AbstractCouple<TLeft, TRight> couple)
        {
            return Tuple.Create(couple.Left, couple.Right);
        }

        public static explicit operator AbstractCouple<TLeft, TRight>(Tuple<TLeft, TRight> tuple)
        {
            if (tuple == null) throw new ArgumentNullException(nameof(tuple));
            return (AbstractCouple<TLeft, TRight>)Activator.CreateInstance(typeof(AbstractCouple<TLeft, TRight>),
                tuple.Item1, tuple.Item2);
        }

        public static implicit operator KeyValuePair<TLeft, TRight>(AbstractCouple<TLeft, TRight> couple)
        {
            return new KeyValuePair<TLeft, TRight>(couple.Left, couple.Right);
        }

        public static explicit operator AbstractCouple<TLeft, TRight>(KeyValuePair<TLeft, TRight> pair)
        {
            return (AbstractCouple<TLeft, TRight>)Activator.CreateInstance(typeof(AbstractCouple<TLeft, TRight>),
                pair.Key, pair.Value);
        }
    }
}
