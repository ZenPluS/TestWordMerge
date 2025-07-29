using System;
using System.Collections.Generic;
using WordMerge.Core;

namespace WordMerge.Abstract
{
    public abstract partial class AbstractCouple<TLeft, TRight>
        : ICouple<TLeft, TRight>
    {
        public TLeft Left { get; }
        public TRight Right { get; }

        protected AbstractCouple(TLeft left, TRight right)
        {
            Left = left;
            Right = right;
        }

        protected bool Equals(AbstractCouple<TLeft, TRight> other)
        {
            return EqualityComparer<TLeft>.Default.Equals(Left, other.Left) &&
                   EqualityComparer<TRight>.Default.Equals(Right, other.Right);
        }

        public int CompareTo(object obj)
        {
            if (ReferenceEquals(this, obj)) return 0;
            if (obj is ICouple<TLeft, TRight> other)
                return CompareTo(other);

            throw new ArgumentException("Object is not a compatible ICouple", nameof(obj));
        }

        public int CompareTo(ICouple<TLeft, TRight> other)
        {
            if (ReferenceEquals(this, other)) return 0;
            if (other == null) return 1;

            var leftCompare = Comparer<TLeft>.Default.Compare(Left, other.Left);
            return leftCompare != 0 ? leftCompare : Comparer<TRight>.Default.Compare(Right, other.Right);
        }

        public bool Equals(ICouple<TLeft, TRight> other)
        {
            if (ReferenceEquals(this, other)) return true;
            if (other == null) return false;

            return EqualityComparer<TLeft>.Default.Equals(Left, other.Left)
                   && EqualityComparer<TRight>.Default.Equals(Right, other.Right);
        }
    }
}