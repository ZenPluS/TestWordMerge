using System;
using System.Collections;
using System.Collections.Generic;
using WordMerge.Core;

namespace WordMerge.Abstract
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

        public int Length => 2;

        public object this[int index]
        {
            get
            {
                switch (index)
                {
                    case 0: return Left;
                    case 1: return Right;
                    default: throw new IndexOutOfRangeException();
                }
            }
        }

        public bool Equals(object other, IEqualityComparer comparer)
        {
            if (ReferenceEquals(this, other))
                return true;

            if (!(other is ICouple<TLeft, TRight> couple))
                return false;

            var leftEquals = comparer.Equals(Left, couple.Left);
            var rightEquals = comparer.Equals(Right, couple.Right);
            return leftEquals && rightEquals;
        }

        public int GetHashCode(IEqualityComparer comparer)
        {
            if (comparer == null)
                throw new ArgumentNullException(nameof(comparer));

            var hashLeft = Left != null ? comparer.GetHashCode(Left) : 0;
            var hashRight = Right != null ? comparer.GetHashCode(Right) : 0;
            unchecked
            {
                return (hashLeft * 397) ^ hashRight;
            }
        }

        public int CompareTo(object other, IComparer comparer)
        {
            if (ReferenceEquals(this, other))
                return 0;

            if (!(other is ICouple<TLeft, TRight> couple))
                throw new ArgumentException("Object is not a compatible ICouple", nameof(other));

            var leftCompare = comparer.Compare(Left, couple.Left);
            return leftCompare != 0
                ? leftCompare
                : comparer.Compare(Right, couple.Right);
        }
    }
}
