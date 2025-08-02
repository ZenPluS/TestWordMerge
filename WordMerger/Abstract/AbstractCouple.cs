using System;
using System.Dynamic;
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
    }
}