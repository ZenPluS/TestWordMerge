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

        public void Deconstruct(out TLeft left, out TRight right)
        {
            left = Left;
            right = Right;
        }
    }
}