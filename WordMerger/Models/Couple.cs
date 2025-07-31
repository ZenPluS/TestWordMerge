using WordMerge.Abstract;

namespace WordMerge.Models
{
    public class Couple<TLeft, TRight>
        : AbstractCouple<TLeft, TRight>
    {
        public Couple(TLeft left, TRight right)
            : base(left, right)
        { }
    }
}
