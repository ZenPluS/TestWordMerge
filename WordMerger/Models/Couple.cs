using System;
using System.Collections.Generic;
using TestWordMerge.Abstract;
using TestWordMerge.Core;

namespace TestWordMerge.Models
{
    public class Couple<TLeft, TRight>
        : AbstractCouple<TLeft, TRight>
    {
        public Couple(TLeft left, TRight right)
            : base(left, right)
        { }
    }
}
