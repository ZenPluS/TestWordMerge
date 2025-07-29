using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordMerge.Core
{
    public interface ICouple<TLeft, TRight> :
        IComparable,
        IComparable<ICouple<TLeft, TRight>>,
        IEquatable<ICouple<TLeft, TRight>>
    {
        TLeft Left { get; }
        TRight Right { get; }
    }
}
