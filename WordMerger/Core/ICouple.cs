using System;
using System.Collections;
using System.Runtime.CompilerServices;

namespace WordMerge.Core
{
    public interface ICouple<TLeft, TRight> :
        IComparable,
        IComparable<ICouple<TLeft, TRight>>,
        IEquatable<ICouple<TLeft, TRight>>,
        ITuple,
        IStructuralEquatable,
        IStructuralComparable
    {
        TLeft Left { get; }
        TRight Right { get; }
    }
}
