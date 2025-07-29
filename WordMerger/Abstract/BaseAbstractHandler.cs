using System;

namespace WordMerge.Abstract
{
    public abstract class BaseAbstractHandler<T>
    {
        protected readonly Action<T> Logger;

        protected BaseAbstractHandler(Action<T> logger)
        {
            Logger = logger ?? (_ => { });
        }
    }
}
