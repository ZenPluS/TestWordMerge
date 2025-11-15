using Microsoft.Xrm.Sdk;

namespace WordMerge.Extensions
{
    internal static class EntityExtension
    {
        internal static Entity CloneEmpty(this Entity source)
            => new Entity(source.LogicalName)
            {
                Id = source.Id
            };
    }
}
