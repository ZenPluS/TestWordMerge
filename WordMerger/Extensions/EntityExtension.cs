using Microsoft.Xrm.Sdk;

namespace WordMerge.Extensions
{
    internal static class EntityExtension
    {
        internal static Entity CloneEmpty(this Entity source)
        {
            var entity = new Entity(source.LogicalName);
            entity.Id = source.Id;

            return entity;
        }
    }
}
