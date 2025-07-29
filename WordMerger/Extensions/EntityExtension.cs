using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace TestWordMerge.Extensions
{
    internal static class EntityExtension
    {
        public static Entity Clone(this Entity source)
        {
            var entity = new Entity(source.LogicalName);
            entity.Id = source.Id;
            foreach (var attribute in source.Attributes)
            {
                entity[attribute.Key] = attribute.Value;
            }

            return entity;
        }

        public static Entity CloneEmpty(this Entity source)
        {
            var entity = new Entity(source.LogicalName);
            entity.Id = source.Id;

            return entity;
        }
    }
}
