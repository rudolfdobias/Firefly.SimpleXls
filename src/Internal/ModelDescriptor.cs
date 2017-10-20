using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Firefly.SimpleXls.Attributes;
using Firefly.SimpleXls.Converters;
using Firefly.SimpleXls.Exceptions;

namespace Firefly.SimpleXls.Internal
{
    /// <summary>
    /// Describes model members for conversion
    /// </summary>
    internal static class ModelDescriptor
    {
        /// <summary>
        /// static reflection cache
        /// </summary>
        private static Dictionary<string, Dictionary<string, ColumnDescriptor>> ReflectionCache { get; } =
            new Dictionary<string, Dictionary<string, ColumnDescriptor>>();

        /// <summary>
        /// Describes model members for conversion
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Dictionary<string, ColumnDescriptor> DescribeModel<T>()
            => DescribeModel(typeof(T));

        /// <summary>
        /// Describes model members for conversion
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public static Dictionary<string, ColumnDescriptor> DescribeModel(Type model)
        {
            if (ReflectionCache.ContainsKey(model.AssemblyQualifiedName))
            {
                return ReflectionCache[model.AssemblyQualifiedName];
            }

            var fields = model.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead).ToList();
            var columns = new Dictionary<string, ColumnDescriptor>();
            foreach (var f in fields)
            {
                var descriptor = new ColumnDescriptor
                {
                    Key = f.Name,
                    Property = f,
                    Attributes = GetAttributes(f)
                };
                if (XlsConverters.Converters.ContainsKey(f.PropertyType))
                {
                    descriptor.CustomValueConverter = XlsConverters.Converters[f.PropertyType];
                }
                columns.Add(f.Name, descriptor);
            }

            ReflectionCache[model.AssemblyQualifiedName] = columns;

            return columns;
        }

        /// <summary>
        /// Resolves data attributes
        /// </summary>
        /// <param name="member"></param>
        /// <returns></returns>
        private static ColumnAttributeInfo GetAttributes(PropertyInfo member)
        {
            var info = new ColumnAttributeInfo
            {
                Heading = member.Name
            };

            var attrs = member.GetCustomAttributes().ToList();
            foreach (var a in attrs)
            {
                switch (a)
                {
                    case XlsIgnoreAttribute _:
                        info.Ignore = true;
                        continue;
                    case XlsHeaderAttribute _:
                        info.Heading = (a as XlsHeaderAttribute)?.Name;
                        break;
                    case XlsTranslateAttribute _:
                        if (typeof(string).IsAssignableFrom(member.PropertyType) == false)
                        {
                            throw new SimpleXlsException("Property " + member.Name +
                                                         " cannot be translated since it's not convertilbe to string.");
                        }
                        info.TranslateValue = true;
                        info.DictionaryPrefix = (a as XlsTranslateAttribute)?.DictPrefix;
                        break;
                }
            }

            return info;
        }
    }
}