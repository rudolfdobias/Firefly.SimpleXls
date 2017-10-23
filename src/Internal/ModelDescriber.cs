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
    internal static class ModelDescriber
    {
        /// <summary>
        /// static reflection cache
        /// </summary>
        private static Dictionary<string, SheetDescriptor> ReflectionCache { get; } =
            new Dictionary<string, SheetDescriptor>();

        /// <summary>
        /// Describes model members for conversion
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static SheetDescriptor DescribeModel<T>()
            => DescribeModel(typeof(T));

        /// <summary>
        /// Describes model members for conversion
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public static SheetDescriptor DescribeModel(Type model)
        {
            if (ReflectionCache.ContainsKey(model.AssemblyQualifiedName))
            {
                return ReflectionCache[model.AssemblyQualifiedName];
            }

            var sheet = CreateSheetDescriptor(model);
            var fields = model.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead).ToList();

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
                sheet.Columns.Add(descriptor);
            }

            ReflectionCache[model.AssemblyQualifiedName] = sheet;

            return sheet;
        }

        /// <summary>
        /// Creates sheet descriptor based on supplied model and type attributes
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private static SheetDescriptor CreateSheetDescriptor(Type model)
        {
            var attr = model.GetTypeInfo().GetCustomAttribute<XlsSheetAttribute>();

            var sheet = new SheetDescriptor
            {
                ModelType = model,
                DictionaryPrefix = attr?.DictionaryPrefix ?? "",
                Name = attr?.Name ?? model.Name
            };

            return sheet;
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
                        info.DictionaryPrefix = (a as XlsTranslateAttribute)?.Prefix;
                        break;
                }
            }

            return info;
        }
    }
}