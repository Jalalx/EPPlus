using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace EPPlus.Utils
{
    internal static class ReflectionUtility
    {
        public static IEnumerable<MemberWithMetadata> GetMembersWithMetadata(MemberInfo[] members)
        {
            var result = new List<MemberWithMetadata>();
            foreach (var memberInfo in members)
            {
                var ignored = memberInfo.GetCustomAttributes(typeof(IgnoreColumnAttribute), false) == null;
                if (ignored)
                {
                    continue;
                }

                var descriptionAttr = memberInfo.GetCustomAttributes(typeof(DescriptionAttribute), false)
                    .FirstOrDefault() as DescriptionAttribute;

                var displayNameAttr = memberInfo.GetCustomAttributes(typeof(DisplayNameAttribute), false)
                    .FirstOrDefault() as DisplayNameAttribute;

                var converterAttr = memberInfo.GetCustomAttributes(typeof(ConvertCellValueAttribute), false)
                    .FirstOrDefault() as ConvertCellValueAttribute;

                ICellValueConverter converter = null;
                var converterType = converterAttr != null ? converterAttr.ConverterType : default;
                if (converterType != null)
                {
                    converter = Activator.CreateInstance(converterType, converterAttr.ConstructorArgs) as ICellValueConverter;
                }

                result.Add(new MemberWithMetadata(memberInfo, converter, descriptionAttr, displayNameAttr));
            }

            return result;
        }

        public static IEnumerable<MemberWithMetadata> GetMembersWithMetadata(Type type, BindingFlags flags)
        {
            var members = type.GetMembers(flags);
            return GetMembersWithMetadata(members);
        }
    }

    internal class MemberWithMetadata
    {
        public MemberWithMetadata(MemberInfo memberInfo, ICellValueConverter converter,
            DescriptionAttribute descriptionAttribute, DisplayNameAttribute displayNameAttribute)
        {
            MemberInfo = memberInfo;
            Converter = converter;
            DescriptionAttribute = descriptionAttribute;
            DisplayNameAttribute = displayNameAttribute;
            ReturnType = TryGetReturnType();
        }

        public DescriptionAttribute DescriptionAttribute { get; }
        public DisplayNameAttribute DisplayNameAttribute { get; }
        public Type ReturnType { get; }
        public ICellValueConverter Converter { get; }
        public MemberInfo MemberInfo { get; }

        private Type TryGetReturnType()
        {
            if (MemberInfo is PropertyInfo prop)
            {
                return prop.PropertyType;
            }
            else if (MemberInfo is FieldInfo field)
            {
                return field.FieldType;
            }
            else if (MemberInfo is MethodInfo method)
            {
                return method.ReturnType;
            }

            return null;
        }

        public object GetValue(object item)
        {
            object rawValue;
            if (MemberInfo is PropertyInfo prop)
            {
                rawValue = prop.GetValue(item, null);
            }
            else if (MemberInfo is FieldInfo field)
            {
                rawValue = field.GetValue(item);
            }
            else if (MemberInfo is MethodInfo method)
            {
                rawValue = method.Invoke(item, null);
            }
            else
            {
                return item;
            }

            if (Converter != null)
            {
                return Converter.Convert(rawValue);
            }
            else if (ReturnType != null && ReturnType.IsEnum)
            {
                return EnumHelper.GetDisplayValue(ReturnType, rawValue);
            }
            else
            {
                return rawValue;
            }
        }
    }
}
