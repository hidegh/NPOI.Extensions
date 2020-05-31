using System;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace NPOI.Extensions.Utils
{
    /// <summary>
    /// Original implementation from TypeUtils.cs (ILSpy);
    /// Some methods were changed to extension methods;
    /// Some additional functions were introduced;
    /// </summary>
    public static class TypeUtilsExtensions
    {
        /*
         * TypeUtils.cs (ILSpy)
         */

        /// <summary>
        /// Returns True if sourceType is Nullable&lt;&gt;.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsNullableType(this Type sourceType)
        {
            return sourceType.IsGenericType && sourceType.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        /// <summary>
        /// If sourceType is nullable, returns the underlying (non-nullable) type.
        /// If sourceType is not nullable, the original type is returned.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static Type GetNonNullableType(this Type sourceType)
        {
            if (sourceType.IsNullableType())
            {
                return sourceType.GetGenericArguments()[0];
            }

            return sourceType;
        }

        /// <summary>
        /// If sourceType is non-nullable a type of Nullable&lt;source&gt; is returned,
        /// If sourceType is nullable, the source type itself is returned,
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static Type GetNullableType(this Type sourceType)
        {
            if (sourceType.IsValueType && !sourceType.IsNullableType())
            {
                return typeof(Nullable<>).MakeGenericType(new Type[]
                {
                    sourceType
                });
            }
            return sourceType;
        }

        /// <summary>
        /// Returns true for nullable and non-nullable booleans.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsBool(this Type sourceType)
        {
            return sourceType.GetNonNullableType() == typeof(bool);
        }

        /// <summary>
        /// Returns true for nullable and non-nullable floating point numbers.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsFloatingPoint(this Type sourceType)
        {
            sourceType = sourceType.GetNonNullableType();
            var typeCode = Type.GetTypeCode(sourceType);
            return typeCode == TypeCode.Single || typeCode == TypeCode.Double || typeCode == TypeCode.Decimal;
        }

        /// <summary>
        /// Returns true for nullable and non-nullable date time.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsDateTime(this Type sourceType)
        {
            return sourceType.GetNonNullableType() == typeof(DateTime);
        }

        /// <summary>
        /// Returns true for nullable and non-nullable time spans.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsTimeSpan(this Type sourceType)
        {
            return sourceType.GetNonNullableType() == typeof(TimeSpan);
        }

        /// <summary>
        /// Returns true for all nullable and non-nullable enum types.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsEnum(this Type sourceType)
        {
            sourceType = sourceType.GetNonNullableType();
            return sourceType.IsEnum;
        }

        /// <summary>
        /// Returns true for all nullable and non-nullable integer types.
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsInteger(this Type sourceType)
        {
            sourceType = sourceType.GetNonNullableType();
            if (sourceType.IsEnum)
            {
                return false;
            }
            switch (Type.GetTypeCode(sourceType))
            {
                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Returns true for any nullable or non nullable numeric type (integer or floating point)...
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsNumeric(this Type sourceType)
        {
            sourceType = sourceType.GetNonNullableType();
            if (!sourceType.IsEnum)
            {
                switch (Type.GetTypeCode(sourceType))
                {
                    case TypeCode.Char:
                    case TypeCode.SByte:
                    case TypeCode.Byte:
                    case TypeCode.Int16:
                    case TypeCode.UInt16:
                    case TypeCode.Int32:
                    case TypeCode.UInt32:
                    case TypeCode.Int64:
                    case TypeCode.UInt64:
                    case TypeCode.Single:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Returns true for nullable and non-nullable unsigned (integer like) values (even for enums)...
        /// </summary>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static bool IsUnsigned(this Type sourceType)
        {
            sourceType = sourceType.GetNonNullableType();
            switch (Type.GetTypeCode(sourceType))
            {
                case TypeCode.Char:
                case TypeCode.Byte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return true;
            }
            return false;
        }

        public static bool AreEquivalent(Type type, Type otherType)
        {
            return type == otherType || type.IsEquivalentTo(otherType);
        }

        public static bool AreReferenceAssignable(Type dest, Type src)
        {
            // WARNING: This actually implements "Is this identity assignable and/or reference assignable?"
            return AreEquivalent(src, dest) || (!dest.IsValueType && !src.IsValueType && dest.IsAssignableFrom(src));
        }

        /// <summary>
        /// Inside the source type interfaces and base types a concrete type implementing the given generic type will be returned - or null if no such type exists.
        /// </summary>
        /// <param name="sourceType">Source type, where we try to find a conrete type (implementation) for the given generic type.</param>
        /// <param name="genericType">Generic type to search for - f.e. IMyInterface&lt;&gt;</param>
        /// <returns></returns>
        public static Type FindGenericType(this Type sourceType, Type genericType)
        {
            while (sourceType != null && sourceType != typeof(object))
            {
                if (sourceType.IsGenericType && AreEquivalent(sourceType.GetGenericTypeDefinition(), genericType))
                {
                    return sourceType;
                }

                if (genericType.IsInterface)
                {
                    var interfaces = sourceType.GetInterfaces();

                    for (int i = 0; i < interfaces.Length; i++)
                    {
                        var interfaceType = interfaces[i];
                        var genericTypeToInterface = interfaceType.FindGenericType(genericType);

                        if (genericTypeToInterface != null)
                        {
                            return genericTypeToInterface;
                        }
                    }
                }

                sourceType = sourceType.BaseType;
            }

            return null;
        }

        public static bool IsAnonymous(this Type type)
        {
            if (type.IsGenericType)
            {
                var d = type.GetGenericTypeDefinition();
                if (d.IsClass && d.IsSealed && d.Attributes.HasFlag(TypeAttributes.NotPublic))
                {
                    var attributes = d.GetCustomAttributes(typeof(CompilerGeneratedAttribute), false);
                    if (attributes != null && attributes.Length > 0)
                        return true;
                }
            }

            return false;
        }

    }

}
