using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq.Expressions;

namespace LinqToExcel.Extensions
{
    public static class CommonExtensions
    {
        /// <summary>
        /// Sets the value of a property
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        /// <param name="value">Value to set the property to</param>
        public static void SetProperty(this object implementation, string propertyName, object value)
        {
            implementation.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, implementation, new[] { value });
        }

        /// <summary>
        /// Calls a method
        /// </summary>
        /// <param name="methodName">Name of the method</param>
        /// <param name="args">Method arguments</param>
        /// <returns>Return value of the method</returns>
        public static object CallMethod(this object implementation, string methodName, params object[] args)
        {
            return implementation.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, implementation, args);
        }

        public static T Cast<T>(this object implementation)
        {
            return (T)implementation.Cast(typeof(T));
        }

        public static object Cast(this object implementation, Type castType)
        {
            //return null for DBNull values
            if (implementation == null || implementation is DBNull)
            {
                return null;
            }

            //checking for nullable types
            if (castType.IsGenericType &&
                castType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                castType = Nullable.GetUnderlyingType(castType);
            }
            if (castType != null)
            {
                return Convert.ChangeType(implementation, castType);
            }
            return null;
        }

        public static IEnumerable<TResult> Cast<TResult>(this IEnumerable<object> list, Func<object, TResult> caster)
        {
            var results = new List<TResult>();
            foreach (var item in list)
            {
                results.Add(caster(item));
            }
            return results;
        }

        public static IEnumerable<TResult> Cast<TResult>(this IEnumerable<object> list)
        {
            var func = new Func<object, TResult>((item) =>
                (TResult)Convert.ChangeType(item, typeof(TResult)));
            return list.Cast(func);
        }

        public static string[] ToArray(this ICollection<string> collection)
        {
            var list = new List<string>();
            foreach (var item in collection)
            {
                list.Add(item);
            }
            return list.ToArray();
        }

        public static bool IsNumber(this string value)
        {
            return Regex.Match(value, @"^\d+$").Success;
        }

        public static bool IsNullValue(this Expression exp)
        {
            return ((exp is ConstantExpression) &&
                (exp.Cast<ConstantExpression>().Value == null));
        }

        public static string RegexReplace(this string source, string regex, string replacement)
        {
            return Regex.Replace(source, regex, replacement);
        }
    }
}
