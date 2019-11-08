using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace ComObject
{
    static class Extensions
    {
        public static object InvokeMethod(
            this Type type,
            string name,
            object obj,
            params object [] args)
        {
            return type.InvokeMember(name, BindingFlags.InvokeMethod, Type.DefaultBinder, obj, args);
        }

        public static void SetProperty(
            this Type type,
            string name,
            object obj,
            object value)
        {
            type.InvokeMember(name, BindingFlags.SetProperty, Type.DefaultBinder, obj, new []{ value });
        }


        public static object GetProperty(
            this Type type,
            string name,
            object obj)
        {
            return type.InvokeMember(name, BindingFlags.GetProperty, Type.DefaultBinder, obj, null);
        }

        public static object GetIndex(
            this Type type,
            object obj,
            object [] indexes)
        {
            return type.InvokeMember("Item", BindingFlags.InvokeMethod, Type.DefaultBinder, obj, indexes);
        }

        public static IEnumerable<object> Select(
            this IEnumerable enumerable,
            Func<object, object> func)
        {
            foreach (var item in enumerable)
            {
                yield return func(item);
            }
        }
    }
}
