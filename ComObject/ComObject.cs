using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Runtime.InteropServices;

namespace ComObject
{
    public class ComObject
        : DynamicObject, IDisposable
    {
        public Type Type { get; }
        public object Object { get; }

        public ComObject(
            Type type,
            object @object)
        {
            Type = type;
            Object = @object;
        }

        public ComObject(
            object instance)
            : this(instance.GetType(), instance)
        {
        }

        public ComObject(
            Type type)
        {
            Type = type;
            Object = Activator.CreateInstance(type);
        }

        public ComObject(
            string typeId)
            : this(Type.GetTypeFromProgID(typeId, true))
        {
        }

        public override bool TrySetMember(
            SetMemberBinder binder,
            object value)
        {
            Type.SetProperty(binder.Name, Object, value is ComObject o ? o.Object : value);

            return true;
        }

        private readonly HashSet<ComObject> _objects = new HashSet<ComObject>();

        public override bool TryGetIndex(
            GetIndexBinder binder,
            object[] indexes,
            out object result)
        {
            result = Type.GetIndex(Object, indexes);

            if (result != null && !IsPrimitive(result))
            {
                _objects.Add((ComObject)(result = new ComObject(result)));
            }

            return true;
        }

        public override bool TryGetMember(
            GetMemberBinder binder, out object result)
        {
            result = Type.GetProperty(binder.Name, Object);

            if (result != null && !IsPrimitive(result))
            {
                _objects.Add((ComObject)(result = new ComObject(result)));
            }

            return true;
        }

        public override bool TryInvokeMember(
            InvokeMemberBinder binder,
            object[] args,
            out object result)
        {
            result = Type.InvokeMethod(binder.Name, Object, TransformArguments(args));

            if (result != null && !IsPrimitive(result))
            {
                _objects.Add((ComObject)(result = new ComObject(result)));
            }

            return true;
        }

        private object[] TransformArguments(
            IEnumerable args)
        {
            return args.Select(s =>
            {
                switch (s)
                {
                    case ComObject o:
                        return o.Object;

                    case IEnumerable a:
                        return IsPrimitive(s) ? s : TransformArguments(a);

                    default:
                        return s;
                }
            }).ToArray();
        }

        public override bool TryConvert(
            ConvertBinder binder,
            out object result)
        {
            result = Object;

            return true;
        }

        private bool IsPrimitive(
            object o)
        {
            var t = o.GetType();

            return t.IsPrimitive || t.IsValueType || (t == typeof(string));
        }

        protected bool Equals(
            ComObject other)
        {
            return Equals(Object, other.Object);
        }

        public override bool Equals(
            object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj.GetType() == GetType() && Equals((ComObject) obj);
        }

        public override int GetHashCode()
        {
            return (Object != null ? Object.GetHashCode() : 0);
        }

        public void Dispose()
        {
            foreach (var comObject in _objects)
            {
                comObject.Dispose();
            }

            Marshal.ReleaseComObject(Object);
        }
    }
}
