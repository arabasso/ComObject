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
            Type.SetProperty(binder.Name, Object,  TransformValue(value));

            return true;
        }

        public override bool TryGetIndex(
            GetIndexBinder binder,
            object[] indexes,
            out object result)
        {
            result = TransformValue(Type.GetIndex(Object, indexes));

            return true;
        }

        public override bool TryGetMember(
            GetMemberBinder binder, out object result)
        {
            result = TransformValue(Type.GetProperty(binder.Name, Object));

            return true;
        }

        public override bool TryInvokeMember(
            InvokeMemberBinder binder,
            object[] args,
            out object result)
        {
            result = TransformValue(Type.InvokeMethod(binder.Name, Object, args.Select(TransformValue).ToArray()));

            return true;
        }

        private object TransformValue(
            object value)
        {
            if (value == null) return null;

            if (IsComObject(value))
            {
                _objects.Add((ComObject)(value = new ComObject(value)));
            }

            else switch (value)
            {
                case IEnumerable a when !IsPrimitive(value):
                    value = a.Select(TransformValue).ToArray();
                    break;

                case ComObject o:
                    value = o.Object;
                    break;
            }

            return value;
        }

        public override bool TryConvert(
            ConvertBinder binder,
            out object result)
        {
            result = Object;

            return true;
        }

        private bool IsComObject(
            object o)
        {
            var t = o.GetType();

            return t.Name == "__ComObject";
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

        private readonly HashSet<ComObject> _objects = new HashSet<ComObject>();
    }
}
