using System;

namespace EPPlus
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Method, AllowMultiple = false)]
    /// <summary>
    /// Ignores a field or property when using <see cref="LoadFromCollection" />.
    /// </summary>
    public class IgnoreColumnAttribute : Attribute
    {
    }
}
