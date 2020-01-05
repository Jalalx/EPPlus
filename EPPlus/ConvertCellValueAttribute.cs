using System;

namespace EPPlus
{
    /// <summary>
    /// Converts value using specified converter.
    /// </summary>
    public class ConvertCellValueAttribute : Attribute
    {
        /// <summary>
        /// Creates an instance of <see cref="ConvertCellValueAttribute" />.
        /// </summary>
        /// <param name="converterType">An implementation of <see cref="ICellValueConverter"/>.</param>
        /// <param name="constructorArgs">Arguments for the constructor. Enter null if there is a default constructor.</param>
        public ConvertCellValueAttribute(Type converterType, params object[] constructorArgs)
        {
            ConverterType = converterType;
            ConstructorArgs = constructorArgs;
        }

        /// <summary>
        /// Gets the implemention of <see cref="ICellValueConverter"./>
        /// </summary>
        public Type ConverterType { get; }
        /// <summary>
        /// An array of arguments for the type constructor.
        /// </summary>
        public object[] ConstructorArgs { get; }
    }
}
