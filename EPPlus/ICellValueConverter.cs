namespace EPPlus
{
    /// <summary>
    /// Represents the base contract for converting cell values using <see cref="ConvertCellValueAttribute"/>.
    /// </summary>
    public interface ICellValueConverter
    {
        /// <summary>
        /// Converts the given value.
        /// </summary>
        /// <param name="source">Source value to convert.</param>
        /// <returns>Returns converted value.</returns>
        object Convert(object source);
    }
}
