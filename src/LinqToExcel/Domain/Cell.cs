namespace LinqToExcel.Domain
{
    using System;

    /// <summary>
    /// Represents a cell and its value in an excel spreadsheet
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Cell's value
        /// </summary>
        public object Value { get; private set; }

        /// <param name="value">Cell's value</param>
        public Cell(object value)
        {
            this.Value = value;
        }

        /// <summary>
        /// Casts the cell's value to the generic argument type
        /// </summary>
        /// <typeparam name="T">Object type to convert to</typeparam>
        public T Cast<T>()
        {
            return (this.Value == null || this.Value is DBNull) ?
                default(T) :
                (T)Convert.ChangeType(this.Value, typeof(T));
        }

        /// <summary>
        /// Cell's value as a string
        /// </summary>
        public override string ToString()
        {
            return this.Value.ToString();
        }

        /// <summary>
        /// Allow cell to be implicitly cast to a string
        /// </summary>
        public static implicit operator string(Cell cell)
        {
            return cell.ToString();
        }
    }
}
