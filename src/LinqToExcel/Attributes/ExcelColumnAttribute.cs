using System;

namespace LinqToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute(string columnName)
        {
            this.ColumnName = columnName;
        }

        public string ColumnName { get; private set; }
    }
}
