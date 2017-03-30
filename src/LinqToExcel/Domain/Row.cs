namespace LinqToExcel.Domain
{
    using System;
    using System.Collections.Generic;
    using Extensions;

    public class Row : List<Cell>
    {
       private readonly IDictionary<string, int> columnIndexMapping;

        public Row() : 
            this(new List<Cell>(),new Dictionary<string, int>())
        { }

        /// <param name="cells">Cells contained within the row</param>
        /// <param name="columnIndexMapping">Column name to cell index mapping</param>
        public Row(IList<Cell> cells, IDictionary<string, int> columnIndexMapping)
        {
            for (int i = 0; i < cells.Count; i++)
            {
                this.Insert(i, cells[i]);
            }
            this.columnIndexMapping = columnIndexMapping;
        }

        /// <param name="columnName">Column Name</param>
        public Cell this[string columnName]
        {
            get 
            {
                if (!this.columnIndexMapping.ContainsKey(columnName))
                {
                    // ReSharper disable UseStringInterpolation
                    throw new ArgumentException(string.Format("'{0}' column name does not exist. Valid column names are '{1}'", 
                        columnName, string.Join("', '", this.columnIndexMapping.Keys.ToArray())));
                    // ReSharper restore UseStringInterpolation

                }
                return base[this.columnIndexMapping[columnName]]; 
            }
        }

        /// <summary>
        /// List of column names in the row object
        /// </summary>
        public IEnumerable<string> ColumnNames => this.columnIndexMapping.Keys;
    }
}
