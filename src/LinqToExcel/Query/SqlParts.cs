namespace LinqToExcel.Query
{
    using System.Collections.Generic;
    using System.Data.OleDb;
    using System.Text;

    public class SqlParts
    {
        public SqlParts()
        {
            this.Aggregate = "*";
            this.Parameters = new List<OleDbParameter>();
            this.OrderByAsc = true;
            this.ColumnNamesUsed = new List<string>();
        }

        public static implicit operator string(SqlParts sql)
        {
            return sql.ToString();
        }

        public override string ToString()
        {
            var sql = new StringBuilder();
            sql.AppendFormat("SELECT {0} FROM {1}", this.Aggregate, this.Table);
            if (!string.IsNullOrEmpty(this.Where))
            {
                sql.AppendFormat(" WHERE {0}", this.Where);
            }
            if (!string.IsNullOrEmpty(this.OrderBy))
            {
                var asc = this.OrderByAsc ? "ASC" : "DESC";
                sql.AppendFormat(" ORDER BY [{0}] {1}", this.OrderBy,
                    asc);
            }
            return sql.ToString();
        }

        public string Aggregate { get; set; }
        public List<string> ColumnNamesUsed { get; set; }
        public string OrderBy { get; set; }
        public bool OrderByAsc { get; set; }
        public IEnumerable<OleDbParameter> Parameters { get; set; }
        public string Table { get; set; }
        public string Where { get; set; }
    }
}