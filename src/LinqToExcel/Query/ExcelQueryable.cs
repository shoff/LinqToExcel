namespace LinqToExcel.Query
{
    using System;
    using System.Linq;
    using System.Linq.Expressions;
    using Attributes;
    using Remotion.Linq;

    public class ExcelQueryable<T> : QueryableBase<T>
    {
        // This constructor is called by users, create a new IQueryExecutor.
        internal ExcelQueryable(ExcelQueryArgs args)
            : base(CreateExecutor(args))
        {
            foreach (var property in typeof(T).GetProperties())
            {
                var att = (ExcelColumnAttribute) Attribute.GetCustomAttribute(property, typeof(ExcelColumnAttribute));
                if (att != null && !args.ColumnMappings.ContainsKey(property.Name))
                {
                    args.ColumnMappings.Add(property.Name, att.ColumnName);
                }
            }
        }

        // This constructor is called indirectly by LINQ's query methods, just pass to base.
        public ExcelQueryable(IQueryProvider provider, Expression expression)
            : base(provider, expression)
        {
        }

        private static IQueryExecutor CreateExecutor(ExcelQueryArgs args)
        {
            return new ExcelQueryExecutor(args);
        }
    }
}