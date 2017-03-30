namespace LinqToExcel.Query
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using Domain;
    using Extensions;
    using NLog;
    using Remotion.Linq;
    using Remotion.Linq.Clauses.ResultOperators;

    internal class ExcelQueryExecutor : IQueryExecutor
    {
        private static readonly ILogger log = LogManager.GetCurrentClassLogger();
        private readonly ExcelQueryArgs args;

        internal ExcelQueryExecutor(ExcelQueryArgs args)
        {
            this.ValidateArgs(args);
            this.args = args;
            log.Debug("Connection String: {0}", ExcelUtilities.GetConnection(args).ConnectionString);

            this.GetWorksheetName();
        }

        /// <summary>
        /// Executes a query with a collection result.
        /// </summary>
        public IEnumerable<T> ExecuteCollection<T>(QueryModel queryModel)
        {
            var sql = this.GetSqlStatement(queryModel);
            this.LogSqlStatement(sql);

            var objectResults = this.GetDataResults(sql, queryModel);
            var enumerable = objectResults as object[] ?? objectResults.ToArray();
            var projector = this.GetSelectProjector<T>(enumerable.FirstOrDefault(), queryModel);
            var returnResults = enumerable.Cast(projector);

            foreach (var resultOperator in queryModel.ResultOperators)
            {
                if (resultOperator is ReverseResultOperator)
                {
                    returnResults = returnResults.Reverse();
                }
                if (resultOperator is SkipResultOperator)
                {
                    returnResults = returnResults.Skip(resultOperator.Cast<SkipResultOperator>().GetConstantCount());
                }
            }

            return returnResults;
        }

        /// <summary>
        /// Executes a query with a scalar result, i.e. a query that ends with a result operator such as Count, Sum, or Average.
        /// </summary>
        public T ExecuteScalar<T>(QueryModel queryModel)
        {
            return this.ExecuteSingle<T>(queryModel, false);
        }

        /// <summary>
        /// Executes a query with a single result object, i.e. a query that ends with a result operator such as First, Last, Single, Min, or Max.
        /// </summary>
        public T ExecuteSingle<T>(QueryModel queryModel, bool returnDefaultWhenEmpty)
        {
            var results = this.ExecuteCollection<T>(queryModel);

            foreach (var resultOperator in queryModel.ResultOperators)
            {
                if (resultOperator is LastResultOperator)
                {
                    return results.LastOrDefault();
                }
            }

            return returnDefaultWhenEmpty ?
                results.FirstOrDefault() :
                results.First();
        }

        private void ValidateArgs(ExcelQueryArgs args)
        {
            log.Debug("ExcelQueryArgs = {0}", args);

            if (args.FileName == null)
            {
                throw new ArgumentNullException(nameof(args));
            }

            if (!string.IsNullOrEmpty(args.StartRange) &&
                !Regex.Match(args.StartRange, "^[a-zA-Z]{1,3}[0-9]{1,7}$").Success)
            {
                throw new ArgumentException($"StartRange argument '{args.StartRange}' is invalid format for cell name");
            }

            if (!string.IsNullOrEmpty(args.EndRange) &&
                !Regex.Match(args.EndRange, "^[a-zA-Z]{1,3}[0-9]{1,7}$").Success)
            {
                throw new ArgumentException($"EndRange argument '{args.EndRange}' is invalid format for cell name");
            }

            if (args.NoHeader &&
                !string.IsNullOrEmpty(args.StartRange) &&
                args.FileName.ToLower().Contains(".csv"))
            {
                throw new ArgumentException("Cannot use WorksheetRangeNoHeader on csv files");
            }
        }

        protected Func<object, T> GetSelectProjector<T>(object firstResult, QueryModel queryModel)
        {
            Func<object, T> projector = result => result.Cast<T>();
            if (this.ShouldBuildResultObjectMapping<T>(firstResult, queryModel))
            {
                var proj = ProjectorBuildingExpressionTreeVisitor.BuildProjector<T>(queryModel.SelectClause.Selector);
                projector = result => proj(new ResultObjectMapping(queryModel.MainFromClause, result));
            }
            return projector;
        }

        protected bool ShouldBuildResultObjectMapping<T>(object firstResult, QueryModel queryModel)
        {
            var ignoredResultOperators = new List<Type>
            {
                typeof(MaxResultOperator),
                typeof(CountResultOperator),
                typeof(LongCountResultOperator),
                typeof(MinResultOperator),
                typeof(SumResultOperator)
            };

            return firstResult != null &&
                firstResult.GetType() != typeof(T) &&
                !queryModel.ResultOperators.Any(x => ignoredResultOperators.Contains(x.GetType()));
        }

        protected SqlParts GetSqlStatement(QueryModel queryModel)
        {
            var sqlVisitor = new SqlGeneratorQueryModelVisitor(this.args);
            sqlVisitor.VisitQueryModel(queryModel);
            return sqlVisitor.SqlStatement;
        }

        private void GetWorksheetName()
        {
            if (this.args.FileName.ToLower().EndsWith("csv"))
            {
                this.args.WorksheetName = Path.GetFileName(this.args.FileName);
            }
            else if (this.args.WorksheetIndex.HasValue)
            {
                var worksheetNames = ExcelUtilities.GetWorksheetNames(this.args);
                var enumerable = worksheetNames as string[] ?? worksheetNames.ToArray();
                if (this.args.WorksheetIndex.Value < enumerable.Count())
                {
                    this.args.WorksheetName = enumerable.ElementAt(this.args.WorksheetIndex.Value);
                }
                else
                {
                    throw new DataException("Worksheet Index Out of Range");
                }
            }
            else if (string.IsNullOrEmpty(this.args.WorksheetName) && string.IsNullOrEmpty(this.args.NamedRangeName))
            {
                this.args.WorksheetName = "Sheet1";
            }
        }

        /// <summary>
        /// Executes the sql query and returns the data results
        /// </summary>
        /// <typeparam name="T">Data type in the main from clause (queryModel.MainFromClause.ItemType)</typeparam>
        /// <param name="queryModel">Linq query model</param>
        protected IEnumerable<object> GetDataResults(SqlParts sql, QueryModel queryModel)
        {
            IEnumerable<object> results;
            OleDbDataReader data = null;

            var conn = ExcelUtilities.GetConnection(this.args);
            var command = conn.CreateCommand();
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                command.CommandText = sql.ToString();
                command.Parameters.AddRange(sql.Parameters.ToArray());
                try
                {
                    data = command.ExecuteReader();
                }
                catch (OleDbException e)
                {
                    if (e.Message.Contains(this.args.WorksheetName))
                    {
                        throw new DataException(
                            string.Format(
                                "'{0}' is not a valid worksheet name in file {3}. Valid worksheet names are: '{1}'. Error received: {2}",
                                this.args.WorksheetName,
                                string.Join("', '", ExcelUtilities.GetWorksheetNames(this.args.FileName).ToArray()), e.Message,
                                this.args.FileName), e);
                    }
                    if (!this.CheckIfInvalidColumnNameUsed(sql))
                    {
                        throw;
                    }
                }

                var columns = ExcelUtilities.GetColumnNames(data);
                var enumerable = columns as string[] ?? columns.ToArray();
                this.LogColumnMappingWarnings(enumerable);
                if (enumerable.Length == 1 && enumerable.First() == "Expr1000")
                {
                    results = this.GetScalarResults(data);
                }
                else if (queryModel.MainFromClause.ItemType == typeof(Row))
                {
                    results = this.GetRowResults(data, enumerable);
                }
                else if (queryModel.MainFromClause.ItemType == typeof(RowNoHeader))
                {
                    results = this.GetRowNoHeaderResults(data);
                }
                else
                {
                    results = this.GetTypeResults(data, enumerable, queryModel);
                }
            }
            finally
            {
                command.Dispose();

                if (!this.args.UsePersistentConnection)
                {
                    conn.Dispose();
                    this.args.PersistentConnection = null;
                }
            }

            return results;
        }

        /// <summary>
        /// Logs a warning for any property to column mappings that do not exist in the excel worksheet
        /// </summary>
        /// <param name="Columns">List of columns in the worksheet</param>
        private void LogColumnMappingWarnings(IEnumerable<string> columns)
        {
            foreach (var kvp in this.args.ColumnMappings)
            {
                var enumerable = columns as string[] ?? columns.ToArray();
                if (!enumerable.Contains(kvp.Value))
                {
                    log.Warn("'{0}' column that is mapped to the '{1}' property does not exist in the '{2}' worksheet",
                        kvp.Value, kvp.Key, this.args.WorksheetName);
                }
            }
        }

        private bool CheckIfInvalidColumnNameUsed(SqlParts sql)
        {
            var usedColumns = sql.ColumnNamesUsed;
            var tableColumns = ExcelUtilities.GetColumnNames(this.args.WorksheetName, this.args.NamedRangeName, this.args.FileName);
            foreach (var column in usedColumns)
            {
                var enumerable = tableColumns as string[] ?? tableColumns.ToArray();
                if (!enumerable.Contains(column))
                {
                    throw new DataException($"'{column}' is not a valid column name. " +
                        $"Valid column names are: '{string.Join("', '", enumerable.ToArray())}'");
                }
            }
            return false;
        }

        private IEnumerable<object> GetRowResults(IDataReader data, IEnumerable<string> columns)
        {
            var results = new List<object>();
            var columnIndexMapping = new Dictionary<string, int>();
            var enumerable = columns as string[] ?? columns.ToArray();
            for (var i = 0; i < enumerable.Length; i++)
            {
                columnIndexMapping[enumerable.ElementAt(i)] = i;
            }

            while (data.Read())
            {
                IList<Cell> cells = new List<Cell>();
                for (var i = 0; i < enumerable.Length; i++)
                {
                    var value = data[i];
                    value = this.TrimStringValue(value);
                    cells.Add(new Cell(value));
                }
                results.CallMethod("Add", new Row(cells, columnIndexMapping));
            }
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetRowNoHeaderResults(OleDbDataReader data)
        {
            var results = new List<object>();
            while (data.Read())
            {
                IList<Cell> cells = new List<Cell>();
                for (var i = 0; i < data.FieldCount; i++)
                {
                    var value = data[i];
                    value = this.TrimStringValue(value);
                    cells.Add(new Cell(value));
                }
                results.CallMethod("Add", new RowNoHeader(cells));
            }
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetTypeResults(IDataReader data, IEnumerable<string> columns, QueryModel queryModel)
        {
            var results = new List<object>();
            var fromType = queryModel.MainFromClause.ItemType;
            var props = fromType.GetProperties();
            var argsStrictMapping = this.args.StrictMapping;
            var enumerable = columns as string[] ?? columns.ToArray();
            if (argsStrictMapping != null && argsStrictMapping.Value != StrictMappingType.None)
            {
                var strictMappingType = this.args.StrictMapping;
                if (strictMappingType != null)
                {
                    this.ConfirmStrictMapping(enumerable, props, strictMappingType.Value);
                }
            }

            while (data.Read())
            {
                var result = Activator.CreateInstance(fromType);
                foreach (var prop in props)
                {
                    var columnName = this.args.ColumnMappings.ContainsKey(prop.Name) ? this.args.ColumnMappings[prop.Name] :
                        prop.Name;
                    if (enumerable.Contains(columnName))
                    {
                        var value = this.GetColumnValue(data, columnName, prop.Name).Cast(prop.PropertyType);
                        value = this.TrimStringValue(value);
                        result.SetProperty(prop.Name, value);
                    }
                }
                results.Add(result);
            }
            return results.AsEnumerable();
        }

        /// <summary>
        /// Trims leading and trailing spaces, based on the value of _args.TrimSpaces
        /// </summary>
        /// <param name="value">Input string value</param>
        /// <returns>Trimmed string value</returns>
        private object TrimStringValue(object value)
        {
            if (value == null || value.GetType() != typeof(string))
            {
                return value;
            }

            switch (this.args.TrimSpaces)
            {
                case TrimSpacesType.Start:
                    return ((string)value).TrimStart();
                case TrimSpacesType.End:
                    return ((string)value).TrimEnd();
                case TrimSpacesType.Both:
                    return ((string)value).Trim();
                default:
                    return value;
            }
        }

        private void ConfirmStrictMapping(IEnumerable<string> columns, PropertyInfo[] properties, StrictMappingType strictMappingType)
        {
            var propertyNames = properties.Select(x => x.Name);
            var enumerable = propertyNames as string[] ?? propertyNames.ToArray();
            var enumerable1 = columns as string[] ?? columns.ToArray();
            if (strictMappingType == StrictMappingType.ClassStrict || strictMappingType == StrictMappingType.Both)
            {
                foreach (var propertyName in enumerable)
                {
                    if (!enumerable1.Contains(propertyName) && this.PropertyIsNotMapped(propertyName))
                    {
                        throw new StrictMappingException("'{0}' property is not mapped to a column", propertyName);
                    }
                }
            }

            if (strictMappingType == StrictMappingType.WorksheetStrict || strictMappingType == StrictMappingType.Both)
            {
                foreach (var column in enumerable1)
                {
                    if (!enumerable.Contains(column) && this.ColumnIsNotMapped(column))
                    {
                        throw new StrictMappingException("'{0}' column is not mapped to a property", column);
                    }
                }
            }
        }

        private bool PropertyIsNotMapped(string propertyName)
        {
            return !this.args.ColumnMappings.Keys.Contains(propertyName);
        }

        private bool ColumnIsNotMapped(string columnName)
        {
            return !this.args.ColumnMappings.Values.Contains(columnName);
        }

        private object GetColumnValue(IDataRecord data, string columnName, string propertyName)
        {
            //Perform the property transformation if there is one
            return this.args.Transformations.ContainsKey(propertyName)
                ? this.args.Transformations[propertyName](data[columnName].ToString()) :
                data[columnName];
        }

        private IEnumerable<object> GetScalarResults(IDataReader data)
        {
            data.Read();
            return new List<object> { data[0] };
        }

        private void LogSqlStatement(SqlParts sqlParts)
        {
            if (log != null && log.IsDebugEnabled)
            {
                var logMessage = new StringBuilder();
                logMessage.AppendFormat("{0};", sqlParts);
                for (var i = 0; i < sqlParts.Parameters.Count(); i++)
                {
                    var paramValue = sqlParts.Parameters.ElementAt(i).Value.ToString();
                    var paramMessage = $" p{i} = '{sqlParts.Parameters.ElementAt(i).Value}';";

                    if (paramValue.IsNumber())
                    {
                        paramMessage = paramMessage.Replace("'", "");
                    }
                    logMessage.Append(paramMessage);
                }

                log.Debug(logMessage.ToString());
            }
        }

 }
}