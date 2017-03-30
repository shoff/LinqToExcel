using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace LinqToExcel.Query
{
    using System.Collections.ObjectModel;
    using Remotion.Linq;
    using Remotion.Linq.Clauses;
    using Remotion.Linq.Clauses.ResultOperators;

    internal class SqlGeneratorQueryModelVisitor : QueryModelVisitorBase
    {
        public SqlParts SqlStatement { get; protected set; }
        private readonly ExcelQueryArgs excelQueryArgs;

        internal SqlGeneratorQueryModelVisitor(ExcelQueryArgs excelQueryArgs)
        {
            this.excelQueryArgs = excelQueryArgs;
            SqlStatement = new SqlParts
            {
                Table = (string.IsNullOrEmpty(this.excelQueryArgs.StartRange)) ?
                    !string.IsNullOrEmpty(this.excelQueryArgs.NamedRangeName) &&
                    string.IsNullOrEmpty(this.excelQueryArgs.WorksheetName) ?
                        $"[{this.excelQueryArgs.NamedRangeName}]" :
                        $"[{this.excelQueryArgs.WorksheetName}${this.excelQueryArgs.NamedRangeName}]" :
                    $"[{this.excelQueryArgs.WorksheetName}${this.excelQueryArgs.StartRange}:{this.excelQueryArgs.EndRange}]"
            };

            if (!string.IsNullOrEmpty(this.excelQueryArgs.WorksheetName) && this.excelQueryArgs.WorksheetName.ToLower().EndsWith(".csv"))
            {
                this.SqlStatement.Table = this.SqlStatement.Table.Replace("$]", "]");
            }
        }

        public override void VisitGroupJoinClause(GroupJoinClause groupJoinClause, QueryModel queryModel, int index)
        {
            throw new NotSupportedException("LinqToExcel does not provide support for group join");
        }

        public override void VisitJoinClause(JoinClause joinClause, QueryModel queryModel, int index)
        {
            throw new NotSupportedException("LinqToExcel does not provide support for the Join() method");
        }

        public override void VisitQueryModel(QueryModel queryModel)
        {
            queryModel.SelectClause.Accept(this, queryModel);
            queryModel.MainFromClause.Accept(this, queryModel);
            VisitBodyClauses(queryModel.BodyClauses, queryModel);
            VisitResultOperators(queryModel.ResultOperators, queryModel);

            if (queryModel.MainFromClause.ItemType.Name == "IGrouping`2")
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Group() method");
            }
        }

        public override void VisitWhereClause(WhereClause whereClause, QueryModel queryModel, int index)
        {
            var where = new WhereClauseExpressionTreeVisitor(queryModel.MainFromClause.ItemType, this.excelQueryArgs.ColumnMappings);
            where.Visit(whereClause.Predicate);
            SqlStatement.Where = where.WhereClause;
            SqlStatement.Parameters = where.Params;
            SqlStatement.ColumnNamesUsed.AddRange(where.ColumnNamesUsed);

            base.VisitWhereClause(whereClause, queryModel, index);
        }

        public override void VisitResultOperator(ResultOperatorBase resultOperator, QueryModel queryModel, int index)
        {
            //Affects SQL result operators
            var roperator = resultOperator as TakeResultOperator;
            if (roperator != null)
            {
                var take = roperator;
                SqlStatement.Aggregate = $"TOP {take.Count} *";
            }
            else if (resultOperator is AverageResultOperator)
            {
                this.UpdateAggregate(queryModel, "AVG");
            }
            else if (resultOperator is CountResultOperator)
            {
                this.SqlStatement.Aggregate = "COUNT(*)";
            }
            else if (resultOperator is LongCountResultOperator)
            {
                this.SqlStatement.Aggregate = "COUNT(*)";
            }
            else if (resultOperator is FirstResultOperator)
            {
                this.SqlStatement.Aggregate = "TOP 1 *";
            }
            else if (resultOperator is MaxResultOperator)
            {
                this.UpdateAggregate(queryModel, "MAX");
            }
            else if (resultOperator is MinResultOperator)
            {
                this.UpdateAggregate(queryModel, "MIN");
            }
            else if (resultOperator is SumResultOperator)
            {
                this.UpdateAggregate(queryModel, "SUM");
            }
            else if (resultOperator is DistinctResultOperator)
            {
                this.ProcessDistinctAggregate(queryModel);
            }

            //Not supported result operators
            else if (resultOperator is ContainsResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Contains() method");
            }
            else if (resultOperator is DefaultIfEmptyResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the DefaultIfEmpty() method");
            }
            else if (resultOperator is ExceptResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Except() method");
            }
            else if (resultOperator is GroupResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Group() method");
            }
            else if (resultOperator is IntersectResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Intersect() method");
            }
            else if (resultOperator is OfTypeResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the OfType() method");
            }
            else if (resultOperator is SingleResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Single() method. Use the First() method instead");
            }
            else if (resultOperator is UnionResultOperator)
            {
                throw new NotSupportedException("LinqToExcel does not provide support for the Union() method");
            }

            base.VisitResultOperator(resultOperator, queryModel, index);
        }

        protected override void VisitBodyClauses(ObservableCollection<IBodyClause> bodyClauses, QueryModel queryModel)
        {
            var orderClause = bodyClauses
                .FirstOrDefault(x => x.GetType() == typeof(OrderByClause))
                as OrderByClause;

            if (orderClause != null)
            {
                var columnName = "";
                var exp = orderClause.Orderings.First().Expression;
                if (exp is MemberExpression)
                {
                    var mExp = exp as MemberExpression;
                    columnName = (this.excelQueryArgs.ColumnMappings.ContainsKey(mExp.Member.Name)) ?
                        this.excelQueryArgs.ColumnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                }
                else if (exp is MethodCallExpression)
                {
                    //row["ColumnName"] is being used in order by statement
                    columnName = ((MethodCallExpression)exp).Arguments.First()
                        .ToString().Replace("\"", "");
                }

                SqlStatement.OrderBy = columnName;
                SqlStatement.ColumnNamesUsed.Add(columnName);
                var orderDirection = orderClause.Orderings.First().OrderingDirection;
                SqlStatement.OrderByAsc = (orderDirection == OrderingDirection.Asc);
            }
            base.VisitBodyClauses(bodyClauses, queryModel);
        }

        protected void UpdateAggregate(QueryModel queryModel, string aggregateName)
        {
            var columnName = GetResultColumnName(queryModel);
            SqlStatement.Aggregate = string.Format("{0}({1})",
                aggregateName,
                columnName);
            SqlStatement.ColumnNamesUsed.Add(columnName);
        }

        protected void ProcessDistinctAggregate(QueryModel queryModel)
        {
            if (queryModel.SelectClause.Selector is MemberExpression)
            {
                this.UpdateAggregate(queryModel, "DISTINCT");
            }
            else
            {
                throw new NotSupportedException("LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]");
            }
        }

        private string GetResultColumnName(QueryModel queryModel)
        {
            var mExp = queryModel.SelectClause.Selector as MemberExpression;
            if (mExp != null)
            {
                return (this.excelQueryArgs.ColumnMappings != null && this.excelQueryArgs.ColumnMappings.ContainsKey(mExp.Member.Name)) ?
                    this.excelQueryArgs.ColumnMappings[mExp.Member.Name] :
                    mExp.Member.Name;
            }
            return "";
        }

    }
}
