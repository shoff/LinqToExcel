namespace LinqToExcel.Query
{
    using System;
    using System.Collections.Generic;
    using System.Data.OleDb;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Text;
    using Domain;
    using Extensions;
    using Remotion.Linq.Parsing;

    public class WhereClauseExpressionTreeVisitor : ThrowingExpressionTreeVisitor
    {
        private readonly Dictionary<string, string> columnMapping;
        private readonly List<string> columnNamesUsed = new List<string>();
        private readonly List<OleDbParameter> oledbParameters = new List<OleDbParameter>();
        private readonly Type sheetType;
        private readonly List<string> validStringMethods;
        private readonly StringBuilder whereClause = new StringBuilder();

        public WhereClauseExpressionTreeVisitor(Type sheetType, Dictionary<string, string> columnMapping)
        {
            this.sheetType = sheetType;
            this.columnMapping = columnMapping;
            this.validStringMethods = new List<string>
            {
                "Equals",
                "Contains",
                "StartsWith",
                "IsNullOrEmpty",
                "EndsWith"
            };
        }

        public void Visit(Expression expression)
        {
            this.VisitExpression(expression);
        }

        protected override Exception CreateUnhandledItemException<T>(T unhandledItem, string visitMethod)
        {
            throw new NotImplementedException(visitMethod + " method is not implemented");
        }

        protected override Expression VisitBinaryExpression(BinaryExpression bExp)
        {
            this.whereClause.Append("(");

            // Patch for vb.net expression that are always considered a MethodCallExpression even if they are not.
            // see http://www.re-motion.org/blogs/mix/archive/2009/10/16/vb.net-specific-text-comparison-in-linq-queries.aspx
            bExp = this.ConvertVbStringCompare(bExp);

            //We always want the MemberAccess (ColumnName) to be on the left side of the statement
            var bLeft = bExp.Left;
            var bRight = bExp.Right;
            if (bExp.Right.NodeType == ExpressionType.MemberAccess &&
                ((MemberExpression) bExp.Right).Member.DeclaringType == this.sheetType)
            {
                bLeft = bExp.Right;
                bRight = bExp.Left;
            }

            this.VisitExpression(bLeft);
            switch (bExp.NodeType)
            {
                case ExpressionType.AndAlso:
                    this.whereClause.Append(" AND ");
                    break;
                case ExpressionType.Equal:
                    this.whereClause.Append(bRight.IsNullValue() ? " IS NULL" : " = ");
                    break;
                case ExpressionType.GreaterThan:
                    this.whereClause.Append(" > ");
                    break;
                case ExpressionType.GreaterThanOrEqual:
                    this.whereClause.Append(" >= ");
                    break;
                case ExpressionType.LessThan:
                    this.whereClause.Append(" < ");
                    break;
                case ExpressionType.LessThanOrEqual:
                    this.whereClause.Append(" <= ");
                    break;
                case ExpressionType.NotEqual:
                    this.whereClause.Append(bRight.IsNullValue() ? " IS NOT NULL" : " <> ");
                    break;
                case ExpressionType.OrElse:
                    this.whereClause.Append(" OR ");
                    break;
                default:
                    throw new NotSupportedException($"{bExp.NodeType} statement is not supported");
            }
            if (!bRight.IsNullValue())
            {
                this.VisitExpression(bRight);
            }
            this.whereClause.Append(")");
            return bExp;
        }

        protected BinaryExpression ConvertVbStringCompare(BinaryExpression exp)
        {
            if (exp.Left.NodeType == ExpressionType.Call)
            {
                var compareStringCall = (MethodCallExpression) exp.Left;
                if (compareStringCall.Method.DeclaringType != null &&
                    compareStringCall.Method.DeclaringType.FullName == "Microsoft.VisualBasic.CompilerServices.Operators" &&
                    compareStringCall.Method.Name == "CompareString")
                {
                    var arg1 = compareStringCall.Arguments[0];
                    var arg2 = compareStringCall.Arguments[1];

                    switch (exp.NodeType)
                    {
                        case ExpressionType.LessThan:
                            return Expression.LessThan(arg1, arg2);
                        case ExpressionType.LessThanOrEqual:
                            return Expression.LessThanOrEqual(arg1, arg2);
                        case ExpressionType.GreaterThan:
                            return Expression.GreaterThan(arg1, arg2);
                        case ExpressionType.GreaterThanOrEqual:
                            return Expression.GreaterThanOrEqual(arg1, arg2);
                        case ExpressionType.NotEqual:
                            return Expression.NotEqual(arg1, arg2);
                        default:
                            return Expression.Equal(arg1, arg2);
                    }
                }
            }
            return exp;
        }

        protected override Expression VisitMemberExpression(MemberExpression mExp)
        {
            //Set the column name to the property mapping if there is one, 
            //else use the property name for the column name
            var columnName = this.columnMapping.ContainsKey(mExp.Member.Name) ?
                this.columnMapping[mExp.Member.Name] :
                mExp.Member.Name;
            this.whereClause.AppendFormat("[{0}]", columnName);
            this.columnNamesUsed.Add(columnName);
            return mExp;
        }

        protected override Expression VisitConstantExpression(ConstantExpression cExp)
        {
            this.oledbParameters.Add(new OleDbParameter("?", cExp.Value));
            this.whereClause.Append("?");
            return cExp;
        }

        /// <summary>
        /// This method is visited when the LinqToExcel.Row type is used in the query
        /// </summary>
        protected override Expression VisitUnaryExpression(UnaryExpression uExp)
        {
            if (this.IsNotStringIsNullOrEmpty(uExp))
            {
                this.AddStringIsNullOrEmptyToWhereClause((MethodCallExpression) uExp.Operand, true);
            }
            else
            {
                this.whereClause.Append(this.GetColumnName(uExp.Operand));
            }
            return uExp;
        }

        private bool IsNotStringIsNullOrEmpty(UnaryExpression uExp)
        {
            return uExp.NodeType == ExpressionType.Not && ((MethodCallExpression) uExp.Operand).Method.Name == "IsNullOrEmpty";
        }

        /// <summary>
        /// Only As<>() method calls on the LinqToExcel.Row type are support
        /// </summary>
        protected override Expression VisitMethodCallExpression(MethodCallExpression mExp)
        {
            if (this.validStringMethods.Contains(mExp.Method.Name))
            {
                this.ProcessStringMethod(mExp);
            }
            else
            {
                var columnName = this.GetColumnName(mExp);
                this.whereClause.Append(columnName);
                this.columnNamesUsed.Add(columnName);
            }
            return mExp;
        }

        private void ProcessStringMethod(MethodCallExpression mExp)
        {
            switch (mExp.Method.Name)
            {
                case "Contains":
                    this.AddStringMethodToWhereClause(mExp, "LIKE", "%{0}%");
                    break;
                case "StartsWith":
                    this.AddStringMethodToWhereClause(mExp, "LIKE", "{0}%");
                    break;
                case "EndsWith":
                    this.AddStringMethodToWhereClause(mExp, "LIKE", "%{0}");
                    break;
                case "Equals":
                    this.AddStringMethodToWhereClause(mExp, "=", "{0}");
                    break;
                case "IsNullOrEmpty":
                    this.AddStringIsNullOrEmptyToWhereClause(mExp);
                    break;
            }
        }

        private void AddStringMethodToWhereClause(MethodCallExpression mExp, string operatorString, string argumentFormat)
        {
            this.whereClause.Append("(");
            this.VisitExpression(mExp.Object);
            this.whereClause.AppendFormat(" {0} ?)", operatorString);

            var value = mExp.Arguments.First().ToString().Replace("\"", "");
            var parameter = string.Format(argumentFormat, value);
            this.oledbParameters.Add(new OleDbParameter("?", parameter));
        }

        private void AddStringIsNullOrEmptyToWhereClause(MethodCallExpression mExp, bool notEqual = false)
        {
            var columnName = this.GetColumnName((MemberExpression) mExp.Arguments[0]);
            this.whereClause.AppendFormat(notEqual ? "(({0} <> '') OR ({0} IS NOT NULL))" : "(({0} = '') OR ({0} IS NULL))", columnName);
        }

        /// <summary>
        /// Retrieves the column name from a member expression or method call expression
        /// </summary>
        /// <param name="exp">Expression</param>
        private string GetColumnName(Expression exp)
        {
            var expression = exp as MemberExpression;
            if (expression != null)
            {
                return this.GetColumnName(expression);
            }
            return this.GetColumnName((MethodCallExpression) exp);
        }

        /// <summary>
        /// Retrieves the column name from a member expression
        /// </summary>
        /// <param name="mExp">Member Expression</param>
        private string GetColumnName(MemberExpression mExp)
        {
            return "[" + mExp.Member.Name + "]";
        }

        /// <summary>
        /// Retrieves the column name from a method call expression
        /// </summary>
        /// <param name="exp">Method Call Expression</param>
        private string GetColumnName(MethodCallExpression mExp)
        {
            var method = mExp;
            var o = mExp.Object as MethodCallExpression;
            if (o != null)
            {
                method = o;
            }

            var arg = method.Arguments.First();
            if (arg.Type == typeof(int))
            {
                if (this.sheetType == typeof(RowNoHeader))
                {
                    return $"F{int.Parse(arg.ToString()) + 1}";
                }
                throw new ArgumentException("Can only use column indexes in WHERE clause when using WorksheetNoHeader");
            }

            var columnName = arg.ToString().ToCharArray();
            columnName[0] = "[".ToCharArray().First();
            columnName[columnName.Length - 1] = "]".ToCharArray().First();
            return new string(columnName);
        }

        public IEnumerable<string> ColumnNamesUsed
        {
            get { return this.columnNamesUsed.Select(x => x.Replace("[", "").Replace("]", "")); }
        }

        public IEnumerable<OleDbParameter> Params => this.oledbParameters;

        public string WhereClause => this.whereClause.ToString();
    }
}