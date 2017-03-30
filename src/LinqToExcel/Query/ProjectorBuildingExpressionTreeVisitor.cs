namespace LinqToExcel.Query
{
    using System;
    using System.Linq.Expressions;
    using System.Reflection;
    using Remotion.Linq.Clauses.Expressions;
    using Remotion.Linq.Parsing;


    public class ProjectorBuildingExpressionTreeVisitor : RelinqExpressionVisitor
    {
        // This is the generic ResultObjectMapping.GetObject<T>() method we'll use to obtain a queried object for an IQuerySource.
        private static readonly MethodInfo getObjectGenericMethodDefinition = typeof(ResultObjectMapping).GetMethod("GetObject");

        private readonly ParameterExpression resultItemParameter;

        private ProjectorBuildingExpressionTreeVisitor(ParameterExpression resultItemParameter)
        {
            this.resultItemParameter = resultItemParameter;
        }

        // Call this method to get the projector. T is the type of the result (after the projection).
        public static Func<ResultObjectMapping, T> BuildProjector<T>(Expression selectExpression)
        {
            // This is the parameter of the delegat we're building. It's the ResultObjectMapping, which holds all the input data needed for the projection.
            var resultItemParameter = Expression.Parameter(typeof(ResultObjectMapping), "resultItem");

            // The visitor gives us the projector's body. It simply replaces all QuerySourceReferenceExpressions with calls to ResultObjectMapping.GetObject<T>().
            var visitor = new ProjectorBuildingExpressionTreeVisitor(resultItemParameter);
            var body = visitor.Visit(selectExpression);

            // Construct a LambdaExpression from parameter and body and compile it into a delegate.
            var projector = Expression.Lambda<Func<ResultObjectMapping, T>>(body, resultItemParameter);
            return projector.Compile();
        }
  

        protected Expression VisitQuerySourceReferenceExpression(QuerySourceReferenceExpression expression)
        {
            // Substitute generic parameter "T" of ResultObjectMapping.GetObject<T>() with type of query source item, then return a call to that method
            // with the query source referenced by the expression.
            var getObjectMethod = getObjectGenericMethodDefinition.MakeGenericMethod(expression.Type);
            return Expression.Call(this.resultItemParameter, getObjectMethod, Expression.Constant(expression.ReferencedQuerySource));
        }
        

        protected Expression VisitSubQueryExpression(SubQueryExpression expression)
        {
            throw new NotSupportedException("This provider does not support subqueries in the select projection.");
        }
    }
}