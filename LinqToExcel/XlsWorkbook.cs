using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;

namespace Excel.Linq
{
    /// <summary>
    /// ワークブックはプロバイダでもある
    /// </summary>
    public class XlsWorkbook : IQueryable<IXLWorksheet>, IQueryProvider, IDisposable
    {
        #region フィールド
        private XLWorkbook m_workbook { get; set; }
        private Expression m_expression { get; set; }
        #endregion

        #region コンストラクタ
        public XlsWorkbook(string fileName)
        {
            m_workbook = new XLWorkbook(fileName);
        }

        public XlsWorkbook(Expression _expression)
        {
            if (m_workbook == null)
            {
                m_workbook = new XLWorkbook();
            }
            m_expression = _expression;
        }
        #endregion

        #region プライベートクラス

        private class ExpressionRebuilder
        {
            /// <summary>
            /// XlsWorksheet型のParameterExpressionと置換するParameter
            /// </summary>
            private ParameterExpression paramExpr;

            public ExpressionRebuilder(string name)
            {
                paramExpr = Expression.Parameter(typeof(IXLWorksheet), name);
            }

            public Expression Rebuild(Expression expr)
            {
                BinaryExpression binExpr;
                UnaryExpression uniExpr;

                switch (expr.NodeType)
                {
                    case ExpressionType.Lambda:
                        return Expression.Lambda<Predicate<IXLWorksheet>>(
                            Rebuild(((LambdaExpression)expr).Body),
                            ((LambdaExpression)expr).Parameters.ToList().ConvertAll<ParameterExpression>(p => (ParameterExpression)Rebuild(p)));
                    case ExpressionType.Equal:
                    case ExpressionType.NotEqual:
                    case ExpressionType.GreaterThan:
                    case ExpressionType.GreaterThanOrEqual:
                    case ExpressionType.LessThan:
                    case ExpressionType.LessThanOrEqual:
                        binExpr = (BinaryExpression)expr;
                        return (Expression)typeof(Expression).InvokeMember(expr.NodeType.ToString(),
                            BindingFlags.Public | BindingFlags.Static | BindingFlags.InvokeMethod,
                            null, null,
                            new object[] {
                                Rebuild(binExpr.Left),
                                Rebuild(binExpr.Right),
                                binExpr.IsLiftedToNull, binExpr.Method
                            });
                    case ExpressionType.Not:
                        uniExpr = (UnaryExpression)expr;
                        return (Expression)typeof(Expression).InvokeMember(expr.NodeType.ToString(),
                            BindingFlags.Public | BindingFlags.Static | BindingFlags.InvokeMethod,
                            null, null,
                            new object[]{
                                Rebuild(uniExpr.Operand),
                                uniExpr.Method
                            });
                    case ExpressionType.And:
                    case ExpressionType.AndAlso:
                    case ExpressionType.Or:
                    case ExpressionType.OrElse:
                        binExpr = (BinaryExpression)expr;
                        return (Expression)typeof(Expression).InvokeMember(expr.NodeType.ToString(),
                            BindingFlags.Public | BindingFlags.Static | BindingFlags.InvokeMethod,
                            null, null,
                            new object[] {
                                Rebuild(binExpr.Left),
                                Rebuild(binExpr.Right),
                                binExpr.Method
                            });
                    case ExpressionType.Convert:
                        uniExpr = (UnaryExpression)expr;
                        return Expression.Convert(Rebuild(uniExpr.Operand), uniExpr.Type);
                    case ExpressionType.Parameter:
                        return ((ParameterExpression)expr).Type == typeof(IXLWorksheet) ? paramExpr : expr;

                    case ExpressionType.MemberAccess:
                        MemberExpression mexpr = (MemberExpression)expr;
                        // IXLWorksheet
                        MemberInfo member = mexpr.Member.DeclaringType != typeof(IXLWorksheet) ?
                            mexpr.Member :
                            ((Func<MemberInfo, MemberInfo>)((m) => {
                                if (m.Name == "Name") return typeof(IXLWorksheet).GetProperty("Name");

                                return typeof(IXLWorksheet).GetProperty(m.Name);

                            }))(mexpr.Member);
                        return Expression.Property(Rebuild(mexpr.Expression), (PropertyInfo)member);

                    case ExpressionType.Call:
                        var call = (MethodCallExpression)expr;

                        var target = call.Object != null ? Rebuild(call.Object) : null;
                        var mi = call.Method;
                        var arg = call.Arguments.Select(Rebuild).ToArray();
                        return Expression.Call(target, mi, arg);
                }
                return expr;
            }
        }

        #endregion

        #region プライベートメソッド
        /// <summary>
        /// 指定した式を条件として、オブジェクトの列挙を行います。
        /// </summary>
        /// <param name="expression">式</param>
        /// <returns>コレクション</returns>
        private IEnumerable<IXLWorksheet> ExecuteExpression(Expression expression)
        {
            Predicate<IXLWorksheet> predicate = ParseExpression(expression);

            foreach (IXLWorksheet sheet in m_workbook.Worksheets)
            {
                if (predicate(sheet)) yield return sheet;
            }
        }

        private Expression RebuildExpression(Expression expression)
        {
            return new ExpressionRebuilder("s").Rebuild(expression);
        }

        /// <summary>
        /// 指定した式を解析して、適切なデリゲートに変換します。
        /// </summary>
        /// <param name="expression">式</param>
        /// <returns>デリゲート</returns>
        private Predicate<IXLWorksheet> ParseExpression(Expression expression)
        {
            Expression lexpr = RebuildExpression(expression);

            return (Predicate<IXLWorksheet>)((LambdaExpression)lexpr).Compile();
        }

        /// <summary>
        /// 式無しでアイテムの列挙のみを行います。
        /// </summary>
        /// <returns>コレクション</returns>
        private IEnumerable<IXLWorksheet> ForEachWithoutExpression()
        {
            foreach (IXLWorksheet sheet in m_workbook.Worksheets) yield return sheet;
        }
        #endregion

        #region IQuerable<IXLWorksheet>
        public Expression Expression => m_expression;

        public Type ElementType => typeof(IXLWorksheet);

        public IQueryProvider Provider => this;

        public IEnumerator<IXLWorksheet> GetEnumerator()
        {
            return Provider.Execute<IEnumerator<IXLWorksheet>>(Expression);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
        #endregion

        #region IQueryProvider

        public IQueryable CreateQuery(Expression expression)
        {
            m_expression = expression;
            return this;
        }

        public IQueryable<TElement> CreateQuery<TElement>(Expression expression)
        {
            return (IQueryable<TElement>)this.CreateQuery(expression);
        }

        public object Execute(Expression expression)
        {
            return (expression != null ? ExecuteExpression(expression) : ForEachWithoutExpression()).GetEnumerator();
        }

        public TResult Execute<TResult>(Expression expression)
        {
            return (TResult)this.Execute(expression);
        }

        #endregion

        #region IDisposable
        public void Dispose()
        {
            this.m_workbook.Dispose();
        }

        #endregion
    }

    /// <summary>
    /// 拡張メソッド
    /// </summary>
    static public class XlsWorkbookExtension
    {
        public static IQueryable<IXLWorksheet> Where(this IQueryable<IXLWorksheet> q, Expression<Predicate<IXLWorksheet>> expression)
        {
            return q.Provider.CreateQuery<IXLWorksheet>(expression);
        }
        public static IQueryable<IXLWorksheet> Where(this XlsWorkbook q, Expression<Predicate<IXLWorksheet>> expression)
        {
            return q.CreateQuery<IXLWorksheet>(expression);
        }
    }
}
