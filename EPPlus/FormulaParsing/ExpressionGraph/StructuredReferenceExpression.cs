namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Represents a structured reference expression that resolves to cells within a table.
	/// </summary>
	public class StructuredReferenceExpression : ExcelAddressExpression
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of an <see cref="StructuredReferenceExpression"/>.
		/// </summary>
		/// <param name="expression">The expression string.</param>
		/// <param name="excelDataProvider">An <see cref="ExcelDataProvider"/> for resolving structured references.</param>
		/// <param name="parsingContext">The current <see cref="ParsingContext"/>.</param>
		/// <param name="negate">A value indicating whether or not to negate the expression.</param>
		public StructuredReferenceExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, bool negate)
			 : base(expression, excelDataProvider, parsingContext, negate) { }
		#endregion

		#region ExcelAddressExpression Overrides
		/// <summary>
		/// Compiles the expression into a value.
		/// </summary>
		/// <returns>The <see cref="CompileResult"/> with the expression value.</returns>
		public override CompileResult Compile()
		{
			var structuredReference = new StructuredReference(this.ExpressionString);
			var c = this._parsingContext.Scopes.Current;
			var result = _excelDataProvider.ResolveStructuredReference(structuredReference, c.Address.Worksheet, c.Address.FromRow, c.Address.FromCol);
			if (result == null)
				return CompileResult.Empty;
			return base.BuildResult(result);
		}
		#endregion
	}
}
