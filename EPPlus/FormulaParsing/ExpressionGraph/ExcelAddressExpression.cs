/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Represents an expression that defines a workbook sheet reference.
	/// </summary>
	public class ExcelAddressExpression : AtomicExpression
	{
		#region Class Variables
		protected readonly ExcelDataProvider _excelDataProvider;
		protected readonly ParsingContext _parsingContext;
		private readonly RangeAddressFactory _rangeAddressFactory;
		private readonly bool _negate;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets a value that indicates whether or not to resolve directly to an <see cref="ExcelDataProvider.IRangeInfo"/>
		/// </summary>
		public bool ResolveAsRange { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an <see cref="ExcelAddressExpression"/>.
		/// </summary>
		/// <param name="expression">The expression string.</param>
		/// <param name="excelDataProvider">An <see cref="ExcelDataProvider"/> for resolving addresses.</param>
		/// <param name="parsingContext">The current <see cref="ParsingContext"/>.</param>
		public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
			 : this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider), false)
		{

		}

		/// <summary>
		/// Creates an instance of an <see cref="ExcelAddressExpression"/>.
		/// </summary>
		/// <param name="expression">The expression string.</param>
		/// <param name="excelDataProvider">An <see cref="ExcelDataProvider"/> for resolving addresses.</param>
		/// <param name="parsingContext">The current <see cref="ParsingContext"/>.</param>
		/// <param name="negate">A value indicating whether or not to negate the expression.</param>
		public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, bool negate)
			 : this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider), negate)
		{

		}

		private ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, RangeAddressFactory rangeAddressFactory, bool negate)
			 : base(expression)
		{
			Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
			Require.That(parsingContext).Named("parsingContext").IsNotNull();
			Require.That(rangeAddressFactory).Named("rangeAddressFactory").IsNotNull();
			_excelDataProvider = excelDataProvider;
			_parsingContext = parsingContext;
			_rangeAddressFactory = rangeAddressFactory;
			_negate = negate;
		}
		#endregion

		#region AtomicExpression Overrides
		/// <summary>
		/// Gets a value indicating whether or not this is a grouped expression.
		/// </summary>
		public override bool IsGroupedExpression
		{
			get { return false; }
		}

		/// <summary>
		/// Compiles the expression into a value.
		/// </summary>
		/// <returns>The <see cref="CompileResult"/> with the expression value.</returns>
		public override CompileResult Compile()
		{
			var c = this._parsingContext.Scopes.Current;
			var result = _excelDataProvider.GetRange(c.Address.Worksheet, c.Address.FromRow, c.Address.FromCol, this.ExpressionString);
			if (result == null)
			{
				var excelAddress = new ExcelAddress(this.ExpressionString);
				// External references are not supported.
				if (!string.IsNullOrEmpty(excelAddress?.Workbook))
					return new CompileResult(eErrorType.Ref);
				else
					return CompileResult.Empty;
			}
			return this.BuildResult(result);
		}
		#endregion

		#region Protected Methods
		/// <summary>
		/// Builds a <see cref="CompileResult"/> based off the result and the expression configuration.
		/// </summary>
		/// <param name="result">The <see cref="ExcelDataProvider.IRangeInfo"/> result.</param>
		/// <returns>The processed result.</returns>
		protected CompileResult BuildResult(ExcelDataProvider.IRangeInfo result)
		{
			if (this.ResolveAsRange || result.Address.Rows > 1 || result.Address.Columns > 1)
				return new CompileResult(result, DataType.Enumerable);
			return CompileSingleCell(result);
		}
		#endregion

		#region Private Methods
		private CompileResult CompileSingleCell(ExcelDataProvider.IRangeInfo result)
		{
			if (result.Address.Address == ExcelErrorValue.Values.Ref)
				return new CompileResult(eErrorType.Ref);
			var cell = result.FirstOrDefault();
			if (cell == null)
				return CompileResult.Empty;
			var factory = new CompileResultFactory();
			var compileResult = factory.Create(cell.Value);
			if (_negate)
			{
				if (compileResult.IsNumeric)
					compileResult = new CompileResult(compileResult.ResultNumeric * -1, compileResult.DataType);
				else if (compileResult.DataType == DataType.String)
				{
					if (compileResult.ResultValue is string resultString && double.TryParse(resultString, out double resultDouble))
						compileResult = new CompileResult((resultDouble * -1).ToString(), compileResult.DataType);
					else
						compileResult = new CompileResult(eErrorType.Value);
				}
			}
			compileResult.IsHiddenCell = cell.IsHiddenRow;
			return compileResult;
		}
		#endregion
	}
}
