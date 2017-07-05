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
using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Represents a converter used to convert values to expressions.
	/// </summary>
	public class ExpressionConverter : IExpressionConverter
	{
		#region Class Variables
		private static IExpressionConverter myInstance;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the current instance of the ExpressionConverter.
		/// </summary>
		public static IExpressionConverter Instance
		{
			get
			{
				if (myInstance == null)
				{
					myInstance = new ExpressionConverter();
				}
				return myInstance;
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Converts the given <see cref="Expression"/> into a <see cref="StringExpression"/>.
		/// </summary>
		/// <param name="expression">The <see cref="Expression"/> to convert.</param>
		/// <returns>Returns the <see cref="StringExpression"/> representation of the given <see cref="Expression"/>.</returns>
		public StringExpression ToStringExpression(Expression expression)
		{
			var result = expression.Compile();
			return new StringExpression(result.Result.ToString()) { Operator = expression.Operator};
		}

		/// <summary>
		/// Converts the given <see cref="CompileResult"/> into an <see cref="Expression"/>.
		/// </summary>
		/// <param name="compileResult">The <see cref="CompileResult"/> to convert.</param>
		/// <returns>Returns the <see cref="Expression"/> representation of the given <see cref="CompileResult"/>.</returns>
		public Expression FromCompileResult(CompileResult compileResult)
		{
			switch (compileResult.DataType)
			{
				case DataType.Integer:
					return compileResult.Result is string
						 ? new IntegerExpression(compileResult.Result.ToString())
						 : new IntegerExpression(Convert.ToDouble(compileResult.Result));
				case DataType.String:
					return new StringExpression(compileResult.Result.ToString());
				case DataType.Decimal:
					return compileResult.Result is string
								  ? new DecimalExpression(compileResult.Result.ToString())
								  : new DecimalExpression(Convert.ToDouble(compileResult.Result));
				case DataType.Boolean:
					return compileResult.Result is string
								  ? new BooleanExpression(compileResult.Result.ToString())
								  : new BooleanExpression((bool)compileResult.Result);
				case DataType.ExcelError:
					return compileResult.Result is string
						 ? new ExcelErrorExpression(compileResult.Result.ToString(),
							  ExcelErrorValue.Parse(compileResult.Result.ToString()))
						 : new ExcelErrorExpression((ExcelErrorValue)compileResult.Result);
				case DataType.Empty:
					return new IntegerExpression(0);
			}
			return null;
		}
		#endregion
	}
}
