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
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Represents a converter used to convert values to expressions.
	/// </summary>
	public class ExpressionConverter : IExpressionConverter
	{
		#region Properties
		private CompileResultFactory ResultFactory { get; } = new CompileResultFactory();
		#endregion

		#region Public Methods
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
				case DataType.Time:
				case DataType.Decimal:
					return compileResult.Result is string
								  ? new DecimalExpression(compileResult.Result.ToString())
								  : new DecimalExpression(Convert.ToDouble(compileResult.Result));
				case DataType.String:
					return new StringExpression(compileResult.Result.ToString());
				case DataType.Boolean:
					return compileResult.Result is string
								  ? new BooleanExpression(compileResult.Result.ToString())
								  : new BooleanExpression((bool)compileResult.Result);
				case DataType.Date:
					if (compileResult.Result is DateTime dateTimeResult || DateTime.TryParse(compileResult.Result.ToString(), out dateTimeResult))
						return new DateExpression(dateTimeResult.ToOADate().ToString());
					if (double.TryParse(compileResult.Result.ToString(), out double oaDate))
						return new DateExpression(oaDate.ToString());
					return new ExcelErrorExpression(ExcelErrorValue.Create(eErrorType.Value));
				case DataType.ExcelError:
					if (compileResult.Result is ExcelErrorValue errorValueResult)
						return new ExcelErrorExpression(errorValueResult);
					else if (compileResult.Result is eErrorType eErrorTypeResult)
						return new ExcelErrorExpression(ExcelErrorValue.Create(eErrorTypeResult));
					else
						return new ExcelErrorExpression(compileResult.Result?.ToString(), ExcelErrorValue.Parse(compileResult.Result?.ToString()));
				case DataType.Empty:
					return new IntegerExpression(0);
				case DataType.ExcelAddress:
					return new StringExpression(compileResult.Result.ToString());
				case DataType.Enumerable:
				case DataType.Unknown:
				default:
					// Enumerable results only end up with the first item in the collection.
					// The result factory will itself return an enumerable CompileResult for List<object> so 
					// in order to prevent infinite recursion there is an explicit check for that specific type.
					// The other form of enumerable result is IRangeInfo which is safely reduced in the result factory.
					var resultToProcess = compileResult.Result;
					if (resultToProcess is List<object> listResult)
						resultToProcess = listResult.FirstOrDefault();
					var result = this.ResultFactory.Create(resultToProcess);
					return this.FromCompileResult(result);
			}
		}
		#endregion
	}
}
