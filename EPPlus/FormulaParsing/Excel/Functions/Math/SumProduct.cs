/* Copyright (C) 2011  Jan Källman
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
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// The class contains the formula for the SUMPRODUCT Excel Function. 
	/// </summary>
	public class SumProduct : ExcelFunction
	{
		/// <summary>
		/// Takes the user arguments multiplies the corresponding components in the given arguments and returns the sum of
		/// those products.
		/// </summary>
		/// <param name="arguments">The user specified arguments, usually arrays for this function.</param>
		/// <param name="context">The current context of the function.</param>
		/// <returns>The sum of the products fo the corresponding components in the given array.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			double result = 0d;
			List<List<double>> results = new List<List<double>>();
			foreach (var arg in arguments)
			{
				results.Add(new List<double>());
				var currentResult = results.Last();
				if (arg.Value is IEnumerable<FunctionArgument>)
				{
					foreach (var val in (IEnumerable<FunctionArgument>)arg.Value)
					{
						this.AddValue(val.Value, currentResult);
					}
				}
				else if (arg.Value is FunctionArgument)
					this.AddValue(arg.Value, currentResult);
				else if (arg.IsExcelRange)
				{
					var r = arg.ValueAsRangeInfo;
					for (int col = r.Address._fromCol; col <= r.Address._toCol; col++)
					{
						for (int row = r.Address._fromRow; row <= r.Address._toRow; row++)
						{
							if (r.GetValue(row, col) is bool)
								this.AddValue(0, currentResult);
							else
								this.AddValue(r.GetValue(row, col), currentResult);
						}
					}
				}
				else if (arg.Value is int || arg.Value is double || arg.Value is System.DateTime)
					this.AddValue(arg.Value, currentResult);
				else
					return new CompileResult(eErrorType.Value);
			}

			var arrayLength = results.First().Count;
			foreach (var list in results)
			{
				if (list.Count != arrayLength)
					throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
			}
			for (var rowIndex = 0; rowIndex < arrayLength; rowIndex++)
			{
				double rowResult = 1;
				for (var colIndex = 0; colIndex < results.Count; colIndex++)
				{
					rowResult *= results[colIndex][rowIndex];
				}
				result += rowResult;
			}
			return this.CreateResult(result, DataType.Decimal);
		}


		#region Private Methods
		/// <summary>
		/// Converts the given object to a double and then adds it to the given list.
		/// </summary>
		/// <param name="convertVal">The object to add to the list. </param>
		/// <param name="currentResult">The list the object will be added to.</param>
		private void AddValue(object convertVal, List<double> currentResult)
		{
			if (IsNumeric(convertVal))
				currentResult.Add(Convert.ToDouble(convertVal));
			else if (convertVal is ExcelErrorValue)
				throw (new ExcelErrorValueException((ExcelErrorValue)convertVal));
			else
				currentResult.Add(0d);
		}
		#endregion
	}
}
