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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Returns the average of the given arguments.
	/// </summary>
	public class Average : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Returns the average of the given arguments.
		/// See https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6
		/// for the documentation on what the AVERAGE function accepts as valid arguments.
		/// </summary>
		/// <param name="arguments">The given arguments to average.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>Returns the average of the given arguments, or an <see cref="ExcelErrorValue"/> if any of the given arguments are invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError, eErrorType.Div0) == false)
				return new CompileResult(argumentError);
			double sumOfAllValues = 0d, numberOfValues = 0d;
			foreach (var argument in arguments)
			{
				var calculationError = this.CalculateComponentsOfAverage(argument, context, ref sumOfAllValues, ref numberOfValues);
				if (calculationError != null)
					return new CompileResult(calculationError.Value);
			}
			return this.CreateResult(this.Divide(sumOfAllValues, numberOfValues), DataType.Decimal);
		}

		/// <summary>
		/// Recursively calculates the sum of all valid cells and the total number of valid cells for the given argument.
		/// </summary>
		/// <param name="argument">The argument to check for a valid numeric value.</param>
		/// <param name="context">The context for this method.</param>
		/// <param name="sumOfAllValues">The running sum of all numeric values processed.</param>
		/// <param name="numberOfValues">The running total number of numeric values processed.</param>
		/// <param name="isInArray">Indicates whether <paramref name="argument"/> is from an array.</param>
		/// <returns>Returns an <see cref="eErrorType"/> if <paramref name="argument"/> is invalid, otherwise returns null.</returns>
		private eErrorType? CalculateComponentsOfAverage(FunctionArgument argument, ParsingContext context, ref double sumOfAllValues, ref double numberOfValues, bool isInArray = false)
		{
			if (argument.Value == null)
				argument = new FunctionArgument(0);
			if (this.ShouldIgnore(argument))
				return null;
			if (argument.Value is IEnumerable<FunctionArgument>)
			{
				foreach (var subArgument in (IEnumerable<FunctionArgument>)argument.Value)
				{
					var errorFromSubArgument = CalculateComponentsOfAverage(subArgument, context, ref sumOfAllValues, ref numberOfValues, true);
					if (errorFromSubArgument != null)
						return errorFromSubArgument;
				}
			}
			else if (argument.IsExcelRange)
			{
				foreach (var cellInfo in argument.ValueAsRangeInfo)
				{
					bool handleAsFormula = (!cellInfo.Formula.Equals(string.Empty));
					if (this.ShouldIgnore(cellInfo, context))
						continue;
					this.CheckForAndHandleExcelError(cellInfo);
					if (this.IsNumeric(cellInfo.Value) && !(cellInfo.Value is bool))
					{
						numberOfValues++;
						sumOfAllValues += cellInfo.ValueDouble;
					}
					else if (!handleAsFormula && cellInfo.Value is string cellAsString && 
						Double.TryParse(cellAsString, out double parsedValue))
					{
						numberOfValues++;
						sumOfAllValues += parsedValue;
					}
				}
			}
			else
			{
				var numericValue = this.GetNumericValue(argument.Value, isInArray);
				if (numericValue.HasValue)
				{
					numberOfValues++;
					sumOfAllValues += numericValue.Value;
				}
				else if (argument.Value is string && !isInArray)
				{
					return eErrorType.Value;
				}
			}
			this.CheckForAndHandleExcelError(argument);
			return null;
		}

		/// <summary>
		/// Converts the given object to its numeric value if possible.
		/// </summary>
		/// <param name="numberCandidate">The given object to convert to a double.</param>
		/// <param name="isInArray">Indicates whether <paramref name="numberCandidate"/> should be handled as if it came from an array.</param>
		/// <returns>Return the double value represented by <paramref name="numberCandidate"/>, 
		///			 otherwise returns the default value for a nullable double.</returns>
		private double? GetNumericValue(object numberCandidate, bool isInArray)
		{
			if (this.IsNumeric(numberCandidate) && !(numberCandidate is bool))
				return ConvertUtil.GetValueDouble(numberCandidate);
			else if (!isInArray)
			{
				if (numberCandidate is bool)
					return ConvertUtil.GetValueDouble(numberCandidate);
				else if (ConvertUtil.TryParseNumericString(numberCandidate, out double number))
					return number;
				else if (ConvertUtil.TryParseDateString(numberCandidate, out System.DateTime date))
					return date.ToOADate();
			}
			return default(double?);
		}
	}
}
