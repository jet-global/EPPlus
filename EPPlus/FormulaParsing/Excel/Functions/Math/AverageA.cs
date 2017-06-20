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
 * Mats Alm   		                Added		                2014-01-06
 *******************************************************************************/
using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class AverageA : HiddenValuesHandlingFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError, eErrorType.Div0) == false)
				return new CompileResult(argumentError);
			double sumOfAllValues = 0d, numberOfValues = 0d;
			foreach (var argument in arguments)
			{
				var calculationError = this.CalculateComponentsOfAverageA(argument, context, ref sumOfAllValues, ref numberOfValues);
				if (calculationError != null)
					return new CompileResult(calculationError.Value);
			}
			return this.CreateResult(this.Divide(sumOfAllValues, numberOfValues), DataType.Decimal);
		}

		private eErrorType? CalculateComponentsOfAverageA(FunctionArgument argument, ParsingContext context, ref double sumOfAllValues, ref double numberOfValues, bool isInArray = false)
		{
			if (argument.Value == null)
				argument = new FunctionArgument(0);
			if (this.ShouldIgnore(argument))
			{
				return null;
			}
			if (argument.Value is IEnumerable<FunctionArgument>)
			{
				foreach (var subArgument in (IEnumerable<FunctionArgument>)argument.Value)
				{
					var calculationError = this.CalculateComponentsOfAverageA(subArgument, context, ref sumOfAllValues, ref numberOfValues, true);
					if (calculationError != null)
						return calculationError;
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
					else if (cellInfo.Value is bool)
					{
						numberOfValues++;
						sumOfAllValues += (bool)cellInfo.Value ? 1 : 0;
					}
					else if (cellInfo.Value is string cellValueAsString)
					{
						bool cellIsNull = cellValueAsString.Equals(string.Empty);
						numberOfValues += (cellIsNull) ? 0 : 1;
						if (handleAsFormula || cellIsNull)
							continue;
						if (Boolean.TryParse(cellValueAsString, out bool cellValueAsBool))
							sumOfAllValues += (cellValueAsBool) ? 1 : 0;
						else if (Double.TryParse(cellValueAsString, out double cellValueAsDouble))
							sumOfAllValues += cellValueAsDouble;
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
				else if ((argument.Value is string))
				{
					if (isInArray)
					{
						numberOfValues++;
					}
					else
					{
						return eErrorType.Value;
					}
				}
			}
			this.CheckForAndHandleExcelError(argument);
			return null;
		}

		private double? GetNumericValue(object obj, bool isInArray)
		{
			double number;
			System.DateTime date;
			if (IsNumeric(obj) && !(obj is bool))
			{
				return ConvertUtil.GetValueDouble(obj);
			}
			if (!isInArray)
			{
				if (obj is bool)
				{
					if (isInArray) return default(double?);
					return ConvertUtil.GetValueDouble(obj);
				}
				else if (ConvertUtil.TryParseNumericString(obj, out number))
				{
					return number;
				}
				else if (ConvertUtil.TryParseDateString(obj, out date))
				{
					return date.ToOADate();
				}
			}
			return default(double?);
		}
	}
}
