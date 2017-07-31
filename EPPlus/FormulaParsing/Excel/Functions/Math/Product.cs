/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
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
*  * Author							Change						Date
* *******************************************************************************
* * Mats Alm   		                Added		                2013-12-03
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for the Product Function in Excel.
	/// </summary>
	public class Product : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Calculates the produce of the user specified arguments. 
		/// </summary>
		/// <param name="arguments">The user specified arguments to multiply.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The product of the arguments given.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var result = 0d;
			var index = 0;
			double valAsDouble;
			foreach (var val in arguments)
			{
				if (val.Value is int || val.Value is double) { }
				else if (val.IsExcelRange) { }
				else if (val.Value is null) { }
				else if (ConvertUtil.TryParseNumericString(val.Value, out valAsDouble)) { }
				else if (ConvertUtil.TryParseObjectToDecimal(val.Value, out valAsDouble)) { }
				else if (val.Value is IEnumerable<FunctionArgument>) { }
				else
					return new CompileResult(eErrorType.Value);
			}
			try
			{
				while (AreEqual(result, 0d) && index < arguments.Count())
				{
					result = CalculateFirstItem(arguments, index++, context);
				}
				result = CalculateCollection(arguments.Skip(index), result, (arg, current) =>
				{
					if (ShouldIgnore(arg)) return current;
					if (arg.ValueIsExcelError)
					{
						throw new Exception(arg.ValueAsExcelErrorValue.ToString());
					}
					if (arg.IsExcelRange)
					{
						foreach (var cell in arg.ValueAsRangeInfo)
						{
							if (ShouldIgnore(cell, context)) return current;
							current *= cell.ValueDouble;
						}
						return current;
					}
					var obj = arg.Value;
					double argAsDouble;
					if (obj != null && IsNumeric(obj))
					{
						var val = Convert.ToDouble(obj);
						current *= val;
					}
					else if (ConvertUtil.TryParseNumericString(obj, out argAsDouble ))
					{
						return current * argAsDouble;
					}
					else if (ConvertUtil.TryParseObjectToDecimal(obj, out argAsDouble))
					{
						return current * argAsDouble;
					}
					return current;
				});
				return CreateResult(result, DataType.Decimal);
			}
			catch
			{
				return new CompileResult(eErrorType.Value);
			}
		}
		/// <summary>
		/// CalculateFirstItem checks the type of the first argument in the list and sets the first value for multiplication accordingly.
		/// </summary>
		/// <param name="arguments">The user specified list of arguments to be multiplied.</param>
		/// <param name="index">The index of the element to check.</param>
		/// <param name="context">The context in which the method is being called.</param>
		/// <returns>The first value in the arguments list as a double.</returns>
		private double CalculateFirstItem(IEnumerable<FunctionArgument> arguments, int index, ParsingContext context)
		{
			var element = arguments.ElementAt(index);
			var argList = new List<FunctionArgument> { element };
			double argAsDouble;
			if(ConvertUtil.TryParseNumericString(arguments.ElementAt(0).Value, out argAsDouble))
			{
				return argAsDouble;
			}
			else if (ConvertUtil.TryParseObjectToDecimal(arguments.ElementAt(0).Value, out argAsDouble))
			{
				return argAsDouble;
			}
			var valueList = ArgsToDoubleEnumerable(false, false, argList, context);
			var result = 0d;
			foreach (var value in valueList)
			{
				if (result == 0d && value != 0d)
				{
					result = value;
				}
				else
				{
					result *= value;
				}
			}
			return result;
		}
	}
}

