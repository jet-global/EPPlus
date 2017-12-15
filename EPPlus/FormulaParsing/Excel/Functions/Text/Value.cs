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
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
	/// <summary>
	/// Represents the Excel VALUE function.
	/// </summary>
	public class Value : ExcelFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Converts a text string that represents a number into a number.
		/// </summary>
		/// <param name="arguments">The arguments containing text number to convert to a number.</param>
		/// <param name="context">The context in which to evaluate the function.</param>
		/// <returns>A <see cref="CompileResult"/> containing a number if the argument was valid, otherwise an error <see cref="CompileResult"/>.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var value = ArgToString(arguments, 0).TrimEnd(' ');
			if (string.IsNullOrEmpty(value))
				return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
			double result = 0d;
			var groupSeparator = Regex.Escape(CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator);
			var decimalSeparator = Regex.Escape(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
			if (Regex.IsMatch(value, $"^[\\d]*({groupSeparator}?[\\d]*)?({decimalSeparator}[\\d]*)?[ ?% ?]?$"))
			{
				if (value.EndsWith("%"))
				{
					value = value.Remove(value.Length - 1, 1);
					result = double.Parse(value) / 100;
				}
				else
					result = double.Parse(value);
				return CreateResult(result, DataType.Decimal);
			}
			if (double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out result))
				return CreateResult(result, DataType.Decimal);
			var timeSeparator = Regex.Escape(CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator);
			if (Regex.IsMatch(value, @"^[\d]{1,2}" + timeSeparator + @"[\d]{2}(" + timeSeparator + @"[\d]{2})?$"))
			{
				var timeResult = new TimeValue().Execute(value);
				if (timeResult.DataType != DataType.ExcelError)
					return timeResult;
			}
			var dateResult = new DateValue().Execute(value);
			if (dateResult.DataType == DataType.Date)
				return dateResult;
			return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
		}
		#endregion
	}
}
