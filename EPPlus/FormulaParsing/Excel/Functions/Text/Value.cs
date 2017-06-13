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
	public class Value : ExcelFunction
	{
		private readonly string _groupSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
		private readonly string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
		private readonly string _timeSeparator = CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator;
		private readonly string _shortTimePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern;
		private readonly DateValue _dateValueFunc = new DateValue();
		private readonly TimeValue _timeValueFunc = new TimeValue();

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var val = ArgToString(arguments, 0).TrimEnd(' ');
			double result = 0d;
			if (Regex.IsMatch(val, $"^[\\d]*({Regex.Escape(_groupSeparator)}?[\\d]*)?({Regex.Escape(_decimalSeparator)}[\\d]*)?[ ?% ?]?$"))
			{
				if (val.EndsWith("%"))
				{
					val = val.TrimEnd('%');
					result = double.Parse(val) / 100;
				}
				else
				{
					result = double.Parse(val);
				}
				return CreateResult(result, DataType.Decimal);
			}
			if (double.TryParse(val, NumberStyles.Float, CultureInfo.CurrentCulture, out result))
			{
				return CreateResult(result, DataType.Decimal);
			}
			var timeSeparator = Regex.Escape(_timeSeparator);
			if (Regex.IsMatch(val, @"^[\d]{1,2}" + timeSeparator + @"[\d]{2}(" + timeSeparator + @"[\d]{2})?$"))
			{
				var timeResult = _timeValueFunc.Execute(val);
				if (timeResult.DataType != DataType.ExcelError)
				{
					return timeResult;
				}
			}
			var dateResult = _dateValueFunc.Execute(val);
			if (dateResult.DataType == DataType.Date)
			{
				return dateResult;
			}
			return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
		}
	}
}
