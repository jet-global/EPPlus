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
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing the number of dates between two given dates. 
	/// </summary>
	public class Days360 : ExcelFunction
	{
		private enum Days360Calctype
		{
			European,
			Us
		}
		#region Public ExcelFunction Overrides
		/// <summary>
		/// Execute returns the number of days between two user specified dates using the US or European date standards.
		/// </summary>
		/// <param name="arguments">The user specified dates and method of calendar they want to use. </param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var startDateObj = arguments.ElementAt(0).Value;
			var endDateObj = arguments.ElementAt(1).Value;
			if (startDateObj == null)
				startDateObj = 0;
			if (endDateObj == null)
				endDateObj = 0;
			if (!ConvertUtil.TryParseObjectToDecimal(startDateObj, out double startOADate) ||
				!ConvertUtil.TryParseObjectToDecimal(endDateObj, out double endOADate))
				return new CompileResult(eErrorType.Value);
			if (startOADate < 0 || endOADate < 0)
				return new CompileResult(eErrorType.Num);
			// The startOADate and endOADate provided by TryParseObjectToDecimal are Excel OADates;
			// they need to be converted back to System.DateTime OADates for the special case
			// of dates before 3/1/1900 (OADate 61 in both Excel and System.DateTime).
			if (startOADate < 61 && startOADate > 0)
				startOADate++;
			if (endOADate < 61 && endOADate > 0)
				endOADate++;
			var startDate = System.DateTime.FromOADate(startOADate);
			var endDate = System.DateTime.FromOADate(endOADate);
			var calcType = Days360Calctype.Us;
			if (arguments.Count() > 2)
			{
				var european = false;
				if (arguments.ElementAt(2).Value is string stringVal)
				{
					if (!stringVal.IsEquivalentTo("true") && !stringVal.IsEquivalentTo("false"))
						return new CompileResult(eErrorType.Value);
				}
				if (arguments.ElementAt(2).Value is string && ConvertUtil.TryParseNumericString(arguments.ElementAt(2).Value, out double val))
					return new CompileResult(eErrorType.Value);
				if(arguments.ElementAt(2).Value is System.DateTime)
				{
					european = true;
				}
				else
				{
					european = this.ArgToBool(arguments.ElementAt(2));
					if (arguments.ElementAt(2).Value is int intval && intval > 0)
						european = true;
				}
				if (european)
					calcType = Days360Calctype.European;
			}

			var startYear = startDate.Year;
			var startMonth = startDate.Month;
			var startDay = startDate.Day;
			var endYear = endDate.Year;
			var endMonth = endDate.Month;
			var endDay = endDate.Day;

			if (calcType == Days360Calctype.European)
			{
				if (startDay == 31)
					startDay = 30;
				if (endDay == 31)
					endDay = 30;
			}
			else
			{
				var calendar = new GregorianCalendar();
				var nDaysInFeb = calendar.IsLeapYear(startDate.Year) ? 29 : 28;
				// The OADate 0 in Excel is treated as the special date 1/0/1900, which requires specific handling.
				var isZeroCase = ((startOADate == 0 || endOADate == 0) && startOADate < endOADate);

				// If the investment is EOM and (Date1 is the last day of February) and (Date2 is the last day of February), then change D2 to 30.
				if (startMonth == 2 && startDay == nDaysInFeb && endMonth == 2 && endDay == nDaysInFeb)
				{
					endDay = (isZeroCase) ? 31 : 30;
				}
				// If the investment is EOM and (Date1 is the last day of February), then change D1 to 30.
				if (startMonth == 2 && startDay == nDaysInFeb)
				{
					startDay = (isZeroCase) ? 31 : 30;
				}
				// If D2 is 31 and D1 is 30 or 31, then change D2 to 30 unless first argument is null (because Excel adds a day in this case).
				if (endDay == 31 && (startDay == 30 || startDay == 31) && arguments.First().Value != null)
				{
					endDay = (isZeroCase) ? 31 : 30;
				}
				// If D1 is 31, then change D1 to 30.
				if (startDay == 31)
				{
					startDay = (isZeroCase) ? 31 : 30;
				}
			}
			var result = (endYear * 12 * 30 + endMonth * 30 + endDay) - (startYear * 12 * 30 + startMonth * 30 + startDay);
			return this.CreateResult(result, DataType.Integer);
		}
		#endregion
	}
}
