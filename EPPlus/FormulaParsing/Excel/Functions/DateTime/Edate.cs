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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Return the Excel OADate for the date that is the indicated number of months before or after
	/// the given start date.
	/// </summary>
	public class Edate : ExcelFunction
	{
		/// <summary>
		/// Given a date represented as a string, int, double, or <see cref="System.DateTime"/> object and a number
		/// represented as a string, double, or int, return the Excel OADate for the date that is the given number
		/// of months before or after the given date.
		/// </summary>
		/// <param name="arguments">The given arguments used to calculate the new date.</param>
		/// <param name="context">Unused in the method, but necessary to override the method.</param>
		/// <returns>Returns the Excel OADate of the calculated date, or an <see cref="ExcelErrorValue"/> if either input is invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var dateObj = this.GetFirstValue(arguments);
			var monthsObj = arguments.ElementAt(1).Value;
			if (dateObj == null || monthsObj== null)
				return new CompileResult(eErrorType.NA);
			// Check for an input date of 0 or a fraction expressed as a string, which require special output.
			if (dateObj is string dateString && Double.TryParse(dateString, out double dateAsNumber))
			{
				if (dateAsNumber < 1 && dateAsNumber >= 0)
					dateObj = dateAsNumber;
			}
			var isZeroDate = ((dateObj is int dateInt && dateInt == 0) ||
				(dateObj is double dateDouble && dateDouble < 1 && dateDouble >= 0));
			if (ConvertUtil.TryParseDateObject(dateObj, out System.DateTime validDate, out eErrorType? error) || isZeroDate)
			{
				var resultDate = new System.DateTime(validDate.Year, validDate.Month, validDate.Day);
				if (isZeroDate)
					resultDate = new System.DateTime(1899, 12, 31);
				if (monthsObj is string monthsString)
				{
					if (Double.TryParse(monthsString, out double numberMonths))
						monthsObj = numberMonths;
					else
						return new CompileResult(eErrorType.Value);
				}
				var monthsToAdd = (monthsObj is double monthsDouble) ? (int)monthsDouble : (int)monthsObj;
				resultDate = resultDate.AddMonths(monthsToAdd);
				var serialNumber = resultDate.ToOADate();
				if (serialNumber < 61)
					serialNumber--;
				// The OADate 0 with 0 months added is a special case.
				if (serialNumber == 0 && isZeroDate)
					return this.CreateResult(0.0, DataType.Date);
				else if (serialNumber < 1)
					return new CompileResult(eErrorType.Num);
				else
					return this.CreateResult(serialNumber, DataType.Date);
			}
			else
				return new CompileResult(error.Value);
		}
	}
}
