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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Returns the sequential serial number that represents a particular date.
	/// </summary>
	public class Date : ExcelFunction
	{
		/// <summary>
		/// Given a year, month, and day, return the Excel OADate representing that date. If the year value is 
		/// between 0 and 1899 inclusive, that value is added to 1900 and uses that sum as the year. If the year value is
		/// between 1900 and 9999 inclusive, that value is used as the year. If the year value is less than 0, an <see cref="ExcelErrorValue"/>
		/// is returned. If the month value is less than 1 or greater than 12, that month value is added from the first month of the specified
		/// year. If the day value is less than 1 or greater than the number of days in the specified month, that number of days is added
		/// to the first day of the specified month.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the Excel OADate.</param>
		/// <param name="context">Unused in the method, but necessary to override the method.</param>
		/// <returns>Returns the Excel OADate of the given date, or an <see cref="ExcelErrorValue"/> if the input is invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 3) == false)
				return new CompileResult(eErrorType.Value);
			var isValidYear = this.TryGetArgAsInt(arguments, 0, out int year);
			var isValidMonth = this.TryGetArgAsInt(arguments, 1, out int month);
			var isValidDay = this.TryGetArgAsInt(arguments, 2, out int day);
			if (!isValidYear || !isValidMonth || !isValidDay)
				return new CompileResult(eErrorType.Value);
			if (year < 0 || year > 9999)
				return new CompileResult(eErrorType.Num);
			else if (year < 1900)
				year += 1900;
			var resultDate = new System.DateTime(year, 1, 1);
			var dateWithMonthsAdded = resultDate;
			month--;
			day--;
			try
			{
				resultDate = resultDate.AddMonths(month);
				dateWithMonthsAdded = resultDate;
				resultDate = resultDate.AddDays(day);
			}
			catch (ArgumentOutOfRangeException e)
			{
				return new CompileResult(eErrorType.Num);
			}
			var resultOADate = resultDate.ToOADate();
			if (resultOADate < 1)
				return new CompileResult(eErrorType.Num);
			if (dateWithMonthsAdded.ToOADate() < 61 && resultOADate != 61)
				resultOADate--;
			return this.CreateResult(resultOADate, DataType.Date);
		}
	}
}
