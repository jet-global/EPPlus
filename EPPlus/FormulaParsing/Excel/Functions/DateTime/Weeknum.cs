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
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing the weeknumber based on a specified date. 
	/// </summary>
	public class Weeknum : ExcelFunction
	{
		/// <summary>
		/// Execute returns the week number based on the user's input.
		/// </summary>
		/// <param name="arguments">The user specified date and optional return type.</param>
		/// <param name="context">Not used, but needed for overriding the method.</param>
		/// <returns>The week number out of 52 based on the date given.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var serialNumberCandidate = arguments.ElementAt(0).Value;
			if (serialNumberCandidate == null)
				return new CompileResult(eErrorType.NA);
			var returnType = 1;
			if (arguments.Count() > 1)
			{
				if (arguments.ElementAt(1).Value == null)
					return new CompileResult(eErrorType.Num);
				else if (!this.TryGetArgAsInt(arguments, 1, out returnType))
					return new CompileResult(eErrorType.Value);
			}
			// Excel treats 0 and fractional dates as special dates and require special output.
			if (ConvertUtil.TryParseObjectToDecimal(serialNumberCandidate, out double serialNumber) && serialNumber < 1 && serialNumber >= 0)
				return this.getZeroWeek(returnType);
			var isValidDate = ConvertUtil.TryParseDateObject(serialNumberCandidate, out System.DateTime date, out eErrorType? error);
			if (!isValidDate)
				return new CompileResult(error.Value);
			var calendarType = CalendarWeekRule.FirstDay;
			var startDay = DayOfWeek.Sunday;
			switch (returnType)
			{
				case 1:
					startDay = DayOfWeek.Sunday;
					break;
				case 2:
					startDay = DayOfWeek.Monday;
					break;
				case 11:
					startDay = DayOfWeek.Monday;
					break;
				case 12:
					startDay = DayOfWeek.Tuesday;
					break;
				case 13:
					startDay = DayOfWeek.Wednesday;
					break;
				case 14:
					startDay = DayOfWeek.Thursday;
					break;
				case 15:
					startDay = DayOfWeek.Friday;
					break;
				case 16:
					startDay = DayOfWeek.Saturday;
					break;
				case 17:
					startDay = DayOfWeek.Sunday;
					break;
				case 21:
					startDay = DayOfWeek.Thursday;
					calendarType = CalendarWeekRule.FirstFullWeek;
					break;
				default:
					return new CompileResult(eErrorType.Num);
			}
			if (DateTimeFormatInfo.CurrentInfo == null)
			{
				throw new InvalidOperationException(
					 "Could not execute Weeknum function because DateTimeFormatInfo.CurrentInfo was null");
			}
			var week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(date, calendarType, startDay);
			return this.CreateResult(week, DataType.Integer);
		}

		/// <summary>
		/// This method handles the output for the special case where the OADate given for the
		/// WEEKNUM function is zero or a fraction.
		/// </summary>
		/// <param name="returnType">The value determining on which day the week begins.</param>
		/// <returns>Returns the <see cref="CompileResult"/> for handling the date-zero case.</returns>
		private CompileResult getZeroWeek(int returnType)
		{
			var zeroWeek = 0;
			switch (returnType)
			{
				case 1:
				case 17:
					zeroWeek = 0;
					break;
				case 2:
				case 11:
				case 12:
				case 13:
				case 14:
				case 15:
				case 16:
					zeroWeek = 1;
					break;
				case 21:
					zeroWeek = 52;
					break;
				default:
					 return new CompileResult(eErrorType.Num);
			}
			return this.CreateResult(zeroWeek, DataType.Integer);
		}
	}
}
