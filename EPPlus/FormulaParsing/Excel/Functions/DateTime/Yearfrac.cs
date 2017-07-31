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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Return the fraction of the year represented by the number of whole days between two dates.
	/// </summary>
	public class Yearfrac : ExcelFunction
	{
		/// <summary>
		/// Given two dates, represented as a string, int, double, or <see cref="System.DateTime"/>,
		/// and an optional type of day count basis to use, return the fraction of the year beteen those two dates.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the fraction of the year.</param>
		/// <param name="context">Used to determine the context for the YEARFRAC function.</param>
		/// <returns>Returns the fraction of the year between the two given dates as a double, or an <see cref="ExcelErrorValue"/> if the input is invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentsAreValid(functionArguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var basis = 0;
			if (functionArguments.Count() > 2)
			{
				if (!this.TryGetArgAsInt(functionArguments, 2, out basis))
					return new CompileResult(eErrorType.Value);
				else if (basis < 0 || basis > 4)
					return new CompileResult(eErrorType.Num);
			}
			var date1Candidate = functionArguments[0].Value;
			var date2Candidate = functionArguments[1].Value;
			if (date1Candidate == null || date2Candidate == null)
				return new CompileResult(eErrorType.NA);
			var date1IsValid = ConvertUtil.TryParseObjectToDecimal(date1Candidate, out double date1Num);
			var date2IsValid = ConvertUtil.TryParseObjectToDecimal(date2Candidate, out double date2Num);
			if (!date1IsValid)
				return new CompileResult(eErrorType.Value);
			else if (date1Num < 0)
				return new CompileResult(eErrorType.Num);
			else if (!date2IsValid)
				return new CompileResult(eErrorType.Value);
			else if (date2Num < 0)
				return new CompileResult(eErrorType.Num);
			// The date1Num and date2Num provided by TryParseObjectToDecimal are Excel OADates;
			// they need to be converted back to System.DateTime OADates for the special case
			// of dates before 3/1/1900 (OADate 61 in both Excel and System.DateTime).
			if (date1Num < 61)
				date1Num++;
			if (date2Num < 61)
				date2Num++;
			if (date1Num > date2Num) // Ensure that date1Num always represents the earliest date.
			{
				var t = date1Num;
				date1Num = date2Num;
				date2Num = t;
			}
			var startDate = System.DateTime.FromOADate(date1Num);
			var endDate = System.DateTime.FromOADate(date2Num);
			var days360Function = context.Configuration.FunctionRepository.GetFunction("days360");
			var daysBetween = (endDate - startDate).TotalDays;
			// Account for the fact that Excel includes an extra day, 2/29/1900 which doesn't actually exist, 
			// between 2/28/1900 and 3/1/1900, which are represented by System.DateTime OADates 60 and 61 respectively.
			if (startDate.ToOADate() < 61 && endDate.ToOADate() >= 61)
				daysBetween++;
			var yearFracResult = 0d;
			// Note: See https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8
			// for more information on the options for the basis parameter for the YEARFRAC function.
			switch (basis)
			{
				case 0: // Use the US 30-day-per-month/360-day-per-year as the Day Count Basis.
					var days360Result = System.Math.Abs(days360Function.Execute(functionArguments, context).ResultNumeric);
					// Reproducing Excel's behavior.
					if (startDate.Month == 2 && endDate.Day == 31)
					{
						var calendar = new GregorianCalendar();
						var daysInFeb = calendar.IsLeapYear(startDate.Year) ? 29 : 28;
						if (startDate.Day == daysInFeb)
							days360Result++;
					}
					yearFracResult = days360Result / 360d;
					break;
				case 1: // Use the actual number of days between dates/actual number of days per year as the Day Count Basis.
					yearFracResult = System.Math.Abs(daysBetween / this.CalculateAverageDaysPerYear(startDate, endDate));
					break;
				case 2: // Use the actual number of days between dates/360-days-per-year as the Day Count Basis
					yearFracResult = System.Math.Abs(daysBetween / 360d);
					break;
				case 3: // Use the actual number of days between dates/365-days-per-year as the Day Count Basis
					yearFracResult = System.Math.Abs(daysBetween / 365d);
					break;
				case 4: // Use the European 30-days-per-month/360-days-per-year as the Day Count Basis.
					var args = functionArguments.ToList();
					args.Add(new FunctionArgument(true));
					double? result = System.Math.Abs(days360Function.Execute(args, context).ResultNumeric / 360d);
					yearFracResult = result.Value;
					break;
				default:
					return new CompileResult(eErrorType.Num);
			}
			yearFracResult = System.Math.Round(yearFracResult, 11);
			return this.CreateResult(yearFracResult, DataType.Decimal);
		}

		/// <summary>
		/// Calculates the average number of days in a year for the years between two given dates.
		/// </summary>
		/// <param name="startDate">The starting date for calculating the actual year value.</param>
		/// <param name="endDate">The ending date for calculating the actual year value.</param>
		/// <returns>Returns the average number of days in a year based on the time between <paramref name="startDate"/> and <paramref name="endDate"/>.</returns>
		private double CalculateAverageDaysPerYear(System.DateTime startDate, System.DateTime endDate)
		{
			var calendar = new GregorianCalendar();
			var totalDaysPerYear = 0d;
			var numberOfYears = endDate.Year - startDate.Year + 1;
			if (new System.DateTime(startDate.Year + 1, startDate.Month, startDate.Day) >= endDate) // Check if the start and end dates are a year or less apart.
			{
				numberOfYears = 1;
				totalDaysPerYear = 365;
				if (calendar.IsLeapYear(startDate.Year) && startDate.Month <= 2)
					totalDaysPerYear = 366;
				else if (calendar.IsLeapYear(endDate.Year) && endDate.Month > 2)
					totalDaysPerYear = 366;
				else if (endDate.Month == 2 && endDate.Day == 29)
					totalDaysPerYear = 366;
			}
			else
			{
				for (var y = startDate.Year; y <= endDate.Year; ++y)
				{
					totalDaysPerYear += calendar.IsLeapYear(y) ? 366 : 365;
				}
			}
			return totalDaysPerYear / numberOfYears;
		}
	}
}
