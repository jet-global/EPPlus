﻿using System;
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
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);

			var serialNumberCandidate = arguments.ElementAt(0).Value;

			if (serialNumberCandidate is null)
				return new CompileResult(eErrorType.NA);
			else if (serialNumberCandidate is string)
			{
				var isDateString = ConvertUtil.TryParseDateString(serialNumberCandidate.ToString(), out System.DateTime date1);
				if (!isDateString)
					return new CompileResult(eErrorType.Value);
			}

			var dateSerial = ArgToDecimal(arguments, 0);

			if(dateSerial < 0)
				return new CompileResult(eErrorType.Num);

			var date = System.DateTime.FromOADate(dateSerial);

			var calendarType = CalendarWeekRule.FirstDay;

			var startDay = DayOfWeek.Sunday;
			if (arguments.Count() > 1)
			{
				var returnType = arguments.ElementAt(1).Value;

				if (returnType is null)
					return new CompileResult(eErrorType.Num);
				else if(returnType is string)
				{
					var isValidReturnType = ConvertUtil.TryParseNumericString(returnType.ToString(), out double date2);
					if (!isValidReturnType)
						return new CompileResult(eErrorType.Value);
				}

				var argStartDay = ArgToInt(arguments, 1);
				switch (argStartDay)
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
						// Not supported return type
						return new CompileResult(eErrorType.Num);
				}
			}

			if (DateTimeFormatInfo.CurrentInfo == null)
			{
				throw new InvalidOperationException(
					 "Could not execute Weeknum function because DateTimeFormatInfo.CurrentInfo was null");
			}

			var week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(date, calendarType, startDay);
			if (serialNumberCandidate is int)
				if (dateSerial == 0)
					week = 0;
			return this.CreateResult(week, DataType.Integer);
		}
	}
}
