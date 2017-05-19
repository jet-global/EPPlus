using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	public class Weeknum : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);

			var serialNumberCandidate = arguments.ElementAt(0).Value;
			if (serialNumberCandidate is string)
			{
				var isDateString = System.DateTime.TryParse(serialNumberCandidate.ToString(), out System.DateTime date1);

				if (!isDateString)
					return new CompileResult(eErrorType.Value);
			}
			else if (serialNumberCandidate is null)
				return new CompileResult(eErrorType.NA);

			var dateSerial = ArgToDecimal(arguments, 0);

			if(dateSerial < 0)
				return new CompileResult(eErrorType.Num);


			var date = System.DateTime.FromOADate(dateSerial);


			var startDay = DayOfWeek.Sunday;
			if (arguments.Count() > 1)
			{

				var returnType = arguments.ElementAt(1).Value;

				if (returnType is null)
					return new CompileResult(eErrorType.Num);
				else if(returnType is string)
				{
					var isValidReturnType = Int32.TryParse(returnType.ToString(), out int result);
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

			var week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay, startDay);
			if (serialNumberCandidate is int)
				if (dateSerial == 0)
					week = (int)dateSerial / 7;
			return CreateResult((int)week, DataType.Integer);
		}
	}
}
