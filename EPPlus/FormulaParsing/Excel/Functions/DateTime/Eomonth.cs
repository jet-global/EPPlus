using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Represents an excel function that gets the end of the month based on a date and a specified month offset. 
	/// </summary>
	public class Eomonth : ExcelFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the EOMONTH function.
		/// </summary>
		/// <param name="arguments">A date object and a number of months to offset from that date.</param>
		/// <param name="context">Unused; an object that contains information about the context in which the function is being run.</param>
		/// <returns>An OADate representing the end of the month the specified number of months before or after the specified date.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var firstArgument = arguments.ElementAt(0).Value;
			var secondArgument = arguments.ElementAt(1).Value;
			System.DateTime dateTime;
			eErrorType? error;

			if (firstArgument == null || secondArgument == null)
				return new CompileResult(eErrorType.NA);
			var date = new System.DateTime(2017,5,6);
			if (firstArgument is int)
			{
				var temp = this.ArgToDecimal(arguments, 0);
				if (temp < 0)
					return new CompileResult(eErrorType.Num);
				if (temp == 0)
					temp = 2;
				date = System.DateTime.FromOADate(temp);
			}
			else if (ConvertUtil.TryParseDateObject(firstArgument, out dateTime, out error))
				date = System.DateTime.FromOADate(dateTime.ToOADate());
			else
				return new CompileResult(eErrorType.Value);

			var monthsToAdd = 0;
			if (secondArgument is double)
			{
				var temp = this.ArgToDecimal(arguments, 1);
				monthsToAdd = (int)temp;
			}
			else if(secondArgument is string)
			{
				if (ConvertUtil.TryParseNumericString(secondArgument, out double result))
					monthsToAdd = (int)result;
				else if (ConvertUtil.TryParseDateString(secondArgument.ToString(), out dateTime))
				{
					//This accounts for the Lotus 1-2-3 error with dates before March 1, 1900.
					if (dateTime.ToOADate() < 61)
					{
						dateTime = System.DateTime.FromOADate(dateTime.ToOADate() - 1);
					}
					monthsToAdd = (int)dateTime.ToOADate();
				}
				else
					return new CompileResult(eErrorType.Value);
			}
			else if (secondArgument is System.DateTime)
			{
				var temp = ConvertUtil.TryParseDateObject(secondArgument, out dateTime, out error);
				//This accounts for the Lotus 1-2-3 error with dates before March 1, 1900.
				if (dateTime.ToOADate() < 61)
				{
					dateTime = System.DateTime.FromOADate(dateTime.ToOADate() - 1);
				}
				monthsToAdd = (int)dateTime.ToOADate();
			}
			else
				monthsToAdd = ArgToInt(arguments, 1);

			if (date.ToOADate() < 60 && monthsToAdd < 0)
				return new CompileResult(eErrorType.Num);

			var resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-1);
			//This accounts for the Lotus 1-2-3 error with dates before March 1, 1900.
			if (date.ToOADate() <= 60)
			{
				resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-2);
			}

			return CreateResult(resultDate.ToOADate(), DataType.Date);
		}

		#endregion
	}
}
