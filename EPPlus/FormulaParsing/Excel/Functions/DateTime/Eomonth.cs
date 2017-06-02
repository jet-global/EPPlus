using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

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
			if (this.ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			System.DateTime date;
			double result;
			var monthsToAdd = 0;
			var dateArgument = arguments.ElementAt(0).Value;
			var monthsToAddArgument = arguments.ElementAt(1).Value;

			if (dateArgument == null || monthsToAddArgument == null)
				return new CompileResult(eErrorType.NA);
			if(ConvertUtil.TryParseDateObjectToOADate(dateArgument, out result))
			{
				if (result < 0)
					return new CompileResult(eErrorType.Num);
				date = System.DateTime.FromOADate(result);
				if (result == 0 || result == 0.0)
					date = System.DateTime.FromOADate(result + 2);
			}
			else
				return new CompileResult(eErrorType.Value);

			if (monthsToAddArgument is double)
			{
				var monthCandidate = this.ArgToDecimal(arguments, 1);
				monthsToAdd = (int)monthCandidate;
			}
			else if (ConvertUtil.TryParseDateObjectToOADate(monthsToAddArgument, out result))
			{
				monthsToAdd = (int)result;
			}
			else
				return new CompileResult(eErrorType.Value);

			if (date.ToOADate() < 60 && monthsToAdd < 0)
				return new CompileResult(eErrorType.Num);

			var resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-1);
			//This accounts for the Lotus 1-2-3 error with dates before March 1, 1900.
			if (date.ToOADate() <= 60)
			{
				resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-2);
			}

			return this.CreateResult(resultDate.ToOADate(), DataType.Date);
		}

		#endregion
	}
}
