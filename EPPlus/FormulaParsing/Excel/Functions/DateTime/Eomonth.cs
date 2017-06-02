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

			if (firstArgument == null || secondArgument == null)
				return new CompileResult(eErrorType.NA);

			if (!ConvertUtil.TryParseDateObject(firstArgument, out System.DateTime dt1, out eErrorType? error))
				return new CompileResult(eErrorType.Value);

			var date = System.DateTime.FromOADate(dt1.ToOADate());
			var monthsToAdd = ArgToInt(arguments, 1);
			var resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-1);
			return CreateResult(resultDate.ToOADate(), DataType.Date);
		}

		#endregion
	}
}
