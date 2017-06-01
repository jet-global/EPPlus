using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	public class Yearfrac : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (ValidateArguments(functionArguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var date1Candidate = functionArguments[0].Value;
			var date2Candidate = functionArguments[1].Value;
			if (date1Candidate == null || date2Candidate == null)
				return new CompileResult(eErrorType.NA);
			var isValidDate1 = ConvertUtil.TryParseDateObjectToOADate(date1Candidate, out double date1Num);
			var isValidDate2 = ConvertUtil.TryParseDateObjectToOADate(date2Candidate, out double date2Num);
			if (!isValidDate1)
				return new CompileResult(eErrorType.Value);
			else if (date1Num < 0)
				return new CompileResult(eErrorType.Num);
			else if (!isValidDate2)
				return new CompileResult(eErrorType.Value);
			else if (date2Num < 0)
				return new CompileResult(eErrorType.Num);
			//var date1Num = parsedDate1.ToOADate();
			//var date2Num = parsedDate2.ToOADate();
			if (date1Num < 61)
				date1Num++;
			if (date2Num < 61)
				date2Num++;
			if (date1Num > date2Num) //Switch to make date1 the lowest date
			{
				var t = date1Num;
				date1Num = date2Num;
				date2Num = t;
				var fa = functionArguments[1];
				functionArguments[1] = functionArguments[0];
				functionArguments[0] = fa;
			}
			var date1 = System.DateTime.FromOADate(date1Num);
			var date2 = System.DateTime.FromOADate(date2Num);

			var basis = 0;
			if (functionArguments.Count() > 2)
			{
				var isValidBasis = this.TryGetArgAsInt(functionArguments, 2, out basis);
				if (!isValidBasis)
					return new CompileResult(eErrorType.Value);
				else if (basis < 0 || basis > 4)
					return new CompileResult(eErrorType.Num);
			}
			var func = context.Configuration.FunctionRepository.GetFunction("days360");
			var calendar = new GregorianCalendar();
			var yearFracResult = 0d;
			switch (basis)
			{
				case 0:
					var d360Result = System.Math.Abs(func.Execute(functionArguments, context).ResultNumeric);
					// reproducing excels behaviour
					if (date1.Month == 2 && date2.Day == 31)
					{
						var daysInFeb = calendar.IsLeapYear(date1.Year) ? 29 : 28;
						if (date1.Day == daysInFeb) d360Result++;
					}
					yearFracResult = d360Result / 360d;
					break;
				case 1:
					yearFracResult = System.Math.Abs((date2 - date1).TotalDays / CalculateAcutalYear(date1, date2));
					break;
				case 2:
					yearFracResult = System.Math.Abs((date2 - date1).TotalDays / 360d);
					break;
				case 3:
					yearFracResult = System.Math.Abs((date2 - date1).TotalDays / 365d);
					break;
				case 4:
					var args = functionArguments.ToList();
					args.Add(new FunctionArgument(true));
					double? result = System.Math.Abs(func.Execute(args, context).ResultNumeric / 360d);
					yearFracResult = result.Value;
					break;
				default:
					return null;
			}

			yearFracResult = System.Math.Round(yearFracResult, 11);
			return this.CreateResult(yearFracResult, DataType.Decimal);
		}

		private double CalculateAcutalYear(System.DateTime dt1, System.DateTime dt2)
		{
			var calendar = new GregorianCalendar();
			var perYear = 0d;
			var nYears = dt2.Year - dt1.Year + 1;
			for (var y = dt1.Year; y <= dt2.Year; ++y)
			{
				perYear += calendar.IsLeapYear(y) ? 366 : 365;
			}
			if (new System.DateTime(dt1.Year + 1, dt1.Month, dt1.Day) >= dt2)
			{
				nYears = 1;
				perYear = 365;
				if (calendar.IsLeapYear(dt1.Year) && dt1.Month <= 2)
					perYear = 366;
				else if (calendar.IsLeapYear(dt2.Year) && dt2.Month > 2)
					perYear = 366;
				else if (dt2.Month == 2 && dt2.Day == 29)
					perYear = 366;
			}
			return perYear / (double)nYears;
		}
	}
}
