﻿using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Simple implementation of TimeValue function, just using .NET built-in
	/// function System.DateTime.TryParse, based on current culture
	/// </summary>
	public class TimeValue : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var dateString = ArgToString(arguments, 0);
			return Execute(dateString);
		}

		internal CompileResult Execute(string dateString)
		{
			System.DateTime result;
			
			
			System.DateTime.TryParse(dateString, out result);
		
				return result != System.DateTime.MinValue ?
					 CreateResult(GetTimeValue(result), DataType.Date) :
					 CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
			
			/*
			if (System.DateTime.TryParse(dateString, out result))
			{
				return result != System.DateTime.MinValue ?
					 CreateResult(GetTimeValue(result), DataType.Date) :
					 CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);

			}//ELse this is where we need to do somthing cray
			else
			{





			}
			*/





			return null;
			
			
		}

		private double GetTimeValue(System.DateTime result)
		{
			
			return (int)result.TimeOfDay.TotalSeconds == 0 ? 0d : result.TimeOfDay.TotalSeconds  / (3600 * 24);//3600 is seconds in an hour. So this is seconds in a day.
		}
	}


}