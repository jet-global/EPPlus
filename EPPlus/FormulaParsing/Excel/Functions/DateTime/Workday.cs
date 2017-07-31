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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing a date based on the given date, number of workdays, and (optional).
	/// dates of holidays
	/// </summary>
	public class Workday : ExcelFunction
	{
		/// <summary>
		/// Execute returns the date based on the user's input.
		/// </summary>
		/// <param name="arguments">The user specified date, number of workdays, and optional dates of holidays</param>
		/// <param name="context">Not used, but needed for overriding the method.</param>
		/// <returns>The date based on the date given and the number of workdays and optional holidays.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			System.DateTime output;
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();

			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var serialNumberCandidate = arguments.ElementAt(0).Value;
			var workDaysCandidate = arguments.ElementAt(1).Value;

			if (workDaysCandidate == null || serialNumberCandidate == null)
				return new CompileResult(eErrorType.NA);

			var serialNumberIsZero = (ConvertUtil.TryParseObjectToDecimal(serialNumberCandidate, out double parsedSerialNumber)
				&& parsedSerialNumber < 1 && parsedSerialNumber >= 0);

			if (serialNumberIsZero || ConvertUtil.TryParseDateObject(serialNumberCandidate, out output, out eErrorType? error))
			{
				if (serialNumberCandidate is int && ArgToInt(functionArguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

				if (workDaysCandidate is string)
				{
					if (!ConvertUtil.TryParseDateString(workDaysCandidate.ToString(), out output))
						return new CompileResult(eErrorType.Value);
				}

				var dateSerial = (int)this.ArgToDecimal(arguments, 0);
				if (dateSerial < 0)
					return new CompileResult(eErrorType.Num);

				if (serialNumberIsZero && this.ArgToInt(arguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

				var startDate = System.DateTime.FromOADate(dateSerial);
				var workDateSerial = this.ArgToDecimal(arguments, 1);
				var resultDate = System.DateTime.MinValue;
				var calculator = new WorkdayCalculator();
				var dateResult = calculator.CalculateWorkday(startDate, (int)workDateSerial);

				if (functionArguments.Length > 2)
				{
					for (int i = 2; i < functionArguments.Length; i++)
					{ 
						var holidayCandidate = arguments.ElementAt(i).Value;
						bool isHolidayZero = (serialNumberCandidate is int holAsint && holAsint == 0);

						if (holidayCandidate is int holAsInt && holAsInt < 0)
							return new CompileResult(eErrorType.Num);
						
						if (holidayCandidate is string && !ConvertUtil.TryParseDateString(holidayCandidate, out output))
							return new CompileResult(eErrorType.Value);

						dateResult = calculator.AdjustResultWithHolidays(dateResult, functionArguments[2]);
					}
				}
				if (serialNumberIsZero)
					return CreateResult(dateResult.EndDate.ToOADate()-1, DataType.Date);
				return CreateResult(dateResult.EndDate.ToOADate(), DataType.Date);
			}
			else
				return new CompileResult(error.Value);
		}
	}
}
