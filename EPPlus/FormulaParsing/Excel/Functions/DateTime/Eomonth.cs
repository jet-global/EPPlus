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
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			System.DateTime date;
			double result;
			var monthsToAdd = 0;
			var dateArgument = arguments.ElementAt(0).Value;
			var monthsToAddArgument = arguments.ElementAt(1).Value;

			if (dateArgument == null || monthsToAddArgument == null)
				return new CompileResult(eErrorType.NA);
			if (ConvertUtil.TryParseObjectToDecimal(dateArgument, out result))
			{
				if (result < 0)
					return new CompileResult(eErrorType.Num);
				int intResult = (int)result;
				date = System.DateTime.FromOADate(intResult);
				if (intResult == 0 || intResult == 0.0)
					date = System.DateTime.FromOADate(result + 2);
			}
			else
				return new CompileResult(eErrorType.Value);

			if (monthsToAddArgument is double)
			{
				var monthCandidate = this.ArgToDecimal(arguments, 1);
				monthsToAdd = (int)monthCandidate;
			}
			else if (ConvertUtil.TryParseObjectToDecimal(monthsToAddArgument, out result))
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
