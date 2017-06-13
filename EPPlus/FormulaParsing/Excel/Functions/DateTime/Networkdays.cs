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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for calculating the number of workdays between a start and end date.
	/// </summary>
	public class Networkdays : ExcelFunction
	{
		/// <summary>
		/// Calculates the number of whole work days from given start and end dates, and an optional list of holiday dates. 
		/// </summary>
		/// <param name="arguments">The user specified start date, end date, and optional holiday dates.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The number of whole work days as an integer value.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var startDate = System.DateTime.FromOADate(ArgToInt(functionArguments, 0));
			var endDate = System.DateTime.FromOADate(ArgToInt(functionArguments, 1));
			var calculator = new WorkdayCalculator();
			var result = calculator.CalculateNumberOfWorkdays(startDate, endDate);
			if (functionArguments.Length > 2)
			{
				result = calculator.ReduceWorkdaysWithHolidays(result, functionArguments[2]);
			}

			return new CompileResult(result.NumberOfWorkdays, DataType.Integer);
		}
	}
}
