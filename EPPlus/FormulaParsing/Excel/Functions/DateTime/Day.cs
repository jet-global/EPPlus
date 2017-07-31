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
*  * Author							Change						Date
* *******************************************************************************
* * Mats Alm   		                Added		                2013-12-03
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Returns the corresponding day of the month (as an int) from the given date.
	/// </summary>
	public class Day : ExcelFunction
	{
		/// <summary>
		/// Checks if the input is valid, and returns the corresponding day of the month if so.
		/// </summary>
		/// <param name="arguments">The given arguments used to calculate the day of the month.</param>
		/// <param name="context">Unused in the method, but necessary to override the method.</param>
		/// <returns>Returns the numeric day of the month for the given date, or an ExcelErrorValue, depending on if the input is valid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var serialNumberCandidate = arguments.ElementAt(0).Value;
			if (ConvertUtil.TryParseObjectToDecimal(serialNumberCandidate, out double serialNumber) &&
				serialNumber < 1 && serialNumber >= 0) // Zero and fractions are special cases and require specific output.
				return this.CreateResult(0, DataType.Integer);
			else if (ConvertUtil.TryParseDateObject(serialNumberCandidate, out System.DateTime date, out eErrorType? error))
				return this.CreateResult(date.Day, DataType.Integer);
			else
				return new CompileResult(error.Value);
		}
	}
}
