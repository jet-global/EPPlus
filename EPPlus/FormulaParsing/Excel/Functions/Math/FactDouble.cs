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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Returns the double factorial of a number.
	/// The double factorial is defined as such:
	/// Given a number n,
	///		if n is even, the double factorial of n is equal to n*(n-2)*(n-4)*...*(4)*(2).
	///		if n is odd, the double factorial of n is equal to n*(n-2)*(n-4)*...*(3)*(1).
	///		
	/// Excel documentation:
	///		https://support.office.com/en-us/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8
	///	
	/// More information on the double factorial:
	///		https://en.wikipedia.org/wiki/Double_factorial
	/// </summary>
	public class FactDouble : ExcelFunction
	{
		/// <summary>
		/// Returns the double factorial of a number.
		/// </summary>
		/// <param name="arguments">Contains the number to perform the double factorial operation on.</param>
		/// <param name="context">Unused in the method.</param>
		/// <returns>Returns the double factorial of the given number, or an <see cref="ExcelErrorValue"/> if the input is invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			if (arguments.ElementAt(0).ValueIsExcelError)
				return new CompileResult(arguments.ElementAt(0).ValueAsExcelErrorValue);
			if (!ConvertUtil.TryParseObjectToDecimal(arguments.ElementAt(0).Value, out double parsedDouble))
				return new CompileResult(eErrorType.Value);
			if (parsedDouble < -1)
				return new CompileResult(eErrorType.Num);
			var number = (int)parsedDouble;
			var result = 1d;
			for (var i = (2 - number % 2); i <= number; i += 2)
			{
				result *= i;
			}
			return this.CreateResult(result, DataType.Integer);
		}
	}
}
