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
	/// This class contains the formula for calculating the natural log of a given input.
	/// </summary>
	public class Ln : ExcelFunction
	{
		/// <summary>
		/// Takes the natural log of a user specified argument. 
		/// </summary>
		/// <param name="arguments">The user specified argument to take the natural log of.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The natural log of the given input as a double value.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var numberCandidate = arguments.ElementAt(0).Value;
		
			if (arguments.ElementAt(0).ValueIsExcelError)
				return new CompileResult(arguments.ElementAt(0).ValueAsExcelErrorValue);
			if (!ConvertUtil.TryParseNumericString(numberCandidate, out _))
				if (!ConvertUtil.TryParseDateString(numberCandidate, out _))
					return new CompileResult(eErrorType.Value);

			var number = this.ArgToDecimal(arguments, 0);
			if (number <= 0)
				return new CompileResult(eErrorType.Num);
			return this.CreateResult(System.Math.Log(number, System.Math.E), DataType.Decimal);
		}
	}
}
