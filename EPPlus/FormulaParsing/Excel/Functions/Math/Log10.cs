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
*
* Code change notes:
* 
* Author							Change						Date
*******************************************************************************
* Mats Alm   		                Added		                2013-12-03
*******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/

using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for computing logarithms with a base of 10.
	/// </summary>
	public class Log10 : ExcelFunction
	{
		/// <summary>
		/// Takes the user's input and calculates the log of that argument with a base of 10.
		/// </summary>
		/// <param name="arguments">The user specified argument.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The log of the argument with a base of 10.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var numberCandidate = arguments.ElementAt(0).Value;

			if (!ConvertUtil.TryParseNumericString(numberCandidate, out _))
				if (!ConvertUtil.TryParseDateString(numberCandidate, out _))
					return new CompileResult(eErrorType.Value);

			var number = ArgToDecimal(arguments, 0);
			if (number < 0)
				return new CompileResult(eErrorType.Num);
			return CreateResult(System.Math.Log10(number), DataType.Decimal);
		}
	}
}
