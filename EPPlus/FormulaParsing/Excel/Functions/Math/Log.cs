/* Copyright (C) 2011  Jan Källman
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for calculating standard logarithms.
	/// </summary>
	public class Log : ExcelFunction
	{
		/// <summary>
		/// Takes two arguments and computes the logarithm of the first argument with the second argument as the base. 
		/// </summary>
		/// <param name="arguments">The user specified arguments, with the first being the number and the second the base.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The log of the first argument with the second argument as the base.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var numberCandidate = arguments.ElementAt(0).Value;
			if (arguments.ElementAt(0).ValueIsExcelError)
				return new CompileResult(arguments.ElementAt(0).ValueAsExcelErrorValue);
			if (numberCandidate == null)
				return new CompileResult(eErrorType.Num);
			if (!ConvertUtil.TryParseNumericString(numberCandidate, out _))
				if (!ConvertUtil.TryParseDateString(numberCandidate, out _))
					return new CompileResult(eErrorType.Value);
			var number = this.ArgToDecimal(arguments, 0);

			if (arguments.Count() == 1)
			{
				if (number <= 0)
					return new CompileResult(eErrorType.Num);
				return this.CreateResult(System.Math.Log(number, 10d), DataType.Decimal);
			}

			var baseCandidate = arguments.ElementAt(1).Value;
			if (arguments.ElementAt(1).ValueIsExcelError)
				return new CompileResult(arguments.ElementAt(0).ValueAsExcelErrorValue);
			if (baseCandidate == null)
				return new CompileResult(eErrorType.Num);
			if (!ConvertUtil.TryParseNumericString(baseCandidate, out _))
				if (!ConvertUtil.TryParseDateString(baseCandidate, out _))
					return new CompileResult(eErrorType.Value);

			var newBase = this.ArgToDecimal(arguments, 1);
			if (number <= 0 || newBase <= 0)
				return new CompileResult(eErrorType.Num);
			return this.CreateResult(System.Math.Log(number, newBase), DataType.Decimal);
		}
	}
}
