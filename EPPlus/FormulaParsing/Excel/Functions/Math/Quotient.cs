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
	/// This class contains the formula for dividing two arguments.
	/// </summary>
	public class Quotient : ExcelFunction
	{
		/// <summary>
		/// Takes two user specified arguments and divides the first by the second. 
		/// </summary>
		/// <param name="arguments">The user specified arguments to divide.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument divided by the second argument as an integer value.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var numeratorCandidate = arguments.ElementAt(0).Value;
			var denominatorCandidate = arguments.ElementAt(1).Value;
			double candidateAsDouble;

			if (numeratorCandidate == null ||denominatorCandidate == null)
				return new CompileResult(eErrorType.NA);
			if (!ConvertUtil.TryParseNumericString(numeratorCandidate, out candidateAsDouble))
				if (!ConvertUtil.TryParseObjectToDecimal(numeratorCandidate, out candidateAsDouble))
					return new CompileResult(eErrorType.Value);
			if (!ConvertUtil.TryParseNumericString(denominatorCandidate, out candidateAsDouble))
				if (!ConvertUtil.TryParseObjectToDecimal(denominatorCandidate, out candidateAsDouble))
					return new CompileResult(eErrorType.Value);

			var num = this.ArgToDecimal(arguments, 0);
			var denominator = this.ArgToDecimal(arguments, 1);
			if (denominator == 0.0)
				return new CompileResult(eErrorType.Div0);
			var result = (int)(num / denominator);
			return this.CreateResult(result, DataType.Integer);
		}
	}
}
