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
	/// Implements the COTH function.
	/// </summary>
	public class Coth : ExcelFunction
	{
		/// <summary>
		/// Calculate the hyperbolic cotangent of a given input.
		/// </summary>
		/// <param name="arguments">Input to have its hyperbolic cotangent calculated.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>Returns the hyperbolic cotangent of an angle.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var argument = arguments.First().Value;
			if (!ConvertUtil.TryParseObjectToDecimal(argument, out double numericValue))
				return new CompileResult(eErrorType.Value);

			if (AdvancedTrigonometry.TryCheckIfCosecantWillHaveADivideByZeroError(numericValue, out double cosecant))
				return new CompileResult(eErrorType.Div0);

			return this.CreateResult(AdvancedTrigonometry.HyperbolicCotangent(numericValue), DataType.Decimal);
		}
	}
}
