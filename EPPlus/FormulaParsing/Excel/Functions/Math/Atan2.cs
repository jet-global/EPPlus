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
 * Mats Alm   		                Added		                2015-01-11
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Implements the ATAN2 function.
	/// </summary>
	public class Atan2 : ExcelFunction
	{
		/// <summary>
		/// Given an (x,y) location, calculate the arctangent.
		/// </summary>
		/// <param name="arguments">Inputs to have its Arctangent calculated.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>Returns the Arctangent of a given pair of x and y coordinates</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);

			var argument = arguments.First().Value;
			var argument2 = arguments.ElementAt(1).Value;

			if ((!ConvertUtil.TryParseObjectToDecimal(argument, out double result)) || ( argument2 is string & !ConvertUtil.TryParseObjectToDecimal(argument2, out double result2)))
				return new CompileResult(eErrorType.Value);

			if(result == 0 && result2 == 0)
				return new CompileResult(eErrorType.Div0);

			var arg1 = ArgToDecimal(arguments, 0);
			var arg2 = ArgToDecimal(arguments, 1);
			// Had to switch order of the arguments to get the same result as in excel /MA

			return this.CreateResult(System.Math.Atan2(arg2, arg1), DataType.Decimal);
		}
	}
}
