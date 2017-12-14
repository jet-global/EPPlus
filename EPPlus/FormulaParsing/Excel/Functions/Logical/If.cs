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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
	/// <summary>
	/// Represents the Excel IF logical function.
	/// </summary>
	public class If : ExcelFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Evaluates the specifed <paramref name="arguments"/> as follows:
		///  - [0]: the condition to evaluate;
		///  - [1]: the value to evaluate if true;
		///  - [2]: the value to evaluate if false.
		/// </summary>
		/// <param name="arguments">The arguments to evaluate.</param>
		/// <param name="context">The context within which to evaluate the function.</param>
		/// <returns>The result of either the second or third argument if the condition evaluates to true or false, respectively.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 3, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var condition = ArgToBool(arguments.ElementAt(0));
			var firstStatement = arguments.ElementAt(1).Value;
			var secondStatement = arguments.ElementAt(2).Value;
			var factory = new CompileResultFactory();
			return condition ? factory.Create(firstStatement) : factory.Create(secondStatement);
		}
		#endregion
	}
}
