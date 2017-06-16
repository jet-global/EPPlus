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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
	/// <summary>
	/// Returns the specified value if the given formula or value evaluates to an error. 
	/// Otherwise, returns the result of the given formula or value.
	/// </summary>
	public class IfError : ExcelFunction
	{
		/// <summary>
		/// Returns the second argument if the first argument evaluates to an error.
		/// Otherwise, returns the result of the first argument.
		/// </summary>
		/// <param name="arguments">The arguments being checked.</param>
		/// <param name="context">Unused in the method.</param>
		/// <returns>Returns the second argument if the first argument evaluates as an <see cref="ExcelErrorValue"/>,
		///				and returns the first argument otherwise.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var firstArgValue = (arguments.ElementAt(0).Value == null) ? 0 : arguments.ElementAt(0).Value;
			var lastArgValue = (arguments.ElementAt(1).Value == null) ? 0 : arguments.ElementAt(1).Value;
			if (arguments.ElementAt(0).ValueIsExcelError)
				return this.GetResultByObject(lastArgValue);
			else
				return this.GetResultByObject(firstArgValue);
		}
	}
}
