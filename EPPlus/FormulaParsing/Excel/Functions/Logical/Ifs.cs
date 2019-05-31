/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Evan Schallerer and others as noted in the source history.
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
	/// Represents the Excel IFS logical function.
	/// </summary>
	public class Ifs : ExcelFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the IFS function.
		/// </summary>
		/// <param name="arguments">The arguments to supply to the function.</param>
		/// <param name="context">The context to evaluate the function in.</param>
		/// <returns>The <see cref="CompileResult"/> result of executing the function.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (arguments.Count() == 0 || arguments.Count() % 2 != 0)
				return new CompileResult(eErrorType.Value);
			for (int i = 0; i + 1 < arguments.Count(); i += 2)
			{
				var condition = base.ArgToBool(arguments.ElementAt(i));
				if (condition)
					return new CompileResultFactory().Create(arguments.ElementAt(i + 1).Value);
			}
			return new CompileResult(eErrorType.NA);
		}
		#endregion
	}
}
