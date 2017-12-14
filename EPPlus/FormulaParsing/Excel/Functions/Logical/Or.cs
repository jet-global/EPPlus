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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
	/// <summary>
	/// Represents the Excel logical OR function.
	/// </summary>
	public class Or : ExcelFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Calculates the logical OR value of the specified <paramref name="arguments"/>.
		/// </summary>
		/// <param name="arguments">The arguments on which to performa a logical OR.</param>
		/// <param name="context">The context in which to evaluate.</param>
		/// <returns>The logical OR value of the specified <paramref name="arguments"/>.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			foreach (var argument in arguments)
			{
				if (argument.IsExcelRange)
				{
					foreach (var cell in argument.Value as IRangeInfo)
					{
						if (this.ArgToBool(cell.Value))
							return new CompileResult(true, DataType.Boolean);
					}
				}
				else if (this.ArgToBool(argument))
					return new CompileResult(true, DataType.Boolean);
			}
			return new CompileResult(false, DataType.Boolean);
		}
	}
	#endregion
}
