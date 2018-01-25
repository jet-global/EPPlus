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
 * Mats Alm   		                Added		                2014-04-13
 *******************************************************************************/
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// Represents the Excel INDIRECT(...) function.
	/// </summary>
	public class Indirect : ExcelFunction
	{
		#region Constants
		/// <summary>
		/// The name of the INDIRECT function.
		/// </summary>
		public const string Name = "INDIRECT";
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Evaluates the INDIRECT function with the specified <paramref name="arguments"/> 
		/// in the specified <paramref name="context"/>.
		/// </summary>
		/// <param name="arguments">The arguments to evaluate the function with.</param>
		/// <param name="context">The context with which to evaluate the function in.</param>
		/// <returns>A <see cref="CompileResult"/> containing the result of evaluation.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var address = base.ArgToString(arguments, 0);
			var adr = new ExcelAddress(address);
			var ws = adr.WorkSheet;
			if (string.IsNullOrEmpty(ws))
			{
				ws = context.Scopes.Current.Address.Worksheet;
			}
			var result = context.ExcelDataProvider.GetRange(ws, adr._fromRow, adr._fromCol, address);
			if (result.IsEmpty)
			{
				return CompileResult.Empty;
			}
			return new CompileResult(result, DataType.Enumerable);
		}
		#endregion
	}
}
