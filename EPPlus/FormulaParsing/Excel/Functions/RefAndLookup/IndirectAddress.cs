/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Evan Schallerer, and others as noted in the source history.
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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// A function which takes the same arguments as the INDIRECT( ) function and calculates a reference.
	/// </summary>
	/// <remarks>
	/// This function is used internally to calculate an INDIRECT( ) function's dependencies before calculation.
	/// </remarks>
	public class IndirectAddress : ExcelFunction
	{
		#region Constants
		/// <summary>
		/// The name of the INDIRECTADDRESS function.
		/// </summary>
		public const string Name = "INDIRECTADDRESS";
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the function with the specified <paramref name="arguments"/> in the specified <paramref name="context"/>.
		/// </summary>
		/// <param name="arguments">The arguments to evaluate the function with.</param>
		/// <param name="context">The context with which to evaluate the function in.</param>
		/// <returns>A <see cref="CompileResult"/> containing the result of evaluation.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var addressArgument = base.ArgToString(arguments, 0);
			var address = new ExcelAddress(addressArgument);
			if (string.IsNullOrEmpty(address.Address))
				return new CompileResult(eErrorType.Value);
			return new CompileResult(address.FullAddress, DataType.String);
		}
		#endregion
	}
}
