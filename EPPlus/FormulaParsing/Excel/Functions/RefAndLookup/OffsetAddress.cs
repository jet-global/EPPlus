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
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// A function which takes the same arguments as the OFFSET( ) function and calculates the address offset.
	/// </summary>
	/// <remarks>
	/// This function is used internally to calculate an OFFSET( ) function's dependencies before calculation.
	/// </remarks>
	public class OffsetAddress : LookupFunction
	{
		#region Constants
		public const string Name = "OFFSETADDRESS";
		#endregion

		#region LookupFunction Members
		/// <summary>
		/// Gets a value representing the indicies of the arguments to the lookup function that
		/// should be compiled as ExcelAddresses instead of being evaluated.
		/// </summary>
		public override List<int> LookupArgumentIndicies { get; } = new List<int> { 0 };
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the function with the specified <paramref name="arguments"/> in the specified <paramref name="context"/>.
		/// </summary>
		/// <param name="arguments">The arguments with which to evaluate the function.</param>
		/// <param name="context">The context in which to evaluate the function.</param>
		/// <returns>An address range <see cref="CompileResult"/> if successful, otherwise an error result.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentsAreValid(functionArguments, 3, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			ExcelAddress offset = base.CalculateOffset(functionArguments, context);
			if (offset == null)
				return new CompileResult(eErrorType.Ref);
			return new CompileResult(offset.FullAddress, DataType.String);
		}
		#endregion
	}
}
