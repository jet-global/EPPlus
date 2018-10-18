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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// An Excel function that returns the the value of a cell at an offset of a given a range.
	/// </summary>
	public class Offset : LookupFunction
	{
		#region Constants
		/// <summary>
		/// The name of the OFFSET function.
		/// </summary>
		public const string Name = "OFFSET";
		#endregion

		#region LookupFunction Members
		/// <summary>
		/// Gets a value representing the indicies of the arguments to the lookup function that
		/// should be compiled as ExcelAddresses instead of being evaluated.
		/// </summary>
		public override List<int> LookupArgumentIndicies { get; } = new List<int> { 0 };
		#endregion

		#region Public LookupFunction Overrides
		/// <summary>
		/// Executes the OFFSET function with the specified <paramref name="arguments"/> in the specified <paramref name="context"/>.
		/// </summary>
		/// <param name="arguments">The arguments with which to evaluate the function.</param>
		/// <param name="context">The context in which to evaluate the function.</param>
		/// <returns>A <see cref="CompileResult"/> result.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentsAreValid(functionArguments, 3, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			ExcelAddress offset = base.CalculateOffset(functionArguments, context);
			if (offset == null)
				return new CompileResult(eErrorType.Ref);
			var newRange = context.ExcelDataProvider.GetRange(offset.WorkSheet, offset._fromRow, offset._fromCol, offset._toRow, offset._toCol);
			if (!newRange.IsMulti)
			{
				if (newRange.IsEmpty)
					return CompileResult.Empty;
				var val = newRange.GetValue(offset._fromRow, offset._fromCol);
				if (IsNumeric(val))
					return CreateResult(val, DataType.Decimal);
				if (val is ExcelErrorValue)
					return CreateResult(val, DataType.ExcelError);
				return CreateResult(val, DataType.String);
			}
			return CreateResult(newRange, DataType.Enumerable);
		}
		#endregion
	}
}
