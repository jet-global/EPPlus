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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// Evaluates the Excel ROW function.
	/// </summary>
	public class Row : LookupFunction
	{
		#region LookupFunction Members
		/// <summary>
		/// Gets a value representing the indicies of the arguments to the lookup function that
		/// should be compiled as ExcelAddresses instead of being evaluated.
		/// </summary>
		public override List<int> LookupArgumentIndicies { get; } = new List<int> { 0 };
		#endregion

		#region Public ExcelFunction overrides
		/// <summary>
		/// Calculates the row of either the given range or the column that the function is executed in.
		/// </summary>
		/// <param name="arguments">The collection of arguments to be used to calculate the row value.</param>
		/// <param name="context">The context of the function when parsed.</param>
		/// <returns>Returns a <see cref="CompileResult"/> containing either the resulting row or an error value.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var rangeAddress = arguments.Count() == 0 ? null : arguments.ElementAt(0).ValueAsRangeInfo?.Address;
			if (arguments == null || arguments.Count() == 0 || rangeAddress == null)
			{
				return CreateResult(context.Scopes.Current.Address.FromRow, DataType.Integer);
			}
			if (!ExcelAddressUtil.IsValidAddress(rangeAddress.Address))
				return new CompileResult(eErrorType.Value);
			return CreateResult(rangeAddress._fromRow, DataType.Integer);
		}
		#endregion
	}
}
