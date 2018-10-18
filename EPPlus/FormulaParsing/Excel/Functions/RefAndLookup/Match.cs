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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// Implements the Excel MATCH function.
	/// </summary>
	public class Match : LookupFunction
	{
		#region LookupFunction Members
		/// <summary>
		/// Gets a value representing the indicies of the arguments to the lookup function that
		/// should be compiled as ExcelAddresses instead of being evaluated.
		/// </summary>
		public override List<int> LookupArgumentIndicies { get; } = new List<int> { 1 };
		#endregion

		#region Enums
		private enum MatchType
		{
			ClosestAbove = -1,
			ExactMatch = 0,
			ClosestBelow = 1
		}
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
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			
			if (!this.TryGetSearchValue(arguments.ElementAt(0), context, out var searchValue, out var error))
				return error;

			var lookupRange = arguments.ElementAt(1).ValueAsRangeInfo;
			var matchType = this.GetMatchType(arguments);
			var args = new LookupArguments(searchValue, lookupRange.Address.Address, 0, 0, false, lookupRange);
			var lookupDirection = this.GetLookupDirection(lookupRange.Address);
			var navigator = LookupNavigatorFactory.Create(lookupDirection, args, context);
			int? lastValidIndex = null;
			do
			{
				if (navigator.CurrentValue == null && searchValue == null)
					return this.CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
				int? matchResult;
				if (matchType == MatchType.ExactMatch)
					matchResult = new WildCardValueMatcher().IsMatch(searchValue, navigator.CurrentValue);
				else
					matchResult = new LookupValueMatcher().IsMatch(searchValue, navigator.CurrentValue);
				// For all match types, if the match result indicated equality, return the index (1 based)
				if (matchResult == 0)
					return this.CreateResult(navigator.Index + 1, DataType.Integer);
				if ((matchType == MatchType.ClosestBelow && matchResult > 0) || (matchType == MatchType.ClosestAbove && matchResult < 0))
					lastValidIndex = navigator.Index + 1;
				// If matchType is ClosestBelow or ClosestAbove and the match result test failed, no more searching is required
				else if (matchType == MatchType.ClosestBelow || matchType == MatchType.ClosestAbove)
					break;
			}
			while (navigator.MoveNext());
			if (lastValidIndex == null)
				return this.CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
			return this.CreateResult(lastValidIndex, DataType.Integer);
		}
		#endregion

		#region Private Methods
		private MatchType GetMatchType(IEnumerable<FunctionArgument> arguments)
		{
			var matchType = MatchType.ClosestBelow;
			if (arguments.Count() > 2)
				matchType = (MatchType)this.ArgToInt(arguments, 2);
			return matchType;
		}

		private bool TryGetSearchValue(FunctionArgument argument, ParsingContext context, out object searchValue, out CompileResult error)
		{
			searchValue = null;
			error = null;
			if (argument.Value is ExcelDataProvider.IRangeInfo rangeInfo)
			{
				// If the first argument is a range, we take the value in the range that is perpendicular to the function.
				// If the lookup range and MATCH function are not perpendicular, #NA is returned. 
				var rangeInfoValue = argument.ValueAsRangeInfo;
				var addr = rangeInfoValue?.Address;
				// The lookup range must be one-dimensional.
				if (addr._fromCol != addr._toCol && addr._fromRow != addr._toRow)
				{
					error = new CompileResult(eErrorType.Value);
					return false;
				}
				var direction = this.GetLookupDirection(addr);
				var functionLocation = context.Scopes.Current.Address;
				if (direction == LookupDirection.Vertical && addr._fromRow <= functionLocation.FromRow && functionLocation.FromRow <= addr._toRow)
					searchValue = rangeInfoValue.GetValue(functionLocation.FromRow, addr._fromCol);
				else if (direction == LookupDirection.Horizontal && addr._fromCol <= functionLocation.FromCol && functionLocation.FromCol <= addr._toCol)
					searchValue = rangeInfoValue.GetValue(addr._fromRow, functionLocation.FromCol);
				else
				{
					error = new CompileResult(eErrorType.NA);
					return false;
				}
			}
			else
				searchValue = argument.Value;
			return true;
		}
		#endregion
	}
}
