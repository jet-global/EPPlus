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
	public class Match : LookupFunction
	{
		public override List<int> LookupArgumentIndicies { get; } = new List<int> { 1 };

		private enum MatchType
		{
			ClosestAbove = -1,
			ExactMatch = 0,
			ClosestBelow = 1
		}

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var searchedValue = arguments.ElementAt(0).Value;
			var address = arguments.ElementAt(1).ValueAsRangeInfo?.Address.Address;
			var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
			var rangeAddress = rangeAddressFactory.Create(address);
			var matchType = this.GetMatchType(arguments);
			var args = new LookupArguments(searchedValue, address, 0, 0, false, arguments.ElementAt(1).ValueAsRangeInfo);
			var lookupDirection = this.GetLookupDirection(rangeAddress);
			var navigator = LookupNavigatorFactory.Create(lookupDirection, args, context);
			int? lastValidIndex = null;
			do
			{
				if (navigator.CurrentValue == null && searchedValue == null)
					return this.CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
				int matchResult;
				if (matchType == MatchType.ExactMatch)
					matchResult = new WildCardValueMatcher().IsMatch(searchedValue, navigator.CurrentValue);
				else
					matchResult = new LookupValueMatcher().IsMatch(searchedValue, navigator.CurrentValue);
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

		private MatchType GetMatchType(IEnumerable<FunctionArgument> arguments)
		{
			var matchType = MatchType.ClosestBelow;
			if (arguments.Count() > 2)
				matchType = (MatchType)this.ArgToInt(arguments, 2);
			return matchType;
		}
	}
}
