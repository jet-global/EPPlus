﻿/* Copyright (C) 2011  Jan Källman
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
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	public class Lookup : LookupFunction
	{
		/// <summary>
		/// Gets a value representing the indicies of the arguments to the lookup function that
		/// should be compiled as ExcelAddresses instead of being evaluated.
		/// </summary>
		public override List<int> LookupArgumentIndicies { get; } = new List<int> { 1, 2 };

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			if (this.HaveTwoRanges(arguments))
				return this.HandleTwoRanges(arguments, context);
			return this.HandleSingleRange(arguments, context);
		}

		private bool HaveTwoRanges(IEnumerable<FunctionArgument> arguments)
		{
			if (arguments.Count() == 2) return false;
			return arguments.ElementAt(2).IsExcelRange;
		}

		private CompileResult HandleSingleRange(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var searchedValue = arguments.ElementAt(0).Value;
			Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
			var firstAddress = arguments.ElementAt(1).ValueAsRangeInfo?.Address.Address;
			var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
			var address = rangeAddressFactory.Create(firstAddress);
			var nRows = address.ToRow - address.FromRow;
			var nCols = address.ToCol - address.FromCol;
			var lookupIndex = nCols + 1;
			var lookupDirection = LookupDirection.Vertical;
			if (nCols > nRows)
			{
				lookupIndex = nRows + 1;
				lookupDirection = LookupDirection.Horizontal;
			}
			var lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, 0, true, arguments.ElementAt(1).ValueAsRangeInfo);
			var navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArgs, context);
			return Lookup(navigator, lookupArgs, new LookupValueMatcher());
		}

		private CompileResult HandleTwoRanges(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var searchedValue = arguments.ElementAt(0).Value;
			Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
			Require.That(arguments.ElementAt(2).Value).Named("secondAddress").IsNotNull();
			var firstAddress = arguments.ElementAt(1).ValueAsRangeInfo?.Address.Address;
			var secondAddress = arguments.ElementAt(2).ValueAsRangeInfo?.Address.Address;
			var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
			var address1 = rangeAddressFactory.Create(firstAddress);
			var address2 = rangeAddressFactory.Create(secondAddress);
			var lookupIndex = (address2.FromCol - address1.FromCol) + 1;
			var lookupOffset = address2.FromRow - address1.FromRow;
			var lookupDirection = GetLookupDirection(address1);
			if (lookupDirection == LookupDirection.Horizontal)
			{
				lookupIndex = (address2.FromRow - address1.FromRow) + 1;
				lookupOffset = address2.FromCol - address1.FromCol;
			}
			var lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, lookupOffset, true, arguments.ElementAt(1).ValueAsRangeInfo);
			var navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArgs, context);
			return Lookup(navigator, lookupArgs, new LookupValueMatcher());
		}
	}
}
