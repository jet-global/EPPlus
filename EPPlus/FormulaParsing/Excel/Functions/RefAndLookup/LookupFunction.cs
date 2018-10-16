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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// The base class for "lookup" type excel functions
	/// </summary>
	public abstract class LookupFunction : ExcelFunction
	{
		#region Abstract Properties
		/// <summary>
		/// Gets a value representing the indicies of the arguments to the lookup function that
		/// should be compiled as ExcelAddresses instead of being evaluated.
		/// </summary>
		public abstract List<int> LookupArgumentIndicies { get; }
		#endregion

		#region Protected Methods
		protected LookupDirection GetLookupDirection(ExcelAddress address)
		{
			var nRows = address._toRow - address._fromRow;
			var nCols = address._toCol - address._fromCol;
			return nCols > nRows ? LookupDirection.Horizontal : LookupDirection.Vertical;
		}

		protected CompileResult Lookup(LookupNavigator navigator, LookupArguments lookupArgs, ValueMatcher valueMatcher)
		{
			object lastValue = null;
			object lastLookupValue = null;
			int? lastMatchResult = null;
			if (lookupArgs.SearchedValue == null)
			{
				return new CompileResult(eErrorType.NA);
			}
			do
			{
				var matchResult = valueMatcher.IsMatch(lookupArgs.SearchedValue, navigator.CurrentValue);
				if (matchResult != 0)
				{
					if (lastValue != null && navigator.CurrentValue == null) break;

					if (!lookupArgs.RangeLookup) continue;
					if (lastValue == null && matchResult < 0)
					{
						return new CompileResult(eErrorType.NA);
					}
					if (lastValue != null && matchResult < 0 && lastMatchResult > 0)
					{
						return new CompileResultFactory().Create(lastLookupValue);
					}
					lastMatchResult = matchResult;
					lastValue = navigator.CurrentValue;
					lastLookupValue = navigator.GetLookupValue();
				}
				else
				{
					return new CompileResultFactory().Create(navigator.GetLookupValue());
				}
			}
			while (navigator.MoveNext());

			return lookupArgs.RangeLookup ? new CompileResultFactory().Create(lastLookupValue) : new CompileResult(eErrorType.NA);
		}

		protected ExcelAddress CalculateOffset(FunctionArgument[] arguments, ParsingContext context)
		{
			var rowOffset = ArgToInt(arguments, 1);
			var columnOffset = ArgToInt(arguments, 2);
			int width = 0, height = 0;
			if (arguments.Length > 3)
				height = ArgToInt(arguments, 3);
			if (arguments.Length > 4)
				width = ArgToInt(arguments, 4);
			if ((arguments.Length > 3 && height == 0) || (arguments.Length > 4 && width == 0))
				return null;
			var address = arguments[0].ValueAsRangeInfo?.Address;
			string targetWorksheetName;
			if (string.IsNullOrEmpty(address.WorkSheet))
				targetWorksheetName = context.Scopes?.Current?.Address?.Worksheet;
			else
				targetWorksheetName = address.WorkSheet;
			var fromRow = address._fromRow + rowOffset;
			var fromCol = address._fromCol + columnOffset;
			var toRow = (height == 0 ? address._toRow : height + address._fromRow - 1) + rowOffset;
			var toCol = (width == 0 ? address._toCol : width + address._fromCol - 1) + columnOffset;
			return new ExcelAddress(targetWorksheetName, fromRow, fromCol, toRow, toCol);
		}
		#endregion
	}
}
