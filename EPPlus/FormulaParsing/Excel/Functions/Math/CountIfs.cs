/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
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
********************************************************************************
* Mats Alm   		                Added		                2015-01-11
********************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This function applies criteria to cells across multiple ranges and counts the number of times all
	/// criteria are met.
	/// </summary>
	public class CountIfs : MultipleRangeCriteriasFunction
	{
		/// <summary>
		/// This function applies criteria to cells across multiple ranges and counts the number of times
		/// all criteria are met. If multiple cell ranges are being compared against criteria, all ranges must
		/// have the same number of rows and columns as the first given range, but the ranges do not have to be
		/// adjacent to each other.
		/// </summary>
		/// <param name="arguments">The arguments being evaluated.</param>
		/// <param name="context">The context for this function.</param>
		/// <returns>Returns the number of times all criteria are met across a row of cells.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var firstRangeToCompare = arguments.ElementAt(0).ValueAsRangeInfo;
			if (firstRangeToCompare == null)
				return new CompileResult(eErrorType.Value);
			var indicesOfValidCells = new List<int>();
			for (var argumentIndex = 0; argumentIndex < arguments.Count(); argumentIndex += 2)
			{
				var currentRangeToCompare = arguments.ElementAt(argumentIndex).ValueAsRangeInfo;
				if (currentRangeToCompare == null || !IfHelper.RangesAreTheSameShape(firstRangeToCompare, currentRangeToCompare))
					return new CompileResult(eErrorType.Value);
				var currentCriterion = IfHelper.ExtractCriterionObject(arguments.ElementAt(argumentIndex + 1), context);

				// This will always look at every cell in the given range of cells to compare. This is done instead of
				// using the iterator provided by the range of cells to compare because the collection of cells that it iterates over
				// does not include empty cells that have not been set since the workbook's creation. This function
				// wants to consider empty cells for comparing with the criterion, but it can be better optimized.
				// A similar problem and optimization opportunity exists in the AverageIf, AverageIfs, SumIf, SumIfs, and CountIf functions.
				var passingIndices = IfHelper.GetIndicesOfCellsPassingCriterion(currentRangeToCompare, currentCriterion);
				if (argumentIndex == 0)
					indicesOfValidCells = passingIndices;
				else
					indicesOfValidCells = indicesOfValidCells.Intersect(passingIndices).ToList();
			}
			double count = indicesOfValidCells.Count();
			return this.CreateResult(count, DataType.Integer);
		}
	}
}
