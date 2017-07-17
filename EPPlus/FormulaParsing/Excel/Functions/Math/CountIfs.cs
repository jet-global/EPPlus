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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class CountIfs : MultipleRangeCriteriasFunction
	{

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var firstRange = arguments.ElementAt(0).ValueAsRangeInfo;
			if (firstRange == null)
				return new CompileResult(eErrorType.Value);
			var indicesOfValidCells = new List<int>();
			for (var argumentIndex = 0; argumentIndex < arguments.Count(); argumentIndex += 2)
			{
				var currentRangeToCompare = arguments.ElementAt(argumentIndex).ValueAsRangeInfo;
				if (currentRangeToCompare == null || !this.RangesAreTheSameShape(firstRange, currentRangeToCompare))
					return new CompileResult(eErrorType.Value);
				var currentCriterion = IfHelper.ExtractCriterionObject(arguments.ElementAt(argumentIndex + 1), context);
				var passingIndices = this.GetIndicesOfCellsPassingCriterion(currentRangeToCompare, currentCriterion);
				if (argumentIndex == 0)
					indicesOfValidCells = passingIndices;
				else
					indicesOfValidCells = indicesOfValidCells.Intersect(passingIndices).ToList();
			}
			double count = indicesOfValidCells.Count();
			return this.CreateResult(count, DataType.Decimal);
		}

		/// <summary>
		/// Returns a list containing the indices of the cells in <paramref name="cellsToCompare"/> that satisfy
		/// the criterion detailed in <paramref name="criterion"/>.
		/// </summary>
		/// <param name="cellsToCompare">The <see cref="ExcelDataProvider.IRangeInfo"/> containing the cells to test against the <paramref name="criteria"/>.</param>
		/// <param name="criterion">The criterion dictating the acceptable contents of a given cell.</param>
		/// <returns>Returns a list of indexes corresponding to each cell that satisfies the given criterion.</returns>
		private List<int> GetIndicesOfCellsPassingCriterion(ExcelDataProvider.IRangeInfo cellsToCompare, object criterion)
		{
			var passingIndices = new List<int>();
			for (var currentCellIndex = 0; currentCellIndex < cellsToCompare.Count(); currentCellIndex++)
			{
				var currentCellValue = cellsToCompare.ElementAt(currentCellIndex).Value;
				if (IfHelper.ObjectMatchesCriterion(this.GetFirstArgument(currentCellValue), criterion))
					passingIndices.Add(currentCellIndex);
			}
			return passingIndices;
		}

		/// <summary>
		/// Checks if the <paramref name="expectedRange"/> is the same width and height as the
		/// <paramref name="actualRange"/>.
		/// </summary>
		/// <param name="expectedRange">The <see cref="ExcelDataProvider.IRangeInfo"/> with the desired cell width and height.</param>
		/// <param name="actualRange">The <see cref="ExcelDataProvider.IRangeInfo"/> with the width and height to be tested.</param>
		/// <returns>Returns true if <paramref name="expectedRange"/> and <paramref name="actualRange"/> have the same width and height values.</returns>
		private bool RangesAreTheSameShape(ExcelDataProvider.IRangeInfo expectedRange, ExcelDataProvider.IRangeInfo actualRange)
		{
			var expectedRangeWidth = expectedRange.Address._toCol - expectedRange.Address._fromCol;
			var expectedRangeHeight = expectedRange.Address._toRow - expectedRange.Address._fromRow;
			var actualRangeWidth = actualRange.Address._toCol - actualRange.Address._fromCol;
			var actualRangeHeight = actualRange.Address._toRow - actualRange.Address._fromRow;
			return (expectedRangeWidth == actualRangeWidth && expectedRangeHeight == actualRangeHeight);
		}
	}
}
