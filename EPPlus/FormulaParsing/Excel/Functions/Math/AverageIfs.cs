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
 * Mats Alm   		                Added		                2015-02-01
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class AverageIfs : MultipleRangeCriteriasFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 3) == false)
				return new CompileResult(eErrorType.Div0);
			var rangeToAverage = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (rangeToAverage == null)
				return new CompileResult(eErrorType.Div0);
			var indexesToAverage = new List<int>();

			for (var argumentIndex = 1; argumentIndex < arguments.Count(); argumentIndex += 2)
			{
				var currentRangeToCompare = arguments.ElementAt(argumentIndex).ValueAsRangeInfo;
				if (currentRangeToCompare == null || !this.rangesAreTheSameShape(rangeToAverage, currentRangeToCompare))
					return new CompileResult(eErrorType.Value);

				var currentCriteriaArgument = arguments.ElementAt(argumentIndex + 1);
				if (!this.tryGetCriteria(currentCriteriaArgument, out string currentCriteria))
					return new CompileResult(eErrorType.Div0);

				var passingIndexes = this.getIndexesOfCellsPassingCriteria(currentRangeToCompare, currentCriteria);
				indexesToAverage = indexesToAverage.Union(passingIndexes).ToList();
			}
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var cellIndex in indexesToAverage)
			{
				var currentCellValue = rangeToAverage.ElementAt(cellIndex).Value;
				if (currentCellValue is ExcelErrorValue cellError)
					return new CompileResult(cellError.Type);
				else if (currentCellValue is string || currentCellValue is bool || currentCellValue == null)
					continue;
				sumOfValidValues += ConvertUtil.GetValueDouble(currentCellValue);
				numberOfValidValues++;
			}

			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}

		private List<int> getIndexesOfCellsPassingCriteria(ExcelDataProvider.IRangeInfo cellsToCompare, string criteria)
		{
			var passingIndexes = new List<int>();
			for (var currentCellIndex = 0; currentCellIndex < cellsToCompare.Count(); currentCellIndex++)
			{
				var currentCellValue = cellsToCompare.ElementAt(currentCellIndex).Value;
				if (IfHelper.objectMatchesCriteria(this.GetFirstArgument(currentCellValue), criteria))
					passingIndexes.Add(currentCellIndex);
			}
			return passingIndexes;
		}

		private bool tryGetCriteria(FunctionArgument criteriaCandidate, out string criteria)
		{
			criteria = null;
			if (criteriaCandidate.Value is ExcelDataProvider.IRangeInfo criteriaAsRange)
			{
				if (criteriaAsRange.IsMulti)
					return false;
				else
					criteria = this.GetFirstArgument(criteriaCandidate.ValueFirst).ToString().ToUpper();
			}
			else
				criteria = this.GetFirstArgument(criteriaCandidate).ValueFirst.ToString().ToUpper();
			return true;
		}

		private bool rangesAreTheSameShape(ExcelDataProvider.IRangeInfo expectedRange, ExcelDataProvider.IRangeInfo actualRange)
		{
			var expectedRangeWidth = expectedRange.Address._toCol - expectedRange.Address._fromCol;
			var expectedRangeHeight = expectedRange.Address._toRow - expectedRange.Address._fromRow;
			var actualRangeWidth = actualRange.Address._toCol - actualRange.Address._fromCol;
			var actualRangeHeight = actualRange.Address._toRow - actualRange.Address._fromRow;
			return (expectedRangeWidth == actualRangeWidth && expectedRangeHeight == actualRangeHeight);
		}
	}
}
