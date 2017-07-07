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
*  * Author							Change						Date
* *******************************************************************************
* * Mats Alm   		                Added		                2013-12-03
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the functionality for the SUMIF Excel Function.
	/// </summary>
	public class SumIf : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Returns the sum of all the cells in a range that meet a given criteria.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the sum.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>Returns the sum of all cells in the given range that passed the given criteria.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);

			var cellRangeToCheck = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (cellRangeToCheck == null)
				return new CompileResult(eErrorType.Value);
			var criteriaString = IfHelper.ExtractCriteriaString(arguments.ElementAt(1), context);
			if (arguments.Count() > 2)
			{
				var cellRangeToSum = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				if (cellRangeToSum == null)
					return new CompileResult(eErrorType.Value);
				else
					return this.CalculateSumUsingSumRange(cellRangeToCheck, criteriaString, cellRangeToSum);
			}
			else
				return this.CalculateSumUsingRange(cellRangeToCheck, criteriaString);
		}

		/// <summary>
		/// Calculates the sum value of all cells that match the given criteria. The sizes/shapes of
		/// <paramref name="cellsToCompare"/> and <paramref name="potentialCellsToSum"/> do not have to be the same;
		/// The size and shape of <paramref name="cellsToCompare"/> is applied to <paramref name="potentialCellsToSum"/>,
		/// using the first cell in <paramref name="potentialCellsToSum"/> as a reference point.
		/// </summary>
		/// <param name="cellsToCompare">The range of cells to compare against the <paramref name="comparisonCriteria"/>.</param>
		/// <param name="comparisonCriteria">The criteria dictating which cells should be included in the sum calculation.</param>
		/// <param name="potentialCellsToSum">
		///		If a cell in <paramref name="cellsToCompare"/> passes the criteria, then its
		///		corresponding cell in this cell range will be included in the sum calculation.</param>
		/// <returns>Returns the sum for all cells that pass the <paramref name="comparisonCriteria"/>.</returns>
		private CompileResult CalculateSumUsingSumRange(ExcelDataProvider.IRangeInfo cellsToCompare, string comparisonCriteria, ExcelDataProvider.IRangeInfo potentialCellsToSum)
		{
			var sumOfValidValues = 0d;
			foreach (var cell in cellsToCompare)
			{
				if (comparisonCriteria != null && IfHelper.ObjectMatchesCriteria(this.GetFirstArgument(cell.Value), comparisonCriteria))
				{
					var relativeRow = cell.Row - cellsToCompare.Address._fromRow;
					var relativeColumn = cell.Column - cellsToCompare.Address._fromCol;
					var valueOfCellToSum = potentialCellsToSum.GetOffset(relativeRow, relativeColumn);
					if (valueOfCellToSum is ExcelErrorValue cellError)
						continue;
					if (valueOfCellToSum is string || valueOfCellToSum is bool || valueOfCellToSum == null)
						continue;
					sumOfValidValues += ConvertUtil.GetValueDouble(valueOfCellToSum, true);
				}
			}
			return this.CreateResult(sumOfValidValues, DataType.Decimal);
		}

		/// <summary>
		/// Calculates the sum value of all cells that match the given criteria.
		/// </summary>
		/// <param name="potentialCellsToSum">
		///		The cell range to compare against the given <paramref name="comparisonCriteria"/>
		///		If a cell passes the criteria, then its value is included in the sum calculation.</param>
		/// <param name="comparisonCriteria">The criteria dictating which cells should be included in the sum calculation.</param>
		/// <returns>Returns the sum value for all cells that pass the <paramref name="comparisonCriteria"/>.</returns>
		private CompileResult CalculateSumUsingRange(ExcelDataProvider.IRangeInfo potentialCellsToSum, string comparisonCriteria)
		{
			var sumOfValidValues = 0d;
			foreach (var cell in potentialCellsToSum)
			{
				if (comparisonCriteria != null && IfHelper.IsNumeric(this.GetFirstArgument(cell.Value), true) &&
						IfHelper.ObjectMatchesCriteria(this.GetFirstArgument(cell.Value), comparisonCriteria))
				{
					sumOfValidValues += cell.ValueDouble;
				}
			}
			return this.CreateResult(sumOfValidValues, DataType.Decimal);
		}
	}
}
