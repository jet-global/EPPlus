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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Count the number of cells in a cell range that meet a criterion.
	/// </summary>
	public class CountIf : ExcelFunction
	{
		/// <summary>
		/// Count the number of cells in a cell range that meet a criterion.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the number of valid cells.</param>
		/// <param name="context">The context for this function.</param>
		/// <returns>Returns the number of cells in the cell range that meet the criterion.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var cellRangeToCheck = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (cellRangeToCheck == null)
				return new CompileResult(eErrorType.Value);

			// This will always look at every cell in the given range of cells to compare. This is done instead of
			// using the iterator provided by the range of cells to compare because the collection of cells that it iterates over
			// does not include empty cells that have not been set since the workbook's creation. This function
			// wants to consider empty cells for comparing with the criterion, but it can be better optimized.
			// A similar problem and optimization opportunity exists in the AverageIf, AverageIfs, SumIf, SumIfs, and CountIfs functions.
			var criteriaObject = IfHelper.ExtractCriterionObject(arguments.ElementAt(1), context);
			double count = cellRangeToCheck.AllValues().Where(cellValue => IfHelper.ObjectMatchesCriterion(cellValue, criteriaObject)).Count();
			return this.CreateResult(count, DataType.Integer);
		}
	}
}
