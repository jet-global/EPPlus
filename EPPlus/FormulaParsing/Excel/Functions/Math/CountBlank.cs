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
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class CountBlank : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var rangeToCount = arguments.ElementAt(0).ValueAsRangeInfo;
			if (rangeToCount == null)
				return new CompileResult(eErrorType.Value);
			// Note that the blank cells should be counted by subtracting the non-blank cells from the total rather
			// than by counting the blank cells directly because, for cells that have not been explicitly set to any value,
			// (null or otherwise), EPPlus will not include those cells in the given range of cells.
			var totalCellsInRange = rangeToCount.GetTotalCellCount();
			var numberOfCellsToIgnore = rangeToCount.Select(cell => this.GetFirstArgument(cell.Value)).Where(cellValue => !(cellValue == null || cellValue.Equals(string.Empty))).Count();
			double count = totalCellsInRange - numberOfCellsToIgnore;
			return this.CreateResult(count, DataType.Integer);
		}
	}
}
