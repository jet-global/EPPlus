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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	public class Index : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var arg1 = arguments.ElementAt(0);
			var arg2 = arguments.ElementAt(1);

			

			var args = arg1.Value as IEnumerable<FunctionArgument>;
			var crf = new CompileResultFactory();
			if (args != null)
			{
				var index = this.ArgToInt(arguments, 1);
				if (index > args.Count())
					throw new ExcelErrorValueException(eErrorType.Ref);
				var candidate = args.ElementAt(index - 1);
				return crf.Create(candidate.Value);
			}
			if (arg2 != null && arg2.Value is ExcelErrorValue && arg2.Value.ToString() == ExcelErrorValue.Values.NA)
				return crf.Create(arg2.Value);
			if (arg1.IsExcelRange)
			{
				
				var row = this.ArgToInt(arguments, 1);

				if (row == 0)
					return new CompileResult(eErrorType.Value);
				if (row < 0)
					return new CompileResult(eErrorType.Value);

				var column = arguments.Count() > 2 ? this.ArgToInt(arguments, 2) : 1;


				if (column == 0 && row == 0)
					return new CompileResult(eErrorType.Value);
				if (column < 0)
					return new CompileResult(eErrorType.Value);


				var rangeInfo = arg1.ValueAsRangeInfo;


				if (rangeInfo.Address.Rows == 1 && arguments.Count() < 3)
				{
					column = row;
					row = 1;
				}
				else if (arguments.ElementAt(2).Value == null)
					return new CompileResult(eErrorType.Value);
						
			

				if (row > rangeInfo.Address.Rows || column > rangeInfo.Address.Columns)
					return new CompileResult(eErrorType.Ref);

				

				if (row > rangeInfo.Address._toRow - rangeInfo.Address._fromRow + 1 || column > rangeInfo.Address._toCol - rangeInfo.Address._fromCol + 1)
					return new CompileResult(eErrorType.Value);
				var candidate = rangeInfo.GetOffset(row - 1, column - 1);
				if (column == 0)
					candidate = rangeInfo.GetOffset(row - 1, column);
				
			
				
				return crf.Create(candidate);
			}
			throw new NotImplementedException();
		}
	}
}
