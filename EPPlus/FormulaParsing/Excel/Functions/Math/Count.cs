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
* Mats Alm   		                Added		                2013-12-03
********************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Count : HiddenValuesHandlingFunction
	{
		private enum ValueContext
		{
			EnteredDirectly,
			FromArray,
			FromCellRange
		}

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var numberOfValues = 0d;
			foreach (var argument in arguments)
			{
				if (argument.Value is IEnumerable<FunctionArgument> subArguments)
				{
					foreach (var subArgument in subArguments)
					{
						if (!this.ShouldIgnore(subArgument) && this.ShouldCount(subArgument.Value, ValueContext.FromArray))
							numberOfValues++;
					}
				}
				else if (argument.Value is ExcelDataProvider.IRangeInfo cellRange)
				{
					foreach (var cell in cellRange)
					{
						if (this.ShouldIgnore(cell, context))
							continue;
						if (cell.Value is IEnumerable<object> array)
						{
							if (this.ShouldCount(array.ElementAt(0), ValueContext.FromArray))
								numberOfValues++;
						}
						else if (this.ShouldCount(cell.Value, ValueContext.FromCellRange))
							numberOfValues++;
					}
				}
				else
				{
					if (this.ShouldCount(argument.Value, ValueContext.EnteredDirectly))
						numberOfValues++;
				}
			}
			return this.CreateResult(numberOfValues, DataType.Integer);
		}

		private bool ShouldCount(object value, ValueContext context)
		{
			switch (context)
			{
				case ValueContext.EnteredDirectly:
					return (value == null || value is decimal || ConvertUtil.TryParseObjectToDecimal(value, out double parsedValue));
				case ValueContext.FromCellRange:
				case ValueContext.FromArray:
					return (!(value is bool) && this.IsNumeric(value));
				default:
					return false;
			}
		}
	}
}
