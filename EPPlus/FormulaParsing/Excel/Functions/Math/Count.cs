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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Count : HiddenValuesHandlingFunction
	{
		private enum ItemContext
		{
			InRange,
			InArray,
			SingleArg
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
						if (IsNumeric(subArgument.Value))
							numberOfValues++;
					}
				}
				else if (argument.Value is ExcelDataProvider.IRangeInfo cellRange)
				{
					foreach (var cell in cellRange)
					{
						if (cell.Value is IEnumerable<object> array)
						{
							if (IsNumeric(array.ElementAt(0)))
								numberOfValues++;
						}
						else if (!(cell.Value is bool) && IsNumeric(cell.Value))
							numberOfValues++;
					}
				}
				else
				{
					if (argument.Value is bool || ConvertUtil.TryParseDateObjectToOADate(argument.Value, out double parsedValue))
						numberOfValues++;
				}
			}
			//Calculate(arguments, ref numberOfValues, context, ItemContext.SingleArg);
			return CreateResult(numberOfValues, DataType.Integer);
		}

		private void Calculate(IEnumerable<FunctionArgument> items, ref double nItems, ParsingContext context, ItemContext itemContext)
		{
			foreach (var item in items)
			{
				var cs = item.Value as ExcelDataProvider.IRangeInfo;
				if (cs != null)
				{
					foreach (var c in cs)
					{
						_CheckForAndHandleExcelError(c, context);
						if (ShouldIgnore(c, context) == false && ShouldCount(c.Value, ItemContext.InRange))
						{
							nItems++;
						}
					}
				}
				else
				{
					var value = item.Value as IEnumerable<FunctionArgument>;
					if (value != null)
					{
						Calculate(value, ref nItems, context, ItemContext.InArray);
					}
					else
					{
						_CheckForAndHandleExcelError(item, context);
						if (ShouldIgnore(item) == false && ShouldCount(item.Value, itemContext))
						{
							nItems++;
						}
					}
				}
			}
		}

		private void _CheckForAndHandleExcelError(FunctionArgument arg, ParsingContext context)
		{
			//if (context.Scopes.Current.IsSubtotal)
			//{
			//    CheckForAndHandleExcelError(arg);
			//}
		}

		private void _CheckForAndHandleExcelError(ExcelDataProvider.ICellInfo cell, ParsingContext context)
		{
			//if (context.Scopes.Current.IsSubtotal)
			//{
			//    CheckForAndHandleExcelError(cell);
			//}
		}

		private bool ShouldCount(object value, ItemContext context)
		{
			switch (context)
			{
				case ItemContext.SingleArg:
					return IsNumeric(value) || IsNumericString(value);
				case ItemContext.InRange:
					return IsNumeric(value);
				case ItemContext.InArray:
					return IsNumeric(value) || IsNumericString(value);
				default:
					throw new ArgumentException("Unknown ItemContext:" + context.ToString());
			}
		}
	}
}
