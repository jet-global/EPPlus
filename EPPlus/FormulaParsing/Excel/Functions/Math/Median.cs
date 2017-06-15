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
* * Mats Alm   		                Added		                2015-01-10
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Median : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			if (arguments.ElementAt(0).Value == null && arguments.Count() == 1)
				return new CompileResult(eErrorType.Num);
			var args = arguments.ElementAt(0);
			var argumentValueList = this.ArgsToObjectEnumerable(false, new List<FunctionArgument> { args }, context);
			//var nums = argumentValueList.Where(arg => ((arg.GetType().IsPrimitive && (arg is bool == false))));

			var nums = ArgsToDoubleEnumerable(arguments, context);
			var arr = nums.ToArray();

			if (!arguments.ElementAt(0).IsExcelRange)
			{
				var tvalues = new List<double> { };
				foreach (var item in arguments)
				{
					if (item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
						continue;
					if (item.Value is string)
					{
						if (ConvertUtil.TryParseNumericString(item.Value, out double relt))
							tvalues.Add(relt);
						else if (ConvertUtil.TryParseDateString(item.Value, out System.DateTime res))
						{
							var temp = res.ToOADate();
							tvalues.Add(temp);
						}
						else if (ConvertUtil.TryParseBooleanString(item.Value, out bool r))
						{
							tvalues.Add(ArgToDecimal(r));
						}
						else if (item.ValueIsExcelError)
							return new CompileResult(item.ValueAsExcelErrorValue);
						else
							return new CompileResult(eErrorType.Value);
					}
					else if(item.Type == null)
					{
						tvalues.Add(0.0);
					}
					else
						tvalues.Add(ArgToDecimal(item.Value));
				}
				
				foreach(var item in argumentValueList)
				{
					if (item is ExcelErrorValue)
						return new CompileResult((ExcelErrorValue)item);
				}

				var tes = tvalues.ToArray();
				Array.Sort(tes);

				double reult;
				if (tes.Length % 2 == 1)
				{
					reult = tes[tes.Length / 2];
				}
				else
				{
					var startIndex = tes.Length / 2 - 1;
					reult = (tes[startIndex] + tes[startIndex + 1]) / 2d;
				}
				return CreateResult(reult, DataType.Decimal);
			}

			foreach (var item in argumentValueList)
			{
				if (item is ExcelErrorValue)
					return new CompileResult((ExcelErrorValue)item);
			}

			Array.Sort(arr);
			if (arr.Length == 0)
				return new CompileResult(eErrorType.Num);
			if (arr.Length > 255)
				return new CompileResult(eErrorType.NA);




			double result;
			if (arr.Length % 2 == 1)
			{
				result = (double)arr[arr.Length / 2];
			}
			else
			{
				var startIndex = arr.Length / 2 - 1;
				result = ((double)arr[startIndex] + (double)arr[startIndex + 1]) / 2d;
			}
			return CreateResult(result, DataType.Decimal);
		}
	}
}
