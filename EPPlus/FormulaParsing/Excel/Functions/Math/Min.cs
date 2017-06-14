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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Min : HiddenValuesHandlingFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var args = arguments.ElementAt(0);
			var argumentValueList = this.ArgsToObjectEnumerable(false, new List<FunctionArgument> { args }, context);
			var values = argumentValueList.Where(arg => ((arg.GetType().IsPrimitive && (arg is bool == false)) || arg is System.DateTime));
			if (arguments.ElementAt(0).Type.Name.Equals("List`1"))
			{
				if (values.Count() > 255)
					return new CompileResult(eErrorType.NA);
				return CreateResult(values.Min(), DataType.Decimal);
			}
			else if (!arguments.ElementAt(0).IsExcelRange)
			{
				var tvalues = new List<double> { };
				foreach (var item in arguments)
				{
					if (item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
						continue;
					if (item.Value is string)
					{
						if (ConvertUtil.TryParseNumericString(item.Value, out double result))
							tvalues.Add(result);
						else if (ConvertUtil.TryParseDateString(item.Value, out System.DateTime res))
						{
							var temp = res.ToOADate();
							tvalues.Add(temp);
						}
					}
					else
						tvalues.Add(ArgToDecimal(item.Value));
				}
				if (tvalues.Count() == 0)
					return new CompileResult(eErrorType.Value);
				if (tvalues.Count() > 255)
					return new CompileResult(eErrorType.NA);
				return CreateResult(tvalues.Min(), DataType.Decimal);
			}

			if (values.Count() > 255)
				return new CompileResult(eErrorType.NA);
			return CreateResult(values.Min(), DataType.Decimal);
		}
	}
}
