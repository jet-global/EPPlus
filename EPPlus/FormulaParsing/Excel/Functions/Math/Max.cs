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
	/// This class contains the formula for calculating the maximum item in an array, list, or excel range.
	/// </summary>
	public class Max : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Takes the user specified arguments and returns the maximum value.
		/// </summary>
		/// <param name="arguments">The user specified array, list, or excel range to take the maximum of.</param>
		/// <param name="context">The context in which the program is being run.</param>
		/// <returns>The maximum item of the user specified arguments.</returns>
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
				return CreateResult(values.Max(), DataType.Decimal);
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
				return CreateResult(tvalues.Max(), DataType.Decimal);
			}

			if (values.Count() > 255)
				return new CompileResult(eErrorType.NA);
			return CreateResult(values.Max(), DataType.Decimal);
		}
	}
}
