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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Maxa : ExcelFunction
	{
		private readonly DoubleEnumerableArgConverter _argConverter;

		public Maxa()
			 : this(new DoubleEnumerableArgConverter())
		{

		}
		public Maxa(DoubleEnumerableArgConverter argConverter)
		{
			Utilities.Require.That(argConverter).Named("argConverter").IsNotNull();
			_argConverter = argConverter;
		}
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType errorValue) == false)
				return new CompileResult(errorValue);
			var args = arguments.ElementAt(0);
			var argumentValueList = this.ArgsToObjectEnumerable(false, new List<FunctionArgument> { args }, context);
			var values = argumentValueList.Where(arg => ((arg.GetType().IsPrimitive && (arg is bool == false)) || arg is System.DateTime));
			foreach (var item in argumentValueList)
			{
				if (item is ExcelErrorValue)
					return new CompileResult((ExcelErrorValue)item);
			}
			//If the input to the Max Function is not an excel range logical values and string representations of numbers
			//are allowed, even though they are not counted in a cell reference.
			if (!arguments.ElementAt(0).IsExcelRange)
			{
				if (arguments.Count() == 1)
					return this.CreateResult(Convert.ToDouble(values.Max()), DataType.Decimal);

				var doublesList = new List<double> { };
				foreach (var item in arguments)
				{
					if (item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
						continue;
					if (item.Value is string)
					{
						if (ConvertUtil.TryParseNumericString(item.Value, out double result))
							doublesList.Add(result);
						else if (ConvertUtil.TryParseDateString(item.Value, out System.DateTime dateResult))
							doublesList.Add(dateResult.ToOADate());
						else if (ConvertUtil.TryParseBooleanString(item.Value, out bool res))
							doublesList.Add(this.ArgToDecimal(res));
						else
							return new CompileResult(eErrorType.Value);
					}
					else
						doublesList.Add(this.ArgToDecimal(item.Value));
				}
				if (doublesList.Count() == 0)
					return new CompileResult(eErrorType.Value);
				if (doublesList.Count() > 255)
					return new CompileResult(eErrorType.NA);
				return this.CreateResult(doublesList.Max(), DataType.Decimal);
			}
			if (argumentValueList.Count() == 0)
				return this.CreateResult(0d, DataType.Decimal);
			var vvalues = _argConverter.ConvertArgsIncludingOtherTypes(arguments);
			if (vvalues.Count() > 255)
				return new CompileResult(eErrorType.NA);
			return this.CreateResult(Convert.ToDouble(vvalues.Max()), DataType.Decimal);
		}
	}
}
