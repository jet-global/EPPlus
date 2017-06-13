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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class AverageIf : HiddenValuesHandlingFunction
	{
		private readonly ExpressionEvaluator _expressionEvaluator;

		public AverageIf()
			 : this(new ExpressionEvaluator())
		{

		}

		public AverageIf(ExpressionEvaluator evaluator)
		{
			Require.That(evaluator).Named("evaluator").IsNotNull();
			_expressionEvaluator = evaluator;
		}

		private bool Evaluate(object obj, string expression)
		{
			double? candidate = default(double?);
			if (IsNumeric(obj))
			{
				candidate = ConvertUtil.GetValueDouble(obj);
			}
			if (candidate.HasValue)
			{
				return _expressionEvaluator.Evaluate(candidate.Value, expression);
			}
			return _expressionEvaluator.Evaluate(obj, expression);
		}

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var args = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			var criteria = GetFirstArgument(arguments.ElementAt(1)).ValueFirst != null ? GetFirstArgument(arguments.ElementAt(1)).ValueFirst.ToString() : string.Empty;
			var retVal = 0d;
			if (args == null)
			{
				var val = GetFirstArgument(arguments.ElementAt(0)).Value;
				if (criteria != null && Evaluate(val, criteria))
				{
					if (arguments.Count() > 2)
					{
						var averageVal = arguments.ElementAt(2).Value;
						var averageRange = averageVal as ExcelDataProvider.IRangeInfo;
						if (averageRange != null)
						{
							retVal = averageRange.First().ValueDouble;
						}
						else
						{
							retVal = ConvertUtil.GetValueDouble(averageVal, true);
						}
					}
					else
					{
						retVal = ConvertUtil.GetValueDouble(val, true);
					}
				}
				else
				{
					throw new ExcelErrorValueException(eErrorType.Div0);
				}
			}
			else if (arguments.Count() > 2)
			{
				var lookupRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				retVal = CalculateWithAverageRange(args, criteria, lookupRange, context);
			}
			else
			{
				retVal = CalculateSingleRange(args, criteria, context);
			}
			return CreateResult(retVal, DataType.Decimal);
		}

		private double CalculateWithAverageRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange, ParsingContext context)
		{
			var retVal = 0d;
			var nMatches = 0;
			foreach (var cell in range)
			{
				if (criteria != null && Evaluate(GetFirstArgument(cell.Value), criteria))
				{
					var or = cell.Row - range.Address._fromRow;
					var oc = cell.Column - range.Address._fromCol;
					if (sumRange.Address._fromRow + or <= sumRange.Address._toRow &&
						sumRange.Address._fromCol + oc <= sumRange.Address._toCol)
					{
						var v = sumRange.GetOffset(or, oc);
						nMatches++;
						retVal += ConvertUtil.GetValueDouble(v, true);
					}
				}
			}
			return Divide(retVal, nMatches);
		}

		private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression, ParsingContext context)
		{
			var retVal = 0d;
			var nMatches = 0;
			foreach (var candidate in range)
			{
				if (expression != null && IsNumeric(GetFirstArgument(candidate.Value)) && Evaluate(GetFirstArgument(candidate.Value), expression))
				{
					retVal += candidate.ValueDouble;
					nMatches++;
				}
			}
			return Divide(retVal, nMatches);
		}
	}
}
