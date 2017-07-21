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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Var : HiddenValuesHandlingFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			List<double> listToDoVarianceOn = new List<double>();

			foreach (var item in arguments)
			{
				if (item.ValueAsRangeInfo != null)
				{
					foreach (var cell in item.ValueAsRangeInfo)
					{
						if (StatisticsFunctionHelper.TryToParseValuesFromInputArgumentByRefrenceOrRange(this.IgnoreHiddenValues, cell, context, false, out double numberToAddToList, out bool onlyStringInputsGiven1))
							listToDoVarianceOn.Add(numberToAddToList);
					}
				}
				else
				{
					if (StatisticsFunctionHelper.TryToParseValuesFromInputArgument(this.IgnoreHiddenValues, item, context, out double numberToAddToList, out bool onlyStringInputsGiven2))
						listToDoVarianceOn.Add(numberToAddToList);
					if (item.ValueFirst == null)
						listToDoVarianceOn.Add(0.0);
				}
			}
			if (listToDoVarianceOn.Count() == 0)
				return new CompileResult(eErrorType.Div0);
			if (!StatisticsFunctionHelper.TryVarSamplePopulationForAValueErrorCheck(listToDoVarianceOn, out double variance))
				return new CompileResult(eErrorType.Value);
			return new CompileResult(variance, DataType.Decimal);
		}
	}
}
