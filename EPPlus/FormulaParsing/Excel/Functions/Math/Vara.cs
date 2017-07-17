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
	/// <summary>
	/// Estimates variance based on a sample (includes logical values and text in the sample).
	/// </summary>
	public class Vara : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Variance measures how far a data set is spread out.
		/// Arguments can be the following: numbers; names, arrays, or references that contain numbers; text representations of numbers; or logical values, such as TRUE and FALSE, in a reference.
		/// Logical values and text representations of numbers that you type directly into the list of arguments are counted.
		/// </summary>
		/// <param name="arguments">Up too 254 individual arguments.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>The variance based on a sample.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			//Note: This follows the Functionality of excel which is diffrent from the excel documentation.
			//If you pass in a null Vara(1,1,1,,) it will treat those emtpy spaces as zeros insted of ignoring them.
			List<double> listToDoVarianceOn = new List<double>();
			bool onlyStringInputsGiven = true;
			foreach (var item in arguments)
			{
				if (item.ValueAsRangeInfo != null)
				{
					foreach (var cell in item.ValueAsRangeInfo)
					{
						if (StdevAndVarHelperClass.TryToParseValuesFromInputArgumentByRefrenceOrRange(this.IgnoreHiddenValues, cell, context, true, out double numberToAddToList, out bool onlyStringInputsGiven1))
							listToDoVarianceOn.Add(numberToAddToList);
						onlyStringInputsGiven = onlyStringInputsGiven1;
					}
				}
				else
				{
					if (StdevAndVarHelperClass.TryToParseValuesFromInputArgument(this.IgnoreHiddenValues, item, context, out double numberToAddToList, out bool onlyStringInputsGiven2))
						listToDoVarianceOn.Add(numberToAddToList);
					onlyStringInputsGiven = onlyStringInputsGiven2;
					if (item.ValueFirst == null)
						listToDoVarianceOn.Add(0.0);
				}
			}
			if (onlyStringInputsGiven)
				return new CompileResult(eErrorType.Value);
			if (listToDoVarianceOn.Count() == 0)
				return new CompileResult(eErrorType.Div0);
			if (!StdevAndVarHelperClass.TryVarSamplePopulationForAValueErrorCheck(listToDoVarianceOn, out double variance))
				return new CompileResult(eErrorType.Value);
			return new CompileResult(variance, DataType.Decimal);
		}
	}
}
