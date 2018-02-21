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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Thi class contains the functionality of the SUMSQ Excel Function.
	/// </summary>
	public class Sum : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Computes the product of the squares of the given arguments.
		/// </summary>
		/// <param name="arguments">The user specified arguments to sum.</param>
		/// <param name="context">The current context of the function.</param>
		/// <returns>The sum of the products of the arguments given.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var sum = 0d;
			if (arguments != null)
			{
				foreach (var arg in arguments)
				{
					if (arg.Value is string && !ConvertUtil.TryParseObjectToDecimal(arg.Value, out _))
						return new CompileResult(eErrorType.Value);
					var valToAdd = this.Calculate(arg, context);
					sum += valToAdd;
				}
			}
			return this.CreateResult(sum, DataType.Decimal);
		}

		#region Private Methods
		/// <summary>
		/// Takes the specified argument and converts it to a double and then squares it. 
		/// </summary>
		/// <param name="arg">The argument to be converted.</param>
		/// <param name="context">The current context of the function.</param>
		/// <param name="isInArray">A boolean to flag if the argument is an an array.</param>
		/// <returns>The given argument squared as a double.</returns>
		private double Calculate(FunctionArgument arg, ParsingContext context, bool isInArray = false)
		{
			var calculatedValue = 0d;
			if (ShouldIgnore(arg))
			{
				return calculatedValue;
			}
			if (arg.Value is IEnumerable<FunctionArgument>)
			{
				foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
				{
					calculatedValue += this.Calculate(item, context, true);
				}
			}
			else
			{
				if (arg.Value is ExcelDataProvider.IRangeInfo rangeVal)
				{
					foreach (var cell in rangeVal)
					{
						if (this.ShouldIgnore(cell, context) == false)
						{
							this.CheckForAndHandleExcelError(cell);
							calculatedValue += cell.ValueDouble;
						}
					}
				}
				else
				{
					this.CheckForAndHandleExcelError(arg);
					if (IsNumericString(arg.Value) && !isInArray)
					{
						ConvertUtil.TryParseObjectToDecimal(arg.Value, out double value);
						return value;
					}
					if (!isInArray)
					{
						if (arg.Value == null)
							return 0;
					}
					calculatedValue += ConvertUtil.GetValueDouble(arg.Value, isInArray);
				}
			}
			return calculatedValue;
		}
		#endregion
	}
}
