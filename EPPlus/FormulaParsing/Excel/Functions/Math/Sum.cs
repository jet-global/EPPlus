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
	/// This class contains the formula for the SUM Excel Function.
	/// </summary>
	public class Sum : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Takes the user specified arguments and returns their sum or an error value if one of the arguments was invalid.
		/// </summary>
		/// <param name="arguments">The arguments to sum.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>The value of the sum of the arguments as a double value.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var sum = 0d;
			if (arguments != null)
			{
				foreach (var arg in arguments)
				{
					sum += Calculate(arg, context);
				}
			}
			return CreateResult(sum, DataType.Decimal);
		}

		#region Private Methods
		/// <summary>
		/// Takes the given argument and converts it into a double to add to the sum.
		/// </summary>
		/// <param name="arg">The argument to be converted.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>The argument as a double.</returns>
		private double Calculate(FunctionArgument arg, ParsingContext context)
		{
			var retVal = 0d;
			if (ShouldIgnore(arg))
			{
				return retVal;
			}
			if (arg.Value is IEnumerable<FunctionArgument>)
			{
				foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
				{
					retVal += Calculate(item, context);
				}
			}
			else if (arg.Value is ExcelDataProvider.IRangeInfo)
			{
				foreach (var c in (ExcelDataProvider.IRangeInfo)arg.Value)
				{
					if (ShouldIgnore(c, context) == false)
					{
						CheckForAndHandleExcelError(c);
						retVal += c.ValueDouble;
					}
				}
			}
			else
			{
				CheckForAndHandleExcelError(arg);
				retVal += ConvertUtil.GetValueDouble(arg.Value, true);
			}
			return retVal;
		}
		#endregion
	}
}
