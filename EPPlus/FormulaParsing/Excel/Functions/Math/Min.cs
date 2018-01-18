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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for calculating the minimum value in a set of arguments. 
	/// </summary>
	public class Min : MinMaxBase
	{
		/// <summary>
		/// Takes the user specified arguments and returns the minimum value.
		/// </summary>
		/// <param name="arguments">The user specified arguments, which can be a list, array, or cell referece.</param>
		/// <param name="context">The context in which the method is being called.</param>
		/// <returns>The minimum item in the user specified argument list.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var argsList = this.GroupArguments(arguments, context, out eErrorType? error);
			if (argsList == null)
				return new CompileResult(error ?? eErrorType.Value);
			else
			{
				double result = argsList.Any() ? Convert.ToDouble(argsList.Min()) : 0;
				return new CompileResult(result, DataType.Decimal);
			}
		}
	}
}
