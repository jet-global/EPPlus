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
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Tis class calculates the nth smallest number in a list of numbers.
	/// </summary>
	public class Small : ExcelFunction
	{
		/// <summary>
		/// Returns the nth smallest number based on user input.
		/// </summary>
		/// <param name="arguments">The user specified list and nth smallest number to look up.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The nth smallest number as specified by the user.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var args = arguments.ElementAt(0);
			if (args.Value is string)
				if (!ConvertUtil.TryParseNumericString(args.Value, out _))
					if (!ConvertUtil.TryParseDateString(args.Value, out _))
						return new CompileResult(eErrorType.Value);
					else
						return new CompileResult(eErrorType.Num);

			var index = this.ArgToInt(arguments, 1) - 1;

			var values = this.ArgsToDoubleEnumerable(new List<FunctionArgument> { args }, context);
			if (index < 0 || index >= values.Count())
				return new CompileResult(eErrorType.Num);
			var result = values.OrderBy(x => x).ElementAt(index);
			return this.CreateResult(result, DataType.Decimal);
		}
	}
}
