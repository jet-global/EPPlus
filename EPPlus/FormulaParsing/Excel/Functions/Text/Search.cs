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
 * Mats Alm   		                Added		                2016-03-28
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
	public class Search : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentsAreValid(functionArguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var search = base.ArgToString(functionArguments, 0);
			var searchIn = base.ArgToString(functionArguments, 1);
			// Subtract 1 because Excel uses 1-based index
			var startIndex = functionArguments.Count() > 2 ? base.ArgToInt(functionArguments, 2) - 1 : 0;
			int result = -1;
			// If the search string contains Excel wildcard values, convert it to regular expression values and find match.
			if (search.Contains('~') || search.Contains('?') || search.Contains('*'))
			{
				var pattern = new WildCardValueMatcher().ExcelWildcardToRegex(search);
				var matches = Regex.Match(searchIn.Substring(startIndex), pattern, RegexOptions.IgnoreCase);
				if (matches == Match.Empty)
					return base.CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
				else
					result = matches.Index + startIndex;
			}
			// Otherwise, find the index of the search string.
			else
			{
				result = searchIn.IndexOf(search, startIndex, System.StringComparison.OrdinalIgnoreCase);
				if (result == -1)
					return base.CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
			}
			// Adding 1 because Excel uses 1-based index
			return base.CreateResult(result + 1, DataType.Integer);
		}
	}
}