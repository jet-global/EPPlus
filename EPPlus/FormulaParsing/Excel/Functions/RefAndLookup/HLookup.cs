﻿/* Copyright (C) 2011  Jan Källman
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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	public class HLookup : LookupFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 3, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var lookupArgs = new LookupArguments(arguments, context);
			if (lookupArgs.LookupIndex < 1)
				return new CompileResult(eErrorType.Value);
			var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, lookupArgs, context);
			if (arguments.Count() > 3 && arguments.ElementAt(3).Value is bool rangeLookup && !rangeLookup)
				return Lookup(navigator, lookupArgs, new WildCardValueMatcher());
			else 
				return Lookup(navigator, lookupArgs, new LookupValueMatcher());
		}
	}
}
