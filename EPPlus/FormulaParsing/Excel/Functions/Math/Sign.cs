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
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for the Sign Excel Function.
	/// </summary>
	public class Sign : ExcelFunction
	{
		/// <summary>
		/// Takes a user specified argument and determines the sign and returns a value accordingly.
		/// </summary>
		/// <param name="arguments">The user specified argument.</param>
		/// <param name="context">Not used, but needed to overrride the method.</param>
		/// <returns>1 if the sign of the argument is positive, -1 if it is negative, and 0 otherwise.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var result = 0d;
			var numberCandidate = arguments.ElementAt(0).Value;

			if (numberCandidate == null)
				return this.CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseObjectToDecimal(numberCandidate, out double doubleArg))
				return new CompileResult(eErrorType.Value);

			var numberValue = doubleArg; 
			if (numberValue < 0)
				result = -1;
			else if (numberValue > 0)
				result = 1;
			return this.CreateResult(result, DataType.Decimal);
		}
	}
}
