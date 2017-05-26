
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
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// 
	/// </summary>
	public class Hour : ExcelFunction
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var dateObj = arguments.ElementAt(0).Value;
			if (ConvertUtil.TryParseDateObjectToOADate(dateObj, out double OADate))
			{
				if (OADate - System.Math.Truncate(OADate) > 0.999988425925926)
					OADate = System.Math.Round(OADate, 5);
				if (OADate < 0)
					return new CompileResult(eErrorType.Num);
				var date = System.DateTime.FromOADate(OADate);
				return this.CreateResult(date.Hour, DataType.Integer);
			}
			else
				return new CompileResult(eErrorType.Value);
		}

	}
}
