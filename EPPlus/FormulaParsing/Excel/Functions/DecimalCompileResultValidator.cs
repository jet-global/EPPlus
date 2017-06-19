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
 * Mats Alm   		                Added		                2013-12-26
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
	public class DecimalCompileResultValidator : CompileResultValidator
	{
		/// <summary>
		/// Checks for a #Num Excel error.
		/// </summary>
		/// <param name="obj">The excel object to check for a num error.</param>
		/// <param name="error">Sends out the Error type.</param>
		/// <returns>Returns true or false.</returns>
		public override bool TryGetValidationError(object obj, out eErrorType error)
		{
			var num = ConvertUtil.GetValueDouble(obj);
			if (double.IsNaN(num) || double.IsInfinity(num))
			{
				error = eErrorType.Num;
				return false;
			}
			//If this returns true ignore the out eError
			error = eErrorType.Null;
			return true;
		}
		/// <summary>
		/// Throws a #Num exception.
		/// </summary>
		/// <param name="obj">The excel object to check for a num error.</param>
		public override void Validate(object obj)
		{
			var num = ConvertUtil.GetValueDouble(obj);
			if (double.IsNaN(num) || double.IsInfinity(num))
			{
				throw new ExcelErrorValueException(eErrorType.Num);
			}
		}

	}
}
