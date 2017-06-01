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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
	/// <summary>
	/// This class contains the formula for parsing objects to boolean type.
	/// </summary>
	public class BoolArgumentParser : ArgumentParser
	{
		/// <summary>
		/// Parse takes an object and attempts to parse it to a boolean type. 
		/// </summary>
		/// <param name="obj">The object to be parsed.</param>
		/// <returns>The resulting boolean object if successful and false if it was not.</returns>
		public override object Parse(object obj)
		{
			if (obj is ExcelDataProvider.IRangeInfo)
			{
				var r = ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault();
				obj = (r == null ? null : r.Value);
			}
			if (obj == null) return false;
			if (obj is bool) return (bool)obj;
			if (obj.IsNumeric()) return Convert.ToBoolean(obj);
			bool result;
			if (bool.TryParse(obj.ToString(), out result))
			{
				return result;
			}
			return result;
		}
	}
}