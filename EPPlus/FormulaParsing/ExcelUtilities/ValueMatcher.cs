/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
	public class ValueMatcher
	{
		#region Private Properties
		private ArgumentParsers _argumentParsers { get; } = new ArgumentParsers();
		#endregion

		#region Public Methods
		/// <summary>
		/// Compares the values of <paramref name="o1"/> and <paramref name="o2"/>.
		/// </summary>
		/// <param name="o1">The first object.</param>
		/// <param name="o2">The second object.</param>
		/// <returns>
		///	A negative integer if the first object is less than the second, 
		///	a positive integer if the first object is greater than the second, 
		///	0 if the two objects are equivalent,
		///	null if the two objects cannot be compared because of incompatible types.
		/// </returns>
		public virtual int? IsMatch(object o1, object o2)
		{
			if (o1 != null && o2 == null) return 1;
			if (o1 == null && o2 != null) return -1;
			if (o1 == null && o2 == null) return 0;
			try
			{
				//Handle ranges and defined names
				o1 = CheckGetRange(o1);
				o2 = CheckGetRange(o2);

				var o1s = o1 as string;
				var o2s = o2 as string;
				if (o1s != null && o2s != null)
					return this.CompareStringToString(o1s.ToLower(), o2s.ToLower());
				else if (o1s != null)
					return this.CompareStringToObject(o1s, o2);
				else if (o2s != null)
					return this.CompareObjectToString(o1, o2s);
				var decimalParser = _argumentParsers.GetParser(DataType.Decimal);
				var o1d = (double)decimalParser.Parse(o1);
				var o2d = (double)decimalParser.Parse(o2);
				return o1d.CompareTo(o2d);
			}
			catch { /* Ignore any parse errors that may have occurred. */}
			return null;
		}
		#endregion

		#region Protected Methods
		protected virtual int? CompareStringToString(string s1, string s2)
		{
			return s1.CompareTo(s2);
		}

		protected virtual int? CompareStringToObject(string o1, object o2)
		{
			double d1;
			if (double.TryParse(o1, out d1))
			{
				var o2d = this._argumentParsers.GetParser(DataType.Decimal).Parse(o2);
				return d1.CompareTo(o2d);
			}
			bool b1;
			if (bool.TryParse(o1, out b1))
			{
				var o2b = this._argumentParsers.GetParser(DataType.Boolean).Parse(o2);
				return b1.CompareTo(o2b);
			}
			DateTime dt1, dt2;
			if (DateTime.TryParse(o1, out dt1) && DateTime.TryParse(o2.ToString(), out dt2))
				return dt1.CompareTo(dt2);
			return null;
		}

		protected virtual int? CompareObjectToString(object o1, string o2)
		{
			double d2;
			if (double.TryParse(o2, out d2))
			{
				var o1d = (double)this._argumentParsers.GetParser(DataType.Decimal).Parse(o1);
				return o1d.CompareTo(d2);
			}
			bool b2;
			if (bool.TryParse(o2, out b2))
			{
				var o1b = (bool)this._argumentParsers.GetParser(DataType.Boolean).Parse(o1);
				return o1b.CompareTo(b2);
			}
			DateTime dt1, dt2;
			if (DateTime.TryParse(o1.ToString(), out dt1) && DateTime.TryParse(o2, out dt2))
				return dt1.CompareTo(dt2);
			return null;
		}
		#endregion

		#region Private Methods
		private static object CheckGetRange(object v)
		{
			if (v is ExcelDataProvider.IRangeInfo)
			{
				var r = ((ExcelDataProvider.IRangeInfo)v);
				if (r.GetTotalCellCount() > 1)
				{
					v = ExcelErrorValue.Create(eErrorType.NA);
				}
				v = r.GetOffset(0, 0);
			}
			else if (v is ExcelDataProvider.INameInfo)
			{
				var n = ((ExcelDataProvider.INameInfo)v);
				v = CheckGetRange(n);
			}
			return v;
		}
		#endregion
	}
}
