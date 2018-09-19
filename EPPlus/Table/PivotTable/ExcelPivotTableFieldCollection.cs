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
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection class for fields in a pivot table.
	/// </summary>
	public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
	{
		#region Properties
		/// <summary>
		/// Gets the field accessed by the name.
		/// </summary>
		/// <param name="name">The name of the field.</param>
		/// <returns>The specified field or null if it does not exist.</returns>
		public ExcelPivotTableField this[string name]
		{
			get
			{
				foreach (var field in myList)
				{
					if (field.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
						return field;
				}
				return null;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="table">The existing pivot table.</param>
		/// <param name="topNode">The text of the top node in the xml.</param>
		internal ExcelPivotTableFieldCollection(ExcelPivotTable table, string topNode) :
			 base(table)
		{

		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Returns the date group field.
		/// </summary>
		/// <param name="groupBy">The type of grouping.</param>
		/// <returns>The matching field or null if none is found.</returns>
		public ExcelPivotTableField GetDateGroupField(eDateGroupBy groupBy)
		{
			foreach (var fld in myList)
			{
				if (fld.Grouping is ExcelPivotTableFieldDateGroup && (((ExcelPivotTableFieldDateGroup)fld.Grouping).GroupBy) == groupBy)
					return fld;
			}
			return null;
		}

		/// <summary>
		/// Returns the numeric group field.
		/// </summary>
		/// <returns>The matching field or null if none is found.</returns>
		public ExcelPivotTableField GetNumericGroupField()
		{
			foreach (var fld in myList)
			{
				if (fld.Grouping is ExcelPivotTableFieldNumericGroup)
					return fld;
			}
			return null;
		}
		#endregion
	}
}