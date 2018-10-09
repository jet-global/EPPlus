/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Jan Källman, Evan Schallerer, and others as noted in the source history.
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
using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Abstract base collection class for pivot table fields.
	/// </summary>
	/// <typeparam name="T">An instance of {T}.</typeparam>
	public abstract class ExcelPivotTableFieldCollectionBase<T> : XmlCollectionBase<T> where T : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// The pivot table.
		/// </summary>
		protected ExcelPivotTable PivotTable { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		internal ExcelPivotTableFieldCollectionBase(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) 
			: base(namespaceManager, node)
		{
			if (table == null)
				throw new ArgumentNullException(nameof(table));
			this.PivotTable = table;
		}
		#endregion
	}
}