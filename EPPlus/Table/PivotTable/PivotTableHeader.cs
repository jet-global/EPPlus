/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A row label or column header object used to calculate pivot table cell values.
	/// </summary>
	internal class PivotTableHeader
	{
		#region Properties
		/// <summary>
		/// Gets the list of <see cref="ExcelPivotCacheRecords"/> indices.
		/// </summary>
		public List<Tuple<int, int>> CacheRecordIndices { get; }
		 
		/// <summary>
		/// Gets the related <see cref="ExcelPivotTableField"/>.
		/// </summary>
		public ExcelPivotTableField PivotTableField { get; }
		
		/// <summary>
		/// Gets the index of the <see cref="ExcelPivotTableDataField"/> in the collection.
		/// </summary>
		public int DataFieldCollectionIndex { get; }

		/// <summary>
		/// Gets the flag indicating if this is a grand total object.
		/// </summary>
		public bool IsGrandTotal { get; }

		/// <summary>
		/// Gets the flag indicating if this is a row label.
		/// </summary>
		public bool IsRowHeader { get; }

		/// <summary>
		/// Gets the value if there is a <see cref="RowColumnItem"/>  with a non-null itemType, typically 'default'.
		/// </summary>
		public string SumType { get; }

		/// <summary>
		/// Gets the flag indicating if this is the inner-most header. Excludes subtotal and grand total nodes.
		/// </summary>
		public bool IsLeafNode { get; }

		/// <summary>
		/// Gets the flag indicating if this is above a data field header. Only used for row <see cref="PivotTableHeader"/>s.
		/// </summary>
		public bool IsAboveDataField { get; }

		/// <summary>
		/// Gets the flag indicating if this is a data field header.
		/// </summary>
		public bool IsDataField { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new <see cref="PivotTableHeader"/> object.
		/// </summary>
		/// <param name="recordIndices">The list of cacheRecord indices.</param>
		/// <param name="field">The pivot table field.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the collection.</param>
		/// <param name="isGrandTotal">A value indicating if it is a grand total.</param>
		/// <param name="isRowHeader">A value indicating if it is a row header.</param>
		/// <param name="isLeafNode">A value indicating if it is a leaf node.</param>
		/// <param name="isDataField">A value indicating if it is a data field node.</param>
		/// <param name="sumType">The itemType value of the <see cref="RowColumnItem"/>.</param>
		/// <param name="isAboveDataField">A value indicating if it is above a data field node.</param>
		public PivotTableHeader(List<Tuple<int, int>> recordIndices, ExcelPivotTableField field, int dataFieldCollectionIndex, bool isGrandTotal,
			bool isRowHeader, bool isLeafNode, bool isDataField, string sumType = null, bool isAboveDataField = false)
		{
			this.CacheRecordIndices = recordIndices;
			this.PivotTableField = field;
			this.DataFieldCollectionIndex = dataFieldCollectionIndex;
			this.IsGrandTotal = isGrandTotal;
			this.IsRowHeader = isRowHeader;
			this.IsLeafNode = isLeafNode;
			this.IsDataField = isDataField;
			this.IsAboveDataField = isAboveDataField;
			this.SumType = sumType;
		}
		#endregion
	}
}