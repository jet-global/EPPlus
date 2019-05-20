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
		/// NOTE: Only used for row headers.
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
		/// Gets the value if there is a <see cref="RowColumnItem"/> with a non-null itemType, typically 'default'.
		/// </summary>
		public string TotalType { get; }

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

		/// <summary>
		/// Gets or sets a value indicating that this header is a placeholder for 
		/// when there are no row or column fields.
		/// </summary>
		public bool IsPlaceHolder { get; set; }

		/// <summary>
		/// Gets the level of indentation of the items that correspond to this header.
		/// </summary>
		public int Indent { get; }

		/// <summary>
		/// Gets the list of cache record indices the header uses.
		/// </summary>
		public List<int> UsedCacheRecordIndices { get; }

		/// <summary>
		/// Gets a list of all the conditional formatting rules applied to this header.
		///		Item1 is the priority value of the conditional format.
		///		Item2 is the data field collection index.
		/// </summary>
		public List<Tuple<int, int>> ConditionalFormattingTupleList { get; } = new List<Tuple<int, int>>();
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new <see cref="PivotTableHeader"/> object.
		/// </summary>
		/// <param name="cacheRecordsUsed">The list of cache record indices the header uses.</param>
		/// <param name="recordIndices">The list tuple of cacheRecord indices and pivot field item index.</param>
		/// <param name="field">The pivot table field.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the collection.</param>
		/// <param name="isGrandTotal">A value indicating if it is a grand total.</param>
		/// <param name="isLeafNode">A value indicating if it is a leaf node.</param>
		/// <param name="isDataField">A value indicating if it is a data field node.</param>
		/// <param name="totalType">The itemType value of the <see cref="RowColumnItem"/>.</param>
		/// <param name="isAboveDataField">A value indicating if it is above a data field node.</param>
		/// <param name="indent">The level of indentation for the items that correspond to this header.</param>
		public PivotTableHeader(List<int> cacheRecordsUsed, List<Tuple<int, int>> recordIndices, ExcelPivotTableField field, int dataFieldCollectionIndex, bool isGrandTotal,
			bool isLeafNode, bool isDataField, string totalType = null, bool isAboveDataField = false, int indent = 0)
		{
			this.UsedCacheRecordIndices = cacheRecordsUsed;
			this.CacheRecordIndices = recordIndices;
			this.PivotTableField = field;
			this.DataFieldCollectionIndex = dataFieldCollectionIndex;
			this.IsGrandTotal = isGrandTotal;
			this.IsLeafNode = isLeafNode;
			this.IsDataField = isDataField;
			this.IsAboveDataField = isAboveDataField;
			this.TotalType = totalType;
			this.Indent = indent;
		}
		#endregion
	}
}