/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau, Evan Schallerer, and others as noted in the source history.
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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// An Excel PivotCacheRecords.
	/// </summary>
	public class ExcelPivotCacheRecords : XmlHelper, IEnumerable<CacheRecordNode>
	{
		#region Constants
		private const string Name = "pivotCacheRecords";
		#endregion

		#region Class Variables
		private List<CacheRecordNode> myRecords = new List<CacheRecordNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets the uri.
		/// </summary>
		public Uri Uri { get; }

		/// <summary>
		/// Gets the cache records xml document.
		/// </summary>
		public XmlDocument CacheRecordsXml { get; }
		
		/// <summary>
		/// Gets the count of total records.
		/// </summary>
		public int Count
		{
			get { return base.GetXmlNodeInt("@count"); }
			set { base.SetXmlNodeString("@count", value.ToString()); }
		}

		/// <summary>
		/// Gets the <see cref="CacheRecordNode"/> at the given index.
		/// </summary>
		/// <param name="Index">The position in the list.</param>
		/// <returns>A <see cref="CacheRecordNode"/>.</returns>
		public CacheRecordNode this[int Index]
		{
			get
			{ return this.Records[Index]; }
		}

		/// <summary>
		/// Gets or sets the reference to the internal package part.
		/// </summary>
		internal Packaging.ZipPackagePart Part { get; set; }

		private ExcelPivotCacheDefinition CacheDefinition { get; }

		private List<CacheRecordNode> Records
		{
			get
			{
				if (myRecords == null)
				{
					var cacheRecordNodes = this.TopNode.SelectNodes("d:r", base.NameSpaceManager);
					foreach (XmlNode recordsNode in cacheRecordNodes)
					{
						myRecords.Add(new CacheRecordNode(base.NameSpaceManager, recordsNode));
					}
				}
				return myRecords;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an existing <see cref="ExcelPivotCacheRecords"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="package">The Excel package.</param>
		/// <param name="cacheRecordsXml">The <see cref="ExcelPivotCacheRecords"/> xml document.</param>
		/// <param name="targetUri">The <see cref="ExcelPivotCacheRecords"/> target uri.</param>
		/// <param name="cacheDefinition">The cache definition of the pivot table.</param>
		public ExcelPivotCacheRecords(XmlNamespaceManager ns, ExcelPackage package, XmlDocument cacheRecordsXml, Uri targetUri, ExcelPivotCacheDefinition cacheDefinition) : base(ns, null)
		{
			if (ns == null)
				throw new ArgumentNullException(nameof(ns));
			if (cacheRecordsXml == null)
				throw new ArgumentNullException(nameof(cacheRecordsXml));
			if (targetUri == null)
				throw new ArgumentNullException(nameof(targetUri));
			if (cacheDefinition == null)
				throw new ArgumentNullException(nameof(cacheDefinition));
			this.CacheRecordsXml = cacheRecordsXml;
			base.TopNode = cacheRecordsXml.SelectSingleNode($"d:{ExcelPivotCacheRecords.Name}", ns);
			this.Uri = targetUri;
			this.Part = package.Package.GetPart(this.Uri);
			this.CacheDefinition = cacheDefinition;

			var cacheRecordNodes = this.TopNode.SelectNodes("d:r", base.NameSpaceManager);
			foreach (XmlNode record in cacheRecordNodes)
			{
				this.Records.Add(new CacheRecordNode(base.NameSpaceManager, record));
			}
		}

		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotCacheRecords"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="package">The <see cref="Packaging.ZipPackage"/> of the Excel package.</param>
		/// <param name="tableId">The <see cref="ExcelPivotTable"/>'s ID.</param>
		/// <param name="cacheDefinition">The cache definition of the pivot table.</param>
		public ExcelPivotCacheRecords(XmlNamespaceManager ns, Packaging.ZipPackage package, ref int tableId, ExcelPivotCacheDefinition cacheDefinition) : base(ns, null)
		{
			if (ns == null)
				throw new ArgumentNullException(nameof(ns));
			if (package == null)
				throw new ArgumentNullException(nameof(package));
			if (cacheDefinition == null)
				throw new ArgumentNullException(nameof(cacheDefinition));
			if (tableId < 1)
				throw new ArgumentOutOfRangeException(nameof(tableId));
			// CacheRecord. Create an empty one.
			this.Uri = XmlHelper.GetNewUri(package, $"/xl/pivotCache/{ExcelPivotCacheRecords.Name}{{0}}.xml", ref tableId);
			var cacheRecord = new XmlDocument();
			cacheRecord.LoadXml($"<{ExcelPivotCacheRecords.Name} xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
			this.Part = package.CreatePart(this.Uri, ExcelPackage.schemaPivotCacheRecords);
			this.CacheRecordsXml = cacheRecord;
			cacheRecord.Save(this.Part.GetStream());

			base.TopNode = cacheRecord.FirstChild;
			this.CacheDefinition = cacheDefinition;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Update the <see cref="CacheItem"/>s.
		/// </summary>
		/// <param name="sourceDataRange">The source range of the data without header row.</param>
		public void UpdateRecords(ExcelRangeBase sourceDataRange)
		{
			// Remove extra records.
			if (sourceDataRange.Rows < this.Records.Count)
			{
				for (int i = this.Records.Count - 1; i >= sourceDataRange.Rows; i--)
				{
					this.Records[i].Remove(base.TopNode);
				}
				var count = this.Records.Count - sourceDataRange.Rows;
				if (count > 0)
					this.Records.RemoveRange(sourceDataRange.Rows, count);
				else
					this.Records.RemoveAt(sourceDataRange.Rows);
			}

			for (int row = sourceDataRange.Start.Row; row < sourceDataRange.Rows + sourceDataRange.Start.Row; row++)
			{
				int recordIndex = row - sourceDataRange.Start.Row;
				var rowCells = new List<object>();
				for (int column = sourceDataRange.Start.Column; column < sourceDataRange.End.Column + 1; column++)
				{
					rowCells.Add(sourceDataRange.Worksheet.Cells[row, column].Value);
				}
				// If the row is within the existing range of cacheRecords, update that cacheRecord. Otherwise, add a new record.
				if (recordIndex < this.Records.Count)
					this.Records[recordIndex].Update(rowCells, this.CacheDefinition);
				else
					this.Records.Add(new CacheRecordNode(this.NameSpaceManager, base.TopNode, rowCells, this.CacheDefinition));
			}
			this.Count = this.Records.Count;
		}

		/// <summary>
		/// Checks if a given row item exists.
		/// </summary>
		/// <param name="nodeIndices">A list of tuples containing the pivotField index and item value.</param>
		/// <param name="pageFieldIndices">A dictionary of page field (filter) indices. Maps a cache field to a list of selected filter item indices.</param>
		/// <returns>True if the item exists, otherwise false.</returns>
		public bool Contains(List<Tuple<int, int>> nodeIndices, Dictionary<int, List<int>> pageFieldIndices)
		{
			return this.Records.Any(r => this.FindCacheRecordIndexAndTupleIndexMatch(nodeIndices, r, pageFieldIndices));
		}

		/// <summary>
		/// Calculate the total of a specified data field for a row/columnn header for custom sorting.
		/// </summary>
		/// <param name="tupleList">A list of tuples containing the pivotField index and item value.</param>
		/// <param name="dataFieldIndex">The index of the referenced data field.</param>
		/// <returns>The calculated total.</returns>
		public double CalculateSortingValues(List<Tuple<int, int>> tupleList, int dataFieldIndex)
		{
			double total = 0;
			foreach (var record in this.Records)
			{
				var findMatch = this.FindCacheRecordIndexAndTupleIndexMatch(tupleList, record);
				if (findMatch)
					total += double.Parse(record.Items[dataFieldIndex].Value);
			}
			return total;
		}

		/// <summary>
		/// Calculate the values for each cell in the pivot table by de-referencing the tuple using the cache definition if a pivot table is given.
		/// Otherwise, calculate the values for each cell in the pivot table for GetPivotData.
		/// </summary>
		/// <param name="rowTuples">The list of rowItem indices.</param>
		/// <param name="columnTuples">The list of columnItem indices.</param>
		/// <param name="filterIndices">A dictionary of page field (filter) indices. Maps a cache field to a list of selected filter item indices.</param>
		/// <param name="dataFieldIndex">The index of the data field.</param>
		/// <param name="pivotTable">The pivot table (optional).</param>
		/// <returns>The subtotal value or null if no values are found.</returns>
		public List<object> FindMatchingValues(List<Tuple<int, int>> rowTuples, List<Tuple<int, int>> columnTuples, 
			Dictionary<int, List<int>> filterIndices, int dataFieldIndex, ExcelPivotTable pivotTable = null)
		{
			if (!string.IsNullOrEmpty(this.CacheDefinition.CacheFields[dataFieldIndex].Formula))
				return null;
			var matchingValues = new List<object>();
			foreach (var record in this.Records)
			{
				bool match = false;
				if (rowTuples != null)
					match = pivotTable == null ? this.FindCacheRecordIndexAndTupleIndexMatch(rowTuples, record) : this.FindCacheRecordValueAndTupleValueMatch(rowTuples, record, pivotTable);
				if (match && columnTuples != null)
					match = pivotTable == null ? this.FindCacheRecordIndexAndTupleIndexMatch(columnTuples, record) : this.FindCacheRecordValueAndTupleValueMatch(columnTuples, record, pivotTable);
				if (match && filterIndices != null)
					match = this.FindCacheRecordValueAndPageFieldTupleValueMatch(filterIndices, record);
				if (match)
					this.AddToList(record, dataFieldIndex, matchingValues);
			}
			return matchingValues;
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Saves the cacheRecords xml.
		/// </summary>
		internal void Save()
		{
			this.CacheRecordsXml.Save(this.Part.GetStream(System.IO.FileMode.Create));
		}
		#endregion

		#region Private Methods
		private bool FindCacheRecordIndexAndTupleIndexMatch(IEnumerable<Tuple<int, int>> indexTupleList, CacheRecordNode record, Dictionary<int, List<int>> pageFieldIndices = null)
		{
			var indexTupleMatch = indexTupleList.All(i => i.Item1 == -2 || int.Parse(record.Items[i.Item1].Value) == i.Item2);
			// If a match was found and page field indices are specified, they must also match the record's values.
			if (indexTupleMatch && (pageFieldIndices == null || this.FindCacheRecordValueAndPageFieldTupleValueMatch(pageFieldIndices, record)))
				return true;
			return false;
		}

		private bool FindCacheRecordValueAndTupleValueMatch(List<Tuple<int, int>> list, CacheRecordNode record, ExcelPivotTable pivotTable)
		{
			foreach (var tuple in list)
			{
				if (tuple.Item1 == -2)
					continue;
				var sharedItems = this.CacheDefinition.CacheFields[tuple.Item1].SharedItems;
				int recordValue = int.Parse(record.Items[tuple.Item1].Value);
				int pivotFieldValue = pivotTable.Fields[tuple.Item1].Items[tuple.Item2].X;
				if (sharedItems[recordValue].Value != sharedItems[pivotFieldValue].Value)
					return false;
			}
			return true;
		}

		private bool FindCacheRecordValueAndPageFieldTupleValueMatch(Dictionary<int, List<int>> pageFieldIndices, CacheRecordNode record)
		{
			// At least one of the page field's items must match to succeed.
			bool allMatch = true;
			foreach (var pageField in pageFieldIndices)
			{
				bool pageFieldMatch = false;
				int recordValue = int.Parse(record.Items[pageField.Key].Value);
				foreach (var item in pageField.Value)
				{
					if (recordValue == item)
					{
						pageFieldMatch = true;
						break;
					}
				}
				if (pageFieldMatch == false)
					return false;
			}
			return allMatch;
		}

		private void AddToList(CacheRecordNode record, int dataFieldIndex, List<object> matchingValues)
		{
			string itemValue = null;
			if (record.Items[dataFieldIndex].Type == PivotCacheRecordType.x)
			{
				int sharedItemIndex = int.Parse(record.Items[dataFieldIndex].Value);
				var cacheField = this.CacheDefinition.CacheFields[dataFieldIndex];
				itemValue = cacheField.SharedItems[sharedItemIndex].Value;
			}
			else
				itemValue = record.Items[dataFieldIndex].Value;
			double.TryParse(itemValue, out var recordData);
			matchingValues.Add(recordData);
		}
		#endregion

		#region IEnumerable Methods
		/// <summary>
		/// Gets the CacheItem enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator<CacheRecordNode> GetEnumerator()
		{
			return myRecords.GetEnumerator();
		}

		/// <summary>
		/// Gets the specified type enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		IEnumerator IEnumerable.GetEnumerator()
		{
			return myRecords.GetEnumerator();
		}
		#endregion
	}
}