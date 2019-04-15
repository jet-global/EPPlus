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
using System.Threading;
using System.Xml;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Extensions;

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
		/// <param name="logger">The logger to use to log method calls.</param>
		public void UpdateRecords(ExcelRangeBase sourceDataRange, IFormulaParserLogger logger)
		{
			logger?.LogFunction(nameof(this.UpdateRecords));
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
				int cacheFieldIndex = 0;
				for (int column = sourceDataRange.Start.Column; column < sourceDataRange.End.Column + 1; column++)
				{
					var cacheField = this.CacheDefinition.CacheFields[cacheFieldIndex];
					var cell = sourceDataRange.Worksheet.Cells[row, column];
					// If the cell value is a DateTime, convert it to an date.
					if (cacheField.SharedItems.ContainsDate == true && cell.Value is double)
						rowCells.Add(DateTime.FromOADate((double)cell.Value));
					else
						rowCells.Add(cell.Value);
					cacheFieldIndex++;
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
		/// Calculate the total of a specified data field for a row/columnn header for custom sorting.
		/// </summary>
		/// <param name="node">The current node that is being evaluated.</param>
		/// <param name="dataFieldIndex">The index of the referenced data field.</param>
		/// <returns>The calculated total.</returns>
		public double CalculateSortingValues(PivotItemTreeNode node, int dataFieldIndex)
		{
			double sortingTotal = 0;
			foreach (var record in node.CacheRecordIndices)
			{
				string dataFieldValue = this.Records[record].Items[dataFieldIndex].Value;
				sortingTotal += double.Parse(dataFieldValue);
			}
			return sortingTotal;
		}

		/// <summary>
		/// Get a list of data field values at the given record indices to use for sorting calculated fields.
		/// </summary>
		/// <param name="indices">A list of record indices.</param>
		/// <param name="dataFieldIndex">The index of the referenced data field.</param>
		/// <returns>A list of data field values.</returns>
		public List<object> GetChildDataFieldValues(List<int> indices, int dataFieldIndex)
		{
			var values = new List<object>();
			for (int i = 0; i < indices.Count; i++)
			{
				var record = this.Records[indices[i]];
				double.TryParse(record.Items[dataFieldIndex].Value, out var recordValue);
				values.Add(recordValue);
			}
			return values;
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
			var matchingValues = new List<object>();
			foreach (var record in this.Records)
			{
				bool match = false;
				if (rowTuples != null)
					match = pivotTable == null ? this.FindCacheRecordIndexAndTupleIndexMatch(rowTuples, record) : this.FindCacheRecordValueAndTupleValueMatch(rowTuples, record, pivotTable);
				if ((match && columnTuples != null) || rowTuples == null)
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
				// Ignore data field tuples or group pivot field tuples.
				if (tuple.Item1 == -2)
					continue;

				var cacheField = this.CacheDefinition.CacheFields[tuple.Item1];
				if (cacheField.IsGroupField)
				{
					bool groupMatch = this.FindGroupingRecordValueAndTupleMatch(cacheField, record, tuple, pivotTable);
					if (!groupMatch)
						return false;
				}
				else
				{
					var sharedItems = this.CacheDefinition.CacheFields[tuple.Item1].SharedItems;
					int recordValue = int.Parse(record.Items[tuple.Item1].Value);
					int pivotFieldValue = pivotTable.Fields[tuple.Item1].Items[tuple.Item2].X;

					if (pivotTable.HasFilters)
					{
						foreach (var filter in pivotTable.Filters)
						{
							int filterFieldIndex = filter.Field;
							int recordReferenceIndex = int.Parse(record.Items[filterFieldIndex].Value);
							string recordReferenceString = this.CacheDefinition.CacheFields[filterFieldIndex].SharedItems[recordReferenceIndex].Value;
							bool isNumeric = this.CacheDefinition.CacheFields[filterFieldIndex].SharedItems[recordReferenceIndex].Type == PivotCacheRecordType.n;
							bool isMatch = filter.MatchesFilterCriteriaResult(recordReferenceString, isNumeric);
							if (!isMatch)
								return false;
						}
					}

					if (!sharedItems[recordValue].Value.IsEquivalentTo(sharedItems[pivotFieldValue].Value))
						return false;
				}
			}
			return true;
		}

		private bool FindGroupingRecordValueAndTupleMatch(CacheFieldNode cacheField, CacheRecordNode record, Tuple<int, int> tuple, ExcelPivotTable pivotTable)
		{
			if (cacheField.FieldGroup.GroupBy != PivotFieldDateGrouping.None)
			{
				// Find record indices for date groupings fields.
				var recordIndices = this.DateGroupingRecordValueTupleMatch(cacheField, tuple.Item2);
				// If the record value is in the list, then the record value and tuple are a match.
				int index = tuple.Item1 < record.Items.Count ? tuple.Item1 : cacheField.FieldGroup.BaseField;
				var itemValue = record.Items[index].Value;
				if (recordIndices.All(i => i != int.Parse(itemValue)))
					return false;
			}
			else
			{
				// Use discrete grouping property collection to determine match.
				int baseIndex = cacheField.FieldGroup.BaseField;
				// Get the pivot field item's x value and the record item's v value.
				int pivotFieldValue = pivotTable.Fields[tuple.Item1].Items[tuple.Item2].X;
				var recordValue = int.Parse(record.Items[baseIndex].Value);
				var fieldGroup = cacheField.FieldGroup;
				// Get the shared item string the pivot field item and record item is pointing to.
				var pivotFieldPtrValue = fieldGroup.GroupItems[pivotFieldValue].Value;
				var recordDiscretePtrValue = int.Parse(fieldGroup.DiscreteGroupingProperties[recordValue].Value);
				var recordPtrValue = fieldGroup.GroupItems[recordDiscretePtrValue].Value;
				// Check if the pivot field item and record item is pointing to the same shared string.
				if (!pivotFieldPtrValue.IsEquivalentTo(recordPtrValue))
					return false;
			}
			return true;
		}

		private List<int> DateGroupingRecordValueTupleMatch(CacheFieldNode cacheField, int tupleItem2)
		{
			var recordIndices = new List<int>();
			// Go through all the shared items and if the item is the targeted value (tuple.Item2 or groupItems[tuple.Item2]), 
			// add the index of the shared item to the list.
			var groupByType = cacheField.FieldGroup.GroupBy;
			string groupFieldItemsValue = cacheField.FieldGroup.GroupItems[tupleItem2].Value;

			int baseFieldIndex = cacheField.FieldGroup.BaseField;
			var sharedItems = this.CacheDefinition.CacheFields[baseFieldIndex].SharedItems;

			for (int i = 0; i < sharedItems.Count; i++)
			{
				var dateTime = DateTime.Parse(sharedItems[i].Value);
				var groupByValue = string.Empty;

				// Get the sharedItem's groupBy value, unless the groupBy value is months.
				if (groupByType == PivotFieldDateGrouping.Months && tupleItem2 == dateTime.Month)
					recordIndices.Add(i);
				else if (groupByType == PivotFieldDateGrouping.Years)
					groupByValue = dateTime.Year.ToString();
				else if (groupByType == PivotFieldDateGrouping.Quarters)
					groupByValue = "Qtr" + ((dateTime.Month - 1) / 3 + 1);
				else if (groupByType == PivotFieldDateGrouping.Days)
					groupByValue = dateTime.Day + "-" + dateTime.ToString("MMM");
				else if (groupByType == PivotFieldDateGrouping.Minutes)
					groupByValue = ":" + dateTime.ToString("mm");
				else if (groupByType == PivotFieldDateGrouping.Seconds)
					groupByValue = ":" + dateTime.ToString("ss");
				else if (groupByType == PivotFieldDateGrouping.Hours)
				{
					int hour = dateTime.Hour == 00 ? 12 : dateTime.Hour;
					groupByValue= hour + " " + dateTime.ToString("tt", Thread.CurrentThread.CurrentCulture);
				}

				// Check if the sharedItem's groupBy value matches the groupFieldItem's value in the cacheField.
				if (groupFieldItemsValue.IsEquivalentTo(groupByValue))
					recordIndices.Add(i);
			}
			return recordIndices;
		}

		private bool FindCacheRecordValueAndPageFieldTupleValueMatch(Dictionary<int, List<int>> pageFieldIndices, CacheRecordNode record)
		{
			// At least one of the page field's items must match to succeed.
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
			return true;
		}

		private void AddToList(CacheRecordNode record, int dataFieldIndex, List<object> matchingValues)
		{
			string itemValue = null;
			PivotCacheRecordType type = record.Items[dataFieldIndex].Type;
			if (type == PivotCacheRecordType.x)
			{
				int sharedItemIndex = int.Parse(record.Items[dataFieldIndex].Value);
				var cacheField = this.CacheDefinition.CacheFields[dataFieldIndex];
				var sharedItem = cacheField.SharedItems[sharedItemIndex];
				type = sharedItem.Type;
				itemValue = sharedItem.Value;
			}
			else
				itemValue = record.Items[dataFieldIndex].Value;
			if (type == PivotCacheRecordType.m)
				matchingValues.Add(null);
			else if (string.IsNullOrWhiteSpace(itemValue))
				matchingValues.Add(itemValue);
			else if (type == PivotCacheRecordType.d)
				matchingValues.Add(DateTime.Parse(itemValue));
			else
			{
				double.TryParse(itemValue, out var recordData);
				matchingValues.Add(recordData);
			}
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