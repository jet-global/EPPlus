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
				{
					if (pivotTable == null)
						match = this.FindCacheRecordIndexAndTupleIndexMatch(rowTuples, record);
					else if (rowTuples.Any(x => x.Item1 >= record.Items.Count) || pivotTable.RowFields.Any(x => x.Index >= record.Items.Count))
						match = this.FindCacheRecordValueAndTupleValueMatch2(rowTuples, record, pivotTable);
					else
						match = this.FindCacheRecordValueAndTupleValueMatch(rowTuples, record, pivotTable);
				}
				if (match && columnTuples != null)
					match = pivotTable == null ? this.FindCacheRecordIndexAndTupleIndexMatch(columnTuples, record) : this.FindCacheRecordValueAndTupleValueMatch(columnTuples, record, pivotTable);
				if (match && filterIndices != null)
					match = this.FindCacheRecordValueAndPageFieldTupleValueMatch(filterIndices, record);
				if (match)
					this.AddToList(record, dataFieldIndex, matchingValues);
			}
			return matchingValues;
			//var matchingValues = new List<object>();
			//foreach (var record in this.Records)
			//{
			//	bool match = true;
			//	if (rowTuples != null)
			//		match = pivotTable == null ? this.FindCacheRecordIndexAndTupleIndexMatch(rowTuples, record) : this.FindCacheRecordValueAndTupleValueMatch(rowTuples, record, pivotTable);
			//	if (match && columnTuples != null)
			//		match = pivotTable == null ? this.FindCacheRecordIndexAndTupleIndexMatch(columnTuples, record) : this.FindCacheRecordValueAndTupleValueMatch(columnTuples, record, pivotTable);
			//	if (match && filterIndices != null)
			//		match = this.FindCacheRecordValueAndPageFieldTupleValueMatch(filterIndices, record);
			//	if (match)
			//		this.AddToList(record, dataFieldIndex, matchingValues);
			//}
			//return matchingValues;
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
			var indexTupleMatch = indexTupleList.All(i => i.Item1 == -2 || i.Item1 >= record.Items.Count || int.Parse(record.Items[i.Item1].Value) == i.Item2);
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

		private bool FindCacheRecordValueAndTupleValueMatch2(List<Tuple<int, int>> list, CacheRecordNode record, ExcelPivotTable pivotTable)
		{
			var indicesMatch = new List<Tuple<int, int>>();
			string yr = "";
			string qtr = "";
			foreach (var tuple in list)
			{
				if (tuple.Item1 == -2)
					continue;
				else if (tuple.Item1 >= this.Records[0].Items.Count)
				{
					// If field groups are used (specifically date field groups for now).
					string cacheFieldGroupName = pivotTable.CacheDefinition.CacheFields[tuple.Item1].Name; // Years
					string value = pivotTable.CacheDefinition.CacheFields[tuple.Item1].FieldGroup.GroupItems[tuple.Item2].Value; // 2016
					int baseFieldIndex = pivotTable.CacheDefinition.CacheFields[tuple.Item1].FieldGroup.BaseField; // 2
					var baseCacheField = pivotTable.CacheDefinition.CacheFields[baseFieldIndex]; // Month
					yr = cacheFieldGroupName.IsEquivalentTo("Years") ? value : yr ;
					qtr = cacheFieldGroupName.IsEquivalentTo("Quarters") ? value : qtr;
					for (int i = 0; i < baseCacheField.SharedItems.Count; i++)
					{
						var dateTimeSplit = baseCacheField.SharedItems[i].Value.Split('-');
						var dateTimeSplit2 = new DateTime(int.Parse(dateTimeSplit[0]), int.Parse(dateTimeSplit[1]), int.Parse(dateTimeSplit[2].Substring(0, 2)));
						if (cacheFieldGroupName.IsEquivalentTo("Years"))
						{
							// If this is an inner node.
							if (tuple == list.Last() && tuple != list.First())
							{
								foreach (var pair in indicesMatch.ToList())
								{
									var itemSplit = baseCacheField.SharedItems[pair.Item2].Value.Split('-');
									// If the year is incorrect, remove it from the list.
									if (!value.IsEquivalentTo(itemSplit[0]))
										indicesMatch.Remove(pair);
								}
								break;
							}
							else
							{
								var tup = new Tuple<int, int>(baseFieldIndex, i);
								int quarter = (dateTimeSplit2.Month - 1) / 3 + 1;
								if (value.IsEquivalentTo(dateTimeSplit[0]) && !indicesMatch.Contains(tup))
								{
									// If Quarters is a parent node of Years, then check that the quarter matches. Otherwise, only check for the year value.
									if (!string.IsNullOrEmpty(qtr) && qtr.IsEquivalentTo("Qtr" + quarter))
									{
										yr = value;
										indicesMatch.Add(tup);
									}
									else if (string.IsNullOrEmpty(qtr) && value.IsEquivalentTo(dateTimeSplit[0]))
									{
										indicesMatch.Add(tup);
									}
								}
								else if (indicesMatch.Contains(tup))
								{
									// If Quarters is a parent node of Years, then check that the quarter matches. Otherwise, only check for the year value.
									if ((!string.IsNullOrEmpty(qtr) && !qtr.IsEquivalentTo("Qtr" + quarter)) || string.IsNullOrEmpty(qtr) && !value.IsEquivalentTo(dateTimeSplit[0]))
									{
										indicesMatch.Remove(tup);
									}
								}
							}
						}
						else if (cacheFieldGroupName.IsEquivalentTo("Quarters"))
						{
							// Since it is the last tuple in the list, go through the indicesMatch list and check that the tuples have the correct month.
							if (tuple == list.Last() && tuple != list.First())
							{
								foreach (var pair in indicesMatch.ToList())
								{
									var itemSplit = baseCacheField.SharedItems[pair.Item2].Value.Split('-');
									int month = int.Parse(itemSplit[1]);
									int quarter = (month - 1) / 3 + 1;
									// If the quarter is incorrect, remove it from the list.
									if (!value.IsEquivalentTo("Qtr" + quarter))
										indicesMatch.Remove(pair);
								}
								break;
							}
							else
							{
								int month = int.Parse(dateTimeSplit[1]);
								int quarter = (month - 1) / 3 + 1;
								var tup = new Tuple<int, int>(baseFieldIndex, i);
								if (value.IsEquivalentTo("Qtr" + quarter) && !indicesMatch.Contains(tup))
								{
									// If Years is a parent node of Quarters, then check that the year matches. Otherwise, only check for the quarter value.
									if (!string.IsNullOrEmpty(yr) && yr.IsEquivalentTo(dateTimeSplit[0]))
									{
										qtr = "Qtr" + quarter;
										indicesMatch.Add(tup);
									}
									else if (string.IsNullOrEmpty(yr) && value.IsEquivalentTo("Qtr" + quarter))
									{
										indicesMatch.Add(tup);
									}
								}
								else if (indicesMatch.Contains(tup))
								{
									if ((!string.IsNullOrEmpty(yr) && !yr.IsEquivalentTo(dateTimeSplit[0])) || string.IsNullOrEmpty(yr) && !value.IsEquivalentTo("Qtr" + quarter))
									{
										indicesMatch.Remove(tup);
									}
								}
							}
						}
					}
				}
				else
				{
					// This is a normal pivot field index.
					if (pivotTable.CacheDefinition.CacheFields[tuple.Item1].Name.IsEquivalentTo("Month"))
					{
						// If Month is the very first node.
						if (tuple == list.First())
						{
							// Get the month from groupItems
							var cacheField = pivotTable.CacheDefinition.CacheFields[tuple.Item1];
							string monthName = cacheField.FieldGroup.GroupItems[tuple.Item2].Value;
							if (list.Count == 1)
							{
								for (int i = 0; i < cacheField.SharedItems.Count; i++)
								{
									var dateSplit = cacheField.SharedItems[i].Value.Split('-');
									var dateTime = new DateTime(int.Parse(dateSplit[0]), int.Parse(dateSplit[1]), int.Parse(dateSplit[2].Substring(0, 2)));
									if (monthName.IsEquivalentTo(dateTime.ToString("MMM")) && int.Parse(record.Items[tuple.Item1].Value) == i)
									{
										return true;
									}
								}
								return false;
							}
							else
							{
								bool matchFound = false;
								for (int i = 0; i < cacheField.SharedItems.Count; i++)
								{
									var dateSplit = cacheField.SharedItems[i].Value.Split('-');
									var dateTime = new DateTime(int.Parse(dateSplit[0]), int.Parse(dateSplit[1]), int.Parse(dateSplit[2].Substring(0, 2)));
									if (monthName.IsEquivalentTo(dateTime.ToString("MMM")) && int.Parse(record.Items[tuple.Item1].Value) == i)
									{
										matchFound = true;
										indicesMatch.Add(new Tuple<int, int>(tuple.Item1, i));
									}
								}
								// If a match is not found, then return false and continue onto the next record.
								if (!matchFound)
									return matchFound;
							}
						}
						// If Month is the last node.
						else if (tuple == list.Last())
						{
							int year = -1;
							string quarter = "";
							foreach (var tup in list)
							{
								if (tup.Item1 >= record.Items.Count)
								{
									if (this.CacheDefinition.CacheFields[tup.Item1].Name.IsEquivalentTo("Years"))
										year = int.Parse(this.CacheDefinition.CacheFields[tup.Item1].FieldGroup.GroupItems[tup.Item2].Value);
									else if (this.CacheDefinition.CacheFields[tup.Item1].Name.IsEquivalentTo("Quarters"))
										quarter = this.CacheDefinition.CacheFields[tup.Item1].FieldGroup.GroupItems[tup.Item2].Value;
								}
							}

							var sharedItems = pivotTable.CacheDefinition.CacheFields[tuple.Item1].SharedItems;

							// If non-date grouping fields are parent nodes of a month field.
							if (indicesMatch.Count == 0)
							{
								string month = pivotTable.CacheDefinition.CacheFields[tuple.Item1].FieldGroup.GroupItems[tuple.Item2].Value; // Mar
								for (int i = 0; i < sharedItems.Count; i++)
								{
									var date = sharedItems[i].Value.Split('-');
									var dateTime = new DateTime(int.Parse(date[0]), int.Parse(date[1]), int.Parse(date[2].Substring(0, 2)));
									int quarterValue = (dateTime.Month - 1) / 3 + 1;
									// Check that the month's are correct and it exists in the cache records.
									// Case 1: If Years and Quarters are a parent node of Month, then also check that the shared item's year and quarter is correct.
									// Case 2: If Quarters is a parent node of Month, then check then also check that the shared item's quarter is correct.
									// Case 3: If Years is a parent node of Month, then check then also check that the shared item's year is correct.
									// Case 4: If both Years and Quarters is not a parent node of Month, then ignore it.
									if (dateTime.Month == tuple.Item2 && int.Parse(record.Items[tuple.Item1].Value) == i &&
										((year == dateTime.Year && quarter.IsEquivalentTo("Qtr" + quarterValue)
										|| (year == -1 & quarter.IsEquivalentTo("Qtr" + quarterValue))
										|| (string.IsNullOrEmpty(quarter) && year == dateTime.Year))
										|| (year == -1 && string.IsNullOrEmpty(quarter))))
									{
										return true;
									}
								}
							}
							else
							{
								for (int i = 0; i < indicesMatch.Count; i++)
								{
									var item = sharedItems[indicesMatch[i].Item2];
									var date = item.Value.Split('-');
									var dateTime = new DateTime(int.Parse(date[0]), int.Parse(date[1]), int.Parse(date[2].Substring(0, 2)));
									int quarterValue = (dateTime.Month - 1) / 3 + 1;
									// Check that the month's are correct and it exists in the cache records.
									// Case 1: If Years and Quarters are a parent node of Month, then also check that the shared item's year and quarter is correct.
									// Case 2: If Quarters is a parent node of Month, then check then also check that the shared item's quarter is correct.
									// Case 3: If Years is a parent node of Month, then check then also check that the shared item's year is correct.
									if (dateTime.Month == tuple.Item2 && int.Parse(record.Items[tuple.Item1].Value) == indicesMatch[i].Item2 &&
										((year == dateTime.Year && quarter.IsEquivalentTo("Qtr" + quarterValue)
										|| (year == -1 & quarter.IsEquivalentTo("Qtr" + quarterValue))
										|| (string.IsNullOrEmpty(quarter) && year == dateTime.Year))))
									{
										return true;
									}
								}
							}
							
							return false;
						}
						// If Month is an inner node.
						else
						{
							var sharedItems = pivotTable.CacheDefinition.CacheFields[tuple.Item1].SharedItems;

							// If Month is a parent node of Years/Quarters and a child node of a non-date grouping field, then add to the indicesMatch list.
							if (indicesMatch.Count == 0)
							{
								string month = pivotTable.CacheDefinition.CacheFields[tuple.Item1].FieldGroup.GroupItems[tuple.Item2].Value; // Mar
								for (int i = 0; i < sharedItems.Count; i++)
								{
									var date = sharedItems[i].Value.Split('-');
									var dateTime = new DateTime(int.Parse(date[0]), int.Parse(date[1]), int.Parse(date[2].Substring(0, 2)));
									if (dateTime.Month == tuple.Item2 && int.Parse(record.Items[tuple.Item1].Value) == i)
									{
										indicesMatch.Add(new Tuple<int, int>(tuple.Item1, i));
									}
								}
							}
							else
							{
								int index = 0;
								while (index < indicesMatch.Count)
								{
									var item = sharedItems[indicesMatch[index].Item2];
									var date = item.Value.Split('-');
									int month = int.Parse(date[1]);
									if (month != tuple.Item2 || int.Parse(record.Items[tuple.Item1].Value) != indicesMatch[index].Item2)
									{
										indicesMatch.Remove(indicesMatch[index]);
									}
									else
										index++;
								}
							}
							
							if (indicesMatch.Count == 0)
								break;
						}
					}
					else
					{
						
						var sharedItems = this.CacheDefinition.CacheFields[tuple.Item1].SharedItems;
						int recordValue = int.Parse(record.Items[tuple.Item1].Value);
						int pivotFieldValue = pivotTable.Fields[tuple.Item1].Items[tuple.Item2].X;
						if (sharedItems[recordValue].Value != sharedItems[pivotFieldValue].Value)
							return false;
						if (list.Count == 1)
							return true;
					}
				}
			}

			var duplicates = indicesMatch.GroupBy(x => x.Item1).Where(g => g.Count() > 1).Select(y => y.Key);
			if (indicesMatch.Count == 0)
				return false;
			foreach (var pair in indicesMatch)
			{
				int item1 = pair.Item1;
				if (duplicates.Contains(item1))
				{
					var item2Values = indicesMatch.FindAll(j => j.Item1 == item1).Select(x => x.Item2);
					int recordValue = int.Parse(record.Items[item1].Value);
					if (!item2Values.Contains(recordValue))
						return false;
				}
				else
				{
					int recordValue = int.Parse(record.Items[pair.Item1].Value);
					if (recordValue != pair.Item2)
						return false;
				}
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