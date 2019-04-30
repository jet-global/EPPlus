/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * EPPlus Copyright (C) 2011 Jan Källman.
 * This File Copyright (C) 2016 Matt Delaney.
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
 * Author                      Change                            Date
 * ******************************************************************************
 * Matt Delaney                Added support for slicers.        11 October 2016
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Sorting;

namespace OfficeOpenXml.Drawing.Slicers
{	
	/// <summary>
	/// Represents an Excel Slicer Cache.
	/// When this part exists, it can be found at /xl/slicerCaches/slicerCacheN.xml.
	/// </summary>
	public class ExcelSlicerCache : XmlHelper
	{
		#region Class Variables
		private List<PivotTableNode> myPivotTables = new List<PivotTableNode>();
		private List<OlapDataNode> myOlapDataSources = new List<OlapDataNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the name of this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public string Name
		{
			get { return base.GetXmlNodeString("@name"); }
			set
			{
				base.SetXmlNodeString("@name", value);
				if (this.Slicer != null)
					this.Slicer.SetXmlNodeString("@cache", value);
			}
		}

		/// <summary>
		/// Gets or sets the source name of this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public string SourceName
		{
			get { return base.GetXmlNodeString("@sourceName"); }
			set { base.SetXmlNodeString("@sourceName", value); }
		}

		/// <summary>
		/// Gets a value indicating that the "Hide items with no data" setting is checked.
		/// </summary>
		public bool HideItemsWithNoData
		{
			get { return base.TopNode.SelectSingleNode("default:extLst/x:ext/x15:slicerCacheHideItemsWithNoData", base.NameSpaceManager) != null; }
		}

		/// <summary>
		/// Gets or sets the <see cref="ExcelSlicer"/> that uses this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public ExcelSlicer Slicer { get; set; }

		/// <summary>
		/// Gets the Uri of the Slicer Cache's associated XML part.
		/// </summary>
		public Uri SlicerCacheUri { get; private set; }

		/// <summary>
		/// Gets a readonly list of <see cref="PivotTableNode"/>s that wrap the <pivotTable /> element.
		/// </summary>
		public IReadOnlyList<PivotTableNode> PivotTables
		{
			get { return myPivotTables; }
		}

		/// <summary>
		/// Gets a readonly list of olap data sources for this slicer cache.
		/// </summary>
		public IReadOnlyList<OlapDataNode> OlapDataSources
		{
			get { return myOlapDataSources; }
		}

		/// <summary>
		/// Gets the table data source for this slicer cache.
		/// </summary>
		public TabularDataNode TabularDataNode { get; }

		private XmlDocument Part { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a new <see cref="ExcelSlicerCache"/> object to represent the slicerCacheN.xml part.
		/// </summary>
		/// <param name="node">The slicerCacheDefinition node to represent.</param>
		/// <param name="namespaceManager">The namespaceManager to use when parsing nodes (This should usually be based on <see cref="ExcelSlicer.SlicerDocumentNamespaceManager"/>).</param>
		/// <param name="slicerCacheUri">The path to this Slicer Cache's part in the package.</param>
		/// <param name="part">The <see cref="XmlDocument"/> based on the <paramref name="slicerCacheUri"/>.</param>
		internal ExcelSlicerCache(XmlNode node, XmlNamespaceManager namespaceManager, Uri slicerCacheUri, XmlDocument part) : base(namespaceManager, node)
		{
			this.SlicerCacheUri = slicerCacheUri;
			this.Part = part;
			foreach (XmlNode pivotTableNode in this.TopNode.SelectNodes("default:pivotTables/default:pivotTable", this.NameSpaceManager))
			{
				myPivotTables.Add(new PivotTableNode(pivotTableNode));
			}
			foreach (XmlNode olapDataNode in this.TopNode.SelectNodes("default:data/default:olap", this.NameSpaceManager))
			{
				myOlapDataSources.Add(new OlapDataNode(olapDataNode, this.NameSpaceManager));
			}

			var tabularDataNode = this.TopNode.SelectSingleNode("default:data/default:tabular", this.NameSpaceManager);
			this.TabularDataNode = new TabularDataNode(tabularDataNode, this.NameSpaceManager);
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Refreshes the slicer cache's values.
		/// </summary>
		internal void Refresh(ExcelPivotCacheDefinition cacheDefinition, List<CacheItem> previouslySelectedItems)
		{
			// If all are selected and a new value is added, it is selected.
			// Otherwise new values are added as deselected.
			bool isFiltered = this.TabularDataNode.Items.Any(i => !i.IsSelected);
			var cacheFieldIndex = cacheDefinition.GetCacheFieldIndex(this.SourceName);
			var cacheField = cacheDefinition.CacheFields[cacheFieldIndex];
			var cacheItems = cacheDefinition.GetCacheItemsForSlicer(cacheField);

			this.TabularDataNode.Items.Clear();

			if (isFiltered)
			{
				for (int i = 0; i < cacheItems.Count; i++)
				{
					var sharedItem = cacheItems[i];
					bool isSelected = previouslySelectedItems.Any(si => si.Value == sharedItem.Value && si.Type == sharedItem.Type);
					this.TabularDataNode.Items.Add(i, isSelected);
				}
			}
			else
				cacheItems.ForEach((c, i) => this.TabularDataNode.Items.Add(i, true));
		}

		/// <summary>
		/// Applies the sort and hide settings to the slice values. This must be called after
		/// pivot tables are refreshed in order to work properly.
		/// </summary>
		/// <param name="cacheDefinition">The backing cache definition.</param>
		internal void ApplySettings(ExcelPivotCacheDefinition cacheDefinition)
		{
			var cacheFieldIndex = cacheDefinition.GetCacheFieldIndex(this.SourceName);
			var cacheField = cacheDefinition.CacheFields[cacheFieldIndex];
			var cacheItems = cacheDefinition.GetCacheItemsForSlicer(cacheField);
			var sortedItems = this.Sort(cacheField, cacheItems);
			sortedItems = this.ApplyNoDataSettings(sortedItems, cacheDefinition, cacheFieldIndex, cacheField, cacheItems);
			this.TabularDataNode.Items.Clear();
			this.TabularDataNode.Items.AddRange(sortedItems);
		}

		/// <summary>
		/// Save this <see cref="ExcelSlicerCache"/> back into the <paramref name="package"/> at.
		/// </summary>
		/// <param name="package">The <see cref="ExcelPackage"/> to save to.</param>
		internal void Save(ExcelPackage package)
		{
			package.SavePart(new Uri("/xl/" + this.SlicerCacheUri, UriKind.Relative), this.Part);
		}
		#endregion

		#region Private Methods
		private List<TabularItemNode> Sort(CacheFieldNode cacheField, SharedItemsCollection cacheItems)
		{
			// Sort the fields according to their types.
			IComparer<string> comparer = null;
			if (this.TabularDataNode.CustomListSort && cacheField.IsDateGrouping)
			{
				if (cacheField.FieldGroup.GroupBy == PivotFieldDateGrouping.Months)
					comparer = new MonthComparer();
				else if (cacheField.FieldGroup.GroupBy == PivotFieldDateGrouping.Days)
					comparer = new DayComparer();
			}
			else
				comparer = new NaturalComparer();

			// Sort the slicer cache items.
			if (this.TabularDataNode.SortOrder == SortOrder.Descending)
				return this.TabularDataNode.Items.OrderByDescending(t => cacheItems[t.AtomIndex].Value, comparer).ToList();
			return this.TabularDataNode.Items.OrderBy(t => cacheItems[t.AtomIndex].Value, comparer).ToList();
		}
		
		private List<TabularItemNode> ApplyNoDataSettings(List<TabularItemNode> sortedItems, ExcelPivotCacheDefinition cacheDefinition, 
			int cacheFieldIndex, CacheFieldNode cacheField, SharedItemsCollection cacheItems)
		{
			var pivotTables = this.GetRelatedPivotTables(cacheDefinition);
			if (this.HideItemsWithNoData)
			{
				foreach (var item in sortedItems)
				{
					// TODO: Task #13685 - Implement hide items with no data settings.
					if (!this.PivotTablesContainItem(item, pivotTables, cacheFieldIndex, cacheItems))
						item.NoData = true;
				}
			}
			else
			{
				if (this.TabularDataNode.CrossFilter == CrossFilter.Both)
				{
					var usedItems = new List<TabularItemNode>();
					var unusedItems = new List<TabularItemNode>();
					foreach (var item in sortedItems)
					{
						bool hasData = this.PivotTablesContainItem(item, pivotTables, cacheFieldIndex, cacheItems);
						if (hasData)
							usedItems.Add(item);
						else
							unusedItems.Add(item);
					}
					sortedItems = usedItems.Concat(unusedItems).ToList();
				}

				if (!this.TabularDataNode.ShowMissing)
				{
					foreach (var item in sortedItems)
					{
						if (cacheItems[item.AtomIndex].Unused)
							item.NoData = true;
					}
				}
			}
			return sortedItems;
		}

		private bool PivotTablesContainItem(TabularItemNode item, List<ExcelPivotTable> pivotTables, 
			int cacheFieldIndex, SharedItemsCollection cacheItems)
		{
			// Cache field items marked as unused are never referenced in a pivot table.
			if (cacheItems[item.AtomIndex].Unused)
				return false;

			foreach (var pivotTable in pivotTables)
			{
				if (pivotTable.ContainsData(cacheFieldIndex, item.AtomIndex))
					return true;
			}
			return false;
		}

		private List<ExcelPivotTable> GetRelatedPivotTables(ExcelPivotCacheDefinition cacheDefinition)
		{
			var relatedPivotTables = cacheDefinition.GetRelatedPivotTables();
			var pivotTableNames = this.PivotTables.Select(p => p.PivotTableName);
			return relatedPivotTables.Where(p => pivotTableNames.Any(n => n.IsEquivalentTo(p.Name))).ToList();
		}
		#endregion
	}
}
