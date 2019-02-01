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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Extensions;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Maintains state for a tree data structure that represents a pivot table.
	/// </summary>
	public class PivotItemTreeNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the cache record item node "v" index.
		/// </summary>
		public int Value { get; set; }
		
		/// <summary>
		/// Gets or sets the list of cache record indices.
		/// </summary>
		public List<int> CacheRecordIndices { get; set; } = new List<int>();

		/// <summary>
		/// Gets or sets the index of the datafield referenced.
		/// </summary>
		public int DataFieldIndex { get; set; }

		/// <summary>
		/// Gets or sets the index of the pivot field referenced.
		/// </summary>
		public int PivotFieldIndex { get; set; } = -2;

		/// <summary>
		/// Gets or sets the index of the referenced pivot field item.
		/// </summary>
		public int PivotFieldItemIndex { get; set; } = -2;

		/// <summary>
		/// Gets or sets whether or not subtotal top is enabled.
		/// </summary>
		public bool SubtotalTop { get; set; } = true; // Excel defaults this to true, so we will too.

		/// <summary>
		/// Gets whether or not this node represents a datafield.
		/// </summary>
		public bool IsDataField => this.Value == -2;

		/// <summary>
		/// Gets a value indicating whether or not this node has any children.
		/// </summary>
		public bool HasChildren => this.Children.Any();

		/// <summary>
		/// Gets or sets the shared item value.
		/// </summary>
		public string SharedItemValue { get; set; }

		/// <summary>
		/// Gets or sets the list of children that this node parents.
		/// </summary>
		public List<PivotItemTreeNode> Children { get; set; } = new List<PivotItemTreeNode>();
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="value">The cache record item node "v" index.</param>
		public PivotItemTreeNode(int value)
		{
			this.Value = value;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new child with the specified values to this node.
		/// </summary>
		/// <param name="value">The cache record item node "v" index of the new child.</param>
		/// <param name="pivotFieldIndex">The index of the pivot field referenced by the new child.</param>
		/// <param name="pivotFieldItemIndex">The index of the pivot field item referenced by the new child.</param>
		/// <param name="sharedItemValue">The shared item "v" value. (Only used for date groupings).</param>
		public PivotItemTreeNode AddChild(int value, int pivotFieldIndex = -2, int pivotFieldItemIndex = -2, string sharedItemValue = null)
		{
			var child = new PivotItemTreeNode(value)
			{
				PivotFieldIndex = pivotFieldIndex,
				PivotFieldItemIndex = pivotFieldItemIndex,
				SharedItemValue = sharedItemValue
			};
			this.Children.Add(child);
			return child;
		}

		/// <summary>
		/// Checks whether or not a child already exists with the specified value.
		/// </summary>
		/// <param name="value">The value to look for in the children list.</param>
		/// <returns>True if the child exists, otherwise false.</returns>
		public bool HasChild(int value)
		{
			return this.Children?.Any(i => i.Value == value) ?? false;
		}

		/// <summary>
		/// Checks whether or not a child already exists with the specified string value.
		/// </summary>
		/// <param name="value">The value to look for in the children list.</param>
		/// <returns>True if the child exists, otherwise false.</returns>
		public bool HasChild(string value)
		{
			return this.Children?.Any(i => i.SharedItemValue.IsEquivalentTo(value)) ?? false;
		}

		/// <summary>
		/// Gets the child node that has the specified value.
		/// </summary>
		/// <param name="value">The value to look for in the children list.</param>
		/// <returns>The child node if it exists.</returns>
		public PivotItemTreeNode GetChildNode(int value)
		{
			return this.Children.Find(i => i.Value == value);
		}

		/// <summary>
		/// Gets the child node that has the specified string value.
		/// Note: Used for date groupings only.
		/// </summary>
		/// <param name="value">The value to look for in the children list.</param>
		/// <returns>The child node if it exists.</returns>
		public PivotItemTreeNode GetChildNode(string value)
		{
			return this.Children.Find(i => i.SharedItemValue.IsEquivalentTo(value));
		}

		/// <summary>
		/// Creates a deep copy of this node.
		/// </summary>
		/// <returns>The newly created node.</returns>
		public PivotItemTreeNode Clone()
		{
			var clone = new PivotItemTreeNode(this.Value);
			clone.DataFieldIndex = this.DataFieldIndex;
			clone.PivotFieldIndex = this.PivotFieldIndex;
			clone.PivotFieldItemIndex = this.PivotFieldItemIndex;
			clone.SubtotalTop = this.SubtotalTop;

			foreach (var child in this.Children)
			{
				clone.Children.Add(child.Clone());
			}
			return clone;
		}

		/// <summary>
		/// Sets the data field index for this node and all it's children.
		/// </summary>
		/// <param name="index">The specified data field index.</param>
		public void RecursivelySetDataFieldIndex(int index)
		{
			this.DataFieldIndex = index;
			foreach (var child in this.Children)
			{
				child.RecursivelySetDataFieldIndex(index);
			}
		}

		/// <summary>
		/// If the only child of this node is a datafield node, expand it by duplicating 
		/// each child for each data field and setting the respective data field indices 
		/// on the duplicated node paths.
		/// </summary>
		/// <param name="dataFieldCount">The number of data fields on the pivot table.</param>
		public void ExpandIfDataFieldNode(int dataFieldCount)
		{
			if (this.Children.Count == 1)
			{
				var onlyChild = this.Children.First();
				if (onlyChild.IsDataField)
				{
					onlyChild.RecursivelySetDataFieldIndex(0);
					// child is a datafield node, create a node for each datafield and update with the index into the datafield collection.
					for (int i = 0; i < dataFieldCount - 1; i++)
					{
						var newChild = onlyChild.Clone();
						newChild.RecursivelySetDataFieldIndex(i + 1);
						this.Children.Add(newChild);
					}
				}
			}
		}

		/// <summary>
		/// Sorts this node's children.
		/// </summary>
		/// <param name="pivotTable">The pivot table that defines sorting.</param>
		public void SortChildren(ExcelPivotTable pivotTable)
		{
			foreach (var child in this.Children.ToList())
			{
				// Set the pivot table field.
				ExcelPivotTableField pivotField = null;
				if (child.PivotFieldIndex == -2 && !child.HasChildren)
					continue;
				else if (child.PivotFieldIndex == -2)
					pivotField = pivotTable.Fields[child.Children[0].PivotFieldIndex];
				else
					pivotField = pivotTable.Fields[child.PivotFieldIndex];

				// Sort the children, so that they are in alphabetical/chronological order.
				var orderNodes = this.Children.OrderBy(x => x.PivotFieldItemIndex).ToList();
				if (!this.Children.SequenceEqual(orderNodes))
				{
					var newChildList = this.Children.ToList();
					this.Children.Clear();
					for (int i = 0; i < orderNodes.Count(); i++)
					{
						int orderListIndex = newChildList.FindIndex(x => x.PivotFieldItemIndex == orderNodes.ElementAt(i).PivotFieldItemIndex);
						this.Children.Add(newChildList[orderListIndex]);
					}
				}

				// Sort items with references to datafields.
				if (pivotField.AutoSortScopeReferences.Count != 0)
					this.SortWithDataFields(pivotTable, pivotField, this);

				// Recursively sort the children of the current node.
				child.SortChildren(pivotTable);
			}
		}
		#endregion

		#region Private Methods
		private void SortWithDataFields(ExcelPivotTable pivotTable, ExcelPivotTableField pivotField, PivotItemTreeNode root)
		{
			int autoScopeIndex = int.Parse(pivotField.AutoSortScopeReferences[0].Value);
			var referenceDataFieldIndex = pivotTable.DataFields[autoScopeIndex].Index;

			// TODO: Implement sorting for calculated fields. Logged in VSTS bug #10277.
			// Skip sorting if this is a calculated field.
			if (!string.IsNullOrEmpty(pivotTable.CacheDefinition.CacheFields[referenceDataFieldIndex].Formula))
				return;

			var orderedList = new List<Tuple<int, double>>();
			// Get the total value for every child at the given data field index.
			foreach (var c in root.Children.ToList())
			{
				double sortingTotal = pivotTable.CacheDefinition.CacheRecords.CalculateSortingValues(c, referenceDataFieldIndex);
				orderedList.Add(new Tuple<int, double>(c.Value, sortingTotal));
			}

			// Sort the list of total values accordingly.
			if (pivotField.Sort == eSortType.Ascending)
				orderedList = orderedList.OrderBy(i => i.Item2).ToList();
			else if (pivotField.Sort == eSortType.Descending)
				orderedList = orderedList.OrderByDescending(i => i.Item2).ToList();

			// If there are duplicated sortingTotal values, sort it based on the value of the first tuple.
			var duplicates = orderedList.GroupBy(x => x.Item2).Where(g => g.Count() > 1).Select(k => k.Key);
			if (duplicates.Count() > 0)
			{
				for (int i = 0; i < duplicates.ToList().Count(); i++)
				{
					var duplicatedList = orderedList.FindAll(j => j.Item2 == duplicates.ElementAt(i));
					int startingIndex = orderedList.FindIndex(y => y == duplicatedList.ElementAt(0));
					if (pivotField.Sort == eSortType.Ascending)
						duplicatedList = duplicatedList.OrderBy(w => w.Item1).ToList();
					else
						duplicatedList = duplicatedList.OrderByDescending(w => w.Item1).ToList();
					orderedList.RemoveRange(startingIndex, duplicatedList.Count());
					orderedList.InsertRange(startingIndex, duplicatedList);
				}
			}

			// Add children back to the root node in sorted order.
			var newChildList = root.Children.ToList();
			root.Children.Clear();
			for (int i = 0; i < orderedList.Count(); i++)
			{
				int index = newChildList.FindIndex(x => x.Value == orderedList.ElementAt(i).Item1);
				root.Children.Add(newChildList[index]);
			}
		}
		#endregion
	}
}
