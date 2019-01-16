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
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	#region Enums
	/// <summary>
	/// Defines the axis for a PivotTable.
	/// </summary>
	public enum ePivotFieldAxis
	{
		/// <summary>
		/// None
		/// </summary>
		None = -1,
		/// <summary>
		/// Column axis
		/// </summary>
		Column,
		/// <summary>
		/// Page axis (Include Count Filter) 
		/// </summary>
		Page,
		/// <summary>
		/// Row axis
		/// </summary>
		Row,
		/// <summary>
		/// Values axis
		/// </summary>
		Values
	}

	/// <summary>
	/// Build-in table row functions.
	/// </summary>
	public enum DataFieldFunctions
	{
		Average,
		Count,
		CountNums,
		Max,
		Min,
		Product,
		None,
		StdDev,
		StdDevP,
		Sum,
		Var,
		VarP
	}

	/// <summary>
	/// Defines the data formats for a field in the PivotTable.
	/// </summary>
	public enum eShowDataAs
	{
		/// <summary>
		/// Indicates the field is shown as the "difference from" a value.
		/// </summary>
		Difference,
		/// <summary>
		/// Indicates the field is shown as the "index.
		/// </summary>
		Index,
		/// <summary>
		/// Indicates that the field is shown as its normal datatype.
		/// </summary>
		Normal,
		/// <summary>
		/// /Indicates the field is show as the "percentage of" a value.
		/// </summary>
		Percent,
		/// <summary>
		/// Indicates the field is shown as the "percentage difference from" a value.
		/// </summary>
		PercentDiff,
		/// <summary>
		/// Indicates the field is shown as the percentage of column.
		/// </summary>
		PercentOfCol,
		/// <summary>
		/// Indicates the field is shown as the percentage of row.
		/// </summary>
		PercentOfRow,
		/// <summary>
		/// Indicates the field is shown as percentage of total.
		/// </summary>
		PercentOfTotal,
		/// <summary>
		/// Indicates the field is shown as running total in the table.
		/// </summary>
		RunTotal,
	}

	/// <summary>
	/// Built-in subtotal functions.
	/// </summary>
	[Flags]
	public enum eSubTotalFunctions
	{
		None = 1,
		Count = 2,
		CountA = 4,
		Avg = 8,
		Default = 16,
		Min = 32,
		Max = 64,
		Product = 128,
		StdDev = 256,
		StdDevP = 512,
		Sum = 1024,
		Var = 2048,
		VarP = 4096
	}

	/// <summary>
	/// Data grouping.
	/// </summary>
	[Flags]
	public enum eDateGroupBy
	{
		Years = 1,
		Quarters = 2,
		Months = 4,
		Days = 8,
		Hours = 16,
		Minutes = 32,
		Seconds = 64
	}

	/// <summary>
	/// Sorting
	/// </summary>
	public enum eSortType
	{
		None,
		Ascending,
		Descending
	}
	#endregion

	/// <summary>
	/// A pivotTable pivotField XML node.
	/// </summary>
	public class ExcelPivotTableField : XmlCollectionItemBase
	{
		#region Class Variables
		/// <summary>
		/// The pivot table.
		/// </summary>
		internal ExcelPivotTable myTable;
		/// <summary>
		/// The cache field xml helper instance object.
		/// </summary>
		internal XmlHelperInstance myCacheFieldHelper;
		/// <summary>
		/// The pivot table field groupings object.
		/// </summary>
		internal ExcelPivotTableFieldGroup myGrouping;
		/// <summary>
		/// The pivot table field items.
		/// </summary>
		internal ExcelPivotTableFieldItemCollection myItems;
		/// <summary>
		/// The pivot field sorting references.
		/// </summary>
		internal AutoSortScopeReferencesCollection mySortingReferences;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the index of the field.
		/// </summary>
		public int Index { get; set; }
		
		/// <summary>
		/// Gets or sets the name of the field.
		/// </summary>
		public string Name
		{
			get
			{
				string v = base.GetXmlNodeString("@name");
				if (v == "")
					return myCacheFieldHelper.GetXmlNodeString("@name");
				else
					return v;
			}
			set
			{
				base.SetXmlNodeString("@name", value);
			}
		}

		/// <summary>
		/// Gets or sets compact mode.
		/// </summary>
		public bool Compact
		{
			get
			{
				return base.GetXmlNodeBool("@compact");
			}
			set
			{
				base.SetXmlNodeBool("@compact", value);
			}
		}

		/// <summary>
		/// Gets or sets whether the items in this field should be shown in Outline form.
		/// </summary>
		public bool Outline
		{
			get
			{
				return base.GetXmlNodeBool("@outline");
			}
			set
			{
				base.SetXmlNodeBool("@outline", value);
			}
		}

		/// <summary>
		/// Gets or sets whether the custom text that is displayed for the subtotals label is shown.
		/// </summary>
		public bool SubtotalTop
		{
			get
			{
				return base.GetXmlNodeBool("@subtotalTop", true);
			}
			set
			{
				base.SetXmlNodeBool("@subtotalTop", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to show all items for this field.
		/// </summary>
		public bool ShowAll
		{
			get
			{
				return base.GetXmlNodeBool("@showAll");
			}
			set
			{
				base.SetXmlNodeBool("@showAll", value);
			}
		}

		/// <summary>
		/// Gets whether to show the default subtotal.
		/// </summary>
		/// <remarks>A blank value in XML indicates true. Setting this value to false will remove the subtotal nodes from 
		/// the <see cref="ExcelPivotTableField"/>.</remarks>
		public bool DefaultSubtotal
		{
			get
			{
				return base.GetXmlNodeBool("@defaultSubtotal", true);
			}
			private set
			{
				base.SetXmlNodeBool("@defaultSubtotal", value);
			}
		}

		/// <summary>
		/// Gets or sets the type of sort that is applied to this field.
		/// </summary>
		public eSortType Sort
		{
			get
			{
				string v = base.GetXmlNodeString("@sortType");
				return v == "" ? eSortType.None : (eSortType)Enum.Parse(typeof(eSortType), v, true);
			}
			set
			{
				if (value == eSortType.None)
					base.DeleteNode("@sortType");
				else
					base.SetXmlNodeString("@sortType", value.ToString().ToLower(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets or sets whether manual filter is in inclusive mode.
		/// </summary>
		public bool IncludeNewItemsInFilter
		{
			get
			{
				return base.GetXmlNodeBool("@includeNewItemsInFilter");
			}
			set
			{
				base.SetXmlNodeBool("@includeNewItemsInFilter", value);
			}
		}

		/// <summary>
		/// Gets or sets the enumeration of the different subtotal operations that can be applied to page, row or column fields.
		/// </summary>
		public eSubTotalFunctions SubTotalFunctions
		{
			get
			{
				eSubTotalFunctions ret = 0;
				XmlNodeList nl = base.TopNode.SelectNodes("d:items/d:item/@t", base.NameSpaceManager);
				if (nl.Count == 0) return eSubTotalFunctions.None;
				foreach (XmlAttribute item in nl)
				{
					try
					{
						ret |= (eSubTotalFunctions)Enum.Parse(typeof(eSubTotalFunctions), item.Value, true);
					}
					catch (ArgumentException ex)
					{
						throw new ArgumentException("Unable to parse value of " + item.Value + " to a valid pivot table subtotal function", ex);
					}
				}
				return ret;
			}
			set
			{
				if ((value & eSubTotalFunctions.None) == eSubTotalFunctions.None && (value != eSubTotalFunctions.None))
					throw (new ArgumentException("Value None can not be combined with other values."));
				if ((value & eSubTotalFunctions.Default) == eSubTotalFunctions.Default && (value != eSubTotalFunctions.Default))
					throw (new ArgumentException("Value Default can not be combined with other values."));
				
				// Remove old attribute                 
				XmlNodeList nl = base.TopNode.SelectNodes("d:items/d:item/@t", base.NameSpaceManager);
				if (nl.Count > 0)
				{
					foreach (XmlAttribute item in nl)
					{
						base.DeleteNode("@" + item.Value + "Subtotal");
						item.OwnerElement.ParentNode.RemoveChild(item.OwnerElement);
					}
				}

				if (value == eSubTotalFunctions.None)
				{
					// For no subtotals, set defaultSubtotal to off
					this.DefaultSubtotal = false;
					base.TopNode.InnerXml = "";
				}
				else
				{
					string innerXml = "";
					int count = 0;
					foreach (eSubTotalFunctions e in Enum.GetValues(typeof(eSubTotalFunctions)))
					{
						if ((value & e) == e)
						{
							var newTotalType = e.ToString();
							var totalType = char.ToLower(newTotalType[0], CultureInfo.InvariantCulture) + newTotalType.Substring(1);
							// Add new attribute
							base.SetXmlNodeBool("@" + totalType + "Subtotal", true);
							innerXml += "<item t=\"" + totalType + "\" />";
							count++;
						}
					}
					this.TopNode.InnerXml = string.Format("<items count=\"{0}\">{1}</items>", count, innerXml);
				}
			}
		}
		
		/// <summary>
		/// Gets or sets the type of axis.
		/// </summary>
		public ePivotFieldAxis Axis
		{
			get
			{
				switch (GetXmlNodeString("@axis"))
				{
					case "axisRow":
						return ePivotFieldAxis.Row;
					case "axisCol":
						return ePivotFieldAxis.Column;
					case "axisPage":
						return ePivotFieldAxis.Page;
					case "axisValues":
						return ePivotFieldAxis.Values;
					default:
						return ePivotFieldAxis.None;
				}
			}
			internal set
			{
				switch (value)
				{
					case ePivotFieldAxis.Row:
						base.SetXmlNodeString("@axis", "axisRow");
						break;
					case ePivotFieldAxis.Column:
						base.SetXmlNodeString("@axis", "axisCol");
						break;
					case ePivotFieldAxis.Values:
						base.SetXmlNodeString("@axis", "axisValues");
						break;
					case ePivotFieldAxis.Page:
						base.SetXmlNodeString("@axis", "axisPage");
						break;
					default:
						base.DeleteNode("@axis");
						break;
				}
			}
		}
		
		/// <summary>
		/// Gets or sets whether the field is a row field.
		/// </summary>
		public bool IsRowField
		{
			get
			{
				return (this.TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) != null);
			}
			internal set
			{
				if (value)
				{
					var rowsNode = this.TopNode.SelectSingleNode("../../d:rowFields", this.NameSpaceManager);
					if (rowsNode == null)
						myTable.CreateNode("d:rowFields");
					rowsNode = this.TopNode.SelectSingleNode("../../d:rowFields", this.NameSpaceManager);

					AppendField(rowsNode, this.Index, "field", "x");
					if (this.BaseIndex == this.Index)
						this.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
					else
						this.TopNode.InnerXml = "<items count=\"0\"></items>";
				}
				else
				{
					if (this.TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) is XmlElement node)
						node.ParentNode.RemoveChild(node);
				}
			}
		}
		
		/// <summary>
		/// Gets or sets whether the field is a column field.
		/// </summary>
		public bool IsColumnField
		{
			get
			{
				return (this.TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) != null);
			}
			internal set
			{
				if (value)
				{
					var columnsNode = this.TopNode.SelectSingleNode("../../d:colFields", this.NameSpaceManager);
					if (columnsNode == null)
						myTable.CreateNode("d:colFields");
					columnsNode = this.TopNode.SelectSingleNode("../../d:colFields", this.NameSpaceManager);

					AppendField(columnsNode, this.Index, "field", "x");
					if (this.BaseIndex == this.Index)
						this.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
					else
						this.TopNode.InnerXml = "<items count=\"0\"></items>";
				}
				else
				{
					if (this.TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) is XmlElement node)
						node.ParentNode.RemoveChild(node);
				}
			}
		}
		
		/// <summary>
		/// Gets or sets whether the field is a data field.
		/// </summary>
		public bool IsDataField
		{
			get
			{
				return base.GetXmlNodeBool("@dataField", false);
			}
		}
		
		/// <summary>
		/// <summary>
		/// Gets the grouping settings. 
		/// Null if the field has no grouping otherwise ExcelPivotTableFieldNumericGroup or ExcelPivotTableFieldNumericGroup.
		/// </summary>        
		public ExcelPivotTableFieldGroup Grouping
		{
			get
			{
				return myGrouping;
			}
		}

		/// <summary>
		/// Gets the pivot table field items that is used for grouping.
		/// </summary>
		public ExcelPivotTableFieldItemCollection Items
		{
			get
			{
				if (myItems == null)
					myItems = new ExcelPivotTableFieldItemCollection(this.NameSpaceManager, this.TopNode.FirstChild, myTable, this);
				return myItems;
			}
		}

		/// <summary>
		/// Gets the pivot field references that is used for custom sorting.
		/// </summary>
		public AutoSortScopeReferencesCollection AutoSortScopeReferences
		{
			get
			{
				if (mySortingReferences == null)
					// Select all 'references' nodes that are descendants of the top node.
					mySortingReferences = new AutoSortScopeReferencesCollection(this.NameSpaceManager, this.TopNode.SelectSingleNode(".//d:references", this.NameSpaceManager), myTable);
				return mySortingReferences;
			}
		}

		/// <summary>
		/// Gets or sets the base index.
		/// </summary>
		internal int BaseIndex { get; set; }

		/// <summary>
		/// Gets or sets the date groupings.
		/// </summary>
		internal eDateGroupBy DateGrouping { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableField"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace of the worksheet.</param>
		/// <param name="topNode">The xml element.</param>
		/// <param name="table">The pivot table.</param>
		/// <param name="index">The index of the field.</param>
		/// <param name="baseIndex">The base index of the field.</param>
		internal ExcelPivotTableField(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelPivotTable table, int index, int baseIndex) 
			: base(namespaceManager, topNode)
		{
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			if (topNode == null)
				throw new ArgumentNullException(nameof(topNode));
			if (table == null)
				throw new ArgumentNullException(nameof(table));
			if (baseIndex < 0)
				throw new ArgumentOutOfRangeException(nameof(baseIndex));
			this.Index = index;
			this.BaseIndex = baseIndex;
			myTable = table;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add numberic grouping to the field.
		/// </summary>
		/// <param name="start">Start value</param>
		/// <param name="end">End value</param>
		/// <param name="interval">Interval</param>
		public void AddNumericGrouping(double start, double end, double interval)
		{
			this.ValidateGrouping();
			this.SetNumericGroup(start, end, interval);
		}

		/// <summary>
		/// Add a date grouping on this field.
		/// </summary>
		/// <param name="groupBy">Group by</param>
		public void AddDateGrouping(eDateGroupBy groupBy)
		{
			this.AddDateGrouping(groupBy, DateTime.MinValue, DateTime.MaxValue, 1);
		}

		/// <summary>
		/// Add a date grouping on this field.
		/// </summary>
		/// <param name="groupBy">Group by</param>
		/// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
		/// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
		public void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate)
		{
			this.AddDateGrouping(groupBy, startDate, endDate, 1);
		}

		/// <summary>
		/// Add a date grouping on this field.
		/// </summary>
		/// <param name="days">Number of days when grouping on days</param>
		/// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
		/// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
		public void AddDateGrouping(int days, DateTime startDate, DateTime endDate)
		{
			this.AddDateGrouping(eDateGroupBy.Days, startDate, endDate, days);
		}

		/// <summary>
		/// Add a field to the pivot table.
		/// </summary>
		/// <param name="rowsNode">The row node in the xml.</param>
		/// <param name="index">The index of the field.</param>
		/// <param name="fieldNodeText">The text of the new node.</param>
		/// <param name="indexAttrText">The text of the index attribute.</param>
		/// <returns>The added xml element.</returns>
		internal XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
		{
			XmlElement prevField = null, newElement;
			foreach (XmlElement field in rowsNode.ChildNodes)
			{
				string x = field.GetAttribute(indexAttrText);
				if (int.TryParse(x, out var fieldIndex))
				{
					if (fieldIndex == index)    //Row already exists
						return field;
				}
				prevField = field;
			}
			newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
			newElement.SetAttribute(indexAttrText, index.ToString());
			rowsNode.InsertAfter(newElement, prevField);
			return newElement;
		}

		/// <summary>
		/// Set the cache field node.
		/// </summary>
		/// <param name="cacheField">The cache field node being set.</param>
		internal void SetCacheFieldNode(XmlNode cacheField)
		{
			myCacheFieldHelper = new XmlHelperInstance(this.NameSpaceManager, cacheField);
			var groupNode = cacheField.SelectSingleNode("d:fieldGroup", this.NameSpaceManager);
			if (groupNode != null)
			{
				var groupBy = groupNode.SelectSingleNode("d:rangePr/@groupBy", this.NameSpaceManager);
				if (groupBy == null)
					myGrouping = new ExcelPivotTableFieldNumericGroup(this.NameSpaceManager, cacheField);
				else
				{
					this.DateGrouping = (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), groupBy.Value, true);
					myGrouping = new ExcelPivotTableFieldDateGroup(this.NameSpaceManager, groupNode);
				}
			}
		}

		/// <summary>
		/// Set the date group.
		/// </summary>
		/// <param name="groupBy">How to group the fields.</param>
		/// <param name="startDate">The start date.</param>
		/// <param name="endDate">The end date.</param>
		/// <param name="interval">The interval of the grouping.</param>
		/// <returns>The new field data group.</returns>
		internal ExcelPivotTableFieldDateGroup SetDateGroup(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int interval)
		{
			ExcelPivotTableFieldDateGroup group;
			group = new ExcelPivotTableFieldDateGroup(this.NameSpaceManager, myCacheFieldHelper.TopNode);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsDate", true);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

			group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", this.BaseIndex, groupBy.ToString().ToLower(CultureInfo.InvariantCulture));

			if (startDate.Year < 1900)
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", "1900-01-01T00:00:00");
			else
			{
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", startDate.ToString("s", CultureInfo.InvariantCulture));
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoStart", "0");
			}

			if (endDate == DateTime.MaxValue)
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", "9999-12-31T00:00:00");
			else
			{
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", endDate.ToString("s", CultureInfo.InvariantCulture));
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
			}

			int items = AddDateGroupItems(group, groupBy, startDate, endDate, interval);
			this.AddFieldItems(items);

			myGrouping = group;
			return group;
		}

		/// <summary>
		/// Set the numeric group.
		/// </summary>
		/// <param name="start">The start value.</param>
		/// <param name="end">The end value.</param>
		/// <param name="interval">The interval value.</param>
		/// <returns>The new field numeric group.</returns>
		internal ExcelPivotTableFieldNumericGroup SetNumericGroup(double start, double end, double interval)
		{
			ExcelPivotTableFieldNumericGroup group;
			group = new ExcelPivotTableFieldNumericGroup(this.NameSpaceManager, myCacheFieldHelper.TopNode);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNumber", true);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsInteger", true);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
			myCacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsString", false);

			group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{1}\" endNum=\"{2}\" groupInterval=\"{3}\"/><groupItems /></fieldGroup>", this.BaseIndex, start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture), interval.ToString(CultureInfo.InvariantCulture));
			int items = AddNumericGroupItems(group, start, end, interval);
			this.AddFieldItems(items);

			myGrouping = group;
			return group;
		}

		/// <summary>
		/// Sets the <see cref="DefaultSubtotal"/> property to false and removes 
		/// the last "default" subtotal item from the <see cref="ExcelPivotTableField"/>.
		/// </summary>
		internal void DisableDefaultSubtotal()
		{
			this.DefaultSubtotal = false;
			this.Items.RemoveLastSubtotalItem();
		}
		#endregion

		#region Private Methods
		private void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int groupInterval)
		{
			if (groupInterval < 1 || groupInterval >= Int16.MaxValue)
				throw (new ArgumentOutOfRangeException("Group interval is out of range"));
			if (groupInterval > 1 && groupBy != eDateGroupBy.Days)
				throw (new ArgumentException("Group interval is can only be used when groupBy is Days"));

			this.ValidateGrouping();

			bool firstField = true;
			List<ExcelPivotTableField> fields = new List<ExcelPivotTableField>();
			//Seconds
			if ((groupBy & eDateGroupBy.Seconds) == eDateGroupBy.Seconds)
				fields.Add(this.AddField(eDateGroupBy.Seconds, startDate, endDate, ref firstField));
			//Minutes
			if ((groupBy & eDateGroupBy.Minutes) == eDateGroupBy.Minutes)
				fields.Add(this.AddField(eDateGroupBy.Minutes, startDate, endDate, ref firstField));
			//Hours
			if ((groupBy & eDateGroupBy.Hours) == eDateGroupBy.Hours)
				fields.Add(this.AddField(eDateGroupBy.Hours, startDate, endDate, ref firstField));
			//Days
			if ((groupBy & eDateGroupBy.Days) == eDateGroupBy.Days)
				fields.Add(this.AddField(eDateGroupBy.Days, startDate, endDate, ref firstField, groupInterval));
			//Month
			if ((groupBy & eDateGroupBy.Months) == eDateGroupBy.Months)
				fields.Add(this.AddField(eDateGroupBy.Months, startDate, endDate, ref firstField));
			//Quarters
			if ((groupBy & eDateGroupBy.Quarters) == eDateGroupBy.Quarters)
				fields.Add(this.AddField(eDateGroupBy.Quarters, startDate, endDate, ref firstField));
			//Years
			if ((groupBy & eDateGroupBy.Years) == eDateGroupBy.Years)
				fields.Add(this.AddField(eDateGroupBy.Years, startDate, endDate, ref firstField));

			if (fields.Count > 1)
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/@par", (myTable.Fields.Count - 1).ToString());
			if (groupInterval != 1)
				myCacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", groupInterval.ToString());
			else
				myCacheFieldHelper.DeleteNode("d:fieldGroup/d:rangePr/@groupInterval");
			myItems = null;
		}

		private void ValidateGrouping()
		{
			if (!(this.IsColumnField || this.IsRowField))
				throw (new Exception("Field must be a row or column field"));
			foreach (var field in myTable.Fields)
			{
				if (field.Grouping != null)
					throw (new Exception("Grouping already exists"));
			}
		}

		private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField)
		{
			return AddField(groupBy, startDate, endDate, ref firstField, 1);
		}

		private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField, int interval)
		{
			if (firstField == false)
			{
				//Pivot field
				var topNode = myTable.PivotTableXml.SelectSingleNode("//d:pivotFields", myTable.NameSpaceManager);
				var fieldNode = myTable.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
				fieldNode.SetAttribute("compact", "0");
				fieldNode.SetAttribute("outline", "0");
				fieldNode.SetAttribute("showAll", "0");
				fieldNode.SetAttribute("defaultSubtotal", "0");
				topNode.AppendChild(fieldNode);

				var field = new ExcelPivotTableField(myTable.NameSpaceManager, fieldNode, myTable, myTable.Fields.Count, this.Index);
				field.DateGrouping = groupBy;

				XmlNode rowColFields;
				if (this.IsRowField)
					rowColFields = this.TopNode.SelectSingleNode("../../d:rowFields", this.NameSpaceManager);
				else
					rowColFields = this.TopNode.SelectSingleNode("../../d:colFields", this.NameSpaceManager);

				int index = 0;
				foreach (XmlElement rowfield in rowColFields.ChildNodes)
				{
					if (int.TryParse(rowfield.GetAttribute("x"), out var fieldIndex))
					{
						if (myTable.Fields[fieldIndex].BaseIndex == this.BaseIndex)
						{
							var newElement = rowColFields.OwnerDocument.CreateElement("field", ExcelPackage.schemaMain);
							newElement.SetAttribute("x", field.Index.ToString());
							rowColFields.InsertBefore(newElement, rowfield);
							break;
						}
					}
					index++;
				}

				if (this.IsRowField)
					myTable.RowFields.Insert(field, index);
				else
					myTable.ColumnFields.Insert(field, index);

				myTable.Fields.Add(field);

				this.AddCacheField(field, startDate, endDate, interval);
				return field;
			}
			else
			{
				firstField = false;
				this.DateGrouping = groupBy;
				this.Compact = false;
				SetDateGroup(groupBy, startDate, endDate, interval);
				return this;
			}
		}

		private void AddCacheField(ExcelPivotTableField field, DateTime startDate, DateTime endDate, int interval)
		{
			//Add Cache definition field.
			var cacheTopNode = myTable.CacheDefinition.CacheDefinitionXml.SelectSingleNode("//d:cacheFields", myTable.NameSpaceManager);
			var cacheFieldNode = myTable.CacheDefinition.CacheDefinitionXml.CreateElement("cacheField", ExcelPackage.schemaMain);

			cacheFieldNode.SetAttribute("name", field.DateGrouping.ToString());
			cacheFieldNode.SetAttribute("databaseField", "0");
			cacheTopNode.AppendChild(cacheFieldNode);
			field.SetCacheFieldNode(cacheFieldNode);

			field.SetDateGroup(field.DateGrouping, startDate, endDate, interval);
		}

		private int AddNumericGroupItems(ExcelPivotTableFieldNumericGroup group, double start, double end, double interval)
		{
			if (interval < 0)
				throw (new Exception("The interval must be a positiv"));
			if (start > end)
				throw (new Exception("Then End number must be larger than the Start number"));

			XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
			int items = 2;
			//First date
			double index = start;
			double nextIndex = start + interval;
			this.AddGroupItem(groupItems, "<" + start.ToString(CultureInfo.InvariantCulture));

			while (index < end)
			{
				this.AddGroupItem(groupItems, string.Format("{0}-{1}", index.ToString(CultureInfo.InvariantCulture), nextIndex.ToString(CultureInfo.InvariantCulture)));
				index = nextIndex;
				nextIndex += interval;
				items++;
			}
			this.AddGroupItem(groupItems, ">" + nextIndex.ToString(CultureInfo.InvariantCulture));
			return items;
		}

		private void AddFieldItems(int items)
		{
			XmlElement prevNode = null;
			XmlElement itemsNode = this.TopNode.SelectSingleNode("d:items", this.NameSpaceManager) as XmlElement;
			for (int x = 0; x < items; x++)
			{
				var itemNode = itemsNode.OwnerDocument.CreateElement("item", ExcelPackage.schemaMain);
				itemNode.SetAttribute("x", x.ToString());
				if (prevNode == null)
					itemsNode.PrependChild(itemNode);
				else
					itemsNode.InsertAfter(itemNode, prevNode);
				prevNode = itemNode;
			}
			itemsNode.SetAttribute("count", (items + 1).ToString());
		}

		private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int interval)
		{
			XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
			int items = 2;
			//First date
			this.AddGroupItem(groupItems, "<" + startDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));

			switch (groupBy)
			{
				case eDateGroupBy.Seconds:
				case eDateGroupBy.Minutes:
					this.AddTimeSerie(60, groupItems);
					items += 60;
					break;
				case eDateGroupBy.Hours:
					this.AddTimeSerie(24, groupItems);
					items += 24;
					break;
				case eDateGroupBy.Days:
					if (interval == 1)
					{
						DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days
						while (dt.Year == 2008)
						{
							this.AddGroupItem(groupItems, dt.ToString("dd-MMM"));
							dt = dt.AddDays(1);
						}
						items += 366;
					}
					else
					{
						DateTime dt = startDate;
						items = 0;
						while (dt < endDate)
						{
							this.AddGroupItem(groupItems, dt.ToString("dd-MMM"));
							dt = dt.AddDays(interval);
							items++;
						}
					}
					break;
				case eDateGroupBy.Months:
					this.AddGroupItem(groupItems, "jan");
					this.AddGroupItem(groupItems, "feb");
					this.AddGroupItem(groupItems, "mar");
					this.AddGroupItem(groupItems, "apr");
					this.AddGroupItem(groupItems, "may");
					this.AddGroupItem(groupItems, "jun");
					this.AddGroupItem(groupItems, "jul");
					this.AddGroupItem(groupItems, "aug");
					this.AddGroupItem(groupItems, "sep");
					this.AddGroupItem(groupItems, "oct");
					this.AddGroupItem(groupItems, "nov");
					this.AddGroupItem(groupItems, "dec");
					items += 12;
					break;
				case eDateGroupBy.Quarters:
					this.AddGroupItem(groupItems, "Qtr1");
					this.AddGroupItem(groupItems, "Qtr2");
					this.AddGroupItem(groupItems, "Qtr3");
					this.AddGroupItem(groupItems, "Qtr4");
					items += 4;
					break;
				case eDateGroupBy.Years:
					if (startDate.Year >= 1900 && endDate != DateTime.MaxValue)
					{
						for (int year = startDate.Year; year <= endDate.Year; year++)
						{
							this.AddGroupItem(groupItems, year.ToString());
						}
						items += endDate.Year - startDate.Year + 1;
					}
					break;
				default:
					throw (new Exception("unsupported grouping"));
			}

			//Lastdate
			this.AddGroupItem(groupItems, ">" + endDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));
			return items;
		}

		private void AddTimeSerie(int count, XmlElement groupItems)
		{
			for (int i = 0; i < count; i++)
			{
				this.AddGroupItem(groupItems, string.Format("{0:00}", i));
			}
		}

		private void AddGroupItem(XmlElement groupItems, string value)
		{
			var s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
			s.SetAttribute("v", value);
			groupItems.AppendChild(s);
		}
		#endregion
	}
}