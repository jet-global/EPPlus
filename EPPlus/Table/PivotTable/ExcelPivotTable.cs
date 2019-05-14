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
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Internationalization;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table.PivotTable.DataCalculation;
using OfficeOpenXml.Table.PivotTable.Filters;
using OfficeOpenXml.Table.PivotTable.Formats;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// An Excel Pivottable
	/// </summary>
	public class ExcelPivotTable : XmlHelper
	{
		#region Constants
		private const string NamePath = "@name";
		private const string DisplayNamePath = "@displayName";
		private const string FirstHeaderRowPath = "d:location/@firstHeaderRow";
		private const string FirstDataRowPath = "d:location/@firstDataRow";
		private const string FirstDataColumnPath = "d:location/@firstDataCol";
		private const string StyleNamePath = "d:pivotTableStyleInfo/@name";
		#endregion

		#region Class Variables
		private ExcelPivotCacheDefinition myCacheDefinition;
		private ExcelPivotTableFieldCollection myFields;
		private ExcelPivotTableRowColumnFieldCollection myRowFields;
		private ExcelPivotTableRowColumnFieldCollection myColumnFields;
		private ExcelPivotTableDataFieldCollection myDataFields;
		private ExcelPageFieldCollection myPageFields;
		private ExcelPivotFieldFiltersCollection myFilters;
		private ExcelFormatsCollection myFormats;
		private ItemsCollection myRowItems;
		private ItemsCollection myColumnItems;
		private TableStyles myTableStyle = Table.TableStyles.Medium6;
		private ExcelAddress myAddress;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the xml data representing the pivot table in the package.
		/// </summary>
		public XmlDocument PivotTableXml { get; private set; }

		/// <summary>
		/// Gets or sets the package internal URI to the pivot table xml Document.
		/// </summary>
		public Uri PivotTableUri { get; internal set; }

		/// <summary>
		/// Gets or sets the name of the pivot table object in Excel.
		/// </summary>
		public string Name
		{
			get { return base.GetXmlNodeString(NamePath); }
			set
			{
				if (this.Worksheet.Workbook.ExistsTableName(value))
					throw (new ArgumentException("PivotTable name is not unique"));
				string prevName = this.Name;
				if (this.Worksheet.Tables.TableNames.ContainsKey(prevName))
				{
					int ix = this.Worksheet.Tables.TableNames[prevName];
					this.Worksheet.Tables.TableNames.Remove(prevName);
					this.Worksheet.Tables.TableNames.Add(value, ix);
				}
				base.SetXmlNodeString(NamePath, value);
				base.SetXmlNodeString(DisplayNamePath, this.CleanDisplayName(value));
			}
		}

		/// <summary>
		/// Gets the reference to the pivot table cache definition object.
		/// </summary>
		public ExcelPivotCacheDefinition CacheDefinition
		{
			get
			{
				if (myCacheDefinition == null)
				{
					if (this.CacheDefinitionRelationship == null)
						throw new InvalidOperationException($"{nameof(this.CacheDefinitionRelationship)} is null.");

					var pivotTableCacheDefinitionPartName = UriHelper.GetUriEndTargetName(this.CacheDefinitionRelationship.TargetUri);
					foreach (var cacheDefinition in this.Worksheet.Workbook.PivotCacheDefinitions)
					{
						var cacheDefinitionPartName = UriHelper.GetUriEndTargetName(cacheDefinition.CacheDefinitionUri);
						if (pivotTableCacheDefinitionPartName.IsEquivalentTo(cacheDefinitionPartName))
						{
							myCacheDefinition = cacheDefinition;
							break;
						}
					}
				}
				return myCacheDefinition;
			}
			private set { myCacheDefinition = value; }
		}

		/// <summary>
		/// Gets the worksheet where the pivot table is located.
		/// </summary>
		public ExcelWorksheet Worksheet
		{
			get { return this.Workbook.Worksheets[this.Address.WorkSheet]; }
		}

		/// <summary>
		/// Gets or sets the location of the pivot table.
		/// </summary>
		public ExcelAddress Address
		{
			get { return myAddress; }
			internal set
			{
				if (string.IsNullOrEmpty(value.WorkSheet))
					throw new InvalidOperationException("PivotTable address must specify a worsheet.");
				myAddress = value;
			}
		}

		/// <summary>
		/// Gets or sets whether multiple datafields are displayed in the row area or the column area.
		/// </summary>
		public bool DataOnRows
		{
			get { return base.GetXmlNodeBool("@dataOnRows"); }
			set { base.SetXmlNodeBool("@dataOnRows", value); }
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat number format properties.
		/// </summary>
		public bool ApplyNumberFormats
		{
			get { return base.GetXmlNodeBool("@applyNumberFormats"); }
			set { base.SetXmlNodeBool("@applyNumberFormats", value); }
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat border properties.
		/// </summary>
		public bool ApplyBorderFormats
		{
			get { return base.GetXmlNodeBool("@applyBorderFormats"); }
			set { base.SetXmlNodeBool("@applyBorderFormats", value); }
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat font properties.
		/// </summary>
		public bool ApplyFontFormats
		{
			get { return base.GetXmlNodeBool("@applyFontFormats"); }
			set { base.SetXmlNodeBool("@applyFontFormats", value); }
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat pattern properties.
		/// </summary>
		public bool ApplyPatternFormats
		{
			get { return base.GetXmlNodeBool("@applyPatternFormats"); }
			set { base.SetXmlNodeBool("@applyPatternFormats", value); }
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat width/height properties.
		/// </summary>
		public bool ApplyWidthHeightFormats
		{
			get { return base.GetXmlNodeBool("@applyWidthHeightFormats"); }
			set { base.SetXmlNodeBool("@applyWidthHeightFormats", value); }
		}

		/// <summary>
		/// Gets or sets whether to show member property information.
		/// </summary>
		/// <remarks>This only applies to pivot tables with an OLAP data source.</remarks>
		public bool ShowMemberPropertyTips
		{
			get { return base.GetXmlNodeBool("@showMemberPropertyTips"); }
			set { base.SetXmlNodeBool("@showMemberPropertyTips", value); }
		}

		/// <summary>
		/// Gets or sets whether to show the drill indicators.
		/// </summary>
		public bool ShowCalcMember
		{
			get { return base.GetXmlNodeBool("@showCalcMbrs"); }
			set { base.SetXmlNodeBool("@showCalcMbrs", value); }
		}

		/// <summary>
		/// Gets or sets if the user can enable drill down on a PivotItem or aggregate value.
		/// </summary>
		public bool EnableDrill
		{
			get { return base.GetXmlNodeBool("@enableDrill", true); }
			set { base.SetXmlNodeBool("@enableDrill", value); }
		}

		/// <summary>
		/// Gets or sets whether to show the drill down buttons (expand/collapse buttons).
		/// </summary>
		public bool ShowDrill
		{
			get { return base.GetXmlNodeBool("@showDrill", true); }
			set { base.SetXmlNodeBool("@showDrill", value); }
		}

		/// <summary>
		/// Gets or sets whether the tooltips should be displayed for PivotTable data cells.
		/// </summary>
		public bool ShowDataTips
		{
			get { return base.GetXmlNodeBool("@showDataTips", true); }
			set { base.SetXmlNodeBool("@showDataTips", value, true); }
		}

		/// <summary>
		/// Gets or sets whether the row and column titles from the PivotTable should be printed.
		/// </summary>
		public bool FieldPrintTitles
		{
			get { return base.GetXmlNodeBool("@fieldPrintTitles"); }
			set { base.SetXmlNodeBool("@fieldPrintTitles", value); }
		}

		/// <summary>
		/// Gets or sets whether the row and column titles from the PivotTable should be printed.
		/// </summary>
		public bool ItemPrintTitles
		{
			get { return base.GetXmlNodeBool("@itemPrintTitles"); }
			set { base.SetXmlNodeBool("@itemPrintTitles", value); }
		}

		/// <summary>
		/// Gets or sets whether the grand totals should be displayed for the PivotTable columns.
		/// </summary>
		/// <remarks>A blank value in XML indicates true.</remarks>
		public bool ColumnGrandTotals
		{
			get { return base.GetXmlNodeBool("@colGrandTotals", true); }
			set { base.SetXmlNodeBool("@colGrandTotals", value); }
		}

		/// <summary>
		///Gets or sets whether the grand totals should be displayed for the PivotTable rows.
		/// </summary>
		/// <remarks>A blank value in XML indicates true.</remarks>
		public bool RowGrandTotals
		{
			get { return base.GetXmlNodeBool("@rowGrandTotals", true); }
			set { base.SetXmlNodeBool("@rowGrandTotals", value); }
		}

		/// <summary>
		/// Gets or sets whether the drill indicators expand collapse buttons should be printed.
		/// </summary>
		public bool PrintDrill
		{
			get { return base.GetXmlNodeBool("@printDrill"); }
			set { base.SetXmlNodeBool("@printDrill", value); }
		}

		/// <summary>
		/// Gets or sets whether to show error messages in cells.
		/// </summary>
		public bool ShowError
		{
			get { return base.GetXmlNodeBool("@showError"); }
			set { base.SetXmlNodeBool("@showError", value); }
		}

		/// <summary>
		/// Gets or sets the string to be displayed in cells that contain errors.
		/// </summary>
		public string ErrorCaption
		{
			get { return base.GetXmlNodeString("@errorCaption"); }
			set { base.SetXmlNodeString("@errorCaption", value); }
		}

		/// <summary>
		/// Gets or sets the name of the value area field header in the PivotTable. 
		/// This caption is shown when the PivotTable when two or more fields are in the values area.
		/// </summary>
		public string DataCaption
		{
			get { return base.GetXmlNodeString("@dataCaption"); }
			set { base.SetXmlNodeString("@dataCaption", value); }
		}

		/// <summary>
		/// Gets or sets whether to show field headers.
		/// </summary>
		public bool ShowHeaders
		{
			get { return base.GetXmlNodeBool("@showHeaders", true); }
			set { base.SetXmlNodeBool("@showHeaders", value, true); }
		}

		/// <summary>
		/// Gets a value indicating whether or not to hide the values row.
		/// </summary>
		public bool HideValuesRow
		{
			get { return base.GetXmlNodeBool("d:extLst/d:ext/x14:pivotTableDefinition/@hideValuesRow", false); }
		}

		/// <summary>
		/// Gets a value indicating whether fields should be shown ascending or in data source order.
		/// </summary>
		public bool FieldListSortAscending
		{
			get { return base.GetXmlNodeBool("fieldListSortAscending"); }
		}

		/// <summary>
		/// Gets or sets the number of page fields to display before starting another row or column.
		/// </summary>
		public int PageWrap
		{
			get { return base.GetXmlNodeIntNull("@pageWrap") ?? 0; }
			set
			{
				if (value < 0)
					throw new Exception("Value can't be negative");
				base.SetXmlNodeString("@pageWrap", value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets whether the legacy auto formatting has been applied to the PivotTable view.
		/// Corresponds to the "Autofit column widths on update" pivot table setting.
		/// </summary>
		public bool UseAutoFormatting
		{
			get { return base.GetXmlNodeBool("@useAutoFormatting", false); }
			set { base.SetXmlNodeBool("@useAutoFormatting", value, false); }
		}

		/// <summary>
		/// Gets a value indicating whether or not cell formatting should be preserved on update.
		/// </summary>
		public bool PreserveFormatting
		{
			get { return base.GetXmlNodeBool("@preserveFormatting", true); }
		}

		/// <summary>
		/// Gets or sets whether the in-grid drop zones should be displayed at runtime, and whether classic layout is applied.
		/// </summary>
		public bool GridDropZones
		{
			get { return base.GetXmlNodeBool("@gridDropZones"); }
			set { base.SetXmlNodeBool("@gridDropZones", value); }
		}

		/// <summary>
		/// Gets or sets the indentation increment for compact axis or can be used to set the Report Layout to Compact Form.
		/// NOTE: For some reason, Excel stores this value as 1 less than it is set in the UI. Also, 0 indent is stored as 127.
		/// </summary>
		public int Indent
		{
			get { return base.GetXmlNodeInt("@indent", 1); }
			set { base.SetXmlNodeString("@indent", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets whether data fields in the PivotTable should be displayed in outline form.
		/// </summary>
		public bool OutlineData
		{
			get { return base.GetXmlNodeBool("@outlineData", false); }
			set { base.SetXmlNodeBool("@outlineData", value); }
		}

		/// <summary>
		/// Gets or sets whether new fields should have their outline flag set to true.
		/// </summary>
		public bool Outline
		{
			get { return base.GetXmlNodeBool("@outline", false); }
			set { base.SetXmlNodeBool("@outline", value); }
		}

		/// <summary>
		/// Gets or sets whether the fields of a PivotTable can have multiple filters set on them.
		/// </summary>
		public bool MultipleFieldFilters
		{
			get { return base.GetXmlNodeBool("@multipleFieldFilters", true); }
			set { base.SetXmlNodeBool("@multipleFieldFilters", value); }
		}

		/// <summary>
		/// Gets the "Totals and Filters" pivot table setting value for "Use Custom Lists when sorting".
		/// </summary>
		public bool CustomListSort
		{
			get { return base.GetXmlNodeBool("@customListSort", true); }
		}

		/// <summary>
		/// Gets or sets whether new fields should have their compact flag set to true.
		/// </summary>
		public bool Compact
		{
			get { return base.GetXmlNodeBool("@compact", true); }
			set { base.SetXmlNodeBool("@compact", value); }
		}

		/// <summary>
		/// Gets or sets whether the field next to the data field in the PivotTable should be displayed in the same column of the spreadsheet.
		/// </summary>
		public bool CompactData
		{
			get { return base.GetXmlNodeBool("@compactData", true); }
			set { base.SetXmlNodeBool("@compactData", value); }
		}

		/// <summary>
		/// Gets or sets the string to be displayed for grand totals.
		/// </summary>
		public string GrandTotalCaption
		{
			get { return base.GetXmlNodeString("@grandTotalCaption"); }
			set { base.SetXmlNodeString("@grandTotalCaption", value); }
		}

		/// <summary>
		/// Gets or sets the string to be displayed in row header in compact mode.
		/// </summary>
		public string RowHeaderCaption
		{
			get { return base.GetXmlNodeString("@rowHeaderCaption"); }
			set { base.SetXmlNodeString("@rowHeaderCaption", value); }
		}

		/// <summary>
		/// Gets or sets whether the "Layout and Format" pivot table setting "For empty values show:" is enabled.
		/// </summary>
		public bool ShowMissing
		{
			get { return base.GetXmlNodeBool("@showMissing", true); }
			set { base.SetXmlNodeBool("@showMissing", value, true); }
		}

		/// <summary>
		/// Gets or sets the string to be displayed in cells with no value.
		/// Corresponds to the "Layout and Format" pivot table setting "For empty values show: [missingCaption]".
		/// </summary>
		public string MissingCaption
		{
			get { return base.GetXmlNodeString("@missingCaption", null); }
			set { base.SetXmlNodeString("@missingCaption", value); }
		}

		/// <summary>
		/// Gets or sets the value that corresponds to the Excel pivot table 
		/// display setting "Show items with no data on columns".
		/// </summary>
		public bool ShowEmptyColumn
		{
			get { return base.GetXmlNodeBool("@showEmptyCol", false); }
			set { base.SetXmlNodeBool("@showEmptyCol", value, false); }
		}

		/// <summary>
		/// Gets or sets the value that corresponds to the Excel pivot table 
		/// display setting "Show items with no data on rows".
		/// </summary>
		public bool ShowEmptyRow
		{
			get { return base.GetXmlNodeBool("@showEmptyRow", false); }
			set { base.SetXmlNodeBool("@showEmptyRow", value, false); }
		}

		/// <summary>
		/// Gets or sets the value that corresponds to the Excel pivot table
		/// display setting "Enable cell editing in the values area".
		/// </summary>
		/// <remarks>This only applies to pivot tables with an OLAP data source.</remarks>
		public bool EnableEdit
		{
			get { return base.GetXmlNodeBool("//x14:pivotTableDefinition/@enableEdit", false); }
			set { base.SetXmlNodeBool("//x14:pivotTableDefinition/@enableEdit", value, false); }
		}

		/// <summary>
		/// Gets or sets the value that corresponds to the Excel pivot table
		/// totals and filters settings "Include filtered items in totals".
		/// </summary>
		/// <remarks>This only applies to pivot tables with an OLAP data source.</remarks>
		public bool VisualTotals
		{
			get { return base.GetXmlNodeBool("@visualTotals", true); }
			set { base.SetXmlNodeBool("@visualTotals", value, true); }
		}

		/// <summary>
		/// Gets or sets the value that correspons to the Excel pivot table
		/// totals and filters settings "Mark totals with *".
		/// </summary>
		/// <remarks>This only applies to pivot tables with an OLAP data source.</remarks>
		public bool AsterisksTotals
		{
			get { return base.GetXmlNodeBool("@asteriskTotals", false); }
			set { base.SetXmlNodeBool("@asteriskTotals", value, false); }
		}

		/// <summary>
		/// Gets or sets the first row of the PivotTable header relative to the top left cell in the ref value.
		/// </summary>
		public int FirstHeaderRow
		{
			get { return base.GetXmlNodeInt(FirstHeaderRowPath); }
			set { base.SetXmlNodeString(FirstHeaderRowPath, value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the first column of the PivotTable data relative to the top left cell in the ref value.
		/// </summary>
		public int FirstDataRow
		{
			get { return base.GetXmlNodeInt(FirstDataRowPath); }
			set { base.SetXmlNodeString(FirstDataRowPath, value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the first column of the PivotTable data relative to the top left cell in the ref value.
		/// </summary>
		public int FirstDataCol
		{
			get { return base.GetXmlNodeInt(FirstDataColumnPath); }
			set { base.SetXmlNodeString(FirstDataColumnPath, value.ToString()); }
		}

		/// <summary>
		/// Gets a value indicating whether or not to merge and center cells with labels.
		/// </summary>
		public bool MergeAndCenterCellsWithLabels
		{
			get { return base.GetXmlNodeBool("@mergeItem", false); }
		}

		/// <summary>
		/// Gets a value corresponding to the pivot table setting "Display fields in report area".
		/// </summary>
		public bool PageOverThenDown
		{
			get { return base.GetXmlNodeBool("@pageOverThenDown", false); }
		}

		/// <summary>
		/// Gets the fields in the table.
		/// </summary>
		public ExcelPivotTableFieldCollection Fields
		{
			get
			{
				if (myFields == null)
				{
					var pivotFieldsNode = this.TopNode.SelectSingleNode("d:pivotFields", this.NameSpaceManager);
					myFields = new ExcelPivotTableFieldCollection(this.NameSpaceManager, pivotFieldsNode, this);
				}
				return myFields;
			}
		}

		/// <summary>
		/// Gets the row label fields.
		/// </summary>
		public ExcelPivotTableRowColumnFieldCollection RowFields
		{
			get
			{
				if (myRowFields == null)
				{
					var rowFieldsNode = this.TopNode.SelectSingleNode("d:rowFields", this.NameSpaceManager);
					myRowFields = new ExcelPivotTableRowColumnFieldCollection(this.NameSpaceManager, rowFieldsNode, this, PivotTableItemType.Row);
				}
				return myRowFields;
			}
		}

		/// <summary>
		/// Gets the column label fields.
		/// </summary>
		public ExcelPivotTableRowColumnFieldCollection ColumnFields
		{
			get
			{
				if (myColumnFields == null)
				{
					var columnFieldsNode = this.TopNode.SelectSingleNode("d:colFields", this.NameSpaceManager);
					myColumnFields = new ExcelPivotTableRowColumnFieldCollection(this.NameSpaceManager, columnFieldsNode, this, PivotTableItemType.Column);
				}
				return myColumnFields;
			}
		}

		/// <summary>
		/// Gets the value fields.
		/// </summary>
		public ExcelPivotTableDataFieldCollection DataFields
		{
			get
			{
				if (myDataFields == null)
				{
					var dataFieldsNode = this.TopNode.SelectSingleNode("d:dataFields", this.NameSpaceManager);
					myDataFields = new ExcelPivotTableDataFieldCollection(this.NameSpaceManager, dataFieldsNode, this);
				}
				return myDataFields;
			}
		}

		/// <summary>
		/// Gets the pivot table filter fields.
		/// </summary>
		public ExcelPageFieldCollection PageFields
		{
			get
			{
				if (myPageFields == null)
				{
					var pageFieldsNode = this.TopNode.SelectSingleNode("d:pageFields", this.NameSpaceManager);
					if (pageFieldsNode == null)
						return null;
					myPageFields = new ExcelPageFieldCollection(this.NameSpaceManager, pageFieldsNode, this);
				}
				return myPageFields;
			}
		}

		/// <summary>
		/// Gets the pivot table label/value filters.
		/// </summary>
		public ExcelPivotFieldFiltersCollection Filters
		{
			get
			{
				if (myFilters == null)
				{
					var filtersNode = this.TopNode.SelectSingleNode("d:filters", this.NameSpaceManager);
					if (filtersNode == null)
						return null;
					myFilters = new ExcelPivotFieldFiltersCollection(this.NameSpaceManager, filtersNode, this);
				}
				return myFilters;
			}
		}

		/// <summary>
		/// Gets a boolean that determines whether any row items exist in this pivot table.
		/// </summary>
		public bool HasRowItems => this.TopNode.SelectSingleNode("d:rowItems", this.NameSpaceManager) != null;

		/// <summary>
		/// Gets the row items.
		/// </summary>
		public ItemsCollection RowItems
		{
			get
			{
				if (myRowItems == null)
					myRowItems = new ItemsCollection(this.NameSpaceManager, this.TopNode.SelectSingleNode("d:rowItems", this.NameSpaceManager));
				return myRowItems;
			}
		}

		/// <summary>
		/// Gets a boolean that determines whether any column items exist in this pivot table.
		/// </summary>
		public bool HasColumnItems => this.TopNode.SelectSingleNode("d:colItems", this.NameSpaceManager) != null;

		/// <summary>
		/// Gets the column items.
		/// </summary>
		public ItemsCollection ColumnItems
		{
			get
			{
				if (myColumnItems == null)
					myColumnItems = new ItemsCollection(this.NameSpaceManager, this.TopNode.SelectSingleNode("d:colItems", this.NameSpaceManager));
				return myColumnItems;
			}
		}

		/// <summary>
		/// Gets the collection of format filters.
		/// </summary>
		public ExcelFormatsCollection Formats
		{
			get
			{
				if (myFormats == null)
				{
					var formatsNode = this.TopNode.SelectSingleNode("d:formats", this.NameSpaceManager);
					if (formatsNode == null)
						return null;
					myFormats = new ExcelFormatsCollection(this.NameSpaceManager, formatsNode);
				}
				return myFormats;
			}
		}

		/// <summary>
		/// Gets or sets the pivot style name that is used for custom styles.
		/// </summary>
		public string StyleName
		{
			get { return base.GetXmlNodeString(StyleNamePath); }
			set
			{
				if (value.StartsWith("PivotStyle"))
				{
					try
					{
						myTableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), true);
					}
					catch
					{
						myTableStyle = TableStyles.Custom;
					}
				}
				else if (value == "None")
				{
					myTableStyle = TableStyles.None;
					value = "";
				}
				else
					myTableStyle = TableStyles.Custom;
				base.SetXmlNodeString(StyleNamePath, value, true);
			}
		}

		/// <summary>
		/// Gets or sets the table style. If this is a custom property, the style from the StyleName propery is used.
		/// </summary>
		public TableStyles TableStyle
		{
			get { return myTableStyle; }
			set
			{
				myTableStyle = value;
				if (value != TableStyles.Custom)
					base.SetXmlNodeString(StyleNamePath, "PivotStyle" + value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the row item tree.
		/// </summary>
		internal PivotItemTreeNode RowItemsRoot { get; set; } = new PivotItemTreeNode(-1);

		/// <summary>
		/// Gets or sets the column item tree.
		/// </summary>
		internal PivotItemTreeNode ColumnItemsRoot { get; set; } = new PivotItemTreeNode(-1);

		/// <summary>
		/// Gets or sets the cache id of the pivot table.
		/// </summary>
		internal int CacheID
		{
			get { return base.GetXmlNodeInt("@cacheId"); }
			set { base.SetXmlNodeString("@cacheId", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the pivot table part.
		/// </summary>
		internal Packaging.ZipPackagePart Part { get; set; }

		/// <summary>
		/// Gets or sets the worksheet-pivot table relationship.
		/// </summary>
		internal Packaging.ZipPackageRelationship WorksheetRelationship { get; set; }

		/// <summary>
		/// Gets or sets the cache definition-pivot table relationship.
		/// </summary>
		internal Packaging.ZipPackageRelationship CacheDefinitionRelationship { get; set; }

		/// <summary>
		/// Gets a list of pivot table row header cell models.
		/// </summary>
		internal List<PivotTableHeader> RowHeaders { get; } = new List<PivotTableHeader>();

		/// <summary>
		/// Gets a list of pivot table column header cell models.
		/// </summary>
		internal List<PivotTableHeader> ColumnHeaders { get; } = new List<PivotTableHeader>();

		/// <summary>
		/// Gets a value indicating whether there is more than one data field in the row fields.
		/// </summary>
		internal bool HasRowDataFields => this.RowFields.Any(c => c.Index == -2);

		/// <summary>
		/// Gets a value indicating whether there is more than one data field in the column fields.
		/// </summary>
		internal bool HasColumnDataFields => this.ColumnFields.Any(c => c.Index == -2);

		/// <summary>
		/// Gets a value indicating whether label/value filters are applied to the pivot table.
		/// </summary>
		internal bool HasFilters => this.Filters != null;

		internal ExcelWorkbook Workbook { get; private set; }

		private int OriginalTableEndRow { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTable"/> from a relationship.
		/// </summary>
		/// <param name="rel">The relationship to create the pivot table from.</param>
		/// <param name="sheet">The worksheet the pivot table is on.</param>
		internal ExcelPivotTable(Packaging.ZipPackageRelationship rel, ExcelWorksheet sheet) :
			 base(sheet.NameSpaceManager)
		{
			this.Workbook = sheet.Workbook;
			this.PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
			this.WorksheetRelationship = rel;
			var pck = sheet.Package.Package;
			this.Part = pck.GetPart(this.PivotTableUri);

			this.PivotTableXml = new XmlDocument();
			XmlHelper.LoadXmlSafe(this.PivotTableXml, this.Part.GetStream());
			this.InitSchemaNodeOrder();
			this.TopNode = this.PivotTableXml.DocumentElement;
			this.Address = new ExcelAddress(sheet.Name, base.GetXmlNodeString("d:location/@ref"));

			var rels = this.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
			if (rels.Count != 1)
				throw new InvalidOperationException($"Pivot table had an unexpected number ({rels.Count}) of pivot cache definitions.");
			this.CacheDefinitionRelationship = rels.FirstOrDefault();

			this.LoadFields();
		}

		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTable"/>.
		/// </summary>
		/// <param name="sheet">The worksheet of the pivot table.</param>
		/// <param name="address">The address of the pivot table.</param>
		/// <param name="sourceAddress">The address of the source data.</param>
		/// <param name="name">The name of the pivot table.</param>
		/// <param name="tblId">The pivot table id.</param>
		internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddress address, ExcelRangeBase sourceAddress, string name, int tblId) :
			 base(sheet.NameSpaceManager)
		{
			this.Workbook = sheet.Workbook;
			this.Address = new ExcelAddress(sheet.Name, address.Address);
			this.Address = address;
			var pck = sheet.Package.Package;

			this.PivotTableXml = new XmlDocument();
			LoadXmlSafe(this.PivotTableXml, this.GetStartXml(name, tblId, address, sourceAddress), Encoding.UTF8);
			this.TopNode = this.PivotTableXml.DocumentElement;
			this.PivotTableUri = GetNewUri(pck, "/xl/pivotTables/pivotTable{0}.xml", ref tblId);
			this.InitSchemaNodeOrder();

			this.Part = pck.CreatePart(this.PivotTableUri, ExcelPackage.schemaPivotTable);
			this.PivotTableXml.Save(this.Part.GetStream());

			// Worksheet-PivotTable relationship
			this.WorksheetRelationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, this.PivotTableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
			bool cacheDefinitionFound = false;
			foreach (var cache in this.Worksheet.Workbook.PivotCacheDefinitions)
			{
				if (cache.GetSourceRangeAddress().IsEquivalentRange(sourceAddress))
				{
					this.CacheDefinition = cache;
					cacheDefinitionFound = true;
					break;
				}
			}
			if (!cacheDefinitionFound)
			{
				this.CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress, tblId);
				sheet.Workbook.PivotCacheDefinitions.Add(this.CacheDefinition);
			}
			// CacheDefinition-PivotTable relationship
			this.CacheDefinitionRelationship = this.Part.CreateRelationship(UriHelper.ResolvePartUri(this.PivotTableUri, this.CacheDefinition.CacheDefinitionUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
			sheet.Workbook.AddPivotTable(this.CacheID.ToString(), this.CacheDefinition.CacheDefinitionUri);
			this.LoadFields();
			using (var range = sheet.Cells[address.Address])
			{
				range.Clear();
			}
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Refresh the <see cref="ExcelPivotTable"/> based on the <see cref="ExcelPivotCacheDefinition"/>.
		/// </summary>
		internal void RefreshFromCache(StringResources stringResources)
		{
			this.Workbook.FormulaParser.Logger?.LogFunction(nameof(this.RefreshFromCache));

			// Update pivot fields and sort pivot field items.
			this.Workbook.FormulaParser.Logger?.LogFunction(nameof(this.UpdatePivotFields));
			this.UpdatePivotFields();

			this.RowHeaders.Clear();
			this.ColumnHeaders.Clear();

			// Update the rowItems.
			if (this.HasRowItems)
			{
				this.Workbook.FormulaParser.Logger?.LogFunction($"{nameof(this.UpdateRowColumnItems)}: Rows");
				this.UpdateRowColumnItems(this.RowFields, this.RowItems, true, stringResources);
			}

			// Update the colItems.
			if (this.HasColumnItems)
			{
				this.Workbook.FormulaParser.Logger?.LogFunction($"{nameof(this.UpdateRowColumnItems)}: Columns");
				this.UpdateRowColumnItems(this.ColumnFields, this.ColumnItems, false, stringResources);
			}

			// Update the pivot table data.
			if (this.HasRowItems || this.HasColumnItems)
			{
				this.Workbook.FormulaParser.Logger?.LogFunction(nameof(this.UpdateWorksheet));
				this.UpdateWorksheet(stringResources);
			}

			// Remove the 'm' (missing) xml attribute from each pivot field item, if it exists, to prevent 
			// corrupting the workbook, since Excel automatically adds them.
			this.RemovePivotFieldItemMAttribute();

			// pivotSelections are causing corruptions when left. Deleting for meow.
			this.Worksheet.View.RemovePivotSelections();

			// Refreshing a pivot table expands all values, so update the field items to make the icons match.
			this.ExpandAllFieldItems();

			// Update conditional formatting rule addresses.
			this.UpdateConditionalFormattingRuleAddresses();
		}

		/// <summary>
		/// Gets a dictionary of field index to a list of items that are to be included in the pivot table.
		/// </summary>
		/// <returns>A dictionary of pivot field index to a list of field item value indices.</returns>
		internal Dictionary<int, List<int>> GetPageFieldIndices()
		{
			if (this.PageFields == null || this.PageFields.Count == 0)
				return null;
			var pageFieldIndices = new Dictionary<int, List<int>>();
			foreach (var pageField in this.PageFields)
			{
				if (pageField.Item != null)
				{
					int pivotFieldItemValue = this.Fields[pageField.Field].Items[pageField.Item.Value].X;
					pageFieldIndices.Add(pageField.Field, new List<int> { pivotFieldItemValue });
				}
				else
				{
					// If page fields are multi-selected, pivot field items that are 
					// not selected are flagged as hidden.
					var pageFieldItems = this.Fields[pageField.Field].Items;
					if (pageFieldItems != null)
					{
						foreach (var item in pageFieldItems.Where(i => !i.Hidden))
						{
							if (!pageFieldIndices.ContainsKey(pageField.Field))
								pageFieldIndices.Add(pageField.Field, new List<int>());
							pageFieldIndices[pageField.Field].Add(item.X);
						}
					}
				}
			}
			return pageFieldIndices;
		}

		/// <summary>
		/// Gets a list of the unsupported features that are enabled on this pivot table.
		/// </summary>
		/// <param name="unsupportedFeatures">The unsupported features enabled on this pivot table.</param>
		/// <returns>True if unsupported features are found, otherwise false.</returns>
		internal bool TryGetUnsupportedFeatures(out List<string> unsupportedFeatures)
		{
			unsupportedFeatures = new List<string>();
			void addUnsupportedFeature(List<string> unsupportedList, string message) => unsupportedList.Add($"[{this.Address.FullAddress}] {message}");

			foreach (var dataField in this.DataFields)
			{
				if (dataField.ShowDataAs == ShowDataAs.Difference || dataField.ShowDataAs == ShowDataAs.PercentDiff
					|| dataField.ShowDataAs == ShowDataAs.RunTotal || dataField.ShowDataAs == ShowDataAs.PercentOfRunningTotal || dataField.ShowDataAs == ShowDataAs.RankAscending
					|| dataField.ShowDataAs == ShowDataAs.RankDescending || dataField.ShowDataAs == ShowDataAs.Index)
				{
					addUnsupportedFeature(unsupportedFeatures, $"Data field '{dataField.Name}' show data as setting '{dataField.ShowDataAs}'");
				}

				// Disallow the '(next)' and '(previous)' options. 
				if (dataField.BaseField == 1048829)
					addUnsupportedFeature(unsupportedFeatures, $"Data field '{dataField.Name}' '(next)' option selected");
				else if (dataField.BaseField == 1048828)
					addUnsupportedFeature(unsupportedFeatures, $"Data field '{dataField.Name}' '(previous)' option selected");
			}
			foreach (var field in this.Fields)
			{
				if (field.RepeatItemLabels)
					addUnsupportedFeature(unsupportedFeatures, $"Field '{field.Name}' repeat item labels enabled");
				if (field.InsertBlankLine)
					addUnsupportedFeature(unsupportedFeatures, $"Field '{field.Name}' insert blank line enabled");
				if (field.ShowAll)
					addUnsupportedFeature(unsupportedFeatures, $"Field '{field.Name}' show items with no data enabled");
				if (field.InsertPageBreak)
					addUnsupportedFeature(unsupportedFeatures, $"Field '{field.Name}' insert page break after each item enabled");
				if (field.IncludeNewItemsInFilter)
					addUnsupportedFeature(unsupportedFeatures, $"Field '{field.Name}' include new items in filter enabled");
			}
			var filters = base.TopNode.SelectSingleNode("d:filters", base.NameSpaceManager);
			if (filters != null)
			{
				foreach (var filter in this.Filters)
				{
					if (filter.FieldFilterType == FieldFilter.Label)
					{
						if (filter.LabelFilterType == LabelFilterType.CaptionBetween || filter.LabelFilterType == LabelFilterType.CaptionNotBetween)
						{
							var stringOne = new WildCardValueMatcher().ExcelWildcardToRegex(filter.StringValueOne);
							var stringTwo = new WildCardValueMatcher().ExcelWildcardToRegex(filter.StringValueTwo);
							bool stringOneHasWildcards = stringOne.Contains("*") || stringOne.Contains(".");
							bool stringTwoHasWildcards = stringTwo.Contains("*") || stringTwo.Contains(".");
							if (stringOneHasWildcards || stringTwoHasWildcards)
								addUnsupportedFeature(unsupportedFeatures, $"'{this.Fields[filter.Field].Name}' field has {filter.LabelFilterType} label filter with wildcards enabled.");
						}
					}
					else if (filter.FieldFilterType == FieldFilter.Value)
						addUnsupportedFeature(unsupportedFeatures, $"'{this.Fields[filter.Field].Name}' field has value filters enabled.");
					else if (filter.FieldFilterType == FieldFilter.Report)
						addUnsupportedFeature(unsupportedFeatures, $"'{this.Fields[filter.Field].Name}' field has report filters enabled.");
				}
			}
			if (this.MergeAndCenterCellsWithLabels)
				addUnsupportedFeature(unsupportedFeatures, "Merge and center cells with labels enabled");
			if (this.PageOverThenDown)
				addUnsupportedFeature(unsupportedFeatures, "Display fields in report filter area over then down enabled");
			if (this.PageWrap != 0)
				addUnsupportedFeature(unsupportedFeatures, "Report filter fields per [row|column] > 0");
			if (!string.IsNullOrEmpty(this.ErrorCaption))
				addUnsupportedFeature(unsupportedFeatures, "Error caption enabled");
			if (this.ShowError)
				addUnsupportedFeature(unsupportedFeatures, "Show error enabled");
			if (!this.PreserveFormatting)
				addUnsupportedFeature(unsupportedFeatures, "Preserve formatting disabled");
			if (this.MultipleFieldFilters)
				addUnsupportedFeature(unsupportedFeatures, "Multiple field filters enabled");
			if (!this.CustomListSort)
				addUnsupportedFeature(unsupportedFeatures, "Use Custom Lists when sorting disabled");
			if (!this.ShowDataTips)
				addUnsupportedFeature(unsupportedFeatures, "Show contextual tooltips disabled");
			if (!this.ShowHeaders)
				addUnsupportedFeature(unsupportedFeatures, "Display field captions and filter dropdowns disabled");
			if (this.GridDropZones)
				addUnsupportedFeature(unsupportedFeatures, "Grid drop zones enabled");
			if (!this.HideValuesRow)
				addUnsupportedFeature(unsupportedFeatures, "Show values row enabled");
			if (this.FieldListSortAscending)
				addUnsupportedFeature(unsupportedFeatures, "Field list sort ascending enabled");
			if (this.ShowEmptyColumn)
				addUnsupportedFeature(unsupportedFeatures, "Show items with no data on columns enabled");
			if (this.ShowEmptyRow)
				addUnsupportedFeature(unsupportedFeatures, "Show items with no data on rows enabled");
			if (this.ShowMemberPropertyTips)
				addUnsupportedFeature(unsupportedFeatures, "Show peroperties in tooltips enabled");
			if (this.EnableEdit)
				addUnsupportedFeature(unsupportedFeatures, "Enable cell editing in the values area enabled");
			if (!this.VisualTotals)
				addUnsupportedFeature(unsupportedFeatures, "Include filtered items in totals enabled");
			if (this.AsterisksTotals)
				addUnsupportedFeature(unsupportedFeatures, "Mark totals with * disabled");
			return unsupportedFeatures.Any();
		}
		#endregion

		#region Private Methods
		private void UpdateConditionalFormattingRuleAddresses()
		{
			foreach (var rule in this.Worksheet.ConditionalFormatting)
			{
				string newAddress = string.Empty;
				var ruleAddress = rule.Address.AddressSpaceSeparated.Split(' ');
				ExcelAddress cellRange = null;
				foreach (var cells in ruleAddress)
				{
					cellRange = new ExcelAddress(cells);
					if (cellRange.IsSingleCell)
					{
						if (this.Address.Start.Row <= cellRange.Start.Row && cellRange.End.Row <= this.Address.End.Row
							&& this.Address.Start.Column <= cellRange.Start.Column && cellRange.End.Column <= this.Address.End.Column)
							newAddress += cells + " ";
					}
					else if (cells.Contains(":"))
					{
						ExcelAddress updatedAddress = null;
						if (cellRange.Start.Row < this.Address.End.Row && cellRange.End.Row > this.Address.End.Row
							&& this.Address.Start.Column <= cellRange.Start.Column && cellRange.End.Column <= this.Address.End.Column)
						{
							// The pivot table's size is smaller than the original, so the conditional formatting range's end row is past the end of the pivot table.
							updatedAddress = new ExcelAddress(cellRange.Start.Row, cellRange.Start.Column, this.Address.End.Row, cellRange.End.Column);
						}
						else if (this.Address.Start.Row > cellRange.Start.Row && cellRange.End.Row > this.Address.Start.Row
							&& this.Address.Start.Column <= cellRange.Start.Column && cellRange.End.Column <= this.Address.End.Column)
						{
							// The pivot table's size is smaller than the original, so the conditional formatting range's start row is before the beginning of the pivot table.
							updatedAddress = new ExcelAddress(this.Address.Start.Row, cellRange.Start.Column, cellRange.End.Row, cellRange.End.Column);
						}
						else if (this.Address.Start.Row <= cellRange.Start.Row && cellRange.End.Row == this.OriginalTableEndRow
							&& this.Address.Start.Column <= cellRange.Start.Column && cellRange.End.Column <= this.Address.End.Column)
						{
							// The pivot table's size is bigger than the original, so expand the conditional formatting rule's end row.
							updatedAddress = new ExcelAddress(cellRange.Start.Row, cellRange.Start.Column, this.Address.End.Row, cellRange.End.Column);
						}
						if (updatedAddress != null)
							newAddress += updatedAddress.AddressSpaceSeparated + " ";
					}
				}
				if (!string.IsNullOrEmpty(newAddress))
					rule.Address = new ExcelAddress(newAddress);
			}
		}


		private void UpdatePivotFields()
		{
			// Update pivotField items to match corresponding cacheField sharedItems.
			foreach (var pivotField in this.Fields)
			{
				var fieldItems = pivotField.Items;
				var cacheField = this.CacheDefinition.CacheFields[pivotField.Index];
				if (!cacheField.HasSharedItems)
					continue;

				if (fieldItems.Count > 0)
				{
					// Only sort pivot field items if the pivot field is not part of a date grouping.
					if (this.CacheDefinition.CacheFields[pivotField.Index].FieldGroup == null)
					{
						// Preserve the "@h" attribute for fields marked as hidden as well as the totals field items.
						var totalsFieldItems = fieldItems.Where(i => !string.IsNullOrEmpty(i.T)).ToList();
						var hiddenFieldItemsDictionary = fieldItems
							.Where(i => string.IsNullOrEmpty(i.T))
							.ToDictionary(i => i.X, i => i.Hidden);
						fieldItems.Clear();
						var sharedItemsList = this.CacheDefinition.CacheFields[pivotField.Index].SharedItems.ToList();

						// Sort the row/column headers.
						var sortedList = this.SortField(pivotField.Sort, pivotField).ToList();

						// Assign the correct index value to each item.
						for (int i = 0; i < sortedList.Count(); i++)
						{
							var field = sortedList[i];
							int index = sharedItemsList.FindIndex(x => x == field);
							fieldItems.AddItem(index);
							if (hiddenFieldItemsDictionary.ContainsKey(index))
								fieldItems[i].Hidden = hiddenFieldItemsDictionary[index];
						}
						// Add back the totals field items.
						fieldItems.AppendItems(totalsFieldItems);
					}
				}
			}
		}

		private int GetRepeatedItemsCount(List<Tuple<int, int>> indices, List<Tuple<int, int>> lastChildIndices)
		{
			int repeatedItemsCount = 0;
			for (; repeatedItemsCount < lastChildIndices.Count; repeatedItemsCount++)
			{
				if (indices[repeatedItemsCount].Item2 != lastChildIndices[repeatedItemsCount].Item2)
					break;
			}
			return repeatedItemsCount;
		}

		private void CreateColumnSubtotalNode(PivotItemTreeNode node, int repeatedItemsCount, List<Tuple<int, int>> indices)
		{
			// Create subtotal nodes if subtotals are enabled and we are not at the root node.
			// Also, if node has a grandchild or the leaf node is not data field create a subtotal node.
			bool defaultSubtotal = node.PivotFieldIndex != -2 && this.Fields[node.PivotFieldIndex].SubtotalLocation != SubtotalLocation.Off;
			if (defaultSubtotal && node.Value != -1 &&
				(node.Children.FirstOrDefault()?.HasChildren == true || this.ColumnFields.Last().Index != -2))
			{
				bool isLastNonDataField = this.ColumnFields.Skip(repeatedItemsCount).All(x => x.Index == -2);
				repeatedItemsCount = this.ColumnFields.ToList().FindIndex(i => i.Index == node.PivotFieldIndex);
				bool isAboveDataField = !indices.Any(x => x.Item1 == -2);
				var pivotField = this.Fields[node.PivotFieldIndex];
				var functionNames = pivotField.GetEnabledSubtotalTypes();
				foreach (var function in functionNames)
				{
					// If the node is above a data field node and there are multiple data fields, then create a subtotal node for each data field. 
					if (this.DataFields.Count > 0 && isAboveDataField && !isLastNonDataField && this.HasColumnDataFields)
						this.CreateTotalNodes(node?.CacheRecordIndices, function, false, indices, null, repeatedItemsCount, true, this.HasColumnDataFields);
					// Otherwise, if the node is not the last non-data field node and is below a data field node, then only create one subtotal node.
					else if (!isLastNonDataField && (!isAboveDataField || !this.HasColumnDataFields))
						this.CreateTotalNodes(node?.CacheRecordIndices, function, false, indices, null, repeatedItemsCount, false, this.HasColumnDataFields, node.DataFieldIndex);
				}
			}
		}

		private void CreateRowSubtotalNodes(PivotItemTreeNode node, int repeatedItemsCount, List<Tuple<int, int>> indices)
		{
			bool defaultSubtotal = node.PivotFieldIndex != -2 && this.Fields[node.PivotFieldIndex].SubtotalLocation != SubtotalLocation.Off;
			if (defaultSubtotal && node.Value != -1 && (node.Children.FirstOrDefault()?.HasChildren == true || this.RowFields.Last().Index != -2))
			{
				repeatedItemsCount = this.RowFields.ToList().FindIndex(i => i.Index == node.PivotFieldIndex);
				bool isAboveDataField = !indices.Any(i => i.Item1 == -2);
				var pivotField = this.Fields[node.PivotFieldIndex];
				bool createTabularSubtotalNodes = true;
				if (this.RowFields.Any(x => x.Outline == false))
				{
					// Create a subtotal node for any of the following cases:
					//		-The pivot field has tabular form enabled.
					//		-The pivot field has tabular form disabled and subtotal bottom is enabled.
					//		-There are multiple datafields and the current node is above a datafield and the pivot field has tabular form enabled.
					createTabularSubtotalNodes = !pivotField.Outline || (pivotField.Outline && !pivotField.SubtotalTop) || (this.HasRowDataFields && pivotField.Outline && isAboveDataField);
				}

				var functionNames = pivotField.GetEnabledSubtotalTypes();
				foreach (var function in functionNames)
				{
					if (createTabularSubtotalNodes && this.RowFields.Any(x => x.Outline == false))
					{
						// Create subtotal nodes for pivot fields with tabular form enabled.
						bool isLastNonDataField = this.RowFields.Skip(repeatedItemsCount).All(x => x.Index == -2);
						// If the node is above a data field node and there are multiple data fields, then create a subtotal node for each data field. 
						if (this.DataFields.Count > 0 && isAboveDataField && !isLastNonDataField && this.HasRowDataFields)
							this.CreateTotalNodes(node?.CacheRecordIndices, function, true, indices, pivotField, repeatedItemsCount, true, this.HasRowDataFields);
						// Otherwise, if the node is not the last non-data field node and is below a data field node, then only create one subtotal node.
						else if (!isLastNonDataField && (!isAboveDataField || !this.HasRowDataFields))
							this.CreateTotalNodes(node?.CacheRecordIndices, function, true, indices, pivotField, repeatedItemsCount, false, this.HasRowDataFields, node.DataFieldIndex);
					}
					else
					{
						// If above a datafield, subtotals are always shown if defaultSubtotal is enabled.
						if (this.HasRowDataFields && isAboveDataField && (node.Children.FirstOrDefault()?.HasChildren == true || this.RowFields.Last().Index != -2))
							this.CreateTotalNodes(node?.CacheRecordIndices, function, true, indices, pivotField, repeatedItemsCount, true, this.HasRowDataFields, node.DataFieldIndex);
						else if (!node.SubtotalTop || functionNames.Count > 1)
						{
							// If this child is only followed by datafields, do not write out a subtotal node.
							// In other words, treat this child as the leaf node.
							if (!isAboveDataField)
								this.CreateTotalNodes(node?.CacheRecordIndices, function, true, indices, pivotField, repeatedItemsCount, false, this.HasRowDataFields, node.DataFieldIndex);
							else if (node.Children.Any(c => !c.IsDataField || c.Children.Any()))
								this.CreateTotalNodes(node?.CacheRecordIndices, function, true, indices, pivotField, repeatedItemsCount, true, this.HasRowDataFields, node.DataFieldIndex);
						}
					}
				}
			}
		}

		private List<Tuple<int, int>> BuildRowItems(PivotItemTreeNode root, List<Tuple<int, int>> indices, List<Tuple<int, int>> lastChildIndices, int indent)
		{
			var pivotField = root.PivotFieldItemIndex != -2 ? this.Fields[root.PivotFieldIndex] : null;
			if (this.IsHidden(root))
				return indices;
			int repeatedItemsCount = 0;
			if (root.Value != -1)
			{
				bool createNode = pivotField != null && (pivotField.Outline && (pivotField.Compact || this.OutlineData || !this.CompactData));
				if (!root.HasChildren || (root.PivotFieldIndex == -2 && (this.CompactData || this.OutlineData)) || createNode)
				{
					repeatedItemsCount = this.GetRepeatedItemsCount(indices, lastChildIndices);
					string subtotalType = root.PivotFieldIndex == -2 ? null : this.GetRowFieldSubtotalType(pivotField);
					bool isDataField = root.PivotFieldIndex == -2;
					indent = indent == -1 ? 0 : indent;
					this.RowHeaders.Add(new PivotTableHeader(root.CacheRecordIndices, indices, pivotField, root.DataFieldIndex, false,
						!root.HasChildren, isDataField, subtotalType, !indices.Any(x => x.Item1 == -2), indent: indent));
					this.RowItems.AddColumnItem(indices.ToList(), repeatedItemsCount, root.DataFieldIndex);
					lastChildIndices = indices.ToList();

					if (!root.HasChildren)
						return indices.ToList();
				}
			}

			root.ExpandIfDataFieldNode(this.DataFields.Count);
			indent = this.GetIndentationLevel(root, indent);
			for (int i = 0; i < root.Children.Count; i++)
			{
				var child = root.Children[i];
				int pivotFieldItemIndex = child.PivotFieldIndex == -2 ? child.DataFieldIndex : child.PivotFieldItemIndex;
				var childIndices = indices.ToList();
				childIndices.Add(new Tuple<int, int>(child.PivotFieldIndex, pivotFieldItemIndex));
				lastChildIndices = this.BuildRowItems(child, childIndices, lastChildIndices, indent);
			}

			this.CreateRowSubtotalNodes(root, repeatedItemsCount, indices);
			return lastChildIndices;
		}
		
		private int GetIndentationLevel(PivotItemTreeNode parent, int parentIndentation)
		{
			if (parent.PivotFieldIndex == -2)
				return parentIndentation + 1;
			else if (parent.Value == -1)  // Children of the root node are not indented.
				return 0;
			var pivotField = this.Fields[parent.PivotFieldIndex];
			if (!pivotField.Compact || parent.IsTabularForm)  // Only children of compact fields are indented.
				return 0;
			return parentIndentation + 1;
		}

		private string GetRowFieldSubtotalType(ExcelPivotTableField pivotField)
		{
			if (pivotField?.SubtotalTop == true)
			{
				// If a pivot field only has one subtotal function type, subtotal top is respected.
				// If there are multiple functions, subtotals will be put at the bottom for each custom function.
				var subtotalTypes = pivotField.GetEnabledSubtotalTypes();
				if (subtotalTypes.Count == 1)
				{
					var type = subtotalTypes.First();
					if (!type.IsEquivalentTo("default") && this.RowFields.Last().Index != pivotField.Index)
						return type;
				}
				else if (subtotalTypes.Count > 1)
					return "none";
			}
			return null;
		}
		
		private bool IsHidden(PivotItemTreeNode node)
		{
			var pivotField = node.PivotFieldItemIndex != -2 ? this.Fields[node.PivotFieldIndex] : null;
			return pivotField != null && pivotField.Items[node.PivotFieldItemIndex].Hidden;
		}

		private List<Tuple<int, int>> BuildColumnItems(PivotItemTreeNode node, List<Tuple<int, int>> indices, List<Tuple<int, int>> lastChildIndices)
		{
			if (this.IsHidden(node))
				return indices;
			int repeatedItemsCount = 0;
			// Base case (leaf node).
			if (!node.HasChildren)
			{
				repeatedItemsCount = this.GetRepeatedItemsCount(indices, lastChildIndices);
				var header = new PivotTableHeader(node.CacheRecordIndices, indices, null, node.DataFieldIndex, false, true, node.IsDataField);
				this.ColumnHeaders.Add(header);
				this.ColumnItems.AddColumnItem(indices.ToList(), repeatedItemsCount, node.DataFieldIndex);
				return indices.ToList();
			}

			node.ExpandIfDataFieldNode(this.DataFields.Count);

			for (int i = 0; i < node.Children.Count; i++)
			{
				var child = node.Children[i];
				int pivotFieldItemIndex = child.PivotFieldIndex == -2 ? child.DataFieldIndex : child.PivotFieldItemIndex;

				var childIndices = indices.ToList();
				childIndices.Add(new Tuple<int, int>(child.PivotFieldIndex, pivotFieldItemIndex));
				lastChildIndices = this.BuildColumnItems(child, childIndices, lastChildIndices);
			}

			this.CreateColumnSubtotalNode(node, repeatedItemsCount, indices);
			return lastChildIndices;
		}

		private PivotItemTreeNode BuildRowColTree(ExcelPivotTableRowColumnFieldCollection rowColFields, Dictionary<int, List<int>> cacheRecordPageFieldIndices, StringResources stringResources)
		{
			// Build a tree using the cache records. Each node in the tree is a cache record that is identified by the row or column field indices.
			var rootNode = new PivotItemTreeNode(-1);
			for (int i = 0; i < this.CacheDefinition.CacheRecords.Count; i++)
			{
				var cacheRecord = this.CacheDefinition.CacheRecords[i];
				var currentNode = rootNode;

				bool createdFilterTreeNode = false;
				for (int j = 0; j < rowColFields.Count; j++)
				{
					int rowColFieldIndex = rowColFields[j].Index;
					var cacheFields = this.myCacheDefinition.CacheFields;

					// These variables are only used for date groupings.
					// Since rowColFieldIndex can be set to the base field index if there are date groupings, this keeps track of the original row field index.
					int originalIndex = rowColFieldIndex;
					var groupBy = rowColFieldIndex == -2 ? null : cacheFields[originalIndex].FieldGroup?.GroupBy;
					// Reset rowColFieldIndex to the base field index if the row/column field index refers to a date grouping field.
					rowColFieldIndex = rowColFieldIndex >= cacheRecord.Items.Count ? cacheFields[rowColFieldIndex].FieldGroup.BaseField : rowColFieldIndex;

					if (currentNode.Children.Any(c => c.IsDataField))
						currentNode = currentNode.GetChildNode(-2);
					else if (rowColFieldIndex == -2)
						currentNode = currentNode.AddChild(-2);  // Create a datafield node.
					else
					{
						int recordItemValue = int.Parse(cacheRecord.Items[rowColFieldIndex].Value);
						var sharedItemValue = cacheFields[rowColFieldIndex].SharedItems[recordItemValue];

						// If filters are enabled, then only create a tree node if it satifies the filter criteria.
						if (this.HasFilters)
							createdFilterTreeNode = this.ShouldCreateTreeNodeWithFilter(cacheRecord, rowColFieldIndex);

						if (this.HasFilters == false || (this.HasFilters && createdFilterTreeNode))
						{
							if (cacheFields[originalIndex].IsGroupField)
								currentNode = this.CreateTreeNodeWithGrouping(sharedItemValue, originalIndex, currentNode, recordItemValue, groupBy, cacheFields, 
									cacheRecordPageFieldIndices, cacheRecord, stringResources, rowColFieldIndex < cacheRecord.Items.Count);
							else
								currentNode = this.CreateTreeNode(false, currentNode, this.Fields[rowColFieldIndex], recordItemValue, rowColFieldIndex, 
									cacheRecordPageFieldIndices, cacheRecord, sharedItemValue.Value, stringResources);
						}

						// This cache record does not contain the page field indices, so continue to the next record.
						if (currentNode == null || (this.HasFilters && !createdFilterTreeNode))
							break;
					}
					if (!currentNode.CacheRecordIndices.Contains(i))
						currentNode.CacheRecordIndices.Add(i);
				}
			}
			return rootNode;
		}

		private bool ShouldCreateTreeNodeWithFilter(CacheRecordNode record, int fieldIndex)
		{
			foreach (var filter in this.Filters)
			{
				int filterFieldIndex = filter.Field;
				// Get the shared item string from the record index.
				int recordItemValue = int.Parse(record.Items[filterFieldIndex].Value);
				var sharedItem = this.myCacheDefinition.CacheFields[filterFieldIndex].SharedItems[recordItemValue];
				var sharedItemValue = sharedItem.Value;
				bool sharedItemIsNumericType = sharedItem.Type == PivotCacheRecordType.n;
				bool isMatch = filter.MatchesFilterCriteriaResult(sharedItemValue, sharedItemIsNumericType);
				if (!isMatch)
					return false;
			}
			return true;
		}

		private PivotItemTreeNode CreateTreeNodeWithGrouping(CacheItem sharedItemValue, int groupingIndex, PivotItemTreeNode currentNode, int recordItemValue, PivotFieldDateGrouping? groupBy,
			IReadOnlyList<CacheFieldNode> cacheFields, Dictionary<int, List<int>> cacheRecordPageFieldIndices, CacheRecordNode cacheRecord, StringResources stringResources, bool baseGroup)
		{
			var pivotField = this.Fields[groupingIndex];
			// A sharedItem value of type DateTime indicates the current pivot field is part of a date grouping. Otherwise, create a new node if necessary.
			if (sharedItemValue.Type == PivotCacheRecordType.d)
			{
				// Handles field date groupings.
				var date = DateTime.Parse(sharedItemValue.Value);
				var searchValue = this.GetItemValueByGroupingType(date, groupBy);
				currentNode = this.CreateTreeNode(true, currentNode, pivotField, recordItemValue, groupingIndex, cacheRecordPageFieldIndices, cacheRecord, searchValue, stringResources, cacheFields[groupingIndex], baseGroup);
			}
			else if (cacheFields[groupingIndex].FieldGroup.DiscreteGroupingProperties != null)
			{
				// Handles custom field groupings.
				var groupingFieldGroup = cacheFields[groupingIndex].FieldGroup;
				int discretePrValue = int.Parse(groupingFieldGroup.DiscreteGroupingProperties[recordItemValue].Value);
				var groupingSearchValue = groupingFieldGroup.GroupItems[discretePrValue].Value;
				currentNode = this.CreateTreeNode(false, currentNode, pivotField, discretePrValue, groupingIndex, cacheRecordPageFieldIndices, cacheRecord, groupingSearchValue, stringResources);
			}
			return currentNode;
		}

		private string GetItemValueByGroupingType(DateTime dateTime, PivotFieldDateGrouping? groupBy)
		{
			string sharedItemGroupingValue = string.Empty;
			if (groupBy == PivotFieldDateGrouping.Years)
				sharedItemGroupingValue = dateTime.Year.ToString();
			else if (groupBy == PivotFieldDateGrouping.Months)
				sharedItemGroupingValue = dateTime.ToString("MMM");
			else if (groupBy == PivotFieldDateGrouping.Days)
				sharedItemGroupingValue = dateTime.Day + "-" + dateTime.ToString("MMM");
			else if (groupBy == PivotFieldDateGrouping.Minutes)
				sharedItemGroupingValue = ":" + dateTime.ToString("mm");
			else if (groupBy == PivotFieldDateGrouping.Seconds)
				sharedItemGroupingValue = ":" + dateTime.ToString("ss");
			else if (groupBy == PivotFieldDateGrouping.Hours)
			{
				int hour = dateTime.Hour == 00 ? 12 : dateTime.Hour;
				sharedItemGroupingValue = hour + " " + dateTime.ToString("tt", Thread.CurrentThread.CurrentCulture);
			}
			else if (groupBy == PivotFieldDateGrouping.Quarters)
			{
				int quarter = (dateTime.Month - 1) / 3 + 1;
				sharedItemGroupingValue = "Qtr" + quarter;
			}
			return sharedItemGroupingValue;
		}

		private PivotItemTreeNode CreateTreeNode(bool isDateGrouping, PivotItemTreeNode currentNode, ExcelPivotTableField pivotField, int recordItemValue, int pivotFieldIndex,
			Dictionary<int, List<int>> cacheRecordPageFieldIndices, CacheRecordNode cacheRecord, string searchValue, StringResources stringResources, CacheFieldNode cacheField = null, bool baseGroup = false)
		{
			// If an identical child already exists, continue. Otherwise, create a new child.
			searchValue = searchValue ?? stringResources.BlankValueHeaderCaption;
			if (currentNode.HasChild(searchValue))
				return currentNode.GetChildNode(searchValue);
			else
			{
				if (cacheRecordPageFieldIndices?.Any() == true && !cacheRecord.ContainsPageFieldIndices(cacheRecordPageFieldIndices))
					return null;
				currentNode.SubtotalTop = pivotField.SubtotalTop;
				int pivotFieldItemIndex = 0;
				if (isDateGrouping)
				{
					var cacheFieldGroupItems = cacheField.FieldGroup.GroupItems;
					if (baseGroup)
					{
						int groupItemIndex = cacheFieldGroupItems.ToList().FindIndex(x => x.Value.IsEquivalentTo(searchValue));
						pivotFieldItemIndex = pivotField.Items.ToList().FindIndex(c => c.X == groupItemIndex);
					}
					else
						pivotFieldItemIndex = cacheFieldGroupItems.ToList().FindIndex(x => x.Value.IsEquivalentTo(searchValue));
				}
				else
					pivotFieldItemIndex = pivotField.Items.ToList().FindIndex(c => c.X == recordItemValue);
				bool isTabularForm = pivotField?.Outline == false;
				return currentNode.AddChild(recordItemValue, pivotFieldIndex, pivotFieldItemIndex, searchValue, isTabularForm);
			}
		}

		private List<CacheItem> SortField(eSortType sortOrder, ExcelPivotTableField pivotField)
		{
			var fieldItems = pivotField.Items;
			var sharedItems = this.CacheDefinition.CacheFields[pivotField.Index].SharedItems;
			bool sortDescending = sortOrder == eSortType.Descending && pivotField.AutoSortScopeReferences.Count == 0;
			bool hasMissingValueType = sharedItems.Any(x => x.Type == PivotCacheRecordType.m);
			var copyList = sharedItems.ToList();
			if (hasMissingValueType)
				copyList.RemoveAll(i => i.Type == PivotCacheRecordType.m);

			IOrderedEnumerable<CacheItem> sortedList = null;
			if (copyList.All(t => t.Type == PivotCacheRecordType.n))
				sortedList = sortDescending ? copyList.OrderByDescending(x => double.Parse(x.Value)) : copyList.OrderBy(x => double.Parse(x.Value));
			else
			{
				bool isDateTime = true;
				foreach (var item in copyList)
				{
					isDateTime = DateTime.TryParseExact(item.Value, "MMMM", Thread.CurrentThread.CurrentCulture, System.Globalization.DateTimeStyles.None, out var result);
					if (!isDateTime)
						break;
					sortedList = sortDescending ? copyList.ToList().OrderByDescending(m => DateTime.ParseExact(m.Value, "MMMM", Thread.CurrentThread.CurrentCulture))
						: copyList.ToList().OrderBy(m => DateTime.ParseExact(m.Value, "MMMM", Thread.CurrentThread.CurrentCulture));
				}
				if (!isDateTime)
					sortedList = sortDescending ? copyList.OrderByDescending(x => x.Value) : copyList.OrderBy(x => x.Value);
			}

			var returnList = sortedList.ToList();
			if (hasMissingValueType)
			{
				// Add "(blank)" value header to the end of the sorted list.
				int index = sharedItems.ToList().FindIndex(i => i.Type == PivotCacheRecordType.m);
				returnList.Add(sharedItems[index]);
			}

			return returnList;
		}

		private void RemovePivotFieldItemMAttribute()
		{
			foreach (var pivotField in this.Fields)
			{
				if (pivotField.Items.Count > 0)
				{
					foreach (var item in pivotField.Items)
					{
						var mAttribute = item.TopNode.Attributes["m"];
						if (mAttribute != null && int.Parse(mAttribute.Value) == 1)
							item.TopNode.Attributes.Remove(mAttribute);
					}
				}
			}
		}

		private void UpdateRowColumnItems(ExcelPivotTableRowColumnFieldCollection rowColFieldCollection, ItemsCollection collection, bool isRowItems, StringResources stringResources)
		{
			// Update the rowItems or colItems.
			if (rowColFieldCollection.Any())
			{
				collection.Clear();
				var pageFieldIndices = this.GetPageFieldIndices();
				if (isRowItems)
				{
					this.RowItemsRoot = this.BuildRowColTree(rowColFieldCollection, pageFieldIndices, stringResources);
					this.RowItemsRoot.SortChildren(this);
					if (this.RowItemsRoot.HasChildren)
					{
						this.BuildRowItems(this.RowItemsRoot, new List<Tuple<int, int>>(), new List<Tuple<int, int>>(), -1);
						this.CreateLeafFieldCustomFieldSettingsNode(true);
					}
				}
				else
				{
					this.ColumnItemsRoot = this.BuildRowColTree(rowColFieldCollection, pageFieldIndices, stringResources);
					this.ColumnItemsRoot.SortChildren(this);
					if (this.ColumnItemsRoot.HasChildren)
					{
						this.BuildColumnItems(this.ColumnItemsRoot, new List<Tuple<int, int>>(), new List<Tuple<int, int>>());
						this.CreateLeafFieldCustomFieldSettingsNode(false);
					}
				}
				// Create grand total items if necessary.
				bool grandTotals = isRowItems ? this.RowGrandTotals : this.ColumnGrandTotals;
				if (grandTotals && isRowItems && !(this.RowFields.Count == 1 && this.RowFields.First().Index == -2))
					this.CreateTotalNodes(null, "grand", true, null, null, 0, false, this.HasRowDataFields);
				else if (grandTotals && !isRowItems && !(this.ColumnFields.Count == 1 && this.ColumnFields.First().Index == -2))
					this.CreateTotalNodes(null, "grand", false, null, null, 0, false, this.HasColumnDataFields);
			}
			else
			{
				string xmlTag = isRowItems ? "d:rowFields" : "d:colFields";
				// If there are no row/column fields, then remove tag or else it will corrupt the workbook.
				this.TopNode.RemoveChild(this.TopNode.SelectSingleNode(xmlTag, this.NameSpaceManager));
				var headerCollection = isRowItems ? this.RowHeaders : this.ColumnHeaders;
				var header = new PivotTableHeader(null, null, null, 0, false, false, false);
				header.IsPlaceHolder = true;
				headerCollection.Add(header);
			}
		}

		private void CreateLeafFieldCustomFieldSettingsNode(bool isRowField)
		{
			ExcelPivotTableRowColumnFieldCollection collection = isRowField ? this.RowFields : this.ColumnFields;
			if (collection.Last().Index == -2)
				return;
			var lastPivotField = isRowField ? this.Fields[collection.Last().Index] : this.Fields[collection.Last().Index];
			var lastItemSubtotalValue = lastPivotField.Items.Last().T;
			if (collection.Count > 1 && lastPivotField.DefaultSubtotal && !lastItemSubtotalValue.IsEquivalentTo("default") && !string.IsNullOrEmpty(lastItemSubtotalValue))
			{
				var functionNames = lastPivotField.GetEnabledSubtotalTypes();
				int itemsCount = lastPivotField.Items.Count(i => string.IsNullOrEmpty(i.T));
				for (int i = 0; i < itemsCount; i++)
				{
					foreach (var function in functionNames)
					{
						List<Tuple<int, int>> indices = new List<Tuple<int, int>>();
						if (i == 0 && function == functionNames.First())
						{
							for (int k = 0; k < collection.Count - 1; k++)
							{
								// Excel stores 1048832 for the member properties ('x' value) for the row/column item.
								indices.Add(new Tuple<int, int>(lastPivotField.Index, 1048832));
							}
							indices.Add(new Tuple<int, int>(lastPivotField.Index, i));
						}
						else
							indices.Add(new Tuple<int, int>(lastPivotField.Index, i));
						var hasDataFields = collection.Any(r => r.Index == -2);
						var list = this.GetCacheRecordList(lastPivotField.Index, lastPivotField.Items[i].X);
						int repeatedItemsCount = i == 0 && function == functionNames.First() ? 0 : collection.Count - 1;
						this.CreateTotalNodes(list, function, isRowField, indices, lastPivotField, repeatedItemsCount, true, hasDataFields, 0, true);
					}
				}
			}
		}

		private List<int> GetCacheRecordList(int pivotFieldIndex, int itemIndex)
		{
			var indicesList = new List<int>();
			for (int i = 0; i < this.CacheDefinition.CacheRecords.Count; i++)
			{
				var record = this.CacheDefinition.CacheRecords[i];
				int.TryParse(record.Items[pivotFieldIndex].Value, out var value);
				if (value == itemIndex)
					indicesList.Add(i);
			}
			return indicesList;
		}

		private void CreateTotalNodes(List<int> cacheRecordIndices, string totalType, bool isRowItem, List<Tuple<int, int>> indices, ExcelPivotTableField pivotField,
			int repeatedItemsCount, bool multipleSubtotalDataFields, bool hasDataFields, int dataFieldIndex = 0, bool customField = false)
		{
			var itemsCollection = isRowItem ? this.RowItems : this.ColumnItems;
			var headerCollection = isRowItem ? this.RowHeaders : this.ColumnHeaders;

			// Variables are set for the default case where item type is a grand total.
			bool hasMultipleDataFields = this.DataFields.Any() && hasDataFields;
			int index = hasMultipleDataFields ? this.DataFields.Count : 1;
			int xMember = 0;
			bool aboveDataField = false;
			bool grandTotal = true;

			// Reset variables if item type is a subtotal.
			if (!totalType.IsEquivalentTo("grand"))
			{
				index = multipleSubtotalDataFields && hasMultipleDataFields ? this.DataFields.Count : 1;
				xMember = indices.Last().Item2;
				aboveDataField = !indices.Any(x => x.Item1 == -2);
				grandTotal = false;
			}

			// Create one xml node if the header is below a data field header with the correct data field index (not necessarily zero).
			if (!aboveDataField && !grandTotal)
			{
				var header = new PivotTableHeader(cacheRecordIndices, indices, pivotField, dataFieldIndex, grandTotal, false, false, totalType, aboveDataField);
				itemsCollection.AddSumNode(totalType, repeatedItemsCount, xMember, dataFieldIndex);
				headerCollection.Add(header);
			}
			else
			{
				// Create the xml node and row/column header.
				for (int i = 0; i < index; i++)
				{
					var header = new PivotTableHeader(cacheRecordIndices, indices, pivotField, i, grandTotal, false, false, totalType, aboveDataField);
					var list = customField ? indices : null;
					itemsCollection.AddSumNode(totalType, repeatedItemsCount, xMember, i, list);
					headerCollection.Add(header);
				}
			}
		}

		private void UpdateWorksheet(StringResources stringResources)
		{
			// Update the row and column header values in the worksheet.
			this.UpdateRowColumnHeaders(stringResources);
			// Update the pivot table's address.
			this.Address = this.GetNewAddress();
			if (this.DataFields.Any())
			{
				var dataManager = new PivotTableDataManager(this);
				dataManager.UpdateWorksheet();
			}
		}

		private ExcelAddress GetNewAddress()
		{
			this.OriginalTableEndRow = this.Address.End.Row;
			int endRow = this.Address.Start.Row + this.FirstDataRow + this.RowHeaders.Count - 1;
			// If there are no data fields, then don't find the offset to obtain the first data column.
			int endColumn = this.Address.Start.Column;
			if (this.DataFields.Any())
				endColumn += this.FirstDataCol + this.ColumnHeaders.Count - 1;
			else
			{
				endColumn += this.RowFields.Count(f => !f.Compact || !f.Outline);
				var lastRowField = this.RowFields.LastOrDefault();
				if (lastRowField != null && lastRowField.Outline == false)
					endColumn--;
			}
			return new ExcelAddress(this.Worksheet.Name, this.Address.Start.Row, this.Address.Start.Column, endRow, endColumn);
		}

		private void ClearTable()
		{
			// Clear out the pivot table in the worksheet except for the "<Row/Column> Labels" cell in the top left-ish corner.
			ExcelRange columnLabelsCell = null, rowLabelsCell = null, singleDataFieldLabelCell = null;
			string columnLabel = null, rowLabel = null, singleDataFieldLabel = null;
			if (this.ColumnFields.Any())
			{
				columnLabelsCell = this.Worksheet.Cells[this.Address.Start.Row, this.Address.Start.Column + this.FirstDataCol];
				columnLabel = columnLabelsCell.Value?.ToString();
			}
			if (this.RowFields.Any())
			{
				rowLabelsCell = this.Worksheet.Cells[this.Address.Start.Row + this.FirstDataRow - 1, this.Address.Start.Column];
				rowLabel = rowLabelsCell.Value?.ToString();
			}
			if (this.DataFields.Count == 1)
			{
				singleDataFieldLabelCell = this.Worksheet.Cells[this.Address.Start.Row, this.Address.Start.Column];
				singleDataFieldLabel = singleDataFieldLabelCell.Value?.ToString();
			}

			this.Worksheet.Cells[this.Address.Address].Clear();

			// Write the labels back into the header cells.
			if (columnLabelsCell != null)
				columnLabelsCell.Value = columnLabel;
			if (rowLabelsCell != null)
				rowLabelsCell.Value = rowLabel;
			if (singleDataFieldLabelCell != null)
				singleDataFieldLabelCell.Value = singleDataFieldLabel;
		}

		private void UpdateRowColumnHeaders(StringResources stringResources)
		{
			// Clear out the pivot table in the worksheet.
			this.ClearTable();

			// Update the row headers in the worksheet.
			this.WriteRowHeaders(stringResources);

			// Update the column headers in the worksheet.
			this.WriteColumnHeaders(stringResources);
		}

		private void WriteRowHeaders(StringResources stringResources)
		{
			int row = this.Address.Start.Row + this.FirstDataRow;
			int column = this.Address.Start.Column;
			int previousColumn = column;
			var columnFieldNames = new List<string>();
			if (this.RowFields.Any())
			{
				for (int i = 0; i < this.RowItems.Count; i++)
				{
					var item = this.RowItems[i];
					var header = this.RowHeaders[i];
					ExcelRange cell = null;
					for (int j = 0; j < item.Count; j++)
					{
						if (item[j] == 1048832)
							continue;
						string totalValue = this.GetTotalCaptionCellValue(this.RowFields, item, header, stringResources, j);
						if (!string.IsNullOrEmpty(totalValue))
						{
							column = this.GetRowHeaderColumn(item, this.Address.Start.Column, header.CacheRecordIndices, j);
							cell = this.Worksheet.Cells[row, column];
							cell.Value = totalValue;
							cell.Style.Indent = this.GetIndent(header.Indent);
						}
						else
						{
							var itemIndex = item.RepeatedItemsCount == 0 ? j : j + item.RepeatedItemsCount;
							string sharedItemValue = this.GetSharedItemValue(this.RowFields, item, itemIndex, j, stringResources);
							column = j == 0 ? this.GetRowHeaderColumn(item, this.Address.Start.Column, header.CacheRecordIndices, j) : column + 1;
							row = this.GetRowHeaderRow(item.RepeatedItemsCount, row, j);
							cell = this.Worksheet.Cells[row, column];
							cell.Value = sharedItemValue;
							cell.Style.Indent = this.GetIndent(header.Indent);

							if (column > previousColumn)
							{
								var rowFieldIndex = this.RowFields[item.RepeatedItemsCount + j].Index;
								var pivotFieldName = rowFieldIndex == -2 ? stringResources.ValuesCaption : this.Fields[rowFieldIndex].Name;
								if (!columnFieldNames.Contains(pivotFieldName))
									columnFieldNames.Add(pivotFieldName);
							}
						}
					}
					row++;
					previousColumn = column;
				}
			}
			else if (this.DataFields.Count == 1)
				this.Worksheet.Cells[row++, this.Address.Start.Column].Value = this.DataFields.First().Name;

			this.WriteHeadersForOutlineAndTabularForm(columnFieldNames, true);
		}

		private int GetRowHeaderRow(int repeatedItemsCount, int row, int j)
		{
			if (j == 0)
				return row;
			int parentRowFieldIndex = this.RowFields[repeatedItemsCount + j - 1].Index;
			if (parentRowFieldIndex != -2)
			{
				var parentPivotField = this.Fields[parentRowFieldIndex];
				if (!parentPivotField.Compact && parentPivotField.Outline)
					return row + 1;
			}
			return row;
		}

		private int GetRowHeaderColumn(RowColumnItem item, int column, List<Tuple<int, int>> cacheRecordIndices, int j)
		{
			bool singleXMemberProperties = item.Count == 0;
			if (item.RepeatedItemsCount == 0)
				return singleXMemberProperties || item.First() == 1048832 ? column : column + j;
			else
			{
				int count = singleXMemberProperties && cacheRecordIndices != null ? cacheRecordIndices.Count - 1 : item.RepeatedItemsCount;
				for (int i = 0; i < count; i++)
				{
					int index = this.RowFields[i].Index;
					if (index == -2)
					{
						if (!this.CompactData)
							column += 1;
					}
					else if (!this.Fields[index].Compact || !this.Fields[index].Outline)
						column += 1;
				}
				return column;
			}
		}

		private void WriteHeadersForOutlineAndTabularForm(List<string> headers, bool isRowHeader)
		{
			int row = isRowHeader ? this.Address.Start.Row + this.FirstDataRow - 1 : this.Address.Start.Row;
			int column = isRowHeader ? this.Address.Start.Column + 1 : this.Address.Start.Column + this.FirstDataCol;
			for (int i = 0; i < headers.Count; i++)
			{
				this.Worksheet.Cells[row, column++].Value = headers[i];
			}
		}

		private int GetIndent(int depth)
		{
			// Excel stores a zero indent as 127, and other values are stored as 1 less than the UI shows.
			var indent = this.Indent == 127 ? 0 : this.Indent + 1;
			return depth * indent;
		}

		private void WriteColumnHeaders(StringResources stringResources)
		{
			int startRow = this.Address.Start.Row + this.FirstHeaderRow;
			int column = this.Address.Start.Column + this.FirstDataCol;
			var columnFieldNames = new List<string>();
			bool hasCompactField = this.RowFields.Where(i => i.Index != -2).Any(x => this.Fields[x.Index].Compact);
			if (this.ColumnFields.Any())
			{
				for (int i = 0; i < this.ColumnItems.Count; i++)
				{
					var item = this.ColumnItems[i];
					int startHeaderRow = startRow;
					for (int j = 0; j < item.Count; j++)
					{
						if (item[j] == 1048832)
						{
							startHeaderRow++;
							continue;
						}
						string itemType = this.GetTotalCaptionCellValue(this.ColumnFields, item, this.ColumnHeaders[i], stringResources, j);
						if (!string.IsNullOrEmpty(itemType))
						{
							int totalCaptionRow = startHeaderRow + item.RepeatedItemsCount;
							this.Worksheet.Cells[totalCaptionRow, column].Value = itemType;
						}
						else
						{
							var columnFieldIndex = item.RepeatedItemsCount == 0 ? j : j + item.RepeatedItemsCount;
							var sharedItem = this.GetSharedItemValue(this.ColumnFields, item, columnFieldIndex, j, stringResources);
							var cellRow = item.RepeatedItemsCount == 0 ? startHeaderRow : startHeaderRow + item.RepeatedItemsCount;
							this.Worksheet.Cells[cellRow, column].Value = sharedItem;
							startHeaderRow++;

							if (!hasCompactField && !this.CompactData && i == 0 && this.ColumnFields.Count > 1)
							{
								var pivotFieldIndex = this.ColumnFields[columnFieldIndex].Index;
								if (pivotFieldIndex == -2)
									columnFieldNames.Add(stringResources.ValuesCaption);
								else
								{
									var pivotField = this.Fields[pivotFieldIndex];
									columnFieldNames.Add(pivotField.Name);
								}
							}
						}
					}
					column++;
				}
			}
			// If there are no column headers and only one data field, print the name of the data field for the column.
			else if (this.DataFields.Count == 1)
				this.Worksheet.Cells[this.Address.Start.Row, column].Value = this.DataFields.First().Name;
			
			// If outline or tabular form is enabled, then write the header values to the correct cell.
			if (columnFieldNames.Any())
				this.WriteHeadersForOutlineAndTabularForm(columnFieldNames, false);
		}

		private string GetTotalCaptionCellValue(ExcelPivotTableRowColumnFieldCollection field, RowColumnItem item, PivotTableHeader header, StringResources stringResources, int j = 0)
		{
			string totalHeader = string.Empty;
			if (!string.IsNullOrEmpty(item.ItemType))
			{
				if (item.ItemType.IsEquivalentTo("grand"))
				{
					// If the pivot table has more than one data field, then use the name of the data field in the total.
					if ((this.HasRowDataFields && field == this.RowFields) || (this.HasColumnDataFields && field == this.ColumnFields))
					{
						string dataFieldName = this.DataFields[item.DataFieldIndex].Name;
						totalHeader = string.Format(stringResources.TotalCaptionWithFollowingValue, dataFieldName);
					}
					else
						totalHeader = stringResources.GrandTotalCaption;
				}
				else
				{
					int repeatedItemsCount = item.First() == 1048832 ? item.Count - 1 : item.RepeatedItemsCount;
					var itemName = this.GetSharedItemValue(field, item, repeatedItemsCount, j, stringResources);
					if (this.DataFields.Count > 1 && header.IsAboveDataField && 
						((this.HasRowDataFields && field == this.RowFields) || (this.HasColumnDataFields && field == this.ColumnFields)))
					{
						string dataFieldName = this.DataFields[item.DataFieldIndex].Name;
						if (item.ItemType.IsEquivalentTo("default"))
							totalHeader = $"{itemName} {dataFieldName}";
						else
						{
							// TODO: This should be translated, but we do not have support for that yet.
							var functionName = ExcelPivotTableField.FunctionTypesToUserFunctionCaptions[item.ItemType];
							var dataFieldPivotFieldName = this.Fields[this.DataFields[item.DataFieldIndex].Index].Name;
							totalHeader = $"{itemName} {functionName} of {dataFieldPivotFieldName}";
						}
					}
					else
					{
						if (item.ItemType.IsEquivalentTo("default"))
							totalHeader = string.Format(stringResources.TotalCaptionWithPrecedingValue, itemName);
						else
						{
							// TODO: Function names should be translated, but we do not have support for that yet.
							var functionName = ExcelPivotTableField.FunctionTypesToUserFunctionCaptions[item.ItemType];
							totalHeader = $"{itemName} {functionName}";
						}
					}
				}
			}
			return totalHeader;
		}

		private string GetSharedItemValue(ExcelPivotTableRowColumnFieldCollection field, RowColumnItem rowColItem, int repeatedItemsCount, int xMemberIndex, StringResources stringResources)
		{
			var sharedItemValue = string.Empty;
			var pivotFieldIndex = field[repeatedItemsCount].Index;
			// A field that has an 'x' attribute equal to -2 is a special row/column field that indicates the
			// pivot table has more than one data field. Excel uses this to display the headings for the data 
			// values and how to group them in relation to other rows/columns. 
			// If a special field alrady exists in that collection, then another one will not be generated.
			if (pivotFieldIndex == -2)
				return this.DataFields[rowColItem.DataFieldIndex].Name;
			var pivotField = this.Fields[pivotFieldIndex];
			var cacheItemIndex = pivotField.Items[rowColItem[xMemberIndex]].X;
			var cacheField = this.CacheDefinition.CacheFields[pivotFieldIndex];
			var item = cacheField.IsGroupField ? cacheField.FieldGroup.GroupItems[cacheItemIndex] : cacheField.SharedItems[cacheItemIndex];
			if (item.Type == PivotCacheRecordType.b)
				sharedItemValue = item.Value.IsEquivalentTo("1") ? stringResources.PivotTableTrueCaption : stringResources.PivotTableFalseCaption;
			else if (item.Type == PivotCacheRecordType.d)
				sharedItemValue = this.GetDateFormat(pivotFieldIndex, item.Value, cacheField.NumFormatId, rowColItem[xMemberIndex], string.IsNullOrEmpty(rowColItem.ItemType));
			else
			{
				sharedItemValue = item.Value;
				if (!string.IsNullOrEmpty(sharedItemValue) && sharedItemValue[0] == '<' && DateTime.TryParse(sharedItemValue.Substring(1), out var date))
					sharedItemValue = this.GetItemValueByGroupingType(date, cacheField.FieldGroup.GroupBy);
			}
			sharedItemValue = sharedItemValue ?? stringResources.BlankValueHeaderCaption;
			return sharedItemValue;
		}

		private string GetDateFormat(int fieldIndex, string sharedItemValue, int numFormatId, int xMemberIndex, bool nonTotalHeader)
		{
			var dateTime = DateTime.Parse(sharedItemValue);
			string date = this.GetTranslatedDate(numFormatId, dateTime);
			if (this.Formats != null && nonTotalHeader)
			{
				foreach (var format in this.Formats)
				{
					int formatId = format.FormatId;
					var formatting = this.Workbook.Styles.Dxfs[formatId];
					if (format.PivotArea.ReferencesCollection != null)
					{
						foreach (var reference in format.PivotArea.ReferencesCollection)
						{
							if (reference.ItemIndexCount == 0)
							{
								if (reference.FieldIndex == fieldIndex)
									return this.GetTranslatedDate(formatting.NumberFormat.NumFmtID, dateTime);
							}
							else
							{
								foreach (var item in reference.SharedItems)
								{
									if (int.Parse(item.Value) == xMemberIndex && reference.FieldIndex == fieldIndex && reference.Selected)
										return this.GetTranslatedDate(formatting.NumberFormat.NumFmtID, dateTime);
								}
							}
						}
					}
				}
			}
			return date;
		}

		private void ExpandAllFieldItems()
		{
			// Set all items to be expanded.
			foreach (var pivotField in this.Fields)
			{
				if (pivotField.Items.Count == 0)
					continue;
				foreach (var item in pivotField.Items)
				{
					if (item.HideDetails == false)
						item.HideDetails = true;
				}
			}
		}

		private string GetTranslatedDate(int nfId, DateTime date)
		{
			var styles = this.Worksheet.Workbook.Styles;
			var nfID = nfId;
			ExcelNumberFormatXml.ExcelFormatTranslator nf = null;
			for (int i = 0; i < styles.NumberFormats.Count; i++)
			{
				if (nfID == styles.NumberFormats[i].NumFmtId)
				{
					nf = styles.NumberFormats[i].FormatTranslator;
					break;
				}
			}
			string format, textFormat;
			format = nf.NetFormat;
			textFormat = nf.NetTextFormat;
			return ExcelRangeBase.FormatValue(date, nf, format, textFormat);
		}

		private void InitSchemaNodeOrder()
		{
			this.SchemaNodeOrder = new string[] { "location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "pageItems", "dataFields", "dataItems", "formats", "pivotTableStyleInfo" };
		}

		private void LoadFields()
		{
			// Add fields.
			int index = 0;
			var fieldNodes = this.CacheDefinition.TopNode.SelectNodes("d:cacheFields/d:cacheField", this.NameSpaceManager);
			if (fieldNodes != null)
			{
				foreach (var pivotField in this.Fields)
				{
					pivotField.SetCacheFieldNode(fieldNodes[index++]);
				}
			}
		}

		private string GetStartXml(string name, int id, ExcelAddress address, ExcelAddress sourceAddress)
		{
			string xml = string.Format("<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{0}\" cacheId=\"{1}\" dataOnRows=\"1\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Data\"  createdVersion=\"4\" showMemberPropertyTips=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" indent=\"0\" compact=\"0\" compactData=\"0\" gridDropZones=\"1\">", name, id);

			xml += string.Format("<location ref=\"{0}\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" /> ", address.FirstAddress);
			xml += string.Format("<pivotFields count=\"{0}\">", sourceAddress._toCol - sourceAddress._fromCol + 1);
			for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
			{
				xml += "<pivotField showAll=\"0\" />";
			}

			xml += "</pivotFields>";
			xml += "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\" />";
			xml += "</pivotTableDefinition>";
			return xml;
		}

		private string CleanDisplayName(string name)
		{
			return Regex.Replace(name, @"[^\w\.-_]", "_");
		}
		#endregion
	}
}