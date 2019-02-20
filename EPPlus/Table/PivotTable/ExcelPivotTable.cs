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
using OfficeOpenXml.Internationalization;
using OfficeOpenXml.Table.PivotTable.DataCalculation;
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
			get
			{	return base.GetXmlNodeBool("@dataOnRows"); }
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
			get
			{
				return base.GetXmlNodeBool("@rowGrandTotals", true);
			}
			set
			{
				base.SetXmlNodeBool("@rowGrandTotals", value);
			}
		}

		/// <summary>
		/// Gets or sets whether the drill indicators expand collapse buttons should be printed.
		/// </summary>
		public bool PrintDrill
		{
			get
			{
				return base.GetXmlNodeBool("@printDrill");
			}
			set
			{
				base.SetXmlNodeBool("@printDrill", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to show error messages in cells.
		/// </summary>
		public bool ShowError
		{
			get
			{
				return base.GetXmlNodeBool("@showError");
			}
			set
			{
				base.SetXmlNodeBool("@showError", value);
			}
		}

		/// <summary>
		/// Gets or sets the string to be displayed in cells that contain errors.
		/// </summary>
		public string ErrorCaption
		{
			get
			{
				return base.GetXmlNodeString("@errorCaption");
			}
			set
			{
				base.SetXmlNodeString("@errorCaption", value);
			}
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
			get
			{
				return base.GetXmlNodeBool("@gridDropZones");
			}
			set
			{
				base.SetXmlNodeBool("@gridDropZones", value);
			}
		}

		/// <summary>
		/// Gets or sets the indentation increment for compact axis or can be used to set the Report Layout to Compact Form.
		/// </summary>
		public int Indent
		{
			get
			{
				return base.GetXmlNodeInt("@indent");
			}
			set
			{
				base.SetXmlNodeString("@indent", value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets whether data fields in the PivotTable should be displayed in outline form.
		/// </summary>
		public bool OutlineData
		{
			get
			{
				return base.GetXmlNodeBool("@outlineData");
			}
			set
			{
				base.SetXmlNodeBool("@outlineData", value);
			}
		}

		/// <summary>
		/// Gets or sets whether new fields should have their outline flag set to true.
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
		/// Gets or sets whether the fields of a PivotTable can have multiple filters set on them.
		/// </summary>
		public bool MultipleFieldFilters
		{
			get
			{
				return base.GetXmlNodeBool("@multipleFieldFilters", true);
			}
			set
			{
				base.SetXmlNodeBool("@multipleFieldFilters", value);
			}
		}

		public bool CustomListSort
		{
			get { return base.GetXmlNodeBool("@customListSort", true); }
		}

		/// <summary>
		/// Gets or sets whether new fields should have their compact flag set to true.
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
		/// Gets or sets whether the field next to the data field in the PivotTable should be displayed in the same column of the spreadsheet.
		/// </summary>
		public bool CompactData
		{
			get
			{
				return base.GetXmlNodeBool("@compactData");
			}
			set
			{
				base.SetXmlNodeBool("@compactData", value);
			}
		}

		/// <summary>
		/// Gets or sets the string to be displayed for grand totals.
		/// </summary>
		public string GrandTotalCaption
		{
			get
			{
				return base.GetXmlNodeString("@grandTotalCaption");
			}
			set
			{
				base.SetXmlNodeString("@grandTotalCaption", value);
			}
		}

		/// <summary>
		/// Gets or sets the string to be displayed in row header in compact mode.
		/// </summary>
		public string RowHeaderCaption
		{
			get
			{
				return base.GetXmlNodeString("@rowHeaderCaption");
			}
			set
			{
				base.SetXmlNodeString("@rowHeaderCaption", value);
			}
		}

		/// <summary>
		/// Gets or sets the string to be displayed in cells with no value.
		/// </summary>
		public string MissingCaption
		{
			get
			{
				return base.GetXmlNodeString("@missingCaption");
			}
			set
			{
				base.SetXmlNodeString("@missingCaption", value);
			}
		}

		/// <summary>
		/// Gets or sets the first row of the PivotTable header relative to the top left cell in the ref value.
		/// </summary>
		public int FirstHeaderRow
		{
			get
			{
				return base.GetXmlNodeInt(FirstHeaderRowPath);
			}
			set
			{
				base.SetXmlNodeString(FirstHeaderRowPath, value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the first column of the PivotTable data relative to the top left cell in the ref value.
		/// </summary>
		public int FirstDataRow
		{
			get
			{
				return base.GetXmlNodeInt(FirstDataRowPath);
			}
			set
			{
				base.SetXmlNodeString(FirstDataRowPath, value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the first column of the PivotTable data relative to the top left cell in the ref value.
		/// </summary>
		public int FirstDataCol
		{
			get
			{
				return base.GetXmlNodeInt(FirstDataColumnPath);
			}
			set
			{
				base.SetXmlNodeString(FirstDataColumnPath, value.ToString());
			}
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
		/// Gets the fields in the table .
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
		/// Gets or sets the pivot style name that is used for custom styles.
		/// </summary>
		public string StyleName
		{
			get
			{
				return base.GetXmlNodeString(StyleNamePath);
			}
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
			get
			{
				return myTableStyle;
			}
			set
			{
				myTableStyle = value;
				if (value != TableStyles.Custom)
					base.SetXmlNodeString(StyleNamePath, "PivotStyle" + value.ToString());
			}
		}

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
		/// Gets a list of headers for tabular row fields. These are basically column headers.
		/// </summary>
		internal List<PivotTableHeader> TabularHeaders { get; } = new List<PivotTableHeader>();

		/// <summary>
		/// Gets a value indicating whether there is more than one data field in the row fields.
		/// </summary>
		internal bool HasRowDataFields => this.RowFields.Any(c => c.Index == -2);

		/// <summary>
		/// Gets a value indicating whether there is more than one data field in the column fields.
		/// </summary>
		internal bool HasColumnDataFields => this.ColumnFields.Any(c => c.Index == -2);

		internal ExcelWorkbook Workbook { get; private set; }
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

			this.RowHeaders.Clear();
			this.ColumnHeaders.Clear();

			// Update the rowItems.
			this.Workbook.FormulaParser.Logger?.LogFunction($"{nameof(this.UpdateRowColumnItems)}: Rows");
			this.UpdateRowColumnItems(this.RowFields, this.RowItems, true);

			// Update the colItems.
			this.Workbook.FormulaParser.Logger?.LogFunction($"{nameof(this.UpdateRowColumnItems)}: Columns");
			this.UpdateRowColumnItems(this.ColumnFields, this.ColumnItems, false);

			// Update the pivot table data.
			this.Workbook.FormulaParser.Logger?.LogFunction(nameof(this.UpdateWorksheet));
			this.UpdateWorksheet(stringResources);

			// Remove the 'm' (missing) xml attribute from each pivot field item, if it exists, to prevent 
			// corrupting the workbook, since Excel automatically adds them.
			this.RemovePivotFieldItemMAttribute();

			// pivotSelections are causing corruptions when left. Deleting for meow.
			this.Worksheet.View.RemovePivotSelections();
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
			foreach (var dataField in this.DataFields)
			{
				if (dataField.ShowDataAs == ShowDataAs.Percent || dataField.ShowDataAs == ShowDataAs.PercentOfParentRow || dataField.ShowDataAs == ShowDataAs.PercentOfParentCol
					|| dataField.ShowDataAs == ShowDataAs.PercentOfParent || dataField.ShowDataAs == ShowDataAs.Difference || dataField.ShowDataAs == ShowDataAs.PercentDiff
					|| dataField.ShowDataAs == ShowDataAs.RunTotal || dataField.ShowDataAs == ShowDataAs.PercentOfRunningTotal || dataField.ShowDataAs == ShowDataAs.RankAscending
					|| dataField.ShowDataAs == ShowDataAs.RankDescending || dataField.ShowDataAs == ShowDataAs.Index)
				{
					unsupportedFeatures.Add($"Data field '{dataField.Name}' show data as setting '{dataField.ShowDataAs}'");
				}
			}
			foreach (var field in this.Fields)
			{
				if (!field.Compact)
					unsupportedFeatures.Add($"Field '{field.Name}' compact disabled");
				if (field.RepeatItemLabels)
					unsupportedFeatures.Add($"Field '{field.Name}' repeat item labels enabled");
				if (field.InsertBlankLine)
					unsupportedFeatures.Add($"Field '{field.Name}' insert blank line enabled");
				if (field.ShowAll)
					unsupportedFeatures.Add($"Field '{field.Name}' show items with no data enabled");
				if (field.InsertPageBreak)
					unsupportedFeatures.Add($"Field '{field.Name}' insert page break after each item enabled");
			}
			var filters = base.TopNode.SelectSingleNode("d:filters", base.NameSpaceManager);
			if (filters != null)
				unsupportedFeatures.Add("Filters enabled");
			if (this.MergeAndCenterCellsWithLabels)
				unsupportedFeatures.Add("Merge and center cells with labels enabled");
			if (this.PageOverThenDown)
				unsupportedFeatures.Add("Display fields in report filter area over then down enabled");
			if (this.PageWrap != 0)
				unsupportedFeatures.Add("Report filter fields per [row|column] > 0");
			if (!string.IsNullOrEmpty(this.ErrorCaption))
				unsupportedFeatures.Add("Error caption enabled");
			if (this.ShowError)
				unsupportedFeatures.Add("Show error enabled");
			if (!string.IsNullOrEmpty(this.MissingCaption))
				unsupportedFeatures.Add("Empty cell caption enabled");
			if (!this.PreserveFormatting)
				unsupportedFeatures.Add("Preserve formatting disabled");
			if (this.MultipleFieldFilters)
				unsupportedFeatures.Add("Multiple field filters enabled");
			if (!this.CustomListSort)
				unsupportedFeatures.Add("Use Custom Lists when sorting disabled");
			if (!this.ShowDataTips)
				unsupportedFeatures.Add("Show contextual tooltips disabled");
			if (!this.ShowHeaders)
				unsupportedFeatures.Add("Display field captions and filter dropdowns disabled");
			if (this.GridDropZones)
				unsupportedFeatures.Add("Grid drop zones enabled");
			if (!this.HideValuesRow)
				unsupportedFeatures.Add("Show values row enabled");
			if (this.FieldListSortAscending)
				unsupportedFeatures.Add("Field list sort ascending enabled");
			return unsupportedFeatures.Any();
		}
		#endregion

		#region Private Methods
		private List<Tuple<int, int>> BuildTabularRowItems(List<int> tabularFieldIndices, PivotItemTreeNode root, List<Tuple<int, int>> indices, List<Tuple<int, int>> lastChildIndices)
		{
			int repeatedItemsCount = 0;
			// Base case: If we're at a leaf node or a non-tabular form field node.
			if (root.Value != -1 && (!root.HasChildren || !root.IsTabularForm))
			{
				repeatedItemsCount = this.GetRepeatedItemsCount(indices, lastChildIndices);
				ExcelPivotTableField pivotField = null;
				string subtotalType = null;
				if (root.PivotFieldIndex != -2)
				{
					pivotField = this.Fields[root.PivotFieldIndex];
					subtotalType = this.GetRowFieldSubtotalType(pivotField);
				}

				bool isLeafNode = !root.HasChildren;
				bool isAboveDataField = !indices.Any(x => x.Item1 == -2);
				bool hasTabularField = tabularFieldIndices.Any(x => indices.Any(y => y.Item1 == x));
				bool isDataField = root.PivotFieldIndex == -2;
				this.RowHeaders.Add(new PivotTableHeader(indices, pivotField, root.DataFieldIndex, false, true, isLeafNode, isDataField, subtotalType, isAboveDataField, hasTabularField));
				this.RowItems.AddColumnItem(indices.ToList(), repeatedItemsCount, root.DataFieldIndex);
				lastChildIndices = indices.ToList();

				if (!root.HasChildren)
					return indices.ToList();
			}

			root.ExpandIfDataFieldNode(this.DataFields.Count);

			for (int i = 0; i < root.Children.Count; i++)
			{
				var child = root.Children[i];
				int pivotFieldItemIndex = child.PivotFieldIndex == -2 ? child.DataFieldIndex : child.PivotFieldItemIndex;
				// If the current tree node's pivot field has tabular form enabled, is not a leaf node and it's children are not datafields,
				// then create a tabular form header if it does not exist already.
				if (root.IsTabularForm && root.HasChildren && !child.IsDataField)
				{
					var pivotField = this.Fields[child.PivotFieldIndex];
					if (!this.TabularHeaders.Any(c => c.IsTabularHeader && c.PivotTableField == pivotField))
						this.TabularHeaders.Add(new PivotTableHeader(pivotField));
				}

				var childIndices = indices.ToList();
				childIndices.Add(new Tuple<int, int>(child.PivotFieldIndex, pivotFieldItemIndex));
				lastChildIndices = this.BuildTabularRowItems(tabularFieldIndices, child, childIndices, lastChildIndices);
			}

			this.CreateColumnAndTabularSubtotalNodes(root, repeatedItemsCount, indices, true);
			return lastChildIndices;
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

		private void CreateColumnAndTabularSubtotalNodes(PivotItemTreeNode node, int repeatedItemsCount, List<Tuple<int, int>> indices, bool tabularFieldEnabled = false)
		{
			var fieldCollection = tabularFieldEnabled ? this.RowFields : this.ColumnFields;
			
			// Create subtotal nodes if default subtotal is enabled and we are not at the root node.
			// Also, if node has a grandchild or the leaf node is not data field create a subtotal node.
			var defaultSubtotal = node.PivotFieldIndex == -2 ? false : this.Fields[node.PivotFieldIndex].DefaultSubtotal;
			if (defaultSubtotal && node.Value != -1 &&
				(node.Children.FirstOrDefault()?.HasChildren == true || fieldCollection.Last().Index != -2))
			{
				repeatedItemsCount = fieldCollection.ToList().FindIndex(x => x.Index == node.PivotFieldIndex);
				bool isAboveDataField = !indices.Any(x => x.Item1 == -2);
				var pivotField = this.Fields[node.PivotFieldIndex];

				bool createTabularSubtotalNode = false;
				if (tabularFieldEnabled)
				{
					// Create a subtotal node for any of the following cases:
					//		-The pivot field has tabular form enabled.
					//		-The pivot field has tabular form disabled and subtotal bottom is enabled.
					//		-There are multiple datafields and the current node is above a datafield and the pivot field has tabular form enabled.
					createTabularSubtotalNode = !pivotField.Outline || (pivotField.Outline && !pivotField.SubtotalTop) || (this.HasRowDataFields && pivotField.Outline && isAboveDataField);
				}

				if (!tabularFieldEnabled || createTabularSubtotalNode)
				{
					bool isLastNonDataField = fieldCollection.Skip(repeatedItemsCount).All(x => x.Index == -2);
					bool hasDataFields = tabularFieldEnabled ? this.HasRowDataFields : this.HasColumnDataFields;
					bool isRowItem = tabularFieldEnabled;
					var headerPivotField = tabularFieldEnabled ? pivotField : null;
					var functionNames = pivotField.GetEnabledSubtotalTypes();
					foreach (var function in functionNames)
					{
						// If the node is above a data field node and there are multiple data fields, then create a subtotal node for each data field. 
						if (this.DataFields.Count > 0 && isAboveDataField && !isLastNonDataField && hasDataFields)
							this.CreateTotalNodes(function, isRowItem, indices, headerPivotField, repeatedItemsCount, true, hasDataFields);
						// Otherwise, if the node is not the last non-data field node and is below a data field node, then only create one subtotal node.
						else if (!isLastNonDataField && (!isAboveDataField || !hasDataFields))
							this.CreateTotalNodes(function, isRowItem, indices, headerPivotField, repeatedItemsCount, false, hasDataFields, node.DataFieldIndex);
					}
				}
			}
		}

		private void BuildRowItems(PivotItemTreeNode root, List<Tuple<int, int>> indices)
		{
			if (!root.HasChildren)
				return;

			root.ExpandIfDataFieldNode(this.DataFields.Count);

			foreach (var child in root.Children)
			{
				var rowDepth = indices.Count;
				ExcelPivotTableField pivotField = null;

				int pivotFieldItemIndex = 0;
				bool isTabularFormat = false;
				if (child.PivotFieldIndex != -2)
				{
					// Child is not a data field.
					pivotField = this.Fields[child.PivotFieldIndex];
					pivotFieldItemIndex = child.PivotFieldItemIndex;
					isTabularFormat = pivotField.Outline == false;
				}
				else
					pivotFieldItemIndex = child.DataFieldIndex;

				var childIndices = indices.ToList();
				childIndices.Add(new Tuple<int, int>(child.PivotFieldIndex, pivotFieldItemIndex));
				this.RowItems.Add(rowDepth, pivotFieldItemIndex, null, child.DataFieldIndex);
				bool isLeafNode = !child.HasChildren;
				bool isAboveDataField = !childIndices.Any(x => x.Item1 == -2);
				string subtotalType = this.GetRowFieldSubtotalType(pivotField);

				this.RowHeaders.Add(new PivotTableHeader(childIndices, pivotField, child.DataFieldIndex, false, true, isLeafNode, false, subtotalType, isAboveDataField));
				this.BuildRowItems(child, childIndices);
				// Create subtotal nodes if default subtotal is enabled.
				this.CreateRowSubtotalNodes(child, childIndices, pivotField, rowDepth, isAboveDataField);
			}
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
					if (!type.IsEquivalentTo("default"))
						return type;
				}
				else if (subtotalTypes.Count > 1)
					return "none";
			}
			return null;
		}

		private void CreateRowSubtotalNodes(PivotItemTreeNode child, List<Tuple<int, int>> childIndices, ExcelPivotTableField pivotField, int rowDepth, bool isAboveDataField)
		{
			if (pivotField == null)
				return;
			if (pivotField.DefaultSubtotal)
			{
				var subtotalTypes = pivotField.GetEnabledSubtotalTypes();
				foreach (var subtotalType in subtotalTypes)
				{
					if (this.HasRowDataFields && isAboveDataField
						&& (child.Children.FirstOrDefault()?.HasChildren == true || this.RowFields.Last().Index != -2))
					{
						// If above a datafield, subtotals are always shown if defaultSubtotal is enabled.
						this.CreateTotalNodes(subtotalType, true, childIndices, pivotField, rowDepth, true, this.HasRowDataFields, child.DataFieldIndex);
					}
					else if (!child.SubtotalTop || subtotalTypes.Count > 1)
					{
						// If this child is only followed by datafields, do not write out a subtotal node.
						// In other words, treat this child as the leaf node.
						if (!isAboveDataField)
							this.CreateTotalNodes(subtotalType, true, childIndices, pivotField, rowDepth, false, this.HasRowDataFields, child.DataFieldIndex);
						else if (child.Children.Any(c => !c.IsDataField || c.Children.Any()))
							this.CreateTotalNodes(subtotalType, true, childIndices, pivotField, rowDepth, true, this.HasRowDataFields, child.DataFieldIndex);
					}
				}
			}
		}

		private List<Tuple<int, int>> BuildColumnItems(PivotItemTreeNode node, List<Tuple<int, int>> indices, List<Tuple<int, int>> lastChildIndices)
		{
			int repeatedItemsCount = 0;
			// Base case (leaf node).
			if (!node.HasChildren)
			{
				repeatedItemsCount = this.GetRepeatedItemsCount(indices, lastChildIndices);
				var header = new PivotTableHeader(indices, null, node.DataFieldIndex, false, false, true, node.IsDataField);
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

			this.CreateColumnAndTabularSubtotalNodes(node, repeatedItemsCount, indices);
			return lastChildIndices;
		}

		private PivotItemTreeNode BuildRowColTree(ExcelPivotTableRowColumnFieldCollection rowColFields, Dictionary<int, List<int>> cacheRecordPageFieldIndices)
		{
			// Build a tree using the cache records. Each node in the tree is a cache record that is 
			// identified by the row or column field indices.
			var rootNode = new PivotItemTreeNode(-1);
			for (int i = 0; i < this.CacheDefinition.CacheRecords.Count; i++)
			{
				var cacheRecord = this.CacheDefinition.CacheRecords[i];
				var currentNode = rootNode;

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
						
						if (cacheFields[originalIndex].IsGroupField)
							currentNode = this.CreateTreeNodeWithGrouping(sharedItemValue, originalIndex, currentNode, recordItemValue, groupBy, cacheFields, cacheRecordPageFieldIndices, cacheRecord);
						else
							currentNode = this.CreateTreeNode(false, currentNode, this.Fields[rowColFieldIndex], recordItemValue, rowColFieldIndex, cacheRecordPageFieldIndices, cacheRecord, sharedItemValue.Value);

						// This cache record does not contain the page field indices, so continue to the next record.
						if (currentNode == null)
							break;
					}
					if (!currentNode.CacheRecordIndices.Contains(i))
						currentNode.CacheRecordIndices.Add(i);
				}
			}
			return rootNode;
		}

		private PivotItemTreeNode CreateTreeNodeWithGrouping(CacheItem sharedItemValue, int groupingIndex, PivotItemTreeNode currentNode, int recordItemValue, PivotFieldDateGrouping? groupBy,
			IReadOnlyList<CacheFieldNode> cacheFields, Dictionary<int, List<int>> cacheRecordPageFieldIndices, CacheRecordNode cacheRecord)
		{
			var pivotField = this.Fields[groupingIndex];
			// A sharedItem value of type DateTime indicates the current pivot field is part of a date grouping. Otherwise, create a new node if necessary.
			if (sharedItemValue.Type == PivotCacheRecordType.d)
			{
				// Handles field date groupings.
				var searchValue = this.GetItemValueByGroupingType(sharedItemValue.Value, groupBy);
				currentNode =  this.CreateTreeNode(true, currentNode, pivotField, recordItemValue, groupingIndex, cacheRecordPageFieldIndices, cacheRecord, searchValue, cacheFields[groupingIndex]);
			}
			else if (cacheFields[groupingIndex].FieldGroup.DiscreteGroupingProperties != null)
			{
				// Handles custom field groupings.
				var groupingFieldGroup = cacheFields[groupingIndex].FieldGroup;
				int discretePrValue = int.Parse(groupingFieldGroup.DiscreteGroupingProperties[recordItemValue].Value);
				var groupingSearchValue = groupingFieldGroup.GroupItems[discretePrValue].Value;
				currentNode = this.CreateTreeNode(false, currentNode, pivotField, discretePrValue, groupingIndex, cacheRecordPageFieldIndices, cacheRecord, groupingSearchValue);
			}
			return currentNode;
		}

		private string GetItemValueByGroupingType(string sharedItemValue, PivotFieldDateGrouping? groupBy)
		{
			var dateSplit = sharedItemValue.Split('-');
			var dateTime = new DateTime(int.Parse(dateSplit[0]), int.Parse(dateSplit[1]), int.Parse(dateSplit[2].Substring(0, 2)));
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
			Dictionary<int, List<int>> cacheRecordPageFieldIndices, CacheRecordNode cacheRecord, string searchValue, CacheFieldNode cacheField = null)
		{
			// If an identical child already exists, continue. Otherwise, create a new child.
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
					pivotFieldItemIndex = cacheFieldGroupItems.ToList().FindIndex(x => x.Value.IsEquivalentTo(searchValue));
				}
				else
					pivotFieldItemIndex = pivotField.Items.ToList().FindIndex(c => c.X == recordItemValue);
				bool isTabularForm = pivotField?.Outline == false;
				return currentNode.AddChild(recordItemValue, pivotFieldIndex, pivotFieldItemIndex, searchValue, isTabularForm);
			}
		}

		private IOrderedEnumerable<CacheItem> SortField(eSortType sortOrder, ExcelPivotTableField pivotField)
		{
			var fieldItems = pivotField.Items;
			var sharedItems = this.CacheDefinition.CacheFields[pivotField.Index].SharedItems;
			var isNumericValues = sharedItems.All(x => double.TryParse(x.Value, out _));

			IOrderedEnumerable<CacheItem> sortedList = null;
			if (sortOrder == eSortType.Descending && pivotField.AutoSortScopeReferences.Count == 0)
			{
				if (isNumericValues)
					sortedList = sharedItems.OrderByDescending(x => double.Parse(x.Value));
				else
					sortedList = sharedItems.ToList().OrderByDescending(x => x.Value);
				if (pivotField.Name.IsEquivalentTo("Month"))
					sortedList = sharedItems.ToList().OrderByDescending(m => DateTime.ParseExact(m.Value, "MMMMM", Thread.CurrentThread.CurrentCulture));
			}
			else
			{
				if (isNumericValues)
					sortedList = sharedItems.OrderBy(x => double.Parse(x.Value));
				else
					sortedList = sharedItems.ToList().OrderBy(x => x.Value);
				if (pivotField.Name.IsEquivalentTo("Month"))
					sortedList = sharedItems.ToList().OrderBy(m => DateTime.ParseExact(m.Value, "MMMMM", Thread.CurrentThread.CurrentCulture));
			}

			return sortedList;
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

		private void UpdateRowColumnItems(ExcelPivotTableRowColumnFieldCollection rowColFieldCollection, ItemsCollection collection, bool isRowItems)
		{
			// Update the rowItems or colItems.
			if (rowColFieldCollection.Any())
			{
				collection.Clear();
				var pageFieldIndices = this.GetPageFieldIndices();
				var root = this.BuildRowColTree(rowColFieldCollection, pageFieldIndices);
				root.SortChildren(this);
				if (isRowItems)
				{
					bool tabularForm = this.Fields.Any(x => x.Outline == false);
					if (!tabularForm)
						this.BuildRowItems(root, new List<Tuple<int, int>>());
					else
					{
						var tabularFieldIndices = new List<int>();
						for (int i = 0; i < this.Fields.Count; i++)
						{
							if (!this.Fields[i].Outline)
								tabularFieldIndices.Add(i);
						}
						this.BuildTabularRowItems(tabularFieldIndices, root, new List<Tuple<int, int>>(), new List<Tuple<int, int>>());
					}
				}
				else
					this.BuildColumnItems(root, new List<Tuple<int, int>>(), new List<Tuple<int, int>>());

				// Create grand total items if necessary.
				bool grandTotals = isRowItems ? this.RowGrandTotals : this.ColumnGrandTotals;
				if (grandTotals && isRowItems && !(this.RowFields.Count == 1 && this.RowFields.First().Index == -2))
					this.CreateTotalNodes("grand", true, null, null, 0, false, this.HasRowDataFields);
				else if (grandTotals && !isRowItems && !(this.ColumnFields.Count == 1 && this.ColumnFields.First().Index == -2))
					this.CreateTotalNodes("grand", false, null, null, 0, false, this.HasColumnDataFields);
			}
			else
			{
				string xmlTag = isRowItems ? "d:rowFields" : "d:colFields";
				// If there are no row/column fields, then remove tag or else it will corrupt the workbook.
				this.TopNode.RemoveChild(this.TopNode.SelectSingleNode(xmlTag, this.NameSpaceManager));
				var headerCollection = isRowItems ? this.RowHeaders : this.ColumnHeaders;
				var header = new PivotTableHeader(null, null, 0, false, false, false, false);
				header.IsPlaceHolder = true;
				headerCollection.Add(header);
			}
		}

		private void CreateTotalNodes(string totalType, bool isRowItem, List<Tuple<int, int>> indices, ExcelPivotTableField pivotField, 
			int repeatedItemsCount, bool multipleSubtotalDataFields, bool hasDataFields, int dataFieldIndex = 0)
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
				var header = new PivotTableHeader(indices, pivotField, dataFieldIndex, grandTotal, isRowItem, false, false, totalType, aboveDataField);
				itemsCollection.AddSumNode(totalType, repeatedItemsCount, xMember, dataFieldIndex);
				headerCollection.Add(header);
			}
			else
			{
				// Create the xml node and row/column header.
				for (int i = 0; i < index; i++)
				{
					var header = new PivotTableHeader(indices, pivotField, i, grandTotal, isRowItem, false, false, totalType, aboveDataField);
					itemsCollection.AddSumNode(totalType, repeatedItemsCount, xMember, i);
					headerCollection.Add(header);
				}
			}
		}

		private void UpdateWorksheet(StringResources stringResources)
		{
			// Update the row and column header values in the worksheet.
			bool tabularTable = this.Fields.Any(x => x.Outline == false);
			this.UpdateRowColumnHeaders(stringResources, tabularTable);
			// Update the pivot table's address.
			int endRow = this.Address.Start.Row + this.FirstDataRow + this.RowHeaders.Count - 1;
			// If there are no data fields, then don't find the offset to obtain the first data column.
			int endColumn = this.Address.Start.Column;
			if (this.DataFields.Any())
				endColumn += this.FirstDataCol + this.ColumnHeaders.Count - 1;
			else if (tabularTable)
				endColumn += this.TabularHeaders.Count;
			this.Address = new ExcelAddress(this.Worksheet.Name, this.Address.Start.Row, this.Address.Start.Column, endRow, endColumn);
			if (this.DataFields.Any())
			{
				var dataManager = new PivotTableDataManager(this);
				dataManager.UpdateWorksheet();
			}
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
		
		private void UpdateRowColumnHeaders(StringResources stringResources, bool tabularTable)
		{
			// Clear out the pivot table in the worksheet.
			this.ClearTable();

			// Update the row headers in the worksheet.
			if (!tabularTable)
				this.WriteRowHeaders(stringResources);
			else
			{
				int dataFieldColumn = this.WriteTabularRowHeaders(stringResources);
				this.WriteTabularHeadersInColumns(dataFieldColumn);
			}

			// Update the column headers in the worksheet.
			this.WriteColumnHeaders(stringResources);
		}

		private int WriteTabularRowHeaders(StringResources stringResources)
		{
			int dataFieldColumn = 0;
			int row = this.Address.Start.Row + this.FirstDataRow;
			int startColumn = this.Address.Start.Column;
			int firstColumnWrittenTo = row;
			int lastColumnWrittenTo = 0;
			RowColumnItem previousNonTotalItem = null;
			for (int i = 0; i < this.RowItems.Count; i++)
			{
				int currentColumn = startColumn;
				var item = this.RowItems[i];
				var header = this.RowHeaders[i];
				for (int j = 0; j < item.Count; j++)
				{
					// Write subtotal headers.
					var totalHeader = this.SetTotalCaptionCellValue(this.RowFields, item, header, stringResources);
					if (!string.IsNullOrEmpty(totalHeader))
					{
						lastColumnWrittenTo = this.GetTabularSubtotalHeaderColumn(item, previousNonTotalItem, startColumn, currentColumn, lastColumnWrittenTo, firstColumnWrittenTo, i, j);
						this.Worksheet.Cells[row, lastColumnWrittenTo].Value = totalHeader;
						continue;
					}

					// Get sharedItem header value.
					var itemIndex = item.RepeatedItemsCount == 0 ? j : j + item.RepeatedItemsCount;
					string headerValue = this.GetSharedItemValue(this.RowFields, item, itemIndex, j);
					int pivotFieldIndex = this.RowFields[itemIndex].Index;

					// Write row headers.
					if (item.Count == 1 && header.IsLeafNode)
					{
						// If the item is a leaf node with pivot fields that have tabular form disabled, then print out the 
						// header values at the last written in column.
						this.Worksheet.Cells[row, lastColumnWrittenTo].Value = headerValue;
					}
					else
					{
						// If the item's repeatedItemsCount is zero, then it does not have a parent, so always set the column to the start column.
						// Otherwise, the item is an inner item, so get the correct column to write to.
						if (item.RepeatedItemsCount != 0)
							currentColumn = this.GetInnerItemColumn(item, previousNonTotalItem, currentColumn, startColumn, ref firstColumnWrittenTo, lastColumnWrittenTo, i, j);
						this.Worksheet.Cells[row, currentColumn].Value = headerValue;
						if (pivotFieldIndex == -2)
							dataFieldColumn = currentColumn;
						lastColumnWrittenTo = currentColumn;
						previousNonTotalItem = item;
					}

					// If the row item contains pivot fields in tabular form, then stay on the same row and increment the column.
					if (item.Count > 1)
						currentColumn++;
				}
				row++;
			}
			return dataFieldColumn;
		}

		private int GetInnerItemColumn(RowColumnItem currentItem, RowColumnItem previousNonTotalItem, int currentColumn, int startColumn, ref int firstColumnWrittenTo, int lastColumnWrittenTo, int i, int j)
		{
			var currentRowFieldIndex = currentItem.RepeatedItemsCount + j;
			var isAboveDataField = this.RowFields.Skip(currentRowFieldIndex).Any(x => x.Index == -2);
			// If we are at an inner node, then get the column based on whether or not we are the the first child.
			if (previousNonTotalItem.RepeatedItemsCount == 0)
			{
				// First child.
				int previousRepeatedItemsCount = previousNonTotalItem.RepeatedItemsCount;
				int previousIndex = previousRepeatedItemsCount + previousNonTotalItem.Count;
				int previousRowFieldIndex = this.RowFields[previousIndex - 1].Index;
				// If the last written in header is a datafield value, then check whether or not this item is above a datafield
				// and get the appropriate column value.
				if (previousRowFieldIndex == -2)
				{
					if (isAboveDataField)
						currentColumn = j == 0 ? startColumn + currentItem.RepeatedItemsCount : currentColumn;
					else
						currentColumn = j == 0 ? lastColumnWrittenTo : currentColumn;
				}
				else
				{
					var previousPivotField = this.Fields[previousRowFieldIndex];
					// If the previous item's pivot field has tabular form disabled and it is not the last field in the row field
					// collection, then write to the last written column. Otherwise, calculate the start column.
					if (previousPivotField.Outline && previousNonTotalItem.Count < this.RowFields.Count)
						currentColumn = j == 0 ? lastColumnWrittenTo : currentColumn;
					else
						currentColumn = j == 0 ? startColumn + currentItem.RepeatedItemsCount : currentColumn;
				}
				firstColumnWrittenTo = j == 0 ? currentColumn : firstColumnWrittenTo;
			}
			else
			{
				// If there are multiple datafields, then calculate the correct column to write to based on the position of the item
				// relative to the datafields. Otherwise, use the firstColumnWrittenTo value as the starting column to write to.
				if (this.HasRowDataFields)
				{
					var rowFieldIndex = this.RowFields[currentRowFieldIndex].Index;
					if (rowFieldIndex == -2 || !isAboveDataField)
						currentColumn = j == 0 ? lastColumnWrittenTo : currentColumn;
					else
					{
						var isLastNonDataField = this.RowFields.Skip(currentRowFieldIndex + 1).All(x => x.Index == -2);
						// If the item is the last non-datafield row field and the previous item was not a data field or the previous item was not a
						// subtotal item, then write to the last column written to. Otherwise, calculate the column.
						if ((isLastNonDataField && !this.RowHeaders[i - 1].IsDataField) || !string.IsNullOrEmpty(this.RowItems[i - 1].ItemType))
							currentColumn = j == 0 ? lastColumnWrittenTo : currentColumn;
						else
							currentColumn = j == 0 ? startColumn + currentItem.RepeatedItemsCount : currentColumn;
					}
				}
				else
					currentColumn = j == 0 ? firstColumnWrittenTo : currentColumn;
			}
			return currentColumn;
		}

		private int GetTabularSubtotalHeaderColumn(RowColumnItem item, RowColumnItem previousNonTotalItem, int startColumn, int currentColumn, int lastColumnWrittenTo,
			int firstColumnWrittenTo, int i, int j)
		{
			int returnColumn = 0;
			if (item.RepeatedItemsCount == 0)
				returnColumn = currentColumn;
			else
			{
				int previousRepeatedItemsCount = previousNonTotalItem.RepeatedItemsCount;
				int rowFieldIndex = previousRepeatedItemsCount == 0 ? previousRepeatedItemsCount : this.RowFields[previousRepeatedItemsCount - 1].Index;
				ExcelPivotTableField field = null;
				if (rowFieldIndex != -2)
					field = this.Fields[rowFieldIndex];

				if (field != null)
				{
					// If the current pivot field has tabular form disabled and we are not a parent item, then get the column to write to using the previous
					// non-total item and header. Otherwise, use the row field collection to get the column.
					if (field.Outline && previousRepeatedItemsCount != 0)
					{
						// If the current item is not a datafield item and the pivot field has tabular form disabled and we're not the first child of a parent.
						// Case 1: If the last non-total header is a leaf node that contains a tabular field and the subtotal's repeatedItemsCount is greater than the 
						//		   last non-default, calculate the correct column to write to.
						// Case 2: If the subtotal's repeatedItemsCount is greater than the last non-default/non-leaf node, print the header in the last column written to.
						// Case 3: Otherwise, always print the header to the first column that was written to by the last non-default/non-leaf node.
						// Note: When a subtotal's repeated item count is greater than the last non-default node, that means it shares values with the last non-default node.
						// AKA, this subtotal node is for a child of the last non-default node.
						if (this.RowHeaders[i - 1].IsLeafNode && previousRepeatedItemsCount < item.RepeatedItemsCount)
							returnColumn = startColumn + item.RepeatedItemsCount - 1;
						else if (previousRepeatedItemsCount < item.RepeatedItemsCount)
							returnColumn = lastColumnWrittenTo;
						else
							returnColumn = firstColumnWrittenTo;
					}
					else
					{
						int currentRowFieldIndex = item.RepeatedItemsCount + j;
						var isAboveDataField = this.RowFields.Skip(currentRowFieldIndex).Any(x => x.Index == -2);
						if (this.HasRowDataFields && !isAboveDataField)
							returnColumn = startColumn + item.RepeatedItemsCount - 1;
						else
						{
							int previousFieldIndex = this.RowFields[item.RepeatedItemsCount - 1].Index;
							var previousField = this.Fields[previousFieldIndex];
							if (previousField.Outline)
								returnColumn = startColumn + item.RepeatedItemsCount - 1;
							else
								returnColumn = startColumn + item.RepeatedItemsCount;
						}
					}
				}
				else
				{
					// If datafields are the first row field, then for every non-leaf field under it, calculate the column to print the total header to.
					// Otherwise, always print it to the firstColumnWrittenTo value.
					if (previousRepeatedItemsCount < item.RepeatedItemsCount)
						returnColumn = startColumn + item.RepeatedItemsCount - 1;
					else
						returnColumn = firstColumnWrittenTo;
				}
			}
			return returnColumn;
		}

		private void WriteRowHeaders(StringResources stringResources)
		{
			int dataRow = this.Address.Start.Row + this.FirstDataRow;
			if (this.RowFields.Any())
			{
				for (int i = 0; i < this.RowItems.Count; i++)
				{
					string itemType = this.SetTotalCaptionCellValue(this.RowFields, this.RowItems[i], this.RowHeaders[i], stringResources);
					if (!string.IsNullOrEmpty(itemType))
					{
						this.Worksheet.Cells[dataRow++, this.Address.Start.Column].Value = itemType;
						continue;
					}
					var sharedItem = this.GetSharedItemValue(this.RowFields, this.RowItems[i], this.RowItems[i].RepeatedItemsCount, 0);
					this.Worksheet.Cells[dataRow++, this.Address.Start.Column].Value = sharedItem;
				}
			}
			// If there are no row headers and only one data field, print the name of the data field for the row.
			else if (this.DataFields.Count == 1)
				this.Worksheet.Cells[dataRow++, this.Address.Start.Column].Value = this.DataFields.First().Name;
		}

		private void WriteColumnHeaders(StringResources stringResources)
		{
			int startRow = this.Address.Start.Row + this.FirstHeaderRow;
			int column = this.Address.Start.Column + this.FirstDataCol;
			if (this.ColumnFields.Any())
			{
				for (int i = 0; i < this.ColumnItems.Count; i++)
				{
					int startHeaderRow = startRow;
					string itemType = this.SetTotalCaptionCellValue(this.ColumnFields, this.ColumnItems[i], this.ColumnHeaders[i], stringResources);
					if (!string.IsNullOrEmpty(itemType))
					{
						int totalCaptionRow = startHeaderRow + this.ColumnItems[i].RepeatedItemsCount;
						this.Worksheet.Cells[totalCaptionRow, column++].Value = itemType;
						continue;
					}

					for (int j = 0; j < this.ColumnItems[i].Count; j++)
					{
						var columnFieldIndex = this.ColumnItems[i].RepeatedItemsCount == 0 ? j : j + this.ColumnItems[i].RepeatedItemsCount;
						var sharedItem = this.GetSharedItemValue(this.ColumnFields, this.ColumnItems[i], columnFieldIndex, j);
						var cellRow = this.ColumnItems[i].RepeatedItemsCount == 0 ? startHeaderRow : startHeaderRow + this.ColumnItems[i].RepeatedItemsCount;
						this.Worksheet.Cells[cellRow, column].Value = sharedItem;
						startHeaderRow++;
					}
					column++;
				}
			}
			// If there are no column headers and only one data field, print the name of the data field for the column.
			else if (this.DataFields.Count == 1)
				this.Worksheet.Cells[this.Address.Start.Row, column].Value = this.DataFields.First().Name;
		}

		private void WriteTabularHeadersInColumns(int dataFieldColumn)
		{
			bool hasValuesLabel = false;
			// Write "Values" string if necessary.
			if (dataFieldColumn > 0 && this.RowFields.First().Index != -2)
			{
				int dataFieldCollectionIndex = this.RowFields.ToList().FindIndex(x => x.Index == -2);
				int precedingRowFieldIndex = this.RowFields[dataFieldCollectionIndex - 1].Index;
				// If the parent field of a datafield has tabular form disabled, then write out to the worksheet at this column.
				if (!this.Fields[precedingRowFieldIndex].Outline)
				{
					this.Worksheet.Cells[this.Address.Start.Row + this.FirstDataRow - 1, dataFieldColumn].Value = "Values";
					hasValuesLabel = true;
				}
			}
			// Update tabular form row headers.
			if (this.TabularHeaders.Any())
			{
				int row = this.Address.Start.Row + this.FirstDataRow - 1;
				int column = this.Address.Start.Column + 1;
				int headerCounter = 0;
				while (headerCounter < this.TabularHeaders.Count)
				{
					if (column == dataFieldColumn && hasValuesLabel)
						column++;
					else
					{
						var value = this.TabularHeaders[headerCounter].PivotTableField.Name;
						this.Worksheet.Cells[row, column++].Value = value;
						headerCounter++;
					}
				}
			}
		}

		private string SetTotalCaptionCellValue(ExcelPivotTableRowColumnFieldCollection field, RowColumnItem item, PivotTableHeader header, StringResources stringResources)
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
					var itemName = this.GetSharedItemValue(field, item, item.RepeatedItemsCount, 0);
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

		private string GetSharedItemValue(ExcelPivotTableRowColumnFieldCollection field, RowColumnItem item, int repeatedItemsCount, int xMemberIndex)
		{
			var pivotFieldIndex = field[repeatedItemsCount].Index;
			// A field that has an 'x' attribute equal to -2 is a special row/column field that indicates the
			// pivot table has more than one data field. Excel uses this to display the headings for the data 
			// values and how to group them in relation to other rows/columns. 
			// If a special field alrady exists in that collection, then another one will not be generated.
			if (pivotFieldIndex == -2)
				return this.DataFields[item.DataFieldIndex].Name;
			var pivotField = this.Fields[pivotFieldIndex];
			var cacheItemIndex = pivotField.Items[item[xMemberIndex]].X;
			// If the pivot field is a part of a grouping, use the groupItems collection. Otherwise, use the sharedItems collection.
			if (this.CacheDefinition.CacheFields[pivotFieldIndex].IsGroupField)
				return this.CacheDefinition.CacheFields[pivotFieldIndex].FieldGroup.GroupItems[cacheItemIndex].Value;
			else
				return this.CacheDefinition.CacheFields[pivotFieldIndex].SharedItems[cacheItemIndex].Value;
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