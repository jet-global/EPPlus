﻿/*******************************************************************************
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
using System.Xml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.Utilities;
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
		private ExcelPivotTableRowColumnFieldCollection myPageFields;
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
			get
			{
				return base.GetXmlNodeString(NamePath);
			}
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
			private set
			{
				myCacheDefinition = value;
			}
		}
		
		/// <summary>
		/// Gets or sets the worksheet where the pivot table is located.
		/// </summary>
		public ExcelWorksheet Worksheet
		{
			get
			{
				return this.Workbook.Worksheets[this.Address.WorkSheet];
			}
		}
		
		/// <summary>
		/// Gets or sets the location of the pivot table.
		/// </summary>
		public ExcelAddress Address
		{
			get
			{
				return myAddress;
			}
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
			{
				return base.GetXmlNodeBool("@dataOnRows");
			}
			set
			{
				base.SetXmlNodeBool("@dataOnRows", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat number format properties.
		/// </summary>
		public bool ApplyNumberFormats
		{
			get
			{
				return base.GetXmlNodeBool("@applyNumberFormats");
			}
			set
			{
				base.SetXmlNodeBool("@applyNumberFormats", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat border properties.
		/// </summary>
		public bool ApplyBorderFormats
		{
			get
			{
				return base.GetXmlNodeBool("@applyBorderFormats");
			}
			set
			{
				base.SetXmlNodeBool("@applyBorderFormats", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat font properties.
		/// </summary>
		public bool ApplyFontFormats
		{
			get
			{
				return base.GetXmlNodeBool("@applyFontFormats");
			}
			set
			{
				base.SetXmlNodeBool("@applyFontFormats", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat pattern properties.
		/// </summary>
		public bool ApplyPatternFormats
		{
			get
			{
				return base.GetXmlNodeBool("@applyPatternFormats");
			}
			set
			{
				base.SetXmlNodeBool("@applyPatternFormats", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to apply the legacy table autoformat width/height properties.
		/// </summary>
		public bool ApplyWidthHeightFormats
		{
			get
			{
				return base.GetXmlNodeBool("@applyWidthHeightFormats");
			}
			set
			{
				base.SetXmlNodeBool("@applyWidthHeightFormats", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to show member property information.
		/// </summary>
		public bool ShowMemberPropertyTips
		{
			get
			{
				return base.GetXmlNodeBool("@showMemberPropertyTips");
			}
			set
			{
				base.SetXmlNodeBool("@showMemberPropertyTips", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to show the drill indicators.
		/// </summary>
		public bool ShowCalcMember
		{
			get
			{
				return base.GetXmlNodeBool("@showCalcMbrs");
			}
			set
			{
				base.SetXmlNodeBool("@showCalcMbrs", value);
			}
		}
		
		/// <summary>
		/// Gets or sets if the user can enable drill down on a PivotItem or aggregate value.
		/// </summary>
		public bool EnableDrill
		{
			get
			{
				return base.GetXmlNodeBool("@enableDrill", true);
			}
			set
			{
				base.SetXmlNodeBool("@enableDrill", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to show the drill down buttons.
		/// </summary>
		public bool ShowDrill
		{
			get
			{
				return base.GetXmlNodeBool("@showDrill", true);
			}
			set
			{
				base.SetXmlNodeBool("@showDrill", value);
			}
		}

		/// <summary>
		/// Gets or sets whether the tooltips should be displayed for PivotTable data cells.
		/// </summary>
		public bool ShowDataTips
		{
			get
			{
				return base.GetXmlNodeBool("@showDataTips", true);
			}
			set
			{
				base.SetXmlNodeBool("@showDataTips", value, true);
			}
		}

		/// <summary>
		/// Gets or sets whether the row and column titles from the PivotTable should be printed.
		/// </summary>
		public bool FieldPrintTitles
		{
			get
			{
				return base.GetXmlNodeBool("@fieldPrintTitles");
			}
			set
			{
				base.SetXmlNodeBool("@fieldPrintTitles", value);
			}
		}

		/// <summary>
		/// Gets or sets whether the row and column titles from the PivotTable should be printed.
		/// </summary>
		public bool ItemPrintTitles
		{
			get
			{
				return base.GetXmlNodeBool("@itemPrintTitles");
			}
			set
			{
				base.SetXmlNodeBool("@itemPrintTitles", value);
			}
		}

		/// <summary>
		/// Gets or sets whether the grand totals should be displayed for the PivotTable columns.
		/// </summary>
		/// <remarks>A blank value in XML indicates true.</remarks>
		public bool ColumnGrandTotals
		{
			get
			{
				return base.GetXmlNodeBool("@colGrandTotals", true);
			}
			set
			{
				base.SetXmlNodeBool("@colGrandTotals", value);
			}
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
			get
			{
				return base.GetXmlNodeString("@dataCaption");
			}
			set
			{
				base.SetXmlNodeString("@dataCaption", value);
			}
		}

		/// <summary>
		/// Gets or sets whether to show field headers.
		/// </summary>
		public bool ShowHeaders
		{
			get
			{
				return base.GetXmlNodeBool("@showHeaders");
			}
			set
			{
				base.SetXmlNodeBool("@showHeaders", value);
			}
		}
		
		/// <summary>
		/// Gets or sets the number of page fields to display before starting another row or column.
		/// </summary>
		public int PageWrap
		{
			get
			{
				return base.GetXmlNodeInt("@pageWrap");
			}
			set
			{
				if (value < 0)
					throw new Exception("Value can't be negative");
				base.SetXmlNodeString("@pageWrap", value.ToString());
			}
		}
		
		/// <summary>
		/// Gets or sets whether the legacy auto formatting has been applied to the PivotTable view.
		/// </summary>
		public bool UseAutoFormatting
		{
			get
			{
				return base.GetXmlNodeBool("@useAutoFormatting");
			}
			set
			{
				base.SetXmlNodeBool("@useAutoFormatting", value);
			}
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
				return base.GetXmlNodeBool("@multipleFieldFilters");
			}
			set
			{
				base.SetXmlNodeBool("@multipleFieldFilters", value);
			}
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
		/// Gets the report filter fields.
		/// </summary>
		public ExcelPivotTableRowColumnFieldCollection PageFields
		{
			get
			{
				if (myPageFields == null)
				{
					var pageFieldsNode = this.TopNode.SelectSingleNode("d:pageFields", this.NameSpaceManager);
					myPageFields = new ExcelPivotTableRowColumnFieldCollection(this.NameSpaceManager, pageFieldsNode, this, PivotTableItemType.Page);
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
			get
			{
				return base.GetXmlNodeInt("@cacheId");
			}
			set
			{
				base.SetXmlNodeString("@cacheId", value.ToString());
			}
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

		private ExcelWorkbook Workbook { get; set; }

		private List<PivotTableHeader> RowHeaders { get; } = new List<PivotTableHeader>();

		private List<PivotTableHeader> ColumnHeaders { get; } = new List<PivotTableHeader>();
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
				if (cache.SourceRange.IsEquivalentRange(sourceAddress))
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
		internal void RefreshFromCache()
		{
			// Update pivotField items to match corresponding cacheField sharedItems.
			foreach (var pivotField in this.Fields)
			{
				var fieldItems = pivotField.Items;
				var sharedItemsCount = this.CacheDefinition.CacheFields[pivotField.Index].SharedItems.Count;

				if (fieldItems.Count > sharedItemsCount + 1)
					throw new InvalidOperationException("There are more pivotField items than cacheField sharedItems.");

				// TODO (Task #8179): Change this to alphabetize the items.
				if (fieldItems.Count > 0)
				{
					for (int fieldIndex = 0; fieldIndex < sharedItemsCount; fieldIndex++)
					{
						if (fieldIndex < fieldItems.Count && string.IsNullOrEmpty(fieldItems[fieldIndex].T))
							fieldItems[fieldIndex].X = fieldIndex;
						else
							fieldItems.AddItem(fieldIndex, pivotField.DefaultSubtotal);
					}
				}
			}

			// Update the rowItems.
			this.RowItems.Clear();
			this.BuildRowItems(0, new List<Tuple<int, int>>());
			if (this.RowGrandTotals)
			{
				var header = new PivotTableHeader(null, null, -1, true, true);
				this.AddSumNodeToCollections(this.RowItems, this.RowHeaders, "grand", 0, 0, header);
			}

			// Update the colItems.
			if (this.ColumnFields.Any())
			{
				this.ColumnItems.Clear();
				this.BuildColumnItems(0, new List<Tuple<int, int>>(), false);
				if (this.ColumnGrandTotals)
				{
					// Only create item nodes and headers if there are column headers.
					if (this.ColumnFields[0].Index != -2)
					{
						if (this.DataFields.Count > 0)
						{
							for (int i = 0; i < this.DataFields.Count; i++)
							{
								var header = new PivotTableHeader(null, null, i, true, false);
								this.AddSumNodeToCollections(this.ColumnItems, this.ColumnHeaders, "grand", 0, 0, header, i);
							}
						}
						else
						{
							var header = new PivotTableHeader(null, null, 0, true, false);
							this.AddSumNodeToCollections(this.ColumnItems, this.ColumnHeaders, "grand", 0, 0, header);
						}
					}
				}
			}
			else
			{
				// If there are no column fields, then remove tag or else it will corrupt the workbook.
				this.TopNode.RemoveChild(this.TopNode.SelectSingleNode("d:colFields", this.NameSpaceManager));
				this.ColumnHeaders.Add(new PivotTableHeader(null, null, 0, false, false));
			}

			this.UpdateWorksheet();
		}
		#endregion

		#region Private Methods
		private void BuildRowItems(int rowDepth, List<Tuple<int, int>> parentNodeIndices)
		{
			// Base case.
			if (rowDepth >= this.RowFields.Count)
				return;

			var pivotFieldIndex = this.RowFields[rowDepth].Index;
			var pivotField = this.Fields[pivotFieldIndex];
			int maxIndex = pivotField.DefaultSubtotal ? pivotField.Items.Count - 1 : pivotField.Items.Count;
			for (int i = 0; i < maxIndex; i++)
			{
				var childList = parentNodeIndices.ToList();
				childList.Add(new Tuple<int, int>(pivotFieldIndex, pivotField.Items[i].X));
				if (this.CacheDefinition.CacheRecords.Contains(childList))
				{
					this.RowItems.Add(rowDepth, i);
					this.RowHeaders.Add(new PivotTableHeader(childList, pivotField, -1, false, true));
					this.BuildRowItems(rowDepth + 1, childList);
				}
			}
			if (pivotField.DefaultSubtotal && !pivotField.SubtotalTop && parentNodeIndices.Any())
			{
				var header = new PivotTableHeader(parentNodeIndices, pivotField, -1, false, true, "default");
				this.AddSumNodeToCollections(this.RowItems, this.RowHeaders, "default", 
					rowDepth - 1, parentNodeIndices.Last().Item2, header);
			}
		}

		private bool BuildColumnItems(int colDepth, List<Tuple<int, int>> parentNodeIndices, bool itemsCreated)
		{
			if (colDepth >= this.ColumnFields.Count)
				return true;

			var pivotFieldIndex = this.ColumnFields[colDepth].Index;
			ExcelPivotTableField pivotField = null; 
			int rValue = itemsCreated ? colDepth - 1 : colDepth;
			if (pivotFieldIndex == -2)
			{
				// This will create iNodes (row/column items) when there are multiple data fields.
				int repeatedItemsCountValue = rValue;
				for (int i = 0; i < this.DataFields.Count; i++)
				{
					var childList = parentNodeIndices.ToList();
					childList.Add(new Tuple<int, int>(-2, i));
					if (this.CreateColumnItemNode(itemsCreated, repeatedItemsCountValue, childList, pivotField, i) != 0)
					{
						if (repeatedItemsCountValue + 1 < this.ColumnFields.Count)
							repeatedItemsCountValue++;
					}
					itemsCreated = true;
				}
			}
			else
			{
				pivotField = this.Fields[pivotFieldIndex];
				int maxIndex = pivotField.DefaultSubtotal ? pivotField.Items.Count - 1 : pivotField.Items.Count;
				for (int i = 0; i < maxIndex; i++)
				{
					var childList = parentNodeIndices.ToList();
					childList.Add(new Tuple<int, int>(pivotFieldIndex, pivotField.Items[i].X));
					if (this.CacheDefinition.CacheRecords.Contains(childList))
					{
						bool result = this.BuildColumnItems(colDepth + 1, childList, itemsCreated);
						if (colDepth == this.ColumnFields.Count - 1)
						{
							// This will create iNodes when there is only one data field.
							this.CreateColumnItemNode(itemsCreated, rValue, childList.ToList(), pivotField, 0);
							itemsCreated = true;
						}
						else if (colDepth == 0)
							itemsCreated = false;
						else if (colDepth < this.ColumnFields.Count - 1)
							itemsCreated = result;
					}
				}

				if (pivotField.DefaultSubtotal && parentNodeIndices.Any())
				{
					int rAttribute = rValue == colDepth ? rValue - 1 : rValue;
					if (this.DataFields.Count > 0)
					{
						// Create a xml subtotal node for each data field.
						for (int i = 0; i < this.DataFields.Count; i++)
						{
							var header = new PivotTableHeader(parentNodeIndices, pivotField, i, false, false, "default");
							this.AddSumNodeToCollections(this.ColumnItems, this.ColumnHeaders, "default",
								rAttribute, parentNodeIndices.Last().Item2, header, i);
						}
					}
					else
					{
						// If there are no data fields, then create a single xml subtotal node.
						var header = new PivotTableHeader(parentNodeIndices, pivotField, 0, false, false, "default");
						this.AddSumNodeToCollections(this.ColumnItems, this.ColumnHeaders, "default",
							rAttribute, parentNodeIndices.Last().Item2, header);
					}
				}
			}

			return itemsCreated;
		}

		private void AddSumNodeToCollections(ItemsCollection collection, List<PivotTableHeader> headerList, 
			string itemType, int repeatedItemsCount, int xMemberValue, PivotTableHeader header, int dataFieldIndex = 0)
		{
			collection.AddSumNode(itemType, repeatedItemsCount, xMemberValue, dataFieldIndex);
			headerList.Add(header);
		}

		private int CreateColumnItemNode(bool itemsCreated, int rValue, List<Tuple<int, int>> recordIndices, ExcelPivotTableField pivotField, int dataFieldIndex)
		{
			int repeatedItemsCount = itemsCreated ? rValue : 0;
			this.ColumnHeaders.Add(new PivotTableHeader(recordIndices.ToList(), pivotField, dataFieldIndex, false, false));
			this.ColumnItems.AddColumnItem(recordIndices, repeatedItemsCount, dataFieldIndex);
			return repeatedItemsCount;
		}

		private void UpdateWorksheet()
		{
			this.UpdateRowColumnHeaders();
			if (this.DataFields.Any())
			{
				this.UpdatePivotTableWorksheetData();
				this.UpdatePivotTableGrandTotalData();
			}
			else
			{
				// If there are no data fields, then remove the xml node to prevent corrupting the workbook.
				this.TopNode.RemoveChild(this.TopNode.SelectSingleNode("d:dataFields", this.NameSpaceManager));
			}
		}

		private void UpdateRowColumnHeaders()
		{
			// Clear out the pivot table in the worksheet.
			int startRow = this.Address.Start.Row + this.FirstHeaderRow;
			int headerColumn = this.Address.Start.Column + this.FirstDataCol;
			int dataRow = this.Address.Start.Row + this.FirstDataRow;
			this.Worksheet.Cells[dataRow, this.Address.Start.Column, this.Address.End.Row, this.Address.Start.Column].Clear();
			this.Worksheet.Cells[startRow, headerColumn, this.Address.End.Row, this.Address.End.Column].Clear();

			// Update the row headers in the worksheet.
			foreach (var rowItem in this.RowItems)
			{
				bool itemType = this.SetTotalCellValue(this.RowFields, rowItem, dataRow, this.Address.Start.Column);
				if (itemType)
				{
					dataRow++;
					continue;
				}
				var sharedItem = this.GetSharedItemValue(this.RowFields, rowItem, rowItem.RepeatedItemsCount, 0);
				this.Worksheet.Cells[dataRow++, this.Address.Start.Column].Value = sharedItem;
			}
			
			// Update the column headers in the worksheet.
			foreach (var colItem in this.ColumnItems)
			{
				int startHeaderRow = startRow;
				bool itemType = this.SetTotalCellValue(this.ColumnFields, colItem, startHeaderRow, headerColumn);
				if (itemType)
				{
					headerColumn++;
					continue;
				}

				// Write in sum column header if the pivot table does not specify columns.
				if (this.TopNode.SelectSingleNode("d:colFields", this.NameSpaceManager) == null)
					this.Worksheet.Cells[startHeaderRow++, headerColumn].Value = this.DataFields[colItem.DataFieldIndex].Name;

				for (int i = 0; i < colItem.Count; i++)
				{
					var columnFieldIndex = colItem.RepeatedItemsCount == 0 ? i : i + colItem.RepeatedItemsCount;
					var sharedItem = this.GetSharedItemValue(this.ColumnFields, colItem, columnFieldIndex, i);
					var cellRow = colItem.RepeatedItemsCount == 0 ? startHeaderRow : startHeaderRow + colItem.RepeatedItemsCount;
					this.Worksheet.Cells[cellRow, headerColumn].Value = sharedItem;
					startHeaderRow++;
				}
				headerColumn++;
			}
		}

		private void UpdatePivotTableWorksheetData()
		{
			int dataColumn = this.Address.Start.Column + this.FirstDataCol;
			var subtotalStack = new List<double?>();
			foreach (var columnHeader in this.ColumnHeaders)
			{
				int dataRow = this.Address.Start.Row + this.FirstDataRow;
				foreach (var rowHeader in this.RowHeaders)
				{
					if (rowHeader.IsGrandTotal || columnHeader.IsGrandTotal)
						continue;

					var subtotal = this.CacheDefinition.CacheRecords.CalculateSubtotal(
						rowHeader.CacheRecordIndices, 
						columnHeader.CacheRecordIndices, 
						this.DataFields[columnHeader.DataFieldCollectionIndex].Index);

					if (rowHeader.CacheRecordIndices.Count == this.RowFields.Count)
						this.Worksheet.Cells[dataRow, dataColumn].Value = subtotal; // At a leaf node, write value.
					else if (rowHeader.PivotTableField.DefaultSubtotal)
					{
						// If subtotals are on and the rowItem's field specifies a sum type, then write out the last calculated sum.
						if (rowHeader.SumType != null)
						{
							this.Worksheet.Cells[dataRow, dataColumn].Value = subtotalStack.Last();
							subtotalStack.RemoveAt(subtotalStack.Count - 1);
						}
						else if (rowHeader.PivotTableField.SubtotalTop)
						{
							// If subtotalTop is true, then write out value.
							this.Worksheet.Cells[dataRow, dataColumn].Value = subtotal;
						}
						else
						{
							// Node is a non-subtotal node.
							subtotalStack.Add(subtotal);
						}
					}
					dataRow++;
				}
				dataColumn++;
			}
		}

		private void UpdatePivotTableGrandTotalData()
		{
			int dataStartRow = this.Address.Start.Row + this.FirstDataRow;
			var rowGrandTotals = new double?[this.RowHeaders.Count, this.DataFields.Count];
			int dataColumn = this.Address.Start.Column + this.FirstDataCol;
			// Calculating row and column grand totals, write in row grand totals.
			for (int i = 0; i < this.ColumnHeaders.Count; i++)
			{
				var columnHeader = this.ColumnHeaders[i];
				int dataRow = dataStartRow;
				double? grandTotal = null;
				int rowItemsCount = this.RowGrandTotals ? this.RowHeaders.Count - 1 : this.RowHeaders.Count;
				int j = 0;
				for (; j < rowItemsCount; j++)
				{
					if (this.Worksheet.Cells[dataRow, dataColumn].Value.IsNumeric())
					{
						double value = this.Worksheet.Cells[dataRow, dataColumn].GetValue<double>();
						// If the subtotals are displayed at the top, then add all the root nodes' values to the column grand totals.
						// Otherwise, add all leaf node values.
						var pivotTableField = this.RowHeaders[j].PivotTableField;
						if (pivotTableField.SubtotalTop && pivotTableField.DefaultSubtotal)
						{
							if (string.IsNullOrEmpty(this.RowHeaders[j].SumType) && this.RowItems[j].RepeatedItemsCount == 0)
								grandTotal = grandTotal == null ? value : value + grandTotal;
						}
						else
						{
							if (string.IsNullOrEmpty(this.RowHeaders[j].SumType))
								grandTotal = grandTotal == null ? value : value + grandTotal;
						}

						// Add the value to the corresponding row grand total if it is not a subtotal node.
						if (string.IsNullOrEmpty(columnHeader.SumType))
						{
							var totalValue = rowGrandTotals[j, columnHeader.DataFieldCollectionIndex];
							rowGrandTotals[j, columnHeader.DataFieldCollectionIndex] = totalValue == null ? value : value + totalValue;
						}
					}
					dataRow++;
				}

				// Write in the row grand totals (grand total for each column).
				if (this.RowGrandTotals) {
					this.Worksheet.Cells[dataRow, dataColumn].Value = grandTotal;
					// Only sum up the non-subtotal column grand total values.
					if (grandTotal != null && string.IsNullOrEmpty(columnHeader.SumType))
					{
						var totalValue = rowGrandTotals[j, columnHeader.DataFieldCollectionIndex];
						rowGrandTotals[j, columnHeader.DataFieldCollectionIndex] = totalValue == null ? grandTotal : totalValue + grandTotal;
					}
				}
				dataColumn++;
			}

			// Write in the column grand totals (grand total for each row).
			if (this.ColumnGrandTotals)
			{
				if (this.ColumnFields.Count > 0 && this.ColumnFields[0].Index != -2)
				{
					for (int dataField = 0; dataField < this.DataFields.Count; dataField++)
					{
						for (int row = 0; row < rowGrandTotals.GetLength(0); row++)
						{
							var column = this.Address.End.Column - (this.DataFields.Count - 1) + dataField;
							this.Worksheet.Cells[dataStartRow + row, column].Value = rowGrandTotals[row, dataField];
						}
					}
				}
			}
		}

		private bool SetTotalCellValue(ExcelPivotTableRowColumnFieldCollection field, RowColumnItem item, int row, int column)
		{
			if (!string.IsNullOrEmpty(item.ItemType))
			{
				int rowLabel = field == this.RowFields ? row : row + item.RepeatedItemsCount;
				if (item.ItemType.IsEquivalentTo("grand"))
				{
					if (field == this.ColumnFields)
					{
						if (this.DataFields.Count == 1)
							this.Worksheet.Cells[rowLabel, column].Value = "Grand Total";
						else if (this.DataFields.Count > 1)
						{
							string dataFieldName = this.DataFields[item.DataFieldIndex].Name;
							this.Worksheet.Cells[rowLabel, column].Value = $"Total {dataFieldName}";
						}
						else
							this.Worksheet.Cells[rowLabel, column].Value = $"Grand Total";
					}
					else
						this.Worksheet.Cells[rowLabel, column].Value = $"Grand Total";
				}
				else if (item.ItemType.IsEquivalentTo("default"))
				{
					var itemName = this.GetSharedItemValue(field, item, item.RepeatedItemsCount, 0);
					if (this.DataFields.Count > 1)
					{
						string dataFieldName = this.DataFields[item.DataFieldIndex].Name;
						this.Worksheet.Cells[rowLabel, column].Value = $"{itemName} {dataFieldName}";
					}
					else
						this.Worksheet.Cells[rowLabel, column].Value = $"{itemName} Total";
				}
				return true;
			}
			return false;
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