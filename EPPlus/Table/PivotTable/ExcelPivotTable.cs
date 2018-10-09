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
using System.Xml;
using OfficeOpenXml.Extensions;
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
				if (this.WorkSheet.Workbook.ExistsTableName(value))
					throw (new ArgumentException("PivotTable name is not unique"));
				string prevName = this.Name;
				if (this.WorkSheet.Tables.TableNames.ContainsKey(prevName))
				{
					int ix = this.WorkSheet.Tables.TableNames[prevName];
					this.WorkSheet.Tables.TableNames.Remove(prevName);
					this.WorkSheet.Tables.TableNames.Add(value, ix);
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
					foreach (var cacheDefinition in this.WorkSheet.Workbook.PivotCacheDefinitions)
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
		public ExcelWorksheet WorkSheet { get; set; }
		
		/// <summary>
		/// Gets or sets the location of the pivot table.
		/// </summary>
		public ExcelAddress Address { get; internal set; }
		
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
		public bool ColumGrandTotals
		{
			get
			{
				return base.GetXmlNodeBool("@colGrandTotals");
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
			this.WorkSheet = sheet;
			this.PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
			this.WorksheetRelationship = rel;
			var pck = sheet.Package.Package;
			this.Part = pck.GetPart(this.PivotTableUri);

			this.PivotTableXml = new XmlDocument();
			XmlHelper.LoadXmlSafe(this.PivotTableXml, this.Part.GetStream());
			this.InitSchemaNodeOrder();
			this.TopNode = this.PivotTableXml.DocumentElement;
			this.Address = new ExcelAddress(base.GetXmlNodeString("d:location/@ref"));

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
			this.WorkSheet = sheet;
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
			foreach (var cache in this.WorkSheet.Workbook.PivotCacheDefinitions)
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
			var previousRowItemsCount = this.RowItems.Sum(r => r.Count);
			this.RowItems.Clear();
			this.BuildRowItems(0, new List<Tuple<int, int>>());
			if (this.RowGrandTotals)
				this.RowItems.AddSumNode("grand");
			this.UpdateWorksheet(previousRowItemsCount);
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
					this.BuildRowItems(rowDepth + 1, childList);
				}
			}
		}

		private void UpdateWorksheet(int previousRowItemsCount)
		{
			int startDataRow = this.Address.Start.Row + this.FirstDataRow;
			var startDataCol = this.Address.Start.Column + this.FirstDataCol - 1;
			int currentDataRow = startDataRow;
			foreach (var rowItem in this.RowItems)
			{
				if (!string.IsNullOrEmpty(rowItem.ItemType))
				{
					this.WorkSheet.Cells[currentDataRow++, startDataCol].Value = "Grand Total";
					continue;
				}

				var fieldIndex = this.RowFields[rowItem.RepeatedItemsCount].Index;

				// TODO 8178: Account for grouped rowItems.
				var fieldItemIndex = rowItem[0];

				var pivotField = this.Fields[fieldIndex]; 
				var cacheItemIndex = pivotField.Items[fieldItemIndex].X;
				var cacheField = this.CacheDefinition.CacheFields[fieldIndex];
				var sharedItem = cacheField.SharedItems[cacheItemIndex].Value;
				this.WorkSheet.Cells[currentDataRow++, startDataCol].Value = sharedItem;
			}
			// Clear out cells after the pivot table if the data was removed.
			int pivotTableHeight = currentDataRow - startDataRow;
			for (int i = pivotTableHeight; i < previousRowItemsCount; i++)
			{
				this.WorkSheet.Cells[currentDataRow++, startDataCol].Value = null;
			}
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