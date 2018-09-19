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
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
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
		private TableStyles myTableStyle = Table.TableStyles.Medium6;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the xml data representing the pivot table in the package.
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
				string prevName = Name;
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
					myCacheDefinition = new ExcelPivotCacheDefinition(this.NameSpaceManager, this, null, 1);
				return myCacheDefinition;
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
		public bool RowGrandTotals
		{
			get
			{
				return base.GetXmlNodeBool("@rowGrandTotals");
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
					myFields = new ExcelPivotTableFieldCollection(this, "");
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
					myRowFields = new ExcelPivotTableRowColumnFieldCollection(this, "rowFields");
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
					myColumnFields = new ExcelPivotTableRowColumnFieldCollection(this, "colFields");
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
					myDataFields = new ExcelPivotTableDataFieldCollection(this);
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
					myPageFields = new ExcelPivotTableRowColumnFieldCollection(this, "pageFields");
				return myPageFields;
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
				return GetXmlNodeInt("@cacheId");
			}
			set
			{
				SetXmlNodeString("@cacheId", value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the pivot table part.
		/// </summary>
		internal Packaging.ZipPackagePart Part { get; set; }

		/// <summary>
		/// Gets or sets the worksheet-pivot table relationship.
		/// </summary>
		internal Packaging.ZipPackageRelationship Relationship { get; set; }
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
			this.Relationship = rel;
			var pck = sheet.Package.Package;
			this.Part = pck.GetPart(this.PivotTableUri);

			this.PivotTableXml = new XmlDocument();
			XmlHelper.LoadXmlSafe(this.PivotTableXml, this.Part.GetStream());
			this.InitSchemaNodeOrder();
			this.TopNode = this.PivotTableXml.DocumentElement;
			this.Address = new ExcelAddress(base.GetXmlNodeString("d:location/@ref"));

			myCacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this);
			this.LoadFields();

			// Add row fields.
			foreach (XmlElement rowElem in this.TopNode.SelectNodes("d:rowFields/d:field", this.NameSpaceManager))
			{
				if (int.TryParse(rowElem.GetAttribute("x"), out var x) && x >= 0)
					this.RowFields.AddInternal(this.Fields[x]);
				else
					rowElem.ParentNode.RemoveChild(rowElem);
			}

			// Add column fields.
			foreach (XmlElement colElem in this.TopNode.SelectNodes("d:colFields/d:field", this.NameSpaceManager))
			{
				if (int.TryParse(colElem.GetAttribute("x"), out var x) && x >= 0)
					this.ColumnFields.AddInternal(this.Fields[x]);
				else
					colElem.ParentNode.RemoveChild(colElem);
			}

			// Add Page elements
			foreach (XmlElement pageElem in this.TopNode.SelectNodes("d:pageFields/d:pageField", this.NameSpaceManager))
			{
				if (int.TryParse(pageElem.GetAttribute("fld"), out var fld) && fld >= 0)
				{
					var field = this.Fields[fld];
					field.myPageFieldSettings = new ExcelPivotTablePageFieldSettings(this.NameSpaceManager, pageElem, field, fld);
					this.PageFields.AddInternal(field);
				}
			}

			// Add data elements
			foreach (XmlElement dataElem in this.TopNode.SelectNodes("d:dataFields/d:dataField", this.NameSpaceManager))
			{
				if (int.TryParse(dataElem.GetAttribute("fld"), out var fld) && fld >= 0)
				{
					var field = this.Fields[fld];
					var dataField = new ExcelPivotTableDataField(this.NameSpaceManager, dataElem, field);
					this.DataFields.AddInternal(dataField);
				}
			}
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
			this.Relationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, this.PivotTableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");

			myCacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress, tblId);
			myCacheDefinition.Relationship = this.Part.CreateRelationship(UriHelper.ResolvePartUri(this.PivotTableUri, myCacheDefinition.CacheDefinitionUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");

			sheet.Workbook.AddPivotTable(this.CacheID.ToString(), myCacheDefinition.CacheDefinitionUri);

			this.LoadFields();

			using (var range = sheet.Cells[address.Address])
			{
				range.Clear();
			}
		}
		#endregion

		#region Private Methods
		private void InitSchemaNodeOrder()
		{
			this.SchemaNodeOrder = new string[] { "location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "pageItems", "dataFields", "dataItems", "formats", "pivotTableStyleInfo" };
		}

		private void LoadFields()
		{
			int index = 0;
			// Add fields.
			foreach (XmlElement fieldElem in this.TopNode.SelectNodes("d:pivotFields/d:pivotField", this.NameSpaceManager))
			{
				var fld = new ExcelPivotTableField(this.NameSpaceManager, fieldElem, this, index, index++);
				this.Fields.AddInternal(fld);
			}

			// Add fields.
			index = 0;
			foreach (XmlElement fieldElem in myCacheDefinition.TopNode.SelectNodes("d:cacheFields/d:cacheField", this.NameSpaceManager))
			{
				var fld = this.Fields[index++];
				fld.SetCacheFieldNode(fieldElem);
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