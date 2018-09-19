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
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	#region Enums
	public enum eSourceType
	{
		/// <summary>
		/// Indicates that the cache contains data that consolidates ranges.
		/// </summary>
		Consolidation,
		/// <summary>
		/// Indicates that the cache contains data from an external data source.
		/// </summary>
		External,
		/// <summary>
		/// Indicates that the cache contains a scenario summary report.
		/// </summary>
		Scenario,
		/// <summary>
		/// Indicates that the cache contains worksheet data.
		/// </summary>
		Worksheet
	}
	#endregion

	/// <summary>
	/// Cache definition. This class defines the source data. Note that one cache definition can be shared between many pivot tables.
	/// </summary>
	public class ExcelPivotCacheDefinition : XmlHelper
	{
		#region Constants
		/// <summary>
		/// The path of the data source worksheet.
		/// </summary>
		internal const string SourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";
		
		/// <summary>
		/// The path of the data source cell range.
		/// </summary>
		internal const string SourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";

		private const string SourceNamePath = "d:cacheSource/d:worksheetSource/@name";
		#endregion

		#region Class Variables
		/// <summary>
		/// The source data range for the pivot table.
		/// </summary>
		internal ExcelRangeBase mySourceRange;

		private List<CacheFieldNode> myCacheFields = new List<CacheFieldNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the XML data representing the cache definition in the package.
		/// </summary>
		public XmlDocument CacheDefinitionXml { get; private set; }
		
		/// <summary>
		/// Gets or sets the package internal URI to the pivot table cache definition Xml Document.
		/// </summary>
		public Uri CacheDefinitionUri { get; internal set; }
		
		/// <summary>
		/// Gets or sets the reference to the PivotTable object.
		/// </summary>
		public ExcelPivotTable PivotTable { get; private set; }

		/// <summary>
		/// Gets the cache fields.
		/// </summary>
		public IReadOnlyList<CacheFieldNode> CacheFields
		{
			get { return myCacheFields; }
		}

		/// <summary>
		/// Gets or sets the source data range when the pivot table has a worksheet data source. 
		/// The number of columns in the range must be intact if this property is changed.
		/// The range must be in the same workbook as the pivot table.
		/// </summary>
		public ExcelRangeBase SourceRange
		{
			get
			{
				if (mySourceRange == null)
				{
					if (this.CacheSource == eSourceType.Worksheet)
					{
						var ws = this.PivotTable.WorkSheet.Workbook.Worksheets[base.GetXmlNodeString(SourceWorksheetPath)];
						if (ws == null)
						{
							var name = base.GetXmlNodeString(SourceNamePath);
							if (this.PivotTable.WorkSheet.Workbook.Names.ContainsKey(name))
							{
								mySourceRange = this.PivotTable.WorkSheet.Workbook.Names[name].GetFormulaAsCellRange();
								return mySourceRange;
							}
							foreach (var worksheet in this.PivotTable.WorkSheet.Workbook.Worksheets)
							{
								if (worksheet.Tables.TableNames.ContainsKey(name))
								{
									mySourceRange = new ExcelRangeBase(worksheet.Workbook, worksheet, name, true);
									break;
								}
								else if (worksheet.Names.ContainsKey(name))
								{
									mySourceRange = worksheet.Names[name].GetFormulaAsCellRange();
									break;
								}
							}
						}
						else
							mySourceRange = ws.Cells[base.GetXmlNodeString(SourceAddressPath)];
					}
					else
						return null;
				}
				return mySourceRange;
			}
			set
			{
				if (this.PivotTable.WorkSheet.Workbook != value.Worksheet.Workbook)
					throw (new ArgumentException("Range must be in the same package as the pivottable"));
				var sourceRange = this.SourceRange;
				if (value.End.Column - value.Start.Column != sourceRange.End.Column - sourceRange.Start.Column)
					throw (new ArgumentException("Can not change the number of columns(fields) in the SourceRange"));
				base.SetXmlNodeString(SourceWorksheetPath, value.Worksheet.Name);
				base.SetXmlNodeString(SourceAddressPath, value.FirstAddress);
				mySourceRange = value;
			}
		}

		/// <summary>
		/// Gets the type of source data.
		/// </summary>
		public eSourceType CacheSource
		{
			get
			{
				var s = base.GetXmlNodeString("d:cacheSource/@type");
				if (s == "")
					return eSourceType.Worksheet;
				else
					return (eSourceType)Enum.Parse(typeof(eSourceType), s, true);
			}
		}

		/// <summary>
		/// Gets or sets the reference to the internal package part.
		/// </summary>
		internal Packaging.ZipPackagePart Part { get; set; }

		/// <summary>
		/// Gets or sets the package internal URI to the pivot table cache record Xml Document.
		/// </summary>
		internal Uri CacheRecordUri { get; set; }

		/// <summary>
		/// Gets or sets the relationship to the pivot table.
		/// </summary>
		internal Packaging.ZipPackageRelationship Relationship { get; set; }

		/// <summary>
		/// Gets or sets the relationship to the cache record.
		/// </summary>
		internal Packaging.ZipPackageRelationship RecordRelationship { get; set; }

		/// <summary>
		/// Gets or sets the record relationship id.
		/// </summary>
		internal string RecordRelationshipID
		{
			get
			{
				return base.GetXmlNodeString("@r:id");
			}
			set
			{
				base.SetXmlNodeString("@r:id", value);
			}
		}
		#endregion
		
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotCacheDefinition"/> from the namespace and pivot table.
		/// </summary>
		/// <param name="ns">The namespace of the sheet.</param>
		/// <param name="pivotTable">The pivot table using this definition.</param>
		internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable) :
			 base(ns, null)
		{
			foreach (var r in pivotTable.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheDefinition"))
			{
				this.Relationship = r;
			}
			this.CacheDefinitionUri = UriHelper.ResolvePartUri(this.Relationship.SourceUri, this.Relationship.TargetUri);

			var pck = pivotTable.WorkSheet.Package.Package;
			this.Part = pck.GetPart(this.CacheDefinitionUri);
			this.CacheDefinitionXml = new XmlDocument();
			XmlHelper.LoadXmlSafe(this.CacheDefinitionXml, this.Part.GetStream());

			this.TopNode = this.CacheDefinitionXml.DocumentElement;
			this.PivotTable = pivotTable;
			if (this.CacheSource == eSourceType.Worksheet)
			{
				var worksheetName = base.GetXmlNodeString(SourceWorksheetPath);
				if (pivotTable.WorkSheet.Workbook.Worksheets.Any(t => t.Name == worksheetName))
				{
					mySourceRange = pivotTable.WorkSheet.Workbook.Worksheets[worksheetName].Cells[base.GetXmlNodeString(SourceAddressPath)];
				}
			}
			foreach (XmlNode cacheFieldNode in this.TopNode.SelectNodes("d:cacheFields/d:cacheField", this.NameSpaceManager))
			{
				myCacheFields.Add(new CacheFieldNode(cacheFieldNode, this.NameSpaceManager));
			}
		}

		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotCacheDefinition"/> using the data source address and pivot table id.
		/// </summary>
		/// <param name="ns">The namespace of the sheet.</param>
		/// <param name="pivotTable">The pivot table using this definition.</param>
		/// <param name="sourceAddress">The address of the data source.</param>
		/// <param name="tableId">The pivot table id.</param>
		internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tableId) :
			 base(ns, null)
		{
			this.PivotTable = pivotTable;

			var pck = pivotTable.WorkSheet.Package.Package;

			// CacheDefinition
			this.CacheDefinitionXml = new XmlDocument();
			XmlHelper.LoadXmlSafe(this.CacheDefinitionXml, this.GetStartXml(sourceAddress), Encoding.UTF8);
			this.CacheDefinitionUri = XmlHelper.GetNewUri(pck, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref tableId);
			this.Part = pck.CreatePart(this.CacheDefinitionUri, ExcelPackage.schemaPivotCacheDefinition);
			this.TopNode = this.CacheDefinitionXml.DocumentElement;

			// CacheRecord. Create an empty one.
			this.CacheRecordUri = XmlHelper.GetNewUri(pck, "/xl/pivotCache/pivotCacheRecords{0}.xml", ref tableId);
			var cacheRecord = new XmlDocument();
			cacheRecord.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
			var recPart = pck.CreatePart(this.CacheRecordUri, ExcelPackage.schemaPivotCacheRecords);
			cacheRecord.Save(recPart.GetStream());

			this.RecordRelationship = this.Part.CreateRelationship(UriHelper.ResolvePartUri(this.CacheDefinitionUri, this.CacheRecordUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
			this.RecordRelationshipID = this.RecordRelationship.Id;

			this.CacheDefinitionXml.Save(this.Part.GetStream());
		}
		#endregion

		#region Private Methods
		private string GetStartXml(ExcelRangeBase sourceAddress)
		{
			string xml = "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"1\" refreshedVersion=\"3\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

			xml += "<cacheSource type=\"worksheet\">";
			xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", sourceAddress.Address, sourceAddress.WorkSheet);
			xml += "</cacheSource>";
			xml += string.Format("<cacheFields count=\"{0}\">", sourceAddress._toCol - sourceAddress._fromCol + 1);
			var sourceWorksheet = this.PivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheet];
			for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
			{
				if (sourceWorksheet == null || sourceWorksheet.GetValueInner(sourceAddress._fromRow, col) == null || sourceWorksheet.GetValueInner(sourceAddress._fromRow, col).ToString().Trim() == "")
					xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - sourceAddress._fromCol + 1);
				else
					xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", sourceWorksheet.GetValueInner(sourceAddress._fromRow, col));
				//xml += "<sharedItems containsNonDate=\"0\" containsString=\"0\" containsBlank=\"1\" /> ";
				xml += "<sharedItems containsBlank=\"1\" /> ";
				xml += "</cacheField>";
			}
			xml += "</cacheFields>";
			xml += "</pivotCacheDefinition>";

			return xml;
		}
		#endregion
	}
}