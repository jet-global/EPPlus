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
using System.IO;
using System.Linq;
using System.Resources;
using System.Text;
using System.Xml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Internationalization;
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

		private const string Name = "pivotCacheDefinition";
		#endregion

		#region Class Variables
		/// <summary>
		/// The source data range for the pivot table.
		/// </summary>
		internal ExcelRangeBase mySourceRange;

		private List<CacheFieldNode> myCacheFields = new List<CacheFieldNode>();

		private ExcelPivotCacheRecords myCacheRecords;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the package internal URI to the pivot table cache definition Xml Document.
		/// </summary>
		public Uri CacheDefinitionUri { get; internal set; }

		/// <summary>
		/// Gets or sets the workbook.
		/// </summary>
		public ExcelWorkbook Workbook { get; set; }

		/// <summary>
		/// Gets the cache fields.
		/// </summary>
		public IReadOnlyList<CacheFieldNode> CacheFields
		{
			get { return myCacheFields; }
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
		/// Gets or sets the XML data representing the cache definition in the package.
		/// </summary>
		internal XmlDocument CacheDefinitionXml { get; private set; }

		/// <summary>
		/// Gets or sets the reference to the internal package part.
		/// </summary>
		internal Packaging.ZipPackagePart Part { get; set; }

		/// <summary>
		/// Gets the pivot table's cache records.
		/// </summary>
		internal ExcelPivotCacheRecords CacheRecords
		{
			get
			{
				if (myCacheRecords == null && this.CacheSource == eSourceType.Worksheet)
				{
					var cacheRecordsRelationship = this.Part.GetRelationshipsByType(ExcelPackage.schemaPivotCacheRecordsRelationship);
					string cacheDefinitionName = UriHelper.GetUriEndTargetName(this.CacheDefinitionUri, ExcelPivotCacheDefinition.Name);
					foreach (var cacheRecordsRel in cacheRecordsRelationship)
					{
						if (UriHelper.GetUriEndTargetName(cacheRecordsRel.SourceUri).IsEquivalentTo(cacheDefinitionName))
						{
							var partUri = new Uri($"xl/pivotCache/{cacheRecordsRel.TargetUri}", UriKind.Relative);
							var possiblePart = this.Workbook.Package.GetXmlFromUri(partUri);
							myCacheRecords = new ExcelPivotCacheRecords(base.NameSpaceManager, this.Workbook.Package, possiblePart, partUri, this);
						}
					}
				}
				return myCacheRecords;
			}
			private set
			{
				myCacheRecords = value;
			}
		}

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

		/// <summary>
		/// Gets the <see cref="StringResources"/> for this <see cref="ExcelPackage"/> that 
		/// can be used to get localized string translations if a <see cref="ResourceManager"/> is loaded.
		/// </summary>
		internal StringResources StringResources { get; } = new StringResources();

		private ExcelRangeBase SourceRange
		{
			// The range must be in the same workbook as the pivot table.
			get
			{
				if (mySourceRange == null)
				{
					if (this.CacheSource == eSourceType.Worksheet)
					{
						var ws = this.Workbook.Worksheets[base.GetXmlNodeString(SourceWorksheetPath)];
						if (ws == null)
						{
							var name = base.GetXmlNodeString(SourceNamePath);
							if (this.Workbook.Names.ContainsKey(name))
							{
								mySourceRange = this.Workbook.Names[name].GetFormulaAsCellRange();
								return mySourceRange;
							}
							foreach (var worksheet in this.Workbook.Worksheets)
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
				if (this.Workbook != value.Worksheet.Workbook)
					throw (new ArgumentException("Range must be in the same package as the pivottable"));
				var sourceRange = this.SourceRange;
				base.SetXmlNodeString(SourceWorksheetPath, value.Worksheet.Name);
				base.SetXmlNodeString(SourceAddressPath, value.FirstAddress);
				mySourceRange = value;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an existing <see cref="ExcelPivotCacheDefinition"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="package">The Excel package.</param>
		/// <param name="xmlDocument">The <see cref="ExcelPivotCacheDefinition"/> xml document.</param>
		/// <param name="cacheUri">The <see cref="ExcelPivotCacheDefinition"/> uri.</param>
		internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPackage package, XmlDocument xmlDocument, Uri cacheUri) :
			 base(ns, null)
		{
			if (ns == null)
				throw new ArgumentNullException(nameof(ns));
			if (package == null)
				throw new ArgumentNullException(nameof(package));
			if (xmlDocument == null)
				throw new ArgumentNullException(nameof(xmlDocument));
			if (cacheUri == null)
				throw new ArgumentNullException(nameof(cacheUri));
			this.CacheDefinitionXml = xmlDocument;
			base.TopNode = xmlDocument.SelectSingleNode($"d:{ExcelPivotCacheDefinition.Name}", ns);

			this.CacheDefinitionUri = cacheUri;
			this.Part = package.Package.GetPart(this.CacheDefinitionUri);
			this.Workbook = package.Workbook;
			if (this.CacheSource == eSourceType.Worksheet)
			{
				var worksheetName = base.GetXmlNodeString(SourceWorksheetPath);
				if (this.Workbook.Worksheets.Any(t => t.Name == worksheetName))
					mySourceRange = this.Workbook.Worksheets[worksheetName].Cells[base.GetXmlNodeString(SourceAddressPath)];
			}
			var cacheFieldNodes = this.TopNode.SelectNodes("d:cacheFields/d:cacheField", base.NameSpaceManager);
			foreach (XmlNode cacheFieldNode in cacheFieldNodes)
			{
				myCacheFields.Add(new CacheFieldNode(base.NameSpaceManager, cacheFieldNode));
			}
		}

		/// <summary>
		/// Creates an instance of an empty <see cref="ExcelPivotCacheDefinition"/> using the data source address and pivot table id.
		/// </summary>
		/// <param name="ns">The namespace of the sheet.</param>
		/// <param name="pivotTable">The pivot table using this definition.</param>
		/// <param name="sourceAddress">The address of the data source.</param>
		/// <param name="tableId">The pivot table id.</param>
		internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tableId) :
			 base(ns, null)
		{
			if (ns == null)
				throw new ArgumentNullException(nameof(ns));
			if (pivotTable == null)
				throw new ArgumentNullException(nameof(pivotTable));
			if (sourceAddress == null)
				throw new ArgumentNullException(nameof(sourceAddress));
			if (tableId < 1)
				throw new ArgumentOutOfRangeException(nameof(tableId));
			this.Workbook = pivotTable.Worksheet.Workbook;

			var pck = pivotTable.Worksheet.Package.Package;

			// CacheDefinition
			this.CacheDefinitionXml = new XmlDocument();
			XmlHelper.LoadXmlSafe(this.CacheDefinitionXml, this.GetStartXml(sourceAddress), Encoding.UTF8);
			this.CacheDefinitionUri = XmlHelper.GetNewUri(pck, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref tableId);
			this.Part = pck.CreatePart(this.CacheDefinitionUri, ExcelPackage.schemaPivotCacheDefinition);
			this.TopNode = this.CacheDefinitionXml.DocumentElement;

			// CacheRecord. Create an empty one.
			this.CacheRecords = new ExcelPivotCacheRecords(ns, pck, ref tableId, this);

			this.RecordRelationship = this.Part.CreateRelationship(UriHelper.ResolvePartUri(this.CacheDefinitionUri, this.CacheRecords.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
			this.RecordRelationshipID = this.RecordRelationship.Id;

			this.CacheDefinitionXml.Save(this.Part.GetStream());
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Update the records in <see cref="ExcelPivotCacheRecords"/> and any referencing <see cref="ExcelPivotTable"/>s.
		/// </summary>
		/// <param name="resourceManager">The <see cref="ResourceManager"/> to retrieve translations from (optional).</param>
		public void UpdateData(ResourceManager resourceManager = null)
		{
			var sourceRange = this.GetSourceRangeAddress();
			// If the source range is an Excel pivot table or named range, resolve the address.
			if (sourceRange.IsName)
				sourceRange = AddressUtility.GetFormulaAsCellRange(this.Workbook, sourceRange.Worksheet, sourceRange.Address);

			// Update all cacheField names assuming the shape of the pivot cache definition source range remains unchanged.
			for (int col = sourceRange.Start.Column; col < sourceRange.Columns + sourceRange.Start.Column; col++)
			{
				int fieldIndex = col - sourceRange.Start.Column;
				this.CacheFields[fieldIndex].Name = sourceRange.Worksheet.Cells[sourceRange.Start.Row, col].Value.ToString();
			}

			// Update all cache record values.
			var worksheet = sourceRange.Worksheet;
			var range = new ExcelRange(worksheet, worksheet.Cells[sourceRange.Start.Row + 1, sourceRange.Start.Column, sourceRange.End.Row, sourceRange.End.Column]);
			this.CacheRecords.UpdateRecords(range);

			this.StringResources.LoadResourceManager(resourceManager);

			// Refresh pivot tables.
			foreach (var pivotTable in this.GetRelatedPivotTables())
			{
				pivotTable.RefreshFromCache(this.StringResources);
			}

			// Remove the 'u' xml attribute from each cache item to prevent corrupting the workbook, since Excel automatically adds them.
			foreach (var cacheField in this.CacheFields)
			{
				if (cacheField.HasSharedItems)
					cacheField.RemoveXmlUAttribute();
			}
		}

		/// <summary>
		/// Gets all the <see cref="ExcelPivotTable"/>s referencing this cache definition.
		/// </summary>
		/// <returns>A list of <see cref="ExcelPivotTable"/>s.</returns>
		public List<ExcelPivotTable> GetRelatedPivotTables()
		{
			var pivotTables = new List<ExcelPivotTable>();
			foreach (var worksheet in this.Workbook.Worksheets)
			{
				foreach (var pivotTable in worksheet.PivotTables)
				{
					if (pivotTable.CacheDefinition == this)
						pivotTables.Add(pivotTable);
				}
			}
			return pivotTables;
		}

		/// <summary>
		/// Gets the source range of the source data table.
		/// </summary>
		/// <returns>The source range.</returns>
		public ExcelRangeBase GetSourceRangeAddress()
		{
			return this.SourceRange;
		}

		/// <summary>
		/// Sets the source range of the source data table.
		/// </summary>
		/// <param name="worksheet">The worksheet that the source data table is on.</param>
		/// <param name="address">The updated address.</param>
		public void SetSourceRangeAddress(ExcelWorksheet worksheet, string address)
		{
			this.SourceRange = new ExcelRangeBase(worksheet, address);
		}

		/// <summary>
		/// Gets the index of the a cache field with the specified <paramref name="fieldName"/>.
		/// </summary>
		/// <param name="fieldName">The name of the cache field to find the index of.</param>
		/// <returns>The index of a cache field matching the specified name, -1 if not found.</returns>
		public int GetCacheFieldIndex(string fieldName)
		{
			for (int i = 0; i < this.CacheFields.Count; i++)
			{
				if (this.CacheFields[i].Name.IsEquivalentTo(fieldName))
					return i;
			}
			return -1;
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Save the cacheDefinition and cacheRecords xml.
		/// </summary>
		internal void Save()
		{
			this.CacheDefinitionXml.Save(this.Part.GetStream(FileMode.Create));
			this.CacheRecords?.Save();
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
			var sourceWorksheet = this.Workbook.Worksheets[sourceAddress.WorkSheet];
			for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
			{
				if (sourceWorksheet == null || sourceWorksheet.GetValueInner(sourceAddress._fromRow, col) == null || sourceWorksheet.GetValueInner(sourceAddress._fromRow, col).ToString().Trim() == "")
					xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - sourceAddress._fromCol + 1);
				else
					xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", sourceWorksheet.GetValueInner(sourceAddress._fromRow, col));
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