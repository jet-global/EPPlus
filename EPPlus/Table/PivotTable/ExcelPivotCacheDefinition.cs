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

		private static readonly IReadOnlyList<string> SecondsMinutes = new List<string>
		{
			":00",
			":01",
			":02",
			":03",
			":04",
			":05",
			":06",
			":07",
			":08",
			":09",
			":10",
			":11",
			":12",
			":13",
			":14",
			":15",
			":16",
			":17",
			":18",
			":19",
			":20",
			":21",
			":22",
			":23",
			":24",
			":25",
			":26",
			":27",
			":28",
			":29",
			":30",
			":31",
			":32",
			":33",
			":34",
			":35",
			":36",
			":37",
			":38",
			":39",
			":40",
			":41",
			":42",
			":43",
			":44",
			":45",
			":46",
			":47",
			":48",
			":49",
			":50",
			":51",
			":52",
			":53",
			":54",
			":55",
			":56",
			":57",
			":58",
			":59"
		};

		private static readonly IReadOnlyList<string> Hours = new List<string>
		{
			"12 AM",
			"1 AM",
			"2 AM",
			"3 AM",
			"4 AM",
			"5 AM",
			"6 AM",
			"7 AM",
			"8 AM",
			"9 AM",
			"10 AM",
			"11 AM",
			"12 PM",
			"1 PM",
			"2 PM",
			"3 PM",
			"4 PM",
			"5 PM",
			"6 PM",
			"7 PM",
			"8 PM",
			"9 PM",
			"10 PM",
			"11 PM"
		};

		private static readonly IReadOnlyList<string> Days = new List<string>
		{
			"1-Jan",
			"2-Jan",
			"3-Jan",
			"4-Jan",
			"5-Jan",
			"6-Jan",
			"7-Jan",
			"8-Jan",
			"9-Jan",
			"10-Jan",
			"11-Jan",
			"12-Jan",
			"13-Jan",
			"14-Jan",
			"15-Jan",
			"16-Jan",
			"17-Jan",
			"18-Jan",
			"19-Jan",
			"20-Jan",
			"21-Jan",
			"22-Jan",
			"23-Jan",
			"24-Jan",
			"25-Jan",
			"26-Jan",
			"27-Jan",
			"28-Jan",
			"29-Jan",
			"30-Jan",
			"31-Jan",
			"1-Feb",
			"2-Feb",
			"3-Feb",
			"4-Feb",
			"5-Feb",
			"6-Feb",
			"7-Feb",
			"8-Feb",
			"9-Feb",
			"10-Feb",
			"11-Feb",
			"12-Feb",
			"13-Feb",
			"14-Feb",
			"15-Feb",
			"16-Feb",
			"17-Feb",
			"18-Feb",
			"19-Feb",
			"20-Feb",
			"21-Feb",
			"22-Feb",
			"23-Feb",
			"24-Feb",
			"25-Feb",
			"26-Feb",
			"27-Feb",
			"28-Feb",
			"29-Feb",
			"1-Mar",
			"2-Mar",
			"3-Mar",
			"4-Mar",
			"5-Mar",
			"6-Mar",
			"7-Mar",
			"8-Mar",
			"9-Mar",
			"10-Mar",
			"11-Mar",
			"12-Mar",
			"13-Mar",
			"14-Mar",
			"15-Mar",
			"16-Mar",
			"17-Mar",
			"18-Mar",
			"19-Mar",
			"20-Mar",
			"21-Mar",
			"22-Mar",
			"23-Mar",
			"24-Mar",
			"25-Mar",
			"26-Mar",
			"27-Mar",
			"28-Mar",
			"29-Mar",
			"30-Mar",
			"31-Mar",
			"1-Apr",
			"2-Apr",
			"3-Apr",
			"4-Apr",
			"5-Apr",
			"6-Apr",
			"7-Apr",
			"8-Apr",
			"9-Apr",
			"10-Apr",
			"11-Apr",
			"12-Apr",
			"13-Apr",
			"14-Apr",
			"15-Apr",
			"16-Apr",
			"17-Apr",
			"18-Apr",
			"19-Apr",
			"20-Apr",
			"21-Apr",
			"22-Apr",
			"23-Apr",
			"24-Apr",
			"25-Apr",
			"26-Apr",
			"27-Apr",
			"28-Apr",
			"29-Apr",
			"30-Apr",
			"1-May",
			"2-May",
			"3-May",
			"4-May",
			"5-May",
			"6-May",
			"7-May",
			"8-May",
			"9-May",
			"10-May",
			"11-May",
			"12-May",
			"13-May",
			"14-May",
			"15-May",
			"16-May",
			"17-May",
			"18-May",
			"19-May",
			"20-May",
			"21-May",
			"22-May",
			"23-May",
			"24-May",
			"25-May",
			"26-May",
			"27-May",
			"28-May",
			"29-May",
			"30-May",
			"31-May",
			"1-Jun",
			"2-Jun",
			"3-Jun",
			"4-Jun",
			"5-Jun",
			"6-Jun",
			"7-Jun",
			"8-Jun",
			"9-Jun",
			"10-Jun",
			"11-Jun",
			"12-Jun",
			"13-Jun",
			"14-Jun",
			"15-Jun",
			"16-Jun",
			"17-Jun",
			"18-Jun",
			"19-Jun",
			"20-Jun",
			"21-Jun",
			"22-Jun",
			"23-Jun",
			"24-Jun",
			"25-Jun",
			"26-Jun",
			"27-Jun",
			"28-Jun",
			"29-Jun",
			"30-Jun",
			"1-Jul",
			"2-Jul",
			"3-Jul",
			"4-Jul",
			"5-Jul",
			"6-Jul",
			"7-Jul",
			"8-Jul",
			"9-Jul",
			"10-Jul",
			"11-Jul",
			"12-Jul",
			"13-Jul",
			"14-Jul",
			"15-Jul",
			"16-Jul",
			"17-Jul",
			"18-Jul",
			"19-Jul",
			"20-Jul",
			"21-Jul",
			"22-Jul",
			"23-Jul",
			"24-Jul",
			"25-Jul",
			"26-Jul",
			"27-Jul",
			"28-Jul",
			"29-Jul",
			"30-Jul",
			"31-Jul",
			"1-Aug",
			"2-Aug",
			"3-Aug",
			"4-Aug",
			"5-Aug",
			"6-Aug",
			"7-Aug",
			"8-Aug",
			"9-Aug",
			"10-Aug",
			"11-Aug",
			"12-Aug",
			"13-Aug",
			"14-Aug",
			"15-Aug",
			"16-Aug",
			"17-Aug",
			"18-Aug",
			"19-Aug",
			"20-Aug",
			"21-Aug",
			"22-Aug",
			"23-Aug",
			"24-Aug",
			"25-Aug",
			"26-Aug",
			"27-Aug",
			"28-Aug",
			"29-Aug",
			"30-Aug",
			"31-Aug",
			"1-Sep",
			"2-Sep",
			"3-Sep",
			"4-Sep",
			"5-Sep",
			"6-Sep",
			"7-Sep",
			"8-Sep",
			"9-Sep",
			"10-Sep",
			"11-Sep",
			"12-Sep",
			"13-Sep",
			"14-Sep",
			"15-Sep",
			"16-Sep",
			"17-Sep",
			"18-Sep",
			"19-Sep",
			"20-Sep",
			"21-Sep",
			"22-Sep",
			"23-Sep",
			"24-Sep",
			"25-Sep",
			"26-Sep",
			"27-Sep",
			"28-Sep",
			"29-Sep",
			"30-Sep",
			"1-Oct",
			"2-Oct",
			"3-Oct",
			"4-Oct",
			"5-Oct",
			"6-Oct",
			"7-Oct",
			"8-Oct",
			"9-Oct",
			"10-Oct",
			"11-Oct",
			"12-Oct",
			"13-Oct",
			"14-Oct",
			"15-Oct",
			"16-Oct",
			"17-Oct",
			"18-Oct",
			"19-Oct",
			"20-Oct",
			"21-Oct",
			"22-Oct",
			"23-Oct",
			"24-Oct",
			"25-Oct",
			"26-Oct",
			"27-Oct",
			"28-Oct",
			"29-Oct",
			"30-Oct",
			"31-Oct",
			"1-Nov",
			"2-Nov",
			"3-Nov",
			"4-Nov",
			"5-Nov",
			"6-Nov",
			"7-Nov",
			"8-Nov",
			"9-Nov",
			"10-Nov",
			"11-Nov",
			"12-Nov",
			"13-Nov",
			"14-Nov",
			"15-Nov",
			"16-Nov",
			"17-Nov",
			"18-Nov",
			"19-Nov",
			"20-Nov",
			"21-Nov",
			"22-Nov",
			"23-Nov",
			"24-Nov",
			"25-Nov",
			"26-Nov",
			"27-Nov",
			"28-Nov",
			"29-Nov",
			"30-Nov",
			"1-Dec",
			"2-Dec",
			"3-Dec",
			"4-Dec",
			"5-Dec",
			"6-Dec",
			"7-Dec",
			"8-Dec",
			"9-Dec",
			"10-Dec",
			"11-Dec",
			"12-Dec",
			"13-Dec",
			"14-Dec",
			"15-Dec",
			"16-Dec",
			"17-Dec",
			"18-Dec",
			"19-Dec",
			"20-Dec",
			"21-Dec",
			"22-Dec",
			"23-Dec",
			"24-Dec",
			"25-Dec",
			"26-Dec",
			"27-Dec",
			"28-Dec",
			"29-Dec",
			"30-Dec",
			"31-Dec"
		};

		private static readonly IReadOnlyList<string> Months = new List<string>
		{
			"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
		};

		private static readonly IReadOnlyList<string> Quarters = new List<string> { "Qtr1", "Qtr2", "Qtr3", "Qtr4" };

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
		/// Gets a value indicating whether to save pivot cache records.
		/// </summary>
		public bool SaveData
		{
			get { return base.GetXmlNodeBool("@saveData", true); }
		}

		/// <summary>
		/// Gets a value indicating the number of items to retain per field.
		/// </summary>
		public int? MissingItemLimit
		{
			get { return base.GetXmlNodeIntNull("@missingItemsLimit"); }
		}

		/// <summary>
		/// Gets or sets a value indicating whether or not the pivot table should refresh on load.
		/// </summary>
		public bool RefreshOnLoad
		{
			get { return base.GetXmlNodeBool("@refreshOnLoad"); }
			set { base.SetXmlNodeBool("@refreshOnLoad", value); }
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
							var partUri = new Uri($"xl/pivotCache/{UriHelper.GetUriEndTargetName(cacheRecordsRel.TargetUri)}", UriKind.Relative);
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
			get { return base.GetXmlNodeString("@r:id"); }
			set { base.SetXmlNodeString("@r:id", value); }
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
				base.SetXmlNodeString(ExcelPivotCacheDefinition.SourceWorksheetPath, value.Worksheet.Name);
				base.SetXmlNodeString(ExcelPivotCacheDefinition.SourceAddressPath, value.FirstAddress);

				// Delete the worksheetSource "name" attribute if it exists.
				if (base.GetXmlNodeString(ExcelPivotCacheDefinition.SourceNamePath) != null)
					base.DeleteNode(ExcelPivotCacheDefinition.SourceNamePath);

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
			var uri = UriHelper.ResolvePartUri(this.CacheDefinitionUri, this.CacheRecords.Uri);
			this.RecordRelationship = this.Part.CreateRelationship(uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
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
			this.Workbook.FormulaParser.Logger?.LogFunction(nameof(this.UpdateData));
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

			// Clear out all of the shared items in order to update them from the source data.
			foreach (var cacheField in this.CacheFields)
			{
				if (cacheField.HasSharedItems)
				{
					// Clear out the shared items, but keep the @minDate attribute in order
					// to correctly parse new values into the shared items.
					var containsDate = cacheField.SharedItems.ContainsDate;
					cacheField.SharedItems.Clear();
					cacheField.SharedItems.ContainsDate = containsDate;
					cacheField.ClearedSharedItems = true;
				}
			}

			this.CacheRecords.UpdateRecords(range, this.Workbook.FormulaParser.Logger);

			this.UpdateCacheFieldSharedItemMetadata();

			this.StringResources.LoadResourceManager(resourceManager);

			this.UpdateCacheFieldFieldGroups();

			// Refresh pivot tables.
			foreach (var pivotTable in this.GetRelatedPivotTables())
			{
				pivotTable.RefreshFromCache(this.StringResources);
			}

			// Remove the 'u' xml attribute from each cache item to prevent corrupting the workbook, since Excel automatically adds them.
			foreach (var cacheField in this.CacheFields)
			{
				if (cacheField.FieldGroup != null || cacheField.HasSharedItems)
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

		/// <summary>
		/// Gets a list of enabled pivot table features that are unsupported.
		/// </summary>
		/// <param name="unsupportedFeatures">A list of unsuppported features.</param>
		/// <returns>True if unsupported features were present, otherwise false.</returns>
		public bool TryGetUnsupportedFeatures(out List<string> unsupportedFeatures)
		{
			unsupportedFeatures = new List<string>();
			foreach (var pivotTable in this.GetRelatedPivotTables())
			{
				if (pivotTable.TryGetUnsupportedFeatures(out var pivotTableUnsupportedFeatures))
					unsupportedFeatures.AddRange(pivotTableUnsupportedFeatures);
			}
			if (this.Workbook.SlicerCaches.Any())
				unsupportedFeatures.Add("Slicer present");
			if (base.TopNode.SelectSingleNode("d:calculatedItems", base.NameSpaceManager) != null)
				unsupportedFeatures.Add("Calculated items present");
			if (!this.SaveData)
				unsupportedFeatures.Add("Save source data with file disabled");
			if (this.MissingItemLimit != null)
				unsupportedFeatures.Add("Missing item limit set");
			if (this.CacheSource != eSourceType.Worksheet)
				unsupportedFeatures.Add($"Unsupported pivot cache source type {this.CacheSource}");
			return unsupportedFeatures.Any();
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

		private void UpdateCacheFieldSharedItemMetadata()
		{
			// See docs:
			// https://web.archive.org/web/20190412204433/https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_sharedItems_topic_ID0EGBCBB.html
			// Update shared item metadata.
			foreach (var cacheField in this.CacheFields)
			{
				if (cacheField.HasSharedItems)
				{
					bool hasDate = false, hasString = false, hasBlank = false, hasNumbers = false, hasInteger = false, 
						hasNonIntegerNumber = false, hasBool = false, hasError = false, hasLongText = false;

					DateTime? minDate = null, maxDate = null;
					foreach (var sharedItem in cacheField.SharedItems)
					{
						if (sharedItem.Type == PivotCacheRecordType.d)
						{
							hasDate = true;
							var dateValue = DateTime.Parse(sharedItem.Value);
							if (minDate == null || dateValue < minDate)
								minDate = dateValue;
							else if (maxDate == null || dateValue > maxDate)
								maxDate = dateValue;
						}
						else if (sharedItem.Type == PivotCacheRecordType.b)
							hasBool = true;
						else if (sharedItem.Type == PivotCacheRecordType.e)
							hasError = true;
						else if (sharedItem.Type == PivotCacheRecordType.m)
							hasBlank = true;
						else if (sharedItem.Type == PivotCacheRecordType.n)
						{
							if (int.TryParse(sharedItem.Value, out _))
								hasInteger = true;
							else
								hasNonIntegerNumber = true;
							hasNumbers = true;
						}
						else if (sharedItem.Type == PivotCacheRecordType.s)
						{
							hasString = true;
							if (sharedItem.Value.Length > 255)
								hasLongText = true;
						}
					}

					// Note that this differs from the documentation, which we believe to be wrong (the false case appears to be worded incorrectly).
					cacheField.SharedItems.ContainsNonDate = (hasString || hasNumbers || hasInteger || hasBool || hasError);

					if (hasDate)
					{
						cacheField.SharedItems.ContainsDate = hasDate;
						cacheField.SharedItems.MinDate = minDate;
						if (maxDate == null)
							maxDate = minDate;

						// If the max date value is a date with no time component (midnight), Excel sets the max date to the next day.
						if (maxDate != null && maxDate.Value.TimeOfDay.TotalMilliseconds == 0)
							maxDate = maxDate.Value.AddDays(1);
						cacheField.SharedItems.MaxDate = maxDate;
					}

					if (hasInteger)
					{
						if (hasDate || hasString || hasBlank || hasNonIntegerNumber || hasBool || hasError)
							cacheField.SharedItems.ContainsInteger = false;  // Indicates non-integer or mixed values.
						else
							cacheField.SharedItems.ContainsInteger = true;  // Indicates strictly integer values.
					}

					// These four fields are only validated if there are shared items.
					// Indicates this field contains more than one data type.
					cacheField.SharedItems.ContainsMixedTypes = this.ContainsMixedDataTypes(hasDate, hasString, hasNumbers, hasBool, hasError);
					cacheField.SharedItems.ContainsBlank = hasBlank;
					cacheField.SharedItems.ContainsSemiMixedTypes = hasString || hasBool || hasBlank;  // Excel appears to think bools are strings.
					cacheField.SharedItems.LongText = hasLongText;

					// "...not validated unless there is more than one item in sharedItems or the one and only item is not a blank item.
					//  If the first item is a blank item the data type the field cannot be verified."
					if (cacheField.SharedItems.Count > 0 || (cacheField.SharedItems.Count == 1 && cacheField.SharedItems.First().Type != PivotCacheRecordType.m))
					{
						cacheField.SharedItems.ContainsNumbers = hasNumbers;
						cacheField.SharedItems.ContainsString = hasString || hasBool;  // Excel appears to think bools are strings.
						// ContainsDate and ContainsInteger are set above.
					}
					else
					{
						// Set to defaults if there are not multiple shared items.
						cacheField.SharedItems.ContainsNumbers = false;
						cacheField.SharedItems.ContainsDate = false;
						cacheField.SharedItems.ContainsString = true;
						cacheField.SharedItems.ContainsInteger = false;
					}
				}
				else
				{
					// NOTE: If there are no shared items but the field contains numeric values, the @minValue and @maxValue
					// attributes should be set here. 
					// Because we do not currently use them and incorrect values do not cause corruptions, it is being left for a later date.
				}
			}
		}

		private bool ContainsMixedDataTypes(params bool[] boolArray)
		{
			return boolArray.Count(b => b) > 1;
		}

		private void UpdateCacheFieldFieldGroups()
		{
			foreach (var cacheField in this.CacheFields)
			{
				if (cacheField.IsGroupField)
				{
					var fieldGroup = cacheField.FieldGroup;

					var baseField = this.CacheFields[fieldGroup.BaseField];
					if (baseField.SharedItems.ContainsDate)
					{
						// Update fieldGroup.RangePr startDate and endDate.
						var startDate = (fieldGroup.RangePr.StartDate = baseField.SharedItems.MinDate).Value;
						var endDate = (fieldGroup.RangePr.EndDate = baseField.SharedItems.MaxDate).Value;
						if (fieldGroup.GroupBy != PivotFieldDateGrouping.None)
						{
							fieldGroup.GroupItems.Clear();

							// Each grouping has a start date as the first element in the form "<11/1/2018".
							fieldGroup.GroupItems.Add($"<{startDate.ToShortDateString()}");

							// Add all of the elements that could possibly make up the grouping.
							if (fieldGroup.GroupBy == PivotFieldDateGrouping.Years)
							{
								Enumerable.Range(startDate.Year, (endDate.Year - startDate.Year) + 1).ToList()
									.ForEach(y => fieldGroup.GroupItems.Add(y.ToString()));
							} 
							else if (fieldGroup.GroupBy == PivotFieldDateGrouping.Quarters)
							{
								// TODO: Quarters should be translated. See bug #13353.
								ExcelPivotCacheDefinition.Quarters.ForEach(q => fieldGroup.GroupItems.Add(q));
							}
							else if (fieldGroup.GroupBy == PivotFieldDateGrouping.Months)
								ExcelPivotCacheDefinition.Months.ForEach(m => fieldGroup.GroupItems.Add(m));
							else if(fieldGroup.GroupBy == PivotFieldDateGrouping.Days)
								ExcelPivotCacheDefinition.Days.ForEach(d => fieldGroup.GroupItems.Add(d));
							else if (fieldGroup.GroupBy == PivotFieldDateGrouping.Hours)
								ExcelPivotCacheDefinition.Hours.ForEach(h => fieldGroup.GroupItems.Add(h));
							else if (fieldGroup.GroupBy == PivotFieldDateGrouping.Minutes)
								ExcelPivotCacheDefinition.SecondsMinutes.ForEach(m => fieldGroup.GroupItems.Add(m));
							else if (fieldGroup.GroupBy == PivotFieldDateGrouping.Seconds)
								ExcelPivotCacheDefinition.SecondsMinutes.ForEach(s => fieldGroup.GroupItems.Add(s));

							// Each grouping has an end date as the last element in the form ">11/1/2018".
							fieldGroup.GroupItems.Add($">{endDate.ToShortDateString()}");
						}
					}
				}
			}
		}
		#endregion
	}
}
