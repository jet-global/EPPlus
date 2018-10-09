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
 * Jan Källman		    Initial Release		       2011-01-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 * Richard Tallent		Fix escaping of quotes					2012-10-31
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Slicers;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml
{
	#region Enums
	/// <summary>
	/// How the application should calculate formulas in the workbook
	/// </summary>
	public enum ExcelCalcMode
	{
		/// <summary>
		/// Indicates that calculations in the workbook are performed automatically when cell values change.
		/// The application recalculates those cells that are dependent on other cells that contain changed values.
		/// This mode of calculation helps to avoid unnecessary calculations.
		/// </summary>
		Automatic,
		/// <summary>
		/// Indicates tables be excluded during automatic calculation
		/// </summary>
		AutomaticNoTable,
		/// <summary>
		/// Indicates that calculations in the workbook be triggered manually by the user.
		/// </summary>
		Manual
	}
	#endregion

	/// <summary>
	/// Represents the Excel workbook and provides access to all the
	/// document properties and worksheets within the workbook.
	/// </summary>
	public sealed class ExcelWorkbook : XmlHelper, IDisposable
	{
		#region Constants
		private const string codeModuleNamePath = "d:workbookPr/@codeName";
		private const string FULL_CALC_ON_LOAD_PATH = "d:calcPr/@fullCalcOnLoad";
		private const string CALC_MODE_PATH = "d:calcPr/@calcMode";
		internal const double date1904Offset = 365.5 * 4;  // offset to fix 1900 and 1904 differences, 4 OLE years
		private const string date1904Path = "d:workbookPr/@date1904";
		#endregion

		#region Nested Classes
		internal class SharedStringItem
		{
			internal int pos;
			internal string Text;
			internal bool isRichText = false;
		}
		#endregion

		#region Class Variables
		private ExcelWorksheets _worksheets;
		private OfficeProperties _properties;
		private ExcelStyles _styles;
		private FormulaParser _formulaParser = null;
		private FormulaParserManager _parserManager;
		private List<ExcelSlicerCache> mySlicerCaches;
		private List<ExcelPivotCacheDefinition> myPivotCacheDefinitions;
		private decimal _standardFontWidth = decimal.MinValue;
		private string _fontID = "";
		private ExcelProtection _protection = null;
		private ExcelWorkbookView _view = null;
		private ExcelVbaProject _vba = null;
		private XmlDocument _workbookXml;
		private bool? date1904Cache = null;
		private XmlDocument _stylesXml;
		#endregion

		#region Public Properties
		/// <summary>
		/// Gets a list of <see cref="ExcelPivotCacheDefinition"/>.
		/// </summary>
		public List<ExcelPivotCacheDefinition> PivotCacheDefinitions
		{
			get
			{
				if (myPivotCacheDefinitions == null)
				{
					myPivotCacheDefinitions = new List<ExcelPivotCacheDefinition>();
					var cacheDefinitions = this.Part.GetRelationshipsByType(ExcelPackage.schemaPivotCacheRelationship);
					foreach (var cache in cacheDefinitions)
					{
						var pivotCacheTargetUri = $"xl/pivotCache/{UriHelper.GetUriEndTargetName(cache.TargetUri)}";
						var uri = new Uri(pivotCacheTargetUri, UriKind.Relative);
						var possiblePart = this.Package.GetXmlFromUri(uri);
						myPivotCacheDefinitions.Add(new ExcelPivotCacheDefinition(this.NameSpaceManager, this.Package, possiblePart, uri));
					}
				}
				return myPivotCacheDefinitions;
			}
		}

		/// <summary>
		/// Gets a list of the slicer caches present in this workbook.
		/// </summary>
		public List<ExcelSlicerCache> SlicerCaches
		{
			get
			{
				if (mySlicerCaches == null)
				{
					var slicerCacheNamespaceManager = ExcelSlicer.SlicerDocumentNamespaceManager;
					mySlicerCaches = new List<ExcelSlicerCache>();
					var slicerCaches = this.Part.GetRelationshipsByType(ExcelPackage.schemaSlicerCache);
					foreach (var cache in slicerCaches)
					{
						var cacheTargetUri = cache.TargetUri.ToString();
						var uri = new Uri($"/xl/{cacheTargetUri}", UriKind.Relative);
						var possiblePart = this.Package.GetXmlFromUri(uri);
						var slicerCacheNode = possiblePart.SelectSingleNode("default:slicerCacheDefinition", slicerCacheNamespaceManager);
						mySlicerCaches.Add(new ExcelSlicerCache(slicerCacheNode, slicerCacheNamespaceManager, cache.TargetUri, possiblePart));
					}
				}
				return mySlicerCaches;
			}
		}

		/// <summary>
		/// Provides access to all the worksheets in the workbook.
		/// </summary>
		public ExcelWorksheets Worksheets
		{
			get
			{
				if (this._worksheets == null)
				{
					var sheetsNode = this._workbookXml.DocumentElement.SelectSingleNode("d:sheets", this.NameSpaceManager);
					if (sheetsNode == null)
					{
						sheetsNode = this.CreateNode("d:sheets");
					}

					_worksheets = new ExcelWorksheets(Package, this.NameSpaceManager, sheetsNode);
				}
				return (this._worksheets);
			}
		}

		/// <summary>
		/// Provides access to named ranges
		/// </summary>
		public ExcelNamedRangeCollection Names { get; }

		/// <summary>
		/// Gets the <see cref="FormulaParserManager"/> to use when parsing formulas in this workbook.
		/// </summary>
		public FormulaParserManager FormulaParserManager
		{
			get
			{
				if (this._parserManager == null)
				{
					_parserManager = new FormulaParserManager(this.FormulaParser);
				}
				return this._parserManager;
			}
		}

		/// <summary>
		/// Max font width for the workbook
		/// <remarks>This method uses GDI. If you use Asure or another environment that does not support GDI, you have to set this value manually if you don't use the standard Calibri font</remarks>
		/// </summary>
		public decimal MaxFontWidth
		{
			get
			{
				var ix = Styles.NamedStyles.FindIndexByID("Normal");
				if (ix >= 0)
				{
					if (this._standardFontWidth == decimal.MinValue || this._fontID != this.Styles.NamedStyles[ix].Style.Font.Id)
					{
						var font = this.Styles.NamedStyles[ix].Style.Font;
						try
						{
							this._standardFontWidth = ExcelWorkbook.GetWidthPixels(font.Name, font.Size);
							this._fontID = this.Styles.NamedStyles[ix].Style.Font.Id;
						}
						catch   //Error, Font missing and Calibri removed in dictionary
						{
							this._standardFontWidth = (int)(font.Size * (2D / 3D)); //Aprox for Calibri.
						}
					}
				}
				else
				{
					this._standardFontWidth = 7; //Calibri 11
				}
				return this._standardFontWidth;
			}
			set
			{
				this._standardFontWidth = value;
			}
		}

		/// <summary>
		/// Access properties to protect or unprotect a workbook
		/// </summary>
		public ExcelProtection Protection
		{
			get
			{
				if (this._protection == null)
				{
					this._protection = new ExcelProtection(this.NameSpaceManager, this.TopNode, this);
					this._protection.SchemaNodeOrder = this.SchemaNodeOrder;
				}
				return this._protection;
			}
		}

		/// <summary>
		/// Access to workbook view properties
		/// </summary>
		public ExcelWorkbookView View
		{
			get
			{
				if (this._view == null)
				{
					this._view = new ExcelWorkbookView(this.NameSpaceManager, this.TopNode, this);
				}
				return this._view;
			}
		}

		/// <summary>
		/// A reference to the VBA project.
		/// Null if no project exists.
		/// Use Workbook.CreateVBAProject to create a new VBA-Project
		/// </summary>
		public ExcelVbaProject VbaProject
		{
			get
			{
				if (this._vba == null)
				{
					if (this.Package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
					{
						this._vba = new ExcelVbaProject(this);
					}
				}
				return this._vba;
			}
		}

		/// <summary>
		/// Provides access to the XML data representing the workbook in the package.
		/// </summary>
		public XmlDocument WorkbookXml
		{
			get
			{
				if (this._workbookXml == null)
				{
					this.CreateWorkbookXml(this.NameSpaceManager);
				}
				return (this._workbookXml);
			}
		}

		/// <summary>
		/// Gets the VBA Code Module contained in this workbook.
		/// </summary>
		public ExcelVBAModule CodeModule
		{
			get
			{
				if (this.VbaProject != null)
				{
					return this.VbaProject.Modules[CodeModuleName];
				}
				else
				{
					return null;
				}
			}
		}

		/// <summary>
		/// The date systems used by Microsoft Excel can be based on one of two different dates. By default, a serial number of 1 in Microsoft Excel represents January 1, 1900.
		/// The default for the serial number 1 can be changed to represent January 2, 1904.
		/// This option was included in Microsoft Excel for Windows to make it compatible with Excel for the Macintosh, which defaults to January 2, 1904.
		/// </summary>
		public bool Date1904
		{
			get
			{
				//return GetXmlNodeBool(date1904Path, false);
				if (this.date1904Cache == null)
				{
					this.date1904Cache = this.GetXmlNodeBool(ExcelWorkbook.date1904Path, false);
				}
				return this.date1904Cache.Value;
			}
			set
			{
				if (this.Date1904 != value)
				{
					// Like Excel when the option it's changed update it all cells with Date format
					foreach (var item in this.Worksheets)
					{
						item.UpdateCellsWithDate1904Setting();
					}
				}
				this.date1904Cache = value;
				this.SetXmlNodeBool(ExcelWorkbook.date1904Path, value, false);
			}
		}

		/// <summary>
		/// Provides access to the XML data representing the styles in the package.
		/// </summary>
		public XmlDocument StylesXml
		{
			get
			{
				if (this._stylesXml == null)
				{
					if (this.Package.Package.PartExists(StylesUri))
						this._stylesXml = this.Package.GetXmlFromUri(StylesUri);
					else
					{
						// create a new styles part and add to the package
						Packaging.ZipPackagePart part = this.Package.Package.CreatePart(StylesUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", Package.Compression);
						// create the style sheet

						StringBuilder xml = new StringBuilder("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
						xml.Append("<numFmts />");
						xml.Append("<fonts count=\"1\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts>");
						xml.Append("<fills><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills>");
						xml.Append("<borders><border><left /><right /><top /><bottom /><diagonal /></border></borders>");
						xml.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs>");
						xml.Append("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" xfId=\"0\" /></cellXfs>");
						xml.Append("<cellStyles><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles>");
						xml.Append("<dxfs count=\"0\" />");
						xml.Append("</styleSheet>");

						this._stylesXml = new XmlDocument();
						this._stylesXml.LoadXml(xml.ToString());

						//Save it to the package
						StreamWriter stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));

						this._stylesXml.Save(stream);
						//stream.Close();
						this.Package.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						this.Package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(this.WorkbookUri, this.StylesUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/styles");
						this.Package.Package.Flush();
					}
				}
				return (this._stylesXml);
			}
			set
			{
				this._stylesXml = value;
			}
		}

		/// <summary>
		/// Package styles collection. Used internally to access style data.
		/// </summary>
		public ExcelStyles Styles
		{
			get
			{
				if (this._styles == null)
				{
					this._styles = new ExcelStyles(this.NameSpaceManager, this.StylesXml, this);
				}
				return this._styles;
			}
		}

		/// <summary>
		/// The office document properties
		/// </summary>
		public OfficeProperties Properties
		{
			get
			{
				if (this._properties == null)
				{
					//  Create a NamespaceManager to handle the default namespace,
					//  and create a prefix for the default namespace:
					this._properties = new OfficeProperties(this.Package, this.NameSpaceManager);
				}
				return this._properties;
			}
		}

		/// <summary>
		/// Should Excel do a full calculation after the workbook has been loaded?
		/// <remarks>This property is always true for both new workbooks and loaded templates(on load). If this is not the wanted behavior set this property to false.</remarks>
		/// </summary>
		public bool FullCalcOnLoad
		{
			get
			{
				return this.GetXmlNodeBool(FULL_CALC_ON_LOAD_PATH);
			}
			set
			{
				this.SetXmlNodeBool(FULL_CALC_ON_LOAD_PATH, value);
			}
		}

		/// <summary>
		/// Calculation mode for the workbook.
		/// </summary>
		public ExcelCalcMode CalcMode
		{
			get
			{
				string calcMode = this.GetXmlNodeString(CALC_MODE_PATH);
				switch (calcMode)
				{
					case "autoNoTable":
						return ExcelCalcMode.AutomaticNoTable;
					case "manual":
						return ExcelCalcMode.Manual;
					default:
						return ExcelCalcMode.Automatic;

				}
			}
			set
			{
				switch (value)
				{
					case ExcelCalcMode.AutomaticNoTable:
						this.SetXmlNodeString(CALC_MODE_PATH, "autoNoTable");
						break;
					case ExcelCalcMode.Manual:
						this.SetXmlNodeString(CALC_MODE_PATH, "manual");
						break;
					default:
						this.SetXmlNodeString(CALC_MODE_PATH, "auto");
						break;

				}
			}
		}

		/// <summary>
		/// Gets the collection of external references in the workbook.
		/// </summary>
		public ExternalReferenceCollection ExternalReferences { get; }
		#endregion

		#region Internal Properties
		/// <summary>
		/// Gets or sets a cellstore containing the formula tokens on the workbook.
		/// </summary>
		internal ICellStore<List<Token>> FormulaTokens { get; set; }

		/// <summary>
		/// Gets or sets the next ID number to use for a drawing.
		/// </summary>
		internal int NextDrawingID { get; set; } = 0;

		/// <summary>
		/// Gets or sets the next ID number to use for a table.
		/// </summary>
		internal int NextTableID { get; set; } = int.MinValue;

		/// <summary>
		/// Gets or sets the next ID number to use for a PivotTable.
		/// </summary>
		internal int NextPivotTableID { get; set; } = int.MinValue;

		/// <summary>
		/// Gets or sets the dictionary containing shared string information.
		/// </summary>
		internal Dictionary<string, SharedStringItem> SharedStrings { get; set; } = new Dictionary<string, SharedStringItem>(); //Used when reading cells.

		/// <summary>
		/// Gets or sets the list of shared string information. Every element in this list should also be in the SharedStrings dictionary.
		/// </summary>
		internal List<SharedStringItem> SharedStringsList { get; set; } = new List<SharedStringItem>(); //Used when reading cells.
		
		/// <summary>
		/// Gets the <see cref="ExcelPackage"/> that this workbook belongs to.
		/// </summary>
		internal ExcelPackage Package { get; set; }

		/// <summary>
		/// Gets a dictionary that stores the next unique ID for each slicer in the workbook.
		/// </summary>
		internal Dictionary<string, int> NextSlicerIdNumber { get; } = new Dictionary<string, int>();

		/// <summary>
		/// Gets the <see cref="FormulaParser"/> to use when parsing formulas in the workbook.
		/// </summary>
		internal FormulaParser FormulaParser
		{
			get
			{
				if (this._formulaParser == null)
				{
					this._formulaParser = new FormulaParser(new EpplusExcelDataProvider(Package));
				}
				return this._formulaParser;
			}
		}

		/// <summary>
		/// URI to the workbook inside the package
		/// </summary>
		internal Uri WorkbookUri { get; private set; }

		/// <summary>
		/// URI to the styles inside the package
		/// </summary>
		internal Uri StylesUri { get; private set; }

		/// <summary>
		/// URI to the shared strings inside the package
		/// </summary>
		internal Uri SharedStringsUri { get; private set; }

		/// <summary>
		/// Returns a reference to the workbook's part within the package
		/// </summary>
		internal Packaging.ZipPackagePart Part { get { return (this.Package.Package.GetPart(WorkbookUri)); } }

		/// <summary>
		/// Gets the name of this workbook's Code Module.
		/// </summary>
		internal string CodeModuleName
		{
			get
			{
				return this.GetXmlNodeString(codeModuleNamePath);
			}
			set
			{
				this.SetXmlNodeString(codeModuleNamePath, value);
			}
		}
		#endregion
		
		#region Constructors
		/// <summary>
		/// Creates a new instance of the ExcelWorkbook class.
		/// </summary>
		/// <param name="package">The parent package</param>
		/// <param name="namespaceManager">NamespaceManager</param>
		internal ExcelWorkbook(ExcelPackage package, XmlNamespaceManager namespaceManager) :
			base(namespaceManager)
		{
			this.Package = package;
			this.WorkbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
			this.SharedStringsUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
			this.StylesUri = new Uri("/xl/styles.xml", UriKind.Relative);

			this.Names = new ExcelNamedRangeCollection(this);
			this.NameSpaceManager = namespaceManager;
			this.TopNode = this.WorkbookXml.DocumentElement;
			this.SchemaNodeOrder = new string[] { "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "functionPrototypes", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", };
			this.FullCalcOnLoad = true;  //Full calculation on load by default, for both new workbooks and templates.
			this.GetSharedStrings();
			var node = this.WorkbookXml.SelectSingleNode("//d:externalReferences", this.NameSpaceManager);
			if (node != null)
				this.ExternalReferences = new ExternalReferenceCollection(this.ResolveExternalReference, node, this.NameSpaceManager);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Create an empty VBA project.
		/// </summary>
		public void CreateVBAProject()
		{
#if !MONO
			if (this._vba != null || this.Package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
			{
				throw (new InvalidOperationException("VBA project already exists."));
			}

			this._vba = new ExcelVbaProject(this);
			this._vba.Create();
#endif
#if MONO
            throw new NotSupportedException("Creating a VBA project is not supported under Mono.");
#endif
		}

		public void Dispose()
		{
			if (this.SharedStrings != null)
			{
				this.SharedStrings.Clear();
				this.SharedStrings = null;
			}
			if (this.SharedStringsList != null)
			{
				this.SharedStringsList.Clear();
				this.SharedStringsList = null;
			}
			this._vba = null;
			if (this._worksheets != null)
			{
				this._worksheets.Dispose();
				this._worksheets = null;
			}
			this.Package = null;
			this._properties = null;
			if (this._formulaParser != null)
			{
				this._formulaParser.Dispose();
				this._formulaParser = null;
			}
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Loads the defined names from the workbook XML.
		/// </summary>
		internal void GetDefinedNames()
		{
			XmlNodeList nodeList = this.WorkbookXml.SelectNodes("//d:definedNames/d:definedName", this.NameSpaceManager);
			if (nodeList != null)
			{
				foreach (XmlElement elem in nodeList)
				{
					string nameFormula = elem.InnerText;
					string comment = elem.GetAttribute("comment");
					bool isHidden = elem.GetAttribute("hidden") == "1";
					if (int.TryParse(elem.GetAttribute("localSheetId"), out int localSheetID))
						this.Worksheets[localSheetID + 1].Names.Add(elem.GetAttribute("name"), nameFormula, isHidden, comment);
					else
						this.Names.Add(elem.GetAttribute("name"), nameFormula, isHidden, comment);
				}
			}
		}

		internal void CodeNameChange(string value)
		{
			CodeModuleName = value;
		}

		/// <summary>
		/// Saves the workbook and all its components to the package.
		/// For internal use only!
		/// </summary>
		internal void Save()  // Workbook Save
		{
			if (this.Worksheets.Count == 0)
				throw new InvalidOperationException("The workbook must contain at least one worksheet");

			this.DeleteCalcChain();

			// An empty <externalReferences /> node corrupts the workbook.
			if (this.ExternalReferences?.References.Count == 0)
			{
				var document = this.ExternalReferences.TopNode.ParentNode;
				document.RemoveChild(this.ExternalReferences.TopNode);
			}

			if (this._vba == null && !this.Package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
			{
				if (this.Part.ContentType != ExcelPackage.contentTypeWorkbookDefault)
				{
					this.Part.ContentType = ExcelPackage.contentTypeWorkbookDefault;
				}
			}
			else
			{
				if (this.Part.ContentType != ExcelPackage.contentTypeWorkbookMacroEnabled)
				{
					this.Part.ContentType = ExcelPackage.contentTypeWorkbookMacroEnabled;
				}
			}

			this.UpdateDefinedNamesXml();

			// save the workbook
			if (this._workbookXml != null)
			{
				this.Package.SavePart(WorkbookUri, _workbookXml);
			}

			// Save any slicer caches
			foreach (var slicerCache in this.SlicerCaches)
			{
				slicerCache.Save(this.Package);
			}

			// Save all the pivotCacheDefinitions
			foreach (var pivotCacheDefinition in this.PivotCacheDefinitions)
			{
				pivotCacheDefinition.Save();
			}

			// save the properties of the workbook
			if (this._properties != null)
			{
				this._properties.Save();
			}

			// save the style sheet
			this.Styles.UpdateXml();
			this.Package.SavePart(StylesUri, _stylesXml);

			// save all the open worksheets
			var isProtected = this.Protection.LockWindows || this.Protection.LockStructure;
			foreach (ExcelWorksheet worksheet in Worksheets)
			{
				if (isProtected && this.Protection.LockWindows)
				{
					worksheet.View.WindowProtection = true;
				}
				worksheet.Save();
				worksheet.Part.SaveHandler = worksheet.SaveHandler;
			}

			// Issue 15252: save SharedStrings only once
			Packaging.ZipPackagePart part;
			if (this.Package.Package.PartExists(SharedStringsUri))
			{
				part = this.Package.Package.GetPart(SharedStringsUri);
			}
			else
			{
				part = this.Package.Package.CreatePart(SharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", this.Package.Compression);
				this.Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, SharedStringsUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/sharedStrings");
			}

			part.SaveHandler = this.SaveSharedStringHandler;

			this.ValidateDataValidations();

			//VBA
			if (this._vba != null)
			{
#if !MONO
				this.VbaProject.Save();
#endif
			}

		}

		/// <summary>
		/// Determine if a table with the specified <paramref name="name"/> exists.
		/// </summary>
		/// <param name="name">The table name to check for.</param>
		/// <returns>True if a table with the name exists in the workbook.</returns>
		internal bool ExistsTableName(string name)
		{
			foreach (var ws in this.Worksheets)
			{
				if (ws.Tables.TableNames.ContainsKey(name))
				{
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// Gets the table with the given <paramref name="name"/>.
		/// </summary>
		/// <param name="name">The name of the table to retrieve.</param>
		/// <returns>The table if it was found; otherwise null.</returns>
		internal ExcelTable GetTable(string name)
		{
			foreach (var ws in this.Worksheets)
			{
				if (ws.Tables.TableNames.ContainsKey(name))
				{
					return ws.Tables[name];
				}
			}
			return null;
		}

		/// <summary>
		/// Determine if a PivotTable with the specified <paramref name="name"/> exists.
		/// </summary>
		/// <param name="name">The table name to check for.</param>
		/// <returns>True if a table with the name exists in the workbook.</returns>
		internal bool ExistsPivotTableName(string name)
		{
			foreach (var ws in this.Worksheets)
			{
				if (ws.PivotTables.myPivotTableNames.ContainsKey(name))
				{
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// Add a pivotTable to the workbook.
		/// </summary>
		/// <param name="cacheID">The PivotTable's CacheID.</param>
		/// <param name="defUri">The URI of the PivotTable.</param>
		internal void AddPivotTable(string cacheID, Uri defUri)
		{
			this.CreateNode("d:pivotCaches");

			XmlElement item = this.WorkbookXml.CreateElement("pivotCache", ExcelPackage.schemaMain);
			item.SetAttribute("cacheId", cacheID);
			var rel = this.Part.CreateRelationship(UriHelper.ResolvePartUri(WorkbookUri, defUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
			item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

			var pivotCaches = this.WorkbookXml.SelectSingleNode("//d:pivotCaches", NameSpaceManager);
			pivotCaches.AppendChild(item);
		}
		
		/// <summary>
		/// Determine the next Table and PivotTable ID numbers.
		/// </summary>
		internal void ReadAllTables()
		{
			if (this.NextTableID > 0)
				return;
			this.NextTableID = 1;
			this.NextPivotTableID = 1;
			foreach (var ws in this.Worksheets)
			{
				if (!(ws is ExcelChartsheet)) //Fixes 15273. Chartsheets should be ignored.
				{
					foreach (var tbl in ws.Tables)
					{
						if (tbl.Id >= this.NextTableID)
						{
							this.NextTableID = tbl.Id + 1;
						}
					}
					foreach (var pt in ws.PivotTables)
					{
						if (pt.CacheID >= this.NextPivotTableID)
						{
							this.NextPivotTableID = pt.CacheID + 1;
						}
					}
				}
			}
		}
		#endregion

		#region Private Methods
		/// <summary>
		/// Create or read the XML for the workbook.
		/// </summary>
		private void CreateWorkbookXml(XmlNamespaceManager namespaceManager)
		{
			if (this.Package.Package.PartExists(WorkbookUri))
				this._workbookXml = this.Package.GetXmlFromUri(WorkbookUri);
			else
			{
				// create a new workbook part and add to the package
				Packaging.ZipPackagePart partWorkbook = this.Package.Package.CreatePart(WorkbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", this.Package.Compression);

				// create the workbook
				this._workbookXml = new XmlDocument(namespaceManager.NameTable);

				this._workbookXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
				// create the workbook element
				XmlElement wbElem = this._workbookXml.CreateElement("workbook", ExcelPackage.schemaMain);

				// Add the relationships namespace
				wbElem.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);

				this._workbookXml.AppendChild(wbElem);

				// create the bookViews and workbooks element
				XmlElement bookViews = this._workbookXml.CreateElement("bookViews", ExcelPackage.schemaMain);
				wbElem.AppendChild(bookViews);
				XmlElement workbookView = this._workbookXml.CreateElement("workbookView", ExcelPackage.schemaMain);
				bookViews.AppendChild(workbookView);

				// save it to the package
				StreamWriter stream = new StreamWriter(partWorkbook.GetStream(FileMode.Create, FileAccess.Write));
				this._workbookXml.Save(stream);
				this.Package.Package.Flush();
			}
		}

		/// <summary>
		/// Read shared strings to list
		/// </summary>
		private void GetSharedStrings()
		{
			if (this.Package.Package.PartExists(SharedStringsUri))
			{
				var xml = this.Package.GetXmlFromUri(SharedStringsUri);
				XmlNodeList nl = xml.SelectNodes("//d:sst/d:si", this.NameSpaceManager);
				this.SharedStringsList = new List<SharedStringItem>();
				if (nl != null)
				{
					foreach (XmlNode node in nl)
					{
						XmlNode n = node.SelectSingleNode("d:t", NameSpaceManager);
						if (n != null)
						{
							this.SharedStringsList.Add(new SharedStringItem() { Text = ConvertUtil.ExcelDecodeString(n.InnerText) });
						}
						else
						{
							this.SharedStringsList.Add(new SharedStringItem() { Text = node.InnerXml, isRichText = true });
						}
					}
				}
				//Delete the shared string part, it will be recreated when the package is saved.
				foreach (var rel in Part.GetRelationships())
				{
					if (rel.TargetUri.OriginalString.EndsWith("sharedstrings.xml", StringComparison.InvariantCultureIgnoreCase))
					{
						this.Part.DeleteRelationship(rel.Id);
						break;
					}
				}
				this.Package.Package.DeletePart(SharedStringsUri); //Remove the part, it is recreated when saved.
			}
		}

		private void DeleteCalcChain()
		{
			//Remove the calc chain if it exists.
			Uri uriCalcChain = new Uri("/xl/calcChain.xml", UriKind.Relative);
			if (this.Package.Package.PartExists(uriCalcChain))
			{
				Uri calcChain = new Uri("calcChain.xml", UriKind.Relative);
				foreach (var relationship in this.Package.Workbook.Part.GetRelationships())
				{
					if (relationship.TargetUri == calcChain)
					{
						this.Package.Workbook.Part.DeleteRelationship(relationship.Id);
						break;
					}
				}
				this.Package.Package.DeletePart(uriCalcChain);
			}
		}

		private void ValidateDataValidations()
		{
			foreach (var sheet in this.Package.Workbook.Worksheets)
			{
				if (!(sheet is ExcelChartsheet))
				{
					sheet.DataValidations.ValidateAll();
				}
			}
		}

		private void SaveSharedStringHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
		{
			stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
			stream.PutNextEntry(fileName);

			var cache = new StringBuilder();
			var sw = new StreamWriter(stream);
			cache.AppendFormat("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", SharedStrings.Count);
			foreach (string t in SharedStrings.Keys)
			{

				SharedStringItem ssi = SharedStrings[t];
				if (ssi.isRichText)
				{
					cache.Append("<si>");
					ConvertUtil.ExcelEncodeString(cache, t);
					cache.Append("</si>");
				}
				else
				{
					if (t.Length > 0 && (t[0] == ' ' || t[t.Length - 1] == ' ' || t.Contains("  ") || t.Contains("\t") || t.Contains("\n") || t.Contains("\n")))   //Fixes issue 14849
					{
						cache.Append("<si><t xml:space=\"preserve\">");
					}
					else
					{
						cache.Append("<si><t>");
					}
					ConvertUtil.ExcelEncodeString(cache, ConvertUtil.ExcelEscapeString(t));
					cache.Append("</t></si>");
				}
				if (cache.Length > 0x600000)
				{
					sw.Write(cache.ToString());
					cache = new StringBuilder();
				}
			}
			cache.Append("</sst>");
			sw.Write(cache.ToString());
			sw.Flush();
		}

		private void UpdateDefinedNamesXml()
		{
			try
			{
				XmlNode top = this.WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
				if (!this.ExistsNames())
				{
					if (top != null) this.TopNode.RemoveChild(top);
					return;
				}
				else
				{
					if (top == null)
					{
						this.CreateNode("d:definedNames");
						top = this.WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
					}
					else
					{
						top.RemoveAll();
					}
					foreach (ExcelNamedRange name in this.Names)
					{

						XmlElement elem = this.WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
						top.AppendChild(elem);
						elem.SetAttribute("name", name.Name);
						if (name.IsNameHidden)
							elem.SetAttribute("hidden", "1");
						if (!string.IsNullOrEmpty(name.NameComment))
							elem.SetAttribute("comment", name.NameComment);
						this.SetNameElement(name, elem);
					}
				}
				foreach (ExcelWorksheet ws in _worksheets)
				{
					if (!(ws is ExcelChartsheet))
					{
						foreach (ExcelNamedRange name in ws.Names)
						{
							XmlElement elem = this.WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
							top.AppendChild(elem);
							elem.SetAttribute("name", name.Name);
							elem.SetAttribute("localSheetId", name.LocalSheetID.ToString());
							if (name.IsNameHidden)
								elem.SetAttribute("hidden", "1");
							if (!string.IsNullOrEmpty(name.NameComment))
								elem.SetAttribute("comment", name.NameComment);
							this.SetNameElement(name, elem);
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Internal error updating named ranges ", ex);
			}
		}

		private void SetNameElement(ExcelNamedRange name, XmlElement elem)
		{
			if (string.IsNullOrEmpty(name.NameFormula))
				throw new InvalidOperationException("Named range formulas cannot be blank.");
			elem.InnerText = name.NameFormula;
		}

		/// <summary>
		/// Is their any names in the workbook or in the sheets.
		/// </summary>
		/// <returns>?</returns>
		private bool ExistsNames()
		{
			if (this.Names.Count == 0)
			{
				foreach (ExcelWorksheet ws in this.Worksheets)
				{
					if (ws is ExcelChartsheet) continue;
					if (ws.Names.Count > 0)
					{
						return true;
					}
				}
			}
			else
			{
				return true;
			}
			return false;
		}

		private string ResolveExternalReference(string referenceId)
		{
			var rel = this.Part.GetRelationship(referenceId);
			var part = this.Package.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
			XmlDocument xmlExtRef = new XmlDocument();
			LoadXmlSafe(xmlExtRef, part.GetStream());
			XmlElement book = xmlExtRef.SelectSingleNode("//d:externalBook", NameSpaceManager) as XmlElement;
			if (book != null)
			{
				string rId_ExtRef = book.GetAttribute("r:id");
				var rel_extRef = part.GetRelationship(rId_ExtRef);
				if (rel_extRef != null)
					return rel_extRef.TargetUri.OriginalString;
			}
			return null;
		}
		#endregion

		#region Internal Static Methods
		/// <summary>
		/// Get the width of a font in pixels.
		/// </summary>
		/// <param name="fontName">The name of the font to measure.</param>
		/// <param name="fontSize">The number of points currently set for the font.</param>
		/// <returns></returns>
		internal static decimal GetWidthPixels(string fontName, float fontSize)
		{
			Dictionary<float, FontSizeInfo> font;
			if (FontSize.FontHeights.ContainsKey(fontName))
			{
				font = FontSize.FontHeights[fontName];
			}
			else
			{
				font = FontSize.FontHeights["Calibri"];
			}

			if (font.ContainsKey(fontSize))
			{
				return Convert.ToDecimal(font[fontSize].Width);
			}
			else
			{
				float min = -1, max = 500;
				foreach (var size in font)
				{
					if (min < size.Key && size.Key < fontSize)
					{
						min = size.Key;
					}
					if (max > size.Key && size.Key > fontSize)
					{
						max = size.Key;
					}
				}
				if (min == max)
				{
					return Convert.ToDecimal(font[min].Width);
				}
				else
				{
					return Convert.ToDecimal(font[min].Height + (font[max].Height - font[min].Height) * ((fontSize - min) / (max - min)));
				}
			}
		}
		#endregion
	}
}
