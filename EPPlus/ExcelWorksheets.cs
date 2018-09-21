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
 * Jan Källman		    Initial Release		       2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Slicers;
using OfficeOpenXml.Drawing.Sparkline;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml
{
	/// <summary>
	/// The collection of worksheets for the workbook
	/// </summary>
	public class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IDisposable
	{
		#region Constants
		private const string ERR_DUP_WORKSHEET = "A worksheet with this name already exists in the workbook";
		internal const string WORKSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
		internal const string CHARTSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
		#endregion

		#region Private Properties
		private ExcelPackage Package { get; set; }
		private Dictionary<int, ExcelWorksheet> Worksheets { get; set; }
		private XmlNamespaceManager NamespaceManager { get; set; }
		#endregion

		#region Public Properties
		/// <summary>
		/// Returns the number of worksheets in the workbook
		/// </summary>
		public int Count
		{
			get { return (this.Worksheets.Count); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a new ExcelWorksheets collection based on the specified <paramref name="topNode"/>.
		/// </summary>
		/// <param name="package">The excel package.</param>
		/// <param name="namespaceManager">The namespace manager with the namespaces for the node.</param>
		/// <param name="topNode">The top XML node of the Worksheets collection.</param>
		internal ExcelWorksheets(ExcelPackage package, XmlNamespaceManager namespaceManager, XmlNode topNode) : base(namespaceManager, topNode)
		{
			this.Package = package;
			this.NamespaceManager = namespaceManager;
			this.Worksheets = new Dictionary<int, ExcelWorksheet>();
			int positionID = 1;

			foreach (XmlNode sheetNode in topNode.ChildNodes)
			{
				if (sheetNode.NodeType == XmlNodeType.Element)
				{
					// Process sheet identity information
					string name = sheetNode.Attributes["name"].Value;
					string relId = sheetNode.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships).Value;
					int sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);

					eWorkSheetHidden hidden = eWorkSheetHidden.Visible;
					XmlNode attr = sheetNode.Attributes["state"];
					if (attr != null)
						hidden = TranslateHidden(attr.Value);

					var sheetRelation = package.Workbook.Part.GetRelationship(relId);
					Uri uriWorksheet = UriHelper.ResolvePartUri(package.Workbook.WorkbookUri, sheetRelation.TargetUri);

					if (sheetRelation.RelationshipType.EndsWith("chartsheet"))
					{
						this.Worksheets.Add(positionID, new ExcelChartsheet(this.NamespaceManager, this.Package, relId, uriWorksheet, name, sheetID, positionID, hidden));
					}
					else
					{
						this.Worksheets.Add(positionID, new ExcelWorksheet(this.NamespaceManager, this.Package, relId, uriWorksheet, name, sheetID, positionID, hidden));
					}
					positionID++;
				}
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Foreach support
		/// </summary>
		/// <returns>An enumerator</returns>
		public IEnumerator<ExcelWorksheet> GetEnumerator()
		{
			return (this.Worksheets.Values.GetEnumerator());
		}

		#region IEnumerable Members
		IEnumerator IEnumerable.GetEnumerator()
		{
			return (this.Worksheets.Values.GetEnumerator());
		}
		#endregion

		/// <summary>
		/// Adds a new blank worksheet.
		/// </summary>
		/// <param name="Name">The name of the workbook</param>
		public ExcelWorksheet Add(string Name)
		{
			ExcelWorksheet worksheet = this.AddSheet(Name, false, null);
			return worksheet;
		}

		/// <summary>
		/// Adds a copy of a worksheet
		/// </summary>
		/// <param name="Name">The name of the workbook</param>
		/// <param name="originalWorksheet">The worksheet to be copied</param>
		public ExcelWorksheet Add(string Name, ExcelWorksheet originalWorksheet)
		{
			lock (this.Worksheets)
			{
				int sheetID;
				Uri uriWorksheet;
				if (originalWorksheet is ExcelChartsheet)
				{
					throw (new ArgumentException("Can not copy a chartsheet"));
				}
				if (this.GetByName(Name) != null)
				{
					throw (new InvalidOperationException(ERR_DUP_WORKSHEET));
				}

				GetSheetURI(ref Name, out sheetID, out uriWorksheet, false);

				//Create a copy of the worksheet XML
				Packaging.ZipPackagePart worksheetPart = this.Package.Package.CreatePart(uriWorksheet, WORKSHEET_CONTENTTYPE, this.Package.Compression);
				StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
				XmlDocument worksheetXml = new XmlDocument();
				worksheetXml.LoadXml(originalWorksheet.WorksheetXml.OuterXml);
				worksheetXml.Save(streamWorksheet);
				this.Package.Package.Flush();

				//Create a relation to the workbook
				string relID = this.CreateWorkbookRel(Name, sheetID, uriWorksheet, false);
				ExcelWorksheet added = new ExcelWorksheet(this.NamespaceManager, this.Package, relID, uriWorksheet, Name, sheetID, this.Worksheets.Count + 1, eWorkSheetHidden.Visible);

				if (originalWorksheet.Comments.Count > 0)
				{
					this.CopyComment(originalWorksheet, added);
				}

				CopyHeaderFooterPictures(originalWorksheet, added);

				if (originalWorksheet.Slicers.Slicers.Count > 0)
				{
					this.CopySlicers(originalWorksheet, added);
				}
				if (originalWorksheet.Drawings.Count > 0)
				{
					this.CopyDrawing(originalWorksheet, added);
				}
				if (originalWorksheet.Tables.Count > 0)
				{
					this.CopyTable(originalWorksheet, added);
				}
				if (originalWorksheet.PivotTables.Count > 0)
				{
					this.CopyPivotTable(originalWorksheet, added);
				}
				if (originalWorksheet.Names.Count > 0)
				{
					this.CopySheetNamedRanges(originalWorksheet, added);
				}
				if (originalWorksheet.SparklineGroups.SparklineGroups.Count > 0)
				{
					this.CopySparklines(originalWorksheet, added);
				}
				this.CloneCells(originalWorksheet, added);

#if !MONO
				if (this.Package.Workbook.VbaProject != null)
				{
					var name = this.Package.Workbook.VbaProject.GetModuleNameFromWorksheet(added);
					this.Package.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(added.CodeNameChange) { Name = name, Code = originalWorksheet.CodeModule.Code, Attributes = this.Package.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
					originalWorksheet.CodeModuleName = name;
				}
#endif
				this.Worksheets.Add(this.Worksheets.Count + 1, added);

				//Remove any relation to printersettings.
				XmlNode pageSetup = added.WorksheetXml.SelectSingleNode("//d:pageSetup", this.NamespaceManager);
				if (pageSetup != null)
				{
					XmlAttribute attr = (XmlAttribute)pageSetup.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
					if (attr != null)
					{
						relID = attr.Value;
						// first delete the attribute from the XML
						pageSetup.Attributes.Remove(attr);
					}
				}

				// Update chart series that reference the same sheet as the chart.
				foreach (ExcelChart chart in added.Drawings.Where(drawing => drawing is ExcelChart))
				{
					if (chart != null)
					{
						foreach (ExcelChartSerie serie in chart.Series)
						{
							string workbook, worksheet, address;
							ExcelRange.SplitAddress(serie.Series, out workbook, out worksheet, out address);
							if (worksheet == originalWorksheet.Name)
							{
								serie.Series = ExcelRange.GetFullAddress(added.Name, address);
							}
							if (!string.IsNullOrEmpty(serie.XSeries))
							{
								ExcelRange.SplitAddress(serie.XSeries, out workbook, out worksheet, out address);
								if (worksheet == originalWorksheet.Name)
								{
									serie.XSeries = ExcelRange.GetFullAddress(added.Name, address);
								}
							}
						}
					}
				}
				return added;
			}
		}

		/// <summary>
		/// Adds a chartsheet to the workbook.
		/// </summary>
		/// <param name="Name">The chart's name.</param>
		/// <param name="chartType">The chart's <see cref="eChartType"/>.</param>
		/// <returns>The newly-created <see cref="ExcelChartsheet"/>.</returns>
		public ExcelChartsheet AddChart(string Name, eChartType chartType)
		{
			return (ExcelChartsheet)this.AddSheet(Name, true, chartType);
		}

		/// <summary>
		/// Deletes a worksheet from the collection.
		/// </summary>
		/// <param name="Index">The position of the worksheet in the workbook to be deleted.</param>
		public void Delete(int Index)
		{
			/*
            * Hack to prefetch all the drawings,
            * so that all the images are referenced,
            * to prevent the deletion of the image file,
            * when referenced more than once
            */
			foreach (var ws in this.Worksheets)
			{
				var drawings = ws.Value.Drawings;
			}

			ExcelWorksheet worksheet = this.Worksheets[Index];
			if (worksheet.Drawings.Count > 0)
				worksheet.Drawings.ClearDrawings();
			if (!(worksheet is ExcelChartsheet) && worksheet.Comments.Count > 0)
				worksheet.Comments.Clear();

			// Update all named range formulas referencing this sheet to #REF!
			foreach (var namedRange in this.Package.Workbook.Names)
			{
				namedRange.NameFormula = this.Package.FormulaManager.UpdateFormulaDeletedSheetReferences(namedRange.NameFormula, worksheet.Name);
			}

			foreach (var sheet in this.Worksheets.Where(w => !w.Value.Name.IsEquivalentTo(worksheet.Name)))
			{
				foreach (var namedRange in sheet.Value.Names)
				{
					namedRange.NameFormula = this.Package.FormulaManager.UpdateFormulaDeletedSheetReferences(namedRange.NameFormula, worksheet.Name);
				}
			}

			//Delete any parts still with relations to the Worksheet.
			this.DeleteRelationsAndParts(worksheet.Part);

			//Delete the worksheet part and relation from the package
			this.Package.Workbook.Part.DeleteRelationship(worksheet.RelationshipID);

			//Delete worksheet from the workbook XML
			XmlNode sheetsNode = this.Package.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", this.NamespaceManager);
			if (sheetsNode != null)
			{
				XmlNode sheetNode = sheetsNode.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetID), this.NamespaceManager);
				if (sheetNode != null)
				{
					sheetsNode.RemoveChild(sheetNode);
				}
			}
			this.Worksheets.Remove(Index);
			if (this.Package.Workbook.VbaProject != null)
			{
				this.Package.Workbook.VbaProject.Modules.Remove(worksheet.CodeModule);
			}
			this.ReindexWorksheetDictionary();
			//If the active sheet is deleted, set the first tab as active.
			if (this.Package.Workbook.View.ActiveTab >= this.Package.Workbook.Worksheets.Count)
			{
				this.Package.Workbook.View.ActiveTab = this.Package.Workbook.View.ActiveTab - 1;
			}
			if (this.Package.Workbook.View.ActiveTab == worksheet.SheetID)
			{
				this.Package.Workbook.Worksheets[1].View.TabSelected = true;
			}
			worksheet = null;
		}

		/// <summary>
		/// Deletes a worksheet from the collection.
		/// </summary>
		/// <param name="name">The name of the worksheet to be deleted.</param>
		public void Delete(string name)
		{
			var sheet = this[name];
			if (sheet == null)
			{
				throw new ArgumentException(string.Format("Could not find worksheet to delete '{0}'", name));
			}
			this.Delete(sheet.PositionID);
		}

		/// <summary>
		/// Delete a worksheet from the collection.
		/// </summary>
		/// <param name="Worksheet">The worksheet to delete.</param>
		public void Delete(ExcelWorksheet Worksheet)
		{
			if (Worksheet.PositionID <= this.Worksheets.Count && Worksheet == this.Worksheets[Worksheet.PositionID])
			{
				this.Delete(Worksheet.PositionID);
			}
			else
			{
				throw (new ArgumentException("Worksheet is not in the collection."));
			}
		}

		/// <summary>
		/// Copies the named worksheet and creates a new worksheet in the same workbook.
		/// </summary>
		/// <param name="originalWorksheetName">The name of the existing worksheet.</param>
		/// <param name="newName">The name of the new worksheet to create.</param>
		/// <returns>The new copy, added to the end of the worksheets collection.</returns>
		public ExcelWorksheet Copy(string originalWorksheetName, string newName)
		{
			ExcelWorksheet original = this[originalWorksheetName];
			if (original == null)
				throw new ArgumentException(string.Format("Copy worksheet error: Could not find worksheet to copy '{0}'", originalWorksheetName));

			ExcelWorksheet added = this.Add(newName, original);
			return added;
		}

		/// <summary>
		/// Moves the source worksheet to the position before the target worksheet
		/// </summary>
		/// <param name="sourceName">The name of the source worksheet</param>
		/// <param name="targetName">The name of the target worksheet</param>
		public void MoveBefore(string sourceName, string targetName)
		{
			this.Move(sourceName, targetName, false);
		}

		/// <summary>
		/// Moves the source worksheet to the position before the target worksheet
		/// </summary>
		/// <param name="sourcePositionId">The id of the source worksheet</param>
		/// <param name="targetPositionId">The id of the target worksheet</param>
		public void MoveBefore(int sourcePositionId, int targetPositionId)
		{
			this.Move(sourcePositionId, targetPositionId, false);
		}

		/// <summary>
		/// Moves the source worksheet to the position after the target worksheet
		/// </summary>
		/// <param name="sourceName">The name of the source worksheet</param>
		/// <param name="targetName">The name of the target worksheet</param>
		public void MoveAfter(string sourceName, string targetName)
		{
			this.Move(sourceName, targetName, true);
		}

		/// <summary>
		/// Moves the source worksheet to the position after the target worksheet
		/// </summary>
		/// <param name="sourcePositionId">The id of the source worksheet</param>
		/// <param name="targetPositionId">The id of the target worksheet</param>
		public void MoveAfter(int sourcePositionId, int targetPositionId)
		{
			this.Move(sourcePositionId, targetPositionId, true);
		}

		/// <summary>
		/// Move a worksheet with the specified name to the front of the workbook.
		/// </summary>
		/// <param name="sourceName">The worksheet to be moved.</param>
		public void MoveToStart(string sourceName)
		{
			var sourceSheet = this[sourceName];
			if (sourceSheet == null)
			{
				throw new ArgumentException(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
			}
			this.Move(sourceSheet.PositionID, 1, false);
		}

		/// <summary>
		/// Move a worksheet with the specified name to the front of the workbook.
		/// </summary>
		/// <param name="sourcePositionId">The worksheet to be moved.</param>
		public void MoveToStart(int sourcePositionId)
		{
			this.Move(sourcePositionId, 1, false);
		}

		/// <summary>
		/// Move a worksheet with the specified name to the end of the workbook.
		/// </summary>
		/// <param name="sourceName">The worksheet to be moved.</param>
		public void MoveToEnd(string sourceName)
		{
			var sourceSheet = this[sourceName];
			if (sourceSheet == null)
			{
				throw new ArgumentException(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
			}
			this.Move(sourceSheet.PositionID, this.Worksheets.Count, true);
		}

		/// <summary>
		/// Move a worksheet from the specified position to the end of the workbook.
		/// </summary>
		/// <param name="sourcePositionId">The index of the worksheet to move.</param>
		public void MoveToEnd(int sourcePositionId)
		{
			Move(sourcePositionId, this.Worksheets.Count, true);
		}

		/// <summary>
		/// Dispose the ExcelWorksheets collection and all its children.
		/// </summary>
		public void Dispose()
		{
			foreach (var sheet in this.Worksheets.Values)
			{
				((IDisposable)sheet).Dispose();
			}
			this.Worksheets = null;
			this.Package = null;
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Validate a possible sheet name, applying standard replacements if necessary.
		/// </summary>
		/// <param name="Name"></param>
		/// <returns></returns>
		internal string ValidateAndFixSheetName(string Name)
		{
			//remove invalid characters
			if (ExcelWorksheets.IsInvalidSheetName(Name))
			{
				if (Name.IndexOf(':') > -1) Name = Name.Replace(":", " ");
				if (Name.IndexOf('/') > -1) Name = Name.Replace("/", " ");
				if (Name.IndexOf('\\') > -1) Name = Name.Replace("\\", " ");
				if (Name.IndexOf('?') > -1) Name = Name.Replace("?", " ");
				if (Name.IndexOf('[') > -1) Name = Name.Replace("[", " ");
				if (Name.IndexOf(']') > -1) Name = Name.Replace("]", " ");
			}

			if (Name.Trim() == "")
			{
				throw new ArgumentException("The worksheet can not have an empty name");
			}
			if (Name.StartsWith("'") || Name.EndsWith("'"))
			{
				throw new ArgumentException("The worksheet name can not start or end with an apostrophe.");
			}
			if (Name.StartsWith("'") || Name.EndsWith("'"))
			{
				throw new ArgumentException("The worksheet name can not start or end with an apostrophe.");
			}
			if (Name.Length > 31) Name = Name.Substring(0, 31);   //A sheet can have max 31 char's
			return Name;
		}

		/// <summary>
		/// Creates the XML document representing a new empty worksheet.
		/// </summary>
		/// <returns>The newly-created XmlDocument representing the worksheet.</returns>
		internal XmlDocument CreateNewWorksheetDocument(bool isChart)
		{
			XmlDocument xmlDoc = new XmlDocument();
			XmlElement elemWs = xmlDoc.CreateElement(isChart ? "chartsheet" : "worksheet", ExcelPackage.schemaMain);
			elemWs.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
			xmlDoc.AppendChild(elemWs);


			if (isChart)
			{
				XmlElement elemSheetPr = xmlDoc.CreateElement("sheetPr", ExcelPackage.schemaMain);
				elemWs.AppendChild(elemSheetPr);

				XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
				elemWs.AppendChild(elemSheetViews);

				XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
				elemSheetView.SetAttribute("workbookViewId", "0");
				elemSheetView.SetAttribute("zoomToFit", "1");

				elemSheetViews.AppendChild(elemSheetView);
			}
			else
			{
				XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
				elemWs.AppendChild(elemSheetViews);

				XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
				elemSheetView.SetAttribute("workbookViewId", "0");
				elemSheetViews.AppendChild(elemSheetView);

				XmlElement elemSheetFormatPr = xmlDoc.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
				elemSheetFormatPr.SetAttribute("defaultRowHeight", "15");
				elemWs.AppendChild(elemSheetFormatPr);

				XmlElement elemSheetData = xmlDoc.CreateElement("sheetData", ExcelPackage.schemaMain);
				elemWs.AppendChild(elemSheetData);
			}
			return xmlDoc;
		}

		/// <summary>
		/// Gets a sheet by its specified <paramref name="sheetID"/>.
		/// </summary>
		/// <param name="sheetID">The sheet Id of the worksheet to retrieve (NOT the localSheetID property of a named range).</param>
		/// <returns>The first worksheet with the specified SheetID.</returns>
		internal ExcelWorksheet GetBySheetID(int sheetID)
		{
			foreach (ExcelWorksheet ws in this)
			{
				if (ws.SheetID == sheetID)
				{
					return ws;
				}
			}
			return null;
		}
		#endregion

		#region Public Operators
		/// <summary>
		/// Returns the worksheet at the specified position.
		/// </summary>
		/// <param name="PositionID">The 1-base position of the worksheet.</param>
		/// <returns>The worksheet at the specified position.</returns>
		public ExcelWorksheet this[int PositionID]
		{
			get
			{
				if (this.Worksheets.ContainsKey(PositionID))
				{
					return this.Worksheets[PositionID];
				}
				else
				{
					throw (new IndexOutOfRangeException("Worksheet position out of range."));
				}
			}
		}

		/// <summary>
		/// Returns the worksheet matching the specified name.
		/// </summary>
		/// <param name="Name">The name of the worksheet to retrieve.</param>
		/// <returns>The worksheet with the specified name.</returns>
		public ExcelWorksheet this[string Name]
		{
			get
			{
				return this.GetByName(Name);
			}
		}
		#endregion

		#region Private Methods
		private eWorkSheetHidden TranslateHidden(string value)
		{
			switch (value)
			{
				case "hidden":
					return eWorkSheetHidden.Hidden;
				case "veryHidden":
					return eWorkSheetHidden.VeryHidden;
				default:
					return eWorkSheetHidden.Visible;
			}
		}

		private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType)
		{
			int sheetID;
			Uri uriWorksheet;
			lock (this.Worksheets)
			{
				Name = this.ValidateAndFixSheetName(Name);
				if (this.GetByName(Name) != null)
				{
					throw (new InvalidOperationException(ERR_DUP_WORKSHEET + " : " + Name));
				}
				this.GetSheetURI(ref Name, out sheetID, out uriWorksheet, isChart);
				Packaging.ZipPackagePart worksheetPart = this.Package.Package.CreatePart(uriWorksheet, isChart ? CHARTSHEET_CONTENTTYPE : WORKSHEET_CONTENTTYPE, this.Package.Compression);

				//Create the new, empty worksheet and save it to the package
				StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
				XmlDocument worksheetXml = this.CreateNewWorksheetDocument(isChart);
				worksheetXml.Save(streamWorksheet);
				this.Package.Package.Flush();

				string rel = this.CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart);

				int positionID = this.Worksheets.Count + 1;
				ExcelWorksheet worksheet;
				if (isChart)
				{
					worksheet = new ExcelChartsheet(this.NamespaceManager, this.Package, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible, (eChartType)chartType);
				}
				else
				{
					worksheet = new ExcelWorksheet(this.NamespaceManager, this.Package, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);
				}

				this.Worksheets.Add(positionID, worksheet);
#if !MONO
				if (this.Package.Workbook.VbaProject != null)
				{
					var name = this.Package.Workbook.VbaProject.GetModuleNameFromWorksheet(worksheet);
					this.Package.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(worksheet.CodeNameChange) { Name = name, Code = "", Attributes = this.Package.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
					worksheet.CodeModuleName = name;

				}
#endif
				return worksheet;
			}
		}

		private void CopySparklines(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			for (int i = 0; i < originalWorksheet.SparklineGroups.SparklineGroups.Count; i++)
			{
				var group = addedWorksheet.SparklineGroups.SparklineGroups[i];
				group.Worksheet = addedWorksheet;
				group.Sparklines.Clear();
				foreach (var originalSparkline in originalWorksheet.SparklineGroups.SparklineGroups[i].Sparklines)
				{
					ExcelAddress newFormula = null;
					if (originalSparkline.Formula != null)
						newFormula = new ExcelAddress(originalSparkline.Formula.Address);
					ExcelAddress newHostCell = new ExcelAddress(originalSparkline.HostCell.Address);
					var sparkline = new ExcelSparkline(newHostCell, newFormula, group, group.NameSpaceManager);
					sparkline.Formula?.ChangeWorksheet(originalWorksheet.Name, addedWorksheet.Name);
					group.Sparklines.Add(sparkline);
				}
			}
		}

		private void CopySheetNamedRanges(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			foreach (var namedRange in originalWorksheet.Names)
			{
				string updatedFormula = this.Package.FormulaManager.UpdateFormulaSheetReferences(namedRange.NameFormula, originalWorksheet.Name, addedWorksheet.Name);
				addedWorksheet.Names.Add(namedRange.Name, updatedFormula, namedRange.IsNameHidden, namedRange.NameComment);
			}
		}

		private void CopyTable(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			string prevName = "";
			//First copy the table XML
			foreach (var tbl in originalWorksheet.Tables)
			{
				string xml = tbl.TableXml.OuterXml;
				string name;
				if (prevName == "")
				{
					name = originalWorksheet.Tables.GetNewTableName();
				}
				else
				{
					int ix = int.Parse(prevName.Substring(5)) + 1;
					name = string.Format("Table{0}", ix);
					while (this.Package.Workbook.ExistsPivotTableName(name))
					{
						name = string.Format("Table{0}", ++ix);
					}
				}
				int Id = this.Package.Workbook.NextTableID++;
				prevName = name;
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xml);
				xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
				xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
				xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
				xml = xmlDoc.OuterXml;

				//var uriTbl = new Uri(string.Format("/xl/tables/table{0}.xml", Id), UriKind.Relative);
				var uriTbl = GetNewUri(this.Package.Package, "/xl/tables/table{0}.xml", ref Id);
				if (this.Package.Workbook.NextTableID < Id) this.Package.Workbook.NextTableID = Id;

				var part = this.Package.Package.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", this.Package.Compression);
				StreamWriter streamTbl = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
				streamTbl.Write(xml);
				//streamTbl.Close();
				streamTbl.Flush();

				//create the relationship and add the ID to the worksheet xml.
				var rel = addedWorksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(addedWorksheet.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");

				if (tbl.RelationshipID == null)
				{
					var topNode = addedWorksheet.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
					if (topNode == null)
					{
						addedWorksheet.CreateNode("d:tableParts");
						topNode = addedWorksheet.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
					}
					XmlElement elem = addedWorksheet.WorksheetXml.CreateElement("tablePart", ExcelPackage.schemaMain);
					topNode.AppendChild(elem);
					elem.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
				}
				else
				{
					XmlAttribute relAtt;
					relAtt = addedWorksheet.WorksheetXml.SelectSingleNode(string.Format("//d:tableParts/d:tablePart/@r:id[.='{0}']", tbl.RelationshipID), tbl.NameSpaceManager) as XmlAttribute;
					relAtt.Value = rel.Id;
				}
			}
		}

		private void CopyPivotTable(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			string prevName = "";
			foreach (var tbl in originalWorksheet.PivotTables)
			{
				string xml = tbl.PivotTableXml.OuterXml;

				string name;
				if (prevName == "")
				{
					name = originalWorksheet.PivotTables.GetNewTableName();
				}
				else
				{
					int ix = int.Parse(prevName.Substring(10)) + 1;
					name = string.Format("PivotTable{0}", ix);
					while (this.Package.Workbook.ExistsPivotTableName(name))
					{
						name = string.Format("PivotTable{0}", ++ix);
					}
				}
				prevName = name;
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xml);
				xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@name", tbl.NameSpaceManager).Value = name;

				var newSheetId = addedWorksheet.SheetID.ToString();
				foreach (var slicerCache in this.Package.Workbook.SlicerCaches)
				{
					foreach (var pivotTable in slicerCache.PivotTables)
					{
						if (pivotTable.TabId == newSheetId && pivotTable.PivotTableName == tbl.Name)
							pivotTable.PivotTableName = name;
					}
				}

				xml = xmlDoc.OuterXml;
				int Id = this.Package.Workbook.NextPivotTableID++;
				var uriTbl = GetNewUri(this.Package.Package, "/xl/pivotTables/pivotTable{0}.xml", ref Id);
				if (this.Package.Workbook.NextPivotTableID < Id) this.Package.Workbook.NextPivotTableID = Id;
				var partTbl = this.Package.Package.CreatePart(uriTbl, ExcelPackage.schemaPivotTable, this.Package.Compression);
				StreamWriter streamTbl = new StreamWriter(partTbl.GetStream(FileMode.Create, FileAccess.Write));
				streamTbl.Write(xml);
				streamTbl.Flush();
				
				// Create the relationship and add the ID to the worksheet xml.
				addedWorksheet.Part.CreateRelationship(UriHelper.ResolvePartUri(addedWorksheet.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
				// Creates the relationship to the original table's cache definition.
				// Copying pivot tables does not duplicate cache definitions.
				partTbl.CreateRelationship(UriHelper.ResolvePartUri(tbl.WorksheetRelationship.SourceUri, tbl.CacheDefinition.CacheDefinitionUri), tbl.CacheDefinitionRelationship.TargetMode, tbl.CacheDefinitionRelationship.RelationshipType);
			}
		}

		private void CopyHeaderFooterPictures(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			if (originalWorksheet.TopNode != null && originalWorksheet.TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager) == null) return;
			//Copy the texts
			if (originalWorksheet.HeaderFooter._oddHeader != null) CopyText(originalWorksheet.HeaderFooter._oddHeader, addedWorksheet.HeaderFooter.OddHeader);
			if (originalWorksheet.HeaderFooter._oddFooter != null) CopyText(originalWorksheet.HeaderFooter._oddFooter, addedWorksheet.HeaderFooter.OddFooter);
			if (originalWorksheet.HeaderFooter._evenHeader != null) CopyText(originalWorksheet.HeaderFooter._evenHeader, addedWorksheet.HeaderFooter.EvenHeader);
			if (originalWorksheet.HeaderFooter._evenFooter != null) CopyText(originalWorksheet.HeaderFooter._evenFooter, addedWorksheet.HeaderFooter.EvenFooter);
			if (originalWorksheet.HeaderFooter._firstHeader != null) CopyText(originalWorksheet.HeaderFooter._firstHeader, addedWorksheet.HeaderFooter.FirstHeader);
			if (originalWorksheet.HeaderFooter._firstFooter != null) CopyText(originalWorksheet.HeaderFooter._firstFooter, addedWorksheet.HeaderFooter.FirstFooter);

			//Copy any images;
			if (originalWorksheet.HeaderFooter.Pictures.Count > 0)
			{
				Uri source = originalWorksheet.HeaderFooter.Pictures.Uri;
				Uri dest = XmlHelper.GetNewUri(this.Package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
				addedWorksheet.DeleteNode("d:legacyDrawingHF");

				//var part = this.Package.Package.CreatePart(dest, "application/vnd.openxmlformats-officedocument.vmlDrawing", this.Package.Compression);
				foreach (ExcelVmlDrawingPicture pic in originalWorksheet.HeaderFooter.Pictures)
				{
					var item = addedWorksheet.HeaderFooter.Pictures.Add(pic.Id, pic.ImageUri, pic.Title, pic.Width, pic.Height);
					foreach (XmlAttribute att in pic.TopNode.Attributes)
					{
						(item.TopNode as XmlElement).SetAttribute(att.Name, att.Value);
					}
					item.TopNode.InnerXml = pic.TopNode.InnerXml;
				}
			}
		}

		private void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
		{
			to.LeftAlignedText = from.LeftAlignedText;
			to.CenteredText = from.CenteredText;
			to.RightAlignedText = from.RightAlignedText;
		}

		private void CloneCells(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			bool sameWorkbook = originalWorksheet.Workbook == this.Package.Workbook;

			bool doAdjust = this.Package.DoAdjustDrawings;
			this.Package.DoAdjustDrawings = false;
			addedWorksheet.MergedCells.List.AddRange(originalWorksheet.MergedCells.List);

			foreach (int key in originalWorksheet._sharedFormulas.Keys)
			{
				addedWorksheet._sharedFormulas.Add(key, originalWorksheet._sharedFormulas[key].Clone());
			}

			Dictionary<int, int> styleCashe = new Dictionary<int, int>();
			int row, col;
			var val = originalWorksheet._values.GetEnumerator();
			while (val.MoveNext())
			{
				row = val.Row;
				col = val.Column;
				int styleID = 0;
				if (row == 0) //Column
				{
					var c = originalWorksheet.GetValueInner(row, col) as ExcelColumn;
					if (c != null)
					{
						var clone = c.Clone(addedWorksheet, c.ColumnMin);
						clone.StyleID = c.StyleID;
						addedWorksheet.SetValueInner(row, col, clone);
						styleID = c.StyleID;
					}
				}
				else if (col == 0) //Row
				{
					var r = originalWorksheet.Row(row);
					if (r != null)
					{
						r.Clone(addedWorksheet);
						styleID = r.StyleID;
					}
				}
				else
				{
					styleID = this.CopyValues(originalWorksheet, addedWorksheet, row, col);
				}
				if (!sameWorkbook)
				{
					if (styleCashe.ContainsKey(styleID))
					{
						addedWorksheet.SetStyleInner(row, col, styleCashe[styleID]);
					}
					else
					{
						var s = addedWorksheet.Workbook.Styles.CloneStyle(originalWorksheet.Workbook.Styles, styleID);
						styleCashe.Add(styleID, s);
						addedWorksheet.SetStyleInner(row, col, s);
					}
				}
			}
			addedWorksheet.Package.DoAdjustDrawings = doAdjust;
		}

		private int CopyValues(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet, int row, int col)
		{
			addedWorksheet.SetValueInner(row, col, originalWorksheet.GetValueInner(row, col));
			byte fl = 0;
			if (originalWorksheet._flags.Exists(row, col, out fl))
			{
				addedWorksheet._flags.SetValue(row, col, fl);
			}

			var v = originalWorksheet._formulas.GetValue(row, col);
			if (v != null)
			{
				addedWorksheet.SetFormula(row, col, v);
			}
			var s = originalWorksheet.GetStyleInner(row, col);
			if (s != 0)
			{
				addedWorksheet.SetStyleInner(row, col, s);
			}
			var f = originalWorksheet._formulas.GetValue(row, col);
			if (f != null)
			{
				addedWorksheet._formulas.SetValue(row, col, f);
			}
			return s;
		}

		private void CopyComment(ExcelWorksheet originalWorksheet, ExcelWorksheet addedWorksheet)
		{
			string xml = originalWorksheet.Comments.CommentXml.InnerXml;
			var uriComment = new Uri(string.Format("/xl/comments{0}.xml", addedWorksheet.SheetID), UriKind.Relative);
			var part = this.Package.Package.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", this.Package.Compression);
			StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
			streamDrawing.Write(xml);
			streamDrawing.Flush();
			//Add the relationship ID to the worksheet xml.
			var commentRelation = addedWorksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(addedWorksheet.WorksheetUri, uriComment), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");
		}

		private void CopySlicers(ExcelWorksheet originalWorksheet, ExcelWorksheet newWorksheet)
		{
			var uriSlicer = XmlHelper.GetNewUri(this.Package.Package, "/xl/slicers/slicer{0}.xml");
			var slicerPart = this.Package.Package.CreatePart(uriSlicer, "application/vnd.ms-excel.slicer+xml", this.Package.Compression);
			var slicerRelation = newWorksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newWorksheet.WorksheetUri, uriSlicer), Packaging.TargetMode.Internal, ExcelPackage.schemaSlicerRelationship);
			var newWorksheetSlicerNode = newWorksheet.TopNode.SelectSingleNode("d:extLst/d:ext/x14:slicerList/x14:slicer", newWorksheet.Workbook.NameSpaceManager);
			newWorksheetSlicerNode.Attributes["r:id"].Value = slicerRelation.Id;
			StreamWriter streamSlicer = new StreamWriter(slicerPart.GetStream(FileMode.Create, FileAccess.Write));
			streamSlicer.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
			streamSlicer.Write(originalWorksheet.Slicers.TopNode.OuterXml);
			streamSlicer.Flush();
		}

		private void CopyDrawing(ExcelWorksheet originalWorksheet, ExcelWorksheet newWorksheet)
		{
			string xml = originalWorksheet.Drawings.DrawingXml.OuterXml;
			var uriDraw = new Uri(string.Format("/xl/drawings/drawing{0}.xml", newWorksheet.SheetID), UriKind.Relative);
			var part = this.Package.Package.CreatePart(uriDraw, "application/vnd.openxmlformats-officedocument.drawing+xml", this.Package.Compression);
			StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
			streamDrawing.Write(xml);
			streamDrawing.Flush();

			XmlDocument drawXml = new XmlDocument();
			drawXml.LoadXml(xml);
			//Add the relationship ID to the worksheet xml.
			var drawRelation = newWorksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newWorksheet.WorksheetUri, uriDraw), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
			XmlElement e = newWorksheet.WorksheetXml.SelectSingleNode("//d:drawing", this.NamespaceManager) as XmlElement;
			e.SetAttribute("id", ExcelPackage.schemaRelationships, drawRelation.Id);

			for (int i = 0; i < originalWorksheet.Drawings.Count; i++)
			{
				ExcelDrawing draw = originalWorksheet.Drawings[i];
				draw.AdjustPositionAndSize();       //Adjust position for any change in normal style font/row size etc.
				if (draw is ExcelChart)
				{
					ExcelChart chart = draw as ExcelChart;
					xml = chart.ChartXml.InnerXml;

					var UriChart = XmlHelper.GetNewUri(this.Package.Package, "/xl/charts/chart{0}.xml");
					var chartPart = this.Package.Package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", this.Package.Compression);
					StreamWriter streamChart = new StreamWriter(chartPart.GetStream(FileMode.Create, FileAccess.Write));
					streamChart.Write(xml);
					streamChart.Flush();
					//Now create the new relationship to the copied chart xml
					var prevRelID = draw.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", originalWorksheet.Drawings.NameSpaceManager).Value;
					var rel = part.CreateRelationship(UriHelper.GetRelativeUri(uriDraw, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
					XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), originalWorksheet.Drawings.NameSpaceManager) as XmlAttribute;
					relAtt.Value = rel.Id;
				}
				else if (draw is ExcelSlicerDrawing)
				{
					this.CopySlicerDrawing(originalWorksheet, newWorksheet, draw);
				}
				else if (draw is ExcelPicture)
				{
					ExcelPicture pic = draw as ExcelPicture;
					var uri = pic.UriPic;
					if (!newWorksheet.Workbook.Package.Package.PartExists(uri))
					{
						var picPart = newWorksheet.Workbook.Package.Package.CreatePart(uri, pic.ContentType, CompressionLevel.None);
						pic.Image.Save(picPart.GetStream(FileMode.Create, FileAccess.Write), ExcelPicture.GetImageFormat(pic.ContentType));
					}

					var rel = part.CreateRelationship(UriHelper.GetRelativeUri(newWorksheet.WorksheetUri, uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
					//Fixes problem with invalid image when the same image is used more than once.
					XmlNode relAtt =
						drawXml.SelectSingleNode(
							string.Format(
								"//xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name[.='{0}']/../../../xdr:blipFill/a:blip/@r:embed",
								pic.Name), originalWorksheet.Drawings.NameSpaceManager);
					if (relAtt != null)
					{
						relAtt.Value = rel.Id;
					}
					if (this.Package.Images.ContainsKey(pic.ImageHash))
					{
						this.Package.Images[pic.ImageHash].RefCount++;
					}
				}
			}
			//rewrite the drawing xml with the new relID's
			streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
			streamDrawing.Write(drawXml.OuterXml);
			streamDrawing.Flush();

			//Copy the size variables to the copy.
			for (int i = 0; i < originalWorksheet.Drawings.Count; i++)
			{
				var draw = originalWorksheet.Drawings[i];
				var c = newWorksheet.Drawings[i];
				if (c != null)
				{
					c._left = draw._left;
					c._top = draw._top;
					c._height = draw._height;
					c._width = draw._width;
					var newSlicerDrawing = c as ExcelSlicerDrawing;
					if (newSlicerDrawing != null)
					{
						var newSlicerNumber = this.Package.Workbook.NextSlicerIdNumber[draw.Name]++;
						var slicer = newWorksheet.Slicers.Slicers.First(excelSlicer => excelSlicer.Name == draw.Name);
						slicer.Name += $" {newSlicerNumber}";
						slicer.SlicerCache.Name += newSlicerNumber.ToString();
						this.Package.Workbook.Names.Add(slicer.SlicerCache.Name, "#N/A");

						newSlicerDrawing.Slicer = slicer;
						newSlicerDrawing.Name = newSlicerDrawing.Slicer.Name;
					}
				}
			}
		}

		private void CopySlicerDrawing(ExcelWorksheet originalWorksheet, ExcelWorksheet newWorksheet, ExcelDrawing draw)
		{
			var uriSlicerCacheFull = XmlHelper.GetNewUri(this.Package.Package, "/xl/slicerCaches/slicerCache{0}.xml");
			var uriSlicerCache = new Uri(uriSlicerCacheFull.ToString().Substring(4), UriKind.Relative);
			var slicerCachePart = this.Package.Package.CreatePart(uriSlicerCacheFull, "application/vnd.ms-excel.slicerCache+xml", this.Package.Compression);
			var slicerCacheRelationship = newWorksheet.Workbook.Part.CreateRelationship(uriSlicerCache, Packaging.TargetMode.Internal, ExcelPackage.schemaSlicerCache);
			StreamWriter streamSlicerCache = new StreamWriter(slicerCachePart.GetStream(FileMode.Create, FileAccess.Write));
			streamSlicerCache.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
			streamSlicerCache.Write(originalWorksheet.Slicers.Slicers.First(originalSlicer => originalSlicer.Name == draw.Name).SlicerCache.TopNode.OuterXml);
			streamSlicerCache.Flush();
			var newCacheDocument = originalWorksheet.Workbook.Package.GetXmlFromUri(uriSlicerCacheFull);
			var slicerCacheNode = newCacheDocument.SelectSingleNode("default:slicerCacheDefinition", ExcelSlicer.SlicerDocumentNamespaceManager);
			var slicerCache = new ExcelSlicerCache(slicerCacheNode, ExcelSlicer.SlicerDocumentNamespaceManager, uriSlicerCache, newCacheDocument);
			// We don't have the copied pivot tables yet, but we do know that any pivot tables on the old sheet will be cloned so the PivotTable assocations
			// need to be updated in the SlicerCache.
			var oldSheetId = originalWorksheet.SheetID.ToString();
			var newSheetId = newWorksheet.SheetID.ToString();
			foreach (var pivotTable in slicerCache.PivotTables)
			{
				if (pivotTable.TabId == oldSheetId)
					pivotTable.TabId = newSheetId;
			}
			originalWorksheet.Workbook.SlicerCaches.Add(slicerCache);
			var newWorkbookSlicerCachesNode = newWorksheet.Workbook.TopNode.SelectSingleNode("d:extLst/d:ext/x14:slicerCaches", newWorksheet.Workbook.NameSpaceManager);
			var newWorkbookSlicerCacheNode = newWorkbookSlicerCachesNode.SelectSingleNode("x14:slicerCache", newWorksheet.Workbook.NameSpaceManager).CloneNode(false);
			newWorkbookSlicerCachesNode.AppendChild(newWorkbookSlicerCacheNode);
			newWorkbookSlicerCacheNode.Attributes["r:id"].Value = slicerCacheRelationship.Id;
			// We don't know the new PivotTableName yet, so updating that must done later, when PivotTables are copied.
		}

		private string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet, bool isChart)
		{
			//Create the relationship between the workbook and the new worksheet
			var rel = this.Package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(this.Package.Workbook.WorkbookUri, uriWorksheet), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/" + (isChart ? "chartsheet" : "worksheet"));
			this.Package.Package.Flush();

			//Create the new sheet node
			XmlElement worksheetNode = this.Package.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
			worksheetNode.SetAttribute("name", Name);
			worksheetNode.SetAttribute("sheetId", sheetID.ToString());
			worksheetNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

			this.TopNode.AppendChild(worksheetNode);
			return rel.Id;
		}

		private void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet, bool isChart)
		{
			Name = this.ValidateAndFixSheetName(Name);
			sheetID = this.Any() ? this.Max(ws => ws.SheetID) + 1 : 1;
			var uriId = sheetID;


			// get the next available worhsheet uri
			do
			{
				if (isChart)
				{
					uriWorksheet = new Uri("/xl/chartsheets/chartsheet" + uriId + ".xml", UriKind.Relative);
				}
				else
				{
					uriWorksheet = new Uri("/xl/worksheets/sheet" + uriId + ".xml", UriKind.Relative);
				}

				uriId++;
			} while (this.Package.Package.PartExists(uriWorksheet));
		}

		private void DeleteRelationsAndParts(Packaging.ZipPackagePart part)
		{
			var rels = part.GetRelationships().ToList();
			for (int i = 0; i < rels.Count; i++)
			{
				var rel = rels[i];
				if (rel.RelationshipType != ExcelPackage.schemaImage &&
						this.Package.Package.TryGetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri), out Packaging.ZipPackagePart relPart))
					this.DeleteRelationsAndParts(relPart);
				part.DeleteRelationship(rel.Id);
			}
			this.Package.Package.DeletePart(part.Uri);
		}

		private void ReindexWorksheetDictionary()
		{
			var index = 1;
			var worksheets = new Dictionary<int, ExcelWorksheet>();
			foreach (var entry in this.Worksheets)
			{
				entry.Value.PositionID = index;
				worksheets.Add(index++, entry.Value);
			}
			this.Worksheets = worksheets;
		}

		private ExcelWorksheet GetByName(string Name)
		{
			if (string.IsNullOrEmpty(Name)) return null;
			ExcelWorksheet xlWorksheet = null;
			foreach (ExcelWorksheet worksheet in this.Worksheets.Values)
			{
				if (worksheet.Name.Equals(Name, StringComparison.InvariantCultureIgnoreCase))
					xlWorksheet = worksheet;
			}
			return (xlWorksheet);
		}

		private void Move(string sourceName, string targetName, bool placeAfter)
		{
			var sourceSheet = this[sourceName];
			if (sourceSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
			}
			var targetSheet = this[targetName];
			if (targetSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", targetName));
			}
			this.Move(sourceSheet.PositionID, targetSheet.PositionID, placeAfter);
		}

		private void Move(int sourcePositionId, int targetPositionId, bool placeAfter)
		{
			// Bugfix: if source and target are the same worksheet the following code will create a duplicate
			//         which will cause a corrupt workbook. /swmal 2014-05-10
			if (sourcePositionId == targetPositionId) return;

			lock (this.Worksheets)
			{
				var sourceSheet = this[sourcePositionId];
				if (sourceSheet == null)
				{
					throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", sourcePositionId));
				}
				var targetSheet = this[targetPositionId];
				if (targetSheet == null)
				{
					throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", targetPositionId));
				}
				if (sourcePositionId == targetPositionId && this.Worksheets.Count < 2)
				{
					return;     //--- no reason to attempt to re-arrange a single item with itself
				}

				var index = 1;
				var newOrder = new Dictionary<int, ExcelWorksheet>();
				foreach (var entry in this.Worksheets)
				{
					if (entry.Key == targetPositionId)
					{
						if (!placeAfter)
						{
							sourceSheet.PositionID = index;
							newOrder.Add(index++, sourceSheet);
						}

						entry.Value.PositionID = index;
						newOrder.Add(index++, entry.Value);

						if (placeAfter)
						{
							sourceSheet.PositionID = index;
							newOrder.Add(index++, sourceSheet);
						}
					}
					else if (entry.Key == sourcePositionId)
					{
						//--- do nothing
					}
					else
					{
						entry.Value.PositionID = index;
						newOrder.Add(index++, entry.Value);
					}
				}
				this.Worksheets = newOrder;

				this.MoveSheetXmlNode(sourceSheet, targetSheet, placeAfter);
			}
		}

		private void MoveSheetXmlNode(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, bool placeAfter)
		{
			lock (this.TopNode.OwnerDocument)
			{
				var sourceNode = this.TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", sourceSheet.SheetID), this.NamespaceManager);
				var targetNode = this.TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", targetSheet.SheetID), this.NamespaceManager);
				if (sourceNode == null || targetNode == null)
				{
					throw new Exception("Source SheetId and Target SheetId must be valid");
				}
				if (placeAfter)
				{
					this.TopNode.InsertAfter(sourceNode, targetNode);
				}
				else
				{
					this.TopNode.InsertBefore(sourceNode, targetNode);
				}
			}
		}
		#endregion

		#region Private Static Methods
		private static bool IsInvalidSheetName(string Name)
		{
			return System.Text.RegularExpressions.Regex.IsMatch(Name, @":|\?|/|\\|\[|\]");
		}
		#endregion
	}
}
