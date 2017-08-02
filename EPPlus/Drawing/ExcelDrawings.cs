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
 * Jan Källman		                Initial Release		        2009-12-22
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing
{
	/// <summary>
	/// Collection for Drawing objects.
	/// </summary>
	public class ExcelDrawings : IEnumerable<ExcelDrawing>, IDisposable
	{
		#region Class Variables
		private XmlDocument _drawingsXml = new XmlDocument();
		private List<ExcelDrawing> _drawings;
		private XmlNamespaceManager _nsManager = null;
		private Packaging.ZipPackagePart _part = null;
		private Uri _uriDrawing = null;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the collection of drawing hash codes.
		/// </summary>
		internal Dictionary<string, string> Hashes { get; set; } = new Dictionary<string, string>();

		/// <summary>
		/// Gets or sets the <see cref="ExcelPackage"/> that these drawings exist within.
		/// </summary>
		internal ExcelPackage Package { get; set; }

		/// <summary>
		/// Gets or sets the relationship between this drawings object and the worksheet.
		/// </summary>
		internal Packaging.ZipPackageRelationship DrawingRelationship { get; set; } = null;

		/// <summary>
		/// Gets or sets the worksheet that these drawings are drawn on.
		/// </summary>
		internal ExcelWorksheet Worksheet { get; set; }

		/// <summary>
		/// Provides access to a namespace manager instance to allow XPath searching
		/// </summary>
		public XmlNamespaceManager NameSpaceManager
		{
			get
			{
				return this._nsManager;
			}
		}

		/// <summary>
		/// A reference to the drawing xml document
		/// </summary>
		public XmlDocument DrawingXml
		{
			get
			{
				return this._drawingsXml;
			}
		}

		public int Count
		{
			get
			{
				if (this._drawings == null)
				{
					return 0;
				}
				else
				{
					return this._drawings.Count;
				}
			}
		}

		internal Packaging.ZipPackagePart Part
		{
			get
			{
				return this._part;
			}
		}

		public Uri UriDrawing
		{
			get
			{
				return this._uriDrawing;
			}
		}
		#endregion

		#region Nested Classes
		internal class ImageCompare
		{
			internal byte[] image { get; set; }
			internal string relID { get; set; }

			internal bool Comparer(byte[] compareImg)
			{
				if (compareImg.Length != image.Length)
				{
					return false;
				}

				for (int i = 0; i < image.Length; i++)
				{
					if (image[i] != compareImg[i])
					{
						return false;
					}
				}
				return true;
			}
		}
		#endregion

		#region Constructors
		internal ExcelDrawings(ExcelPackage xlPackage, ExcelWorksheet sheet)
		{
			this._drawingsXml = new XmlDocument();
			this._drawingsXml.PreserveWhitespace = false;
			this._drawings = new List<ExcelDrawing>();
			this.Package = xlPackage;
			this.Worksheet = sheet;
			XmlNode node = sheet.WorksheetXml.SelectSingleNode("//d:drawing", sheet.NameSpaceManager);
			this.CreateNSM();
			if (node != null)
			{
				this.DrawingRelationship = sheet.Part.GetRelationship(node.Attributes["r:id"].Value);
				this._uriDrawing = UriHelper.ResolvePartUri(sheet.WorksheetUri, this.DrawingRelationship.TargetUri);

				this._part = xlPackage.Package.GetPart(this._uriDrawing);
				XmlHelper.LoadXmlSafe(this._drawingsXml, this._part.GetStream());

				this.AddDrawings();
			}
		}
		#endregion

		#region Private Methods
		private XmlElement CreateDrawingXml()
		{
			if (this.DrawingXml.OuterXml == "")
			{
				this.DrawingXml.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{0}\" xmlns:a=\"{1}\" />", ExcelPackage.schemaSheetDrawings, ExcelPackage.schemaDrawings));
				this._uriDrawing = new Uri(string.Format("/xl/drawings/drawing{0}.xml", Worksheet.SheetID), UriKind.Relative);

				Packaging.ZipPackage package = Worksheet.Package.Package;
				this._part = package.CreatePart(_uriDrawing, "application/vnd.openxmlformats-officedocument.drawing+xml", Package.Compression);

				StreamWriter streamChart = new StreamWriter(this._part.GetStream(FileMode.Create, FileAccess.Write));
				this.DrawingXml.Save(streamChart);
				streamChart.Close();
				package.Flush();

				this.DrawingRelationship = this.Worksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(Worksheet.WorksheetUri, _uriDrawing), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
				XmlElement e = Worksheet.WorksheetXml.CreateElement("drawing", ExcelPackage.schemaMain);
				e.SetAttribute("id", ExcelPackage.schemaRelationships, DrawingRelationship.Id);

				this.Worksheet.WorksheetXml.DocumentElement.AppendChild(e);
				package.Flush();
			}
			XmlNode colNode = this._drawingsXml.SelectSingleNode("//xdr:wsDr", NameSpaceManager);
			XmlElement drawNode;
			if (this.Worksheet is ExcelChartsheet)
			{
				drawNode = this._drawingsXml.CreateElement("xdr", "absoluteAnchor", ExcelPackage.schemaSheetDrawings);
				XmlElement posNode = this._drawingsXml.CreateElement("xdr", "pos", ExcelPackage.schemaSheetDrawings);
				posNode.SetAttribute("y", "0");
				posNode.SetAttribute("x", "0");
				drawNode.AppendChild(posNode);
				XmlElement extNode = this._drawingsXml.CreateElement("xdr", "ext", ExcelPackage.schemaSheetDrawings);
				extNode.SetAttribute("cy", "6072876");
				extNode.SetAttribute("cx", "9299263");
				drawNode.AppendChild(extNode);
				colNode.AppendChild(drawNode);
			}
			else
			{
				drawNode = this._drawingsXml.CreateElement("xdr", "twoCellAnchor", ExcelPackage.schemaSheetDrawings);
				colNode.AppendChild(drawNode);
				//Add from position Element;
				XmlElement fromNode = this._drawingsXml.CreateElement("xdr", "from", ExcelPackage.schemaSheetDrawings);
				drawNode.AppendChild(fromNode);
				fromNode.InnerXml = "<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>";

				//Add to position Element;
				XmlElement toNode = this._drawingsXml.CreateElement("xdr", "to", ExcelPackage.schemaSheetDrawings);
				drawNode.AppendChild(toNode);
				toNode.InnerXml = "<xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>";
			}

			return drawNode;
		}

		private void AddDrawings()
		{
			// Look inside all children for the drawings because they could be inside
			// Markup Compatibility AlternativeContent/Choice or AlternativeContent/Fallback nodes.
			// The code below currently pretends that loading all Choice alternative drawings doesn't cause a problem
			// elsewhere. This seems to be ok for the time being as encountered drawing files so far only seem to have
			// one Choice node (and no Fallback) underneath the AlternativeContent node. (Excel 2013 that is.)
			// This change prevents CodePlex issue #15028 from occurring. 
			// (the drawing xml part (that ONLY contained AlternativeContent nodes) was incorrectly being garbage collected when the package was saved)
			XmlNodeList list = this.DrawingXml.SelectNodes("//*[self::xdr:twoCellAnchor or self::xdr:oneCellAnchor or self::xdr:absoluteAnchor]", NameSpaceManager);

			foreach (XmlNode node in list)
			{
				ExcelDrawing drawing;
				switch (node.LocalName)
				{
					case "oneCellAnchor":
					//dr = new ExcelDrawing(this, node, "xdr:sp/xdr:nvSpPr/xdr:cNvPr/@name");                        
					case "twoCellAnchor":
					case "absoluteAnchor":
						drawing = ExcelDrawing.GetDrawing(this, node);
						break;
					default: //"absoluteCellAnchor":
						drawing = null;
						break;
				}
				if (drawing != null)
				{
					this._drawings.Add(drawing);
					if (!this.Worksheet.Workbook.NextSlicerIdNumber.ContainsKey(drawing.Name))
						this.Worksheet.Workbook.NextSlicerIdNumber[drawing.Name] = 1;
				}
			}
		}

		/// <summary>
		/// Creates the NamespaceManager. 
		/// </summary>
		private void CreateNSM()
		{
			NameTable nt = new NameTable();
			this._nsManager = new XmlNamespaceManager(nt);
			this._nsManager.AddNamespace("a", ExcelPackage.schemaDrawings);
			this._nsManager.AddNamespace("xdr", ExcelPackage.schemaSheetDrawings);
			this._nsManager.AddNamespace("c", ExcelPackage.schemaChart);
			this._nsManager.AddNamespace("r", ExcelPackage.schemaRelationships);
			this._nsManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
			this._nsManager.AddNamespace("sle", ExcelPackage.schemaSlicerDrawing);
		}

		#endregion

		#region Public Methods
		public IEnumerator GetEnumerator()
		{
			return (this._drawings.GetEnumerator());
		}

		IEnumerator<ExcelDrawing> IEnumerable<ExcelDrawing>.GetEnumerator()
		{
			return (this._drawings.GetEnumerator());
		}

		/// <summary>
		/// Add a new chart to the worksheet.
		/// Do not support Bubble-, Radar-, Stock- or Surface charts. 
		/// </summary>
		/// <param name="name"></param>
		/// <param name="chartType">Type of chart</param>
		/// <param name="pivotTableSource">The pivottable source for a pivotchart</param>    
		/// <returns>The chart</returns>
		public ExcelChart AddChart(string name, eChartType chartType, ExcelPivotTable pivotTableSource)
		{
			if (chartType == eChartType.StockHLC || chartType == eChartType.StockOHLC || chartType == eChartType.StockVOHLC)
				throw (new NotImplementedException("Chart type is not supported in the current version"));
			if (this.Worksheet is ExcelChartsheet && this._drawings.Count > 0)
				throw new InvalidOperationException("Chart Worksheets can't have more than one chart");
			XmlElement drawNode = this.CreateDrawingXml();
			ExcelChart chart = ExcelChart.GetNewChart(this, drawNode, chartType, null, pivotTableSource);
			chart.Name = name;
			this._drawings.Add(chart);
			return chart;
		}

		/// <summary>
		/// Add a new chart to the worksheet.
		/// Do not support Bubble-, Radar-, Stock- or Surface charts. 
		/// </summary>
		/// <param name="name"></param>
		/// <param name="chartType">Type of chart</param>
		/// <returns>The chart</returns>
		public ExcelChart AddChart(string name, eChartType chartType)
		{
			return this.AddChart(name, chartType, null);
		}

		/// <summary>
		/// Add a picure to the worksheet
		/// </summary>
		/// <param name="name"></param>
		/// <param name="image">An image. Allways saved in then JPeg format</param>
		/// <returns></returns>
		public ExcelPicture AddPicture(string name, Image image)
		{
			return this.AddPicture(name, image, null);
		}

		/// <summary>
		/// Add a picure to the worksheet
		/// </summary>
		/// <param name="name"></param>
		/// <param name="image">An image. Allways saved in then JPeg format</param>
		/// <param name="hyperlink">Picture Hyperlink</param>
		/// <returns></returns>
		public ExcelPicture AddPicture(string name, Image image, Uri hyperlink)
		{
			if (image != null)
			{
				XmlElement drawNode = CreateDrawingXml();
				drawNode.SetAttribute("editAs", "oneCell");
				ExcelPicture picture = new ExcelPicture(this, drawNode, image, hyperlink);
				picture.Name = name;
				this._drawings.Add(picture);
				return picture;
			}
			throw (new Exception("AddPicture: Image can't be null"));
		}

		/// <summary>
		/// Add a picure to the worksheet
		/// </summary>
		/// <param name="name"></param>
		/// <param name="imageFile">The image file</param>
		/// <returns></returns>
		public ExcelPicture AddPicture(string name, FileInfo imageFile)
		{
			return this.AddPicture(name, imageFile, null);
		}

		/// <summary>
		/// Add a picure to the worksheet
		/// </summary>
		/// <param name="name"></param>
		/// <param name="imageFile">The image file</param>
		/// <param name="hyperlink">Picture Hyperlink</param>
		/// <returns></returns>
		public ExcelPicture AddPicture(string name, FileInfo imageFile, Uri hyperlink)
		{
			if (this.Worksheet is ExcelChartsheet && this._drawings.Count > 0)
				throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
			if (imageFile != null)
			{
				XmlElement drawNode = this.CreateDrawingXml();
				drawNode.SetAttribute("editAs", "oneCell");
				ExcelPicture picture = new ExcelPicture(this, drawNode, imageFile, hyperlink);
				picture.Name = name;
				this._drawings.Add(picture);
				return picture;
			}
			throw (new Exception("AddPicture: ImageFile can't be null"));
		}

		/// <summary>
		/// Add a new shape to the worksheet
		/// </summary>
		/// <param name="name">Name</param>
		/// <param name="style">Shape style</param>
		/// <returns>The shape object.</returns>
		public ExcelShape AddShape(string name, eShapeStyle style)
		{
			if (this.Worksheet is ExcelChartsheet && this._drawings.Count > 0)
				throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
			XmlElement drawNode = this.CreateDrawingXml();
			ExcelShape shape = new ExcelShape(this, drawNode, style);
			shape.Name = name;
			shape.Style = style;
			this._drawings.Add(shape);
			return shape;
		}

		/// <summary>
		/// Add a new shape to the worksheet
		/// </summary>
		/// <param name="name">Name</param>
		/// <param name="source">Source shape</param>
		/// <returns>The shape object</returns>
		public ExcelShape AddShape(string name, ExcelShape source)
		{
			if (this.Worksheet is ExcelChartsheet && this._drawings.Count > 0)
				throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
			XmlElement drawNode = this.CreateDrawingXml();
			drawNode.InnerXml = source.TopNode.InnerXml;
			ExcelShape shape = new ExcelShape(this, drawNode);
			shape.Name = name;
			shape.Style = source.Style;
			this._drawings.Add(shape);
			return shape;
		}

		/// <summary>
		/// Removes a drawing.
		/// </summary>
		/// <param name="index">The index of the <see cref="ExcelDrawing"/> to remove.</param>
		public void Remove(int index)
		{
			this.RemoveDrawing(this._drawings[index]);
		}

		/// <summary>
		/// Removes a drawing.
		/// </summary>
		/// <param name="drawing">The <see cref="ExcelDrawing"/> to remove.</param>
		public void Remove(ExcelDrawing drawing)
		{
			this.RemoveDrawing(drawing);
		}

		/// <summary>
		/// Removes a drawing.
		/// </summary>
		/// <param name="name">The name of the <see cref="ExcelDrawing"/> to remove.</param>
		public void Remove(string name)
		{
			this.RemoveDrawing(this[name]);
		}

		/// <summary>
		/// Removes all drawings from the collection
		/// </summary>
		public void Clear()
		{
			if (this.Worksheet is ExcelChartsheet && this._drawings.Count > 0)
				throw new InvalidOperationException("Can't remove charts from chart worksheets.");
			this.ClearDrawings();
		}

		public void Dispose()
		{
			this._drawingsXml = null;
			this.Hashes.Clear();
			this.Hashes = null;
			this._part = null;
			this.DrawingRelationship = null;
			foreach (var drawing in this._drawings)
			{
				drawing.Dispose();
			}
			this._drawings.Clear();
			this._drawings = null;
		}
		#endregion

		#region Internal Methods
		internal void RemoveDrawing(ExcelDrawing drawing)
		{
			if (this.Worksheet is ExcelChartsheet && this._drawings.Count > 0)
				throw new InvalidOperationException("Can't remove charts from chart worksheets.");
			else if (this._drawings.Contains(drawing))
			{
				drawing.DeleteMe();
				this._drawings.Remove(drawing);
			}
		}

		internal void ClearDrawings()
		{
			while (this.Count > 0)
			{
				this.RemoveDrawing(this._drawings[0]);
			}
		}

		internal void AdjustWidth(int[,] pos)
		{
			var ix = 0;
			//Now set the size for all drawings depending on the editAs property.
			foreach (OfficeOpenXml.Drawing.ExcelDrawing d in this)
			{
				if (d.EditAs != Drawing.eEditAs.TwoCell)
				{
					if (d.EditAs == Drawing.eEditAs.Absolute)
					{
						d.SetPixelLeft(pos[ix, 0]);
					}
					d.SetPixelWidth(pos[ix, 1]);

				}
				ix++;
			}
		}

		internal void AdjustHeight(int[,] pos)
		{
			var ix = 0;
			//Now set the size for all drawings depending on the editAs property.
			foreach (OfficeOpenXml.Drawing.ExcelDrawing d in this)
			{
				if (d.EditAs != Drawing.eEditAs.TwoCell)
				{
					if (d.EditAs == Drawing.eEditAs.Absolute)
						d.SetPixelTop(pos[ix, 0]);
					d.SetPixelHeight(pos[ix, 1]);
				}
				ix++;
			}
		}

		internal int[,] GetDrawingWidths()
		{
			int[,] pos = new int[Count, 2];
			int ix = 0;
			//Save the size for all drawings
			foreach (ExcelDrawing d in this)
			{
				pos[ix, 0] = d.GetPixelLeft();
				pos[ix++, 1] = d.GetPixelWidth();
			}
			return pos;
		}

		internal int[,] GetDrawingHeight()
		{
			int[,] pos = new int[Count, 2];
			int ix = 0;
			//Save the size for all drawings
			foreach (ExcelDrawing d in this)
			{
				pos[ix, 0] = d.GetPixelTop();
				pos[ix++, 1] = d.GetPixelHeight();
			}
			return pos;
		}
		#endregion

		#region Public Operators
		/// <summary>
		/// Returns the drawing at the specified position.  
		/// </summary>
		/// <param name="PositionID">The position of the drawing. 0-base</param>
		/// <returns></returns>
		public ExcelDrawing this[int PositionID]
		{
			get
			{
				return (this._drawings[PositionID]);
			}
		}

		/// <summary>
		/// Returns the <see cref="ExcelDrawing"/> matching the specified name.
		/// </summary>
		/// <param name="Name">The name of the desired <see cref="ExcelDrawing"/>.</param>
		/// <returns>
		///		Returns the first <see cref="ExcelDrawing"/> with the given <paramref name="Name"/>,
		///		or null if no <see cref="ExcelDrawing"/> exists with the given <paramref name="Name"/>.
		///	</returns>
		public ExcelDrawing this[string Name]
		{
			get
			{
				var drawingsWithGivenName = this._drawings.Where((ExcelDrawing drawing) => drawing.Name.Equals(Name));
				if (drawingsWithGivenName.Count() > 0)
					return drawingsWithGivenName.First();
				else
					return null;
			}
		}
		#endregion
	}
}
