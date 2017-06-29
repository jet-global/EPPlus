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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
	internal class ExcelVmlDrawingCommentCollection : ExcelVmlDrawingBaseCollection, IEnumerable
	{
		private Dictionary<ulong, ExcelVmlDrawingComment> myDrawings;

		private Dictionary<ulong, ExcelVmlDrawingComment> Drawings
		{
			get
			{
				if (this.myDrawings == null)
					this.myDrawings = new Dictionary<ulong, ExcelVmlDrawingComment>();
				return this.myDrawings;
			}
		}

		private ExcelWorksheet Worksheet { get; set; }
		internal ExcelVmlDrawingCommentCollection(ExcelPackage package, ExcelWorksheet worksheet, Uri uri) :
			 base(package, worksheet, uri)
		{
			this.Worksheet = worksheet;
			if (uri == null)
				this.VmlDrawingXml.LoadXml(this.CreateVmlDrawings());
			else
				this.AddDrawingsFromXml(worksheet);
		}
		protected void AddDrawingsFromXml(ExcelWorksheet worksheet)
		{
			XmlNodeList nodes = this.VmlDrawingXml.SelectNodes("//v:shape", this.NameSpaceManager);
			foreach (XmlNode node in nodes)
			{
				XmlNode rowNode = node.SelectSingleNode("x:ClientData/x:Row", this.NameSpaceManager);
				XmlNode colNode = node.SelectSingleNode("x:ClientData/x:Column", this.NameSpaceManager);
				if (rowNode != null && colNode != null)
				{
					int row = int.Parse(rowNode.InnerText) + 1;
					int col = int.Parse(colNode.InnerText) + 1;
					ulong key = ExcelCellBase.GetCellID(worksheet.SheetID, row, col);
					this.Drawings.Add(key, new ExcelVmlDrawingComment(node, worksheet.Cells[row, col], base.NameSpaceManager));
				}
				else
				{
					ulong cellId = ExcelCellBase.GetCellID(worksheet.SheetID, 1, 1);
					this.Drawings.Add(cellId, new ExcelVmlDrawingComment(node, worksheet.Cells[1, 1], base.NameSpaceManager));
				}
			}
		}
		private string CreateVmlDrawings()
		{
			string vml = string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">",
				 ExcelPackage.schemaMicrosoftVml,
				 ExcelPackage.schemaMicrosoftOffice,
				 ExcelPackage.schemaMicrosoftExcel);

			vml += "<o:shapelayout v:ext=\"edit\">";
			vml += "<o:idmap v:ext=\"edit\" data=\"1\"/>";
			vml += "</o:shapelayout>";

			vml += "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">";
			vml += "<v:stroke joinstyle=\"miter\" />";
			vml += "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />";
			vml += "</v:shapetype>";
			vml += "</xml>";

			return vml;
		}
		internal ExcelVmlDrawingComment Add(ExcelRangeBase cell, XmlNode drawingNode = null)
		{
			if (drawingNode == null)
				drawingNode = this.CreateDrawing(cell);
			else
				this.AddDrawing(drawingNode, cell);
			var drawing = new ExcelVmlDrawingComment(drawingNode, cell, this.NameSpaceManager);
			ulong cellId = ExcelCellBase.GetCellID(cell.Worksheet.SheetID, cell._fromRow, cell._fromCol);
			this.Drawings.Add(cellId, drawing);
			return drawing;
		}
		internal void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			for (int i = fromRow; i < (fromRow + rows); i++)
			{
				for (int j = fromCol; j < (fromCol + columns); j++)
				{
					ulong key = ExcelCellBase.GetCellID(this.Worksheet.SheetID, i, j);
					if (this.Drawings.ContainsKey(key))
						this.Drawings.Remove(key);
				}
			}
		}
		private XmlNode CreateDrawing(ExcelRangeBase cell)
		{
			int row = cell.Start.Row, col = cell.Start.Column;
			var node = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);

			this.InsertDrawingNode(cell, node);

			node.SetAttribute("id", GetNewId());
			node.SetAttribute("type", "#_x0000_t202");
			node.SetAttribute("style", "position:absolute;z-index:1; visibility:hidden");
			node.SetAttribute("fillcolor", "#ffffe1");
			node.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");

			string vml = "<v:fill color2=\"#ffffe1\" />";
			vml += "<v:shadow on=\"t\" color=\"black\" obscured=\"t\" />";
			vml += "<v:path o:connecttype=\"none\" />";
			vml += "<v:textbox style=\"mso-direction-alt:auto\">";
			vml += "<div style=\"text-align:left\" />";
			vml += "</v:textbox>";
			vml += "<x:ClientData ObjectType=\"Note\">";
			vml += "<x:MoveWithCells />";
			vml += "<x:SizeWithCells />";
			vml += string.Format("<x:Anchor>{0}, 15, {1}, 2, {2}, 31, {3}, 1</x:Anchor>", col, row - 1, col + 2, row + 3);
			vml += "<x:AutoFill>False</x:AutoFill>";
			vml += string.Format("<x:Row>{0}</x:Row>", row - 1); ;
			vml += string.Format("<x:Column>{0}</x:Column>", col - 1);
			vml += "</x:ClientData>";

			node.InnerXml = vml;
			return node;
		}
		private void AddDrawing(XmlNode node, ExcelRangeBase cell)
		{
			node.Attributes["id"].Value = this.GetNewId();
			node = VmlDrawingXml.ImportNode(node, true);
			this.InsertDrawingNode(cell, node);
		}
		private void InsertDrawingNode(ExcelRangeBase cell, XmlNode node)
		{
			ulong cellId = ExcelCellBase.GetCellID(cell.Worksheet.SheetID, cell._fromRow, cell._fromCol);
			List<ulong> nodes = this.Drawings.Keys.ToList();
			nodes.Sort((cellId1, cellId2) => cellId1.CompareTo(cellId2));
			int index = nodes.BinarySearch(cellId);
			if ((index < 0) && (~index < nodes.Count))
			{
				index = ~index;
				ExcelVmlDrawingBase previousDrawing = this.Drawings[nodes[index]];
				previousDrawing.TopNode.ParentNode.InsertBefore(node, previousDrawing.TopNode);
			}
			else
				this.VmlDrawingXml.DocumentElement.AppendChild(node);
		}
		int _nextID = 1024;
		/// <summary>
		/// returns the next drawing id.
		/// </summary>
		/// <returns></returns>
		internal string GetNewId()
		{
			const string idFormat = "_x0000_s";
			if (_nextID == 1024)
			{
				foreach (ExcelVmlDrawingComment draw in this)
				{
					if (draw.Id.Length >= 11 && draw.Id.StartsWith(idFormat))
					{
						int id;
						if (int.TryParse(draw.Id.Substring(draw.Id.Length - 4, 4), out id) && id > _nextID)
							_nextID = id;
					}
				}
			}
			_nextID++;
			return idFormat + _nextID.ToString();
		}
		internal ExcelVmlDrawingBase this[ulong cellId]
		{
			get
			{
				return this.Drawings[cellId] as ExcelVmlDrawingComment;
			}
		}
		internal bool ContainsKey(ulong cellId)
		{
			return this.Drawings.ContainsKey(cellId);
		}
		internal int Count
		{
			get
			{
				return this.Drawings.Count;
			}
		}
		#region IEnumerable Members

		#endregion

		public IEnumerator GetEnumerator()
		{
			return this.Drawings.Values.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.Drawings.Values.GetEnumerator();
		}
	}
}
