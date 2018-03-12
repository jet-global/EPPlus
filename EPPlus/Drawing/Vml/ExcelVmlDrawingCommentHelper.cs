using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Vml
{
	/// <summary>
	/// A helper class that creates VML drawings for comments on a worksheet.
	/// </summary>
	public static class ExcelVmlDrawingCommentHelper
	{
		#region Public Static Methods
		/// <summary>
		/// Create and/or update vmlDrawings[n].vml for the specified <paramref name="commentCollection"/> on
		/// the specified <paramref name="sheet"/>.
		/// </summary>
		/// <param name="sheet">The worksheet that the comments belong to.</param>
		/// <param name="commentCollection">The comments to add drawings for.</param>
		public static void AddCommentDrawings(ExcelWorksheet sheet, ExcelCommentCollection commentCollection)
		{
			if (sheet == null)
				throw new ArgumentNullException(nameof(sheet));
			if (commentCollection == null)
				throw new ArgumentNullException(nameof(commentCollection));
			var vmlDrawingsUri = XmlHelper.GetNewUri(sheet.Package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
			XmlDocument vmlDocumentXml = new XmlDocument();
			if (sheet.Package.Package.TryGetPart(vmlDrawingsUri, out var vmlDrawingsPart))
			{
				vmlDocumentXml.Load(vmlDrawingsPart.GetStream());
				var nodesToDelete = new List<XmlNode>();
				var shapeNodes = vmlDocumentXml.SelectNodes("v:shape");
				foreach (XmlNode node in shapeNodes)
				{
					if (node.Attributes?["type"].Value == "#_x0000_t202")
						nodesToDelete.Add(node);
				}
				foreach (var node in nodesToDelete)
				{
					vmlDocumentXml.RemoveChild(node);
				}
				if (commentCollection.Count > 0)
					ExcelVmlDrawingCommentHelper.RemoveLegacyDrawingRel(sheet);
			}
			else
			{
				vmlDrawingsPart = sheet.Package.Package.CreatePart(vmlDrawingsUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", sheet.Package.Compression);
				var rel = sheet.Part.CreateRelationship(UriHelper.GetRelativeUri(sheet.WorksheetUri, vmlDrawingsUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
				sheet.SetXmlNodeString("d:legacyDrawing/@r:id", rel.Id);
				vmlDocumentXml.LoadXml(ExcelVmlDrawingCommentHelper.CreateVmlDrawings());
			}
			int id = 1024;
			foreach (ExcelComment comment in commentCollection)
			{
				ExcelVmlDrawingCommentHelper.CreateDrawing(vmlDocumentXml, comment.Range, id++);
			}
			vmlDocumentXml.Save(vmlDrawingsPart.GetStream(FileMode.Create));
		}
		#endregion

		#region Private Methods
		private static string CreateVmlDrawings()
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

		private static void CreateDrawing(XmlDocument xmlDocument, ExcelRangeBase cell, int id)
		{
			int row = cell.Start.Row, col = cell.Start.Column;
			var node = xmlDocument.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
			xmlDocument.DocumentElement.AppendChild(node);
			node.SetAttribute("id", $"_x0000_s{id}");
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
			vml += $"<x:Anchor>{col}, 15, { row - 1}, 2, {col + 2}, 31, {row + 3}, 1</x:Anchor>";
			vml += "<x:AutoFill>False</x:AutoFill>";
			vml += $"<x:Row>{row - 1}</x:Row>";
			vml += $"<x:Column>{col - 1}</x:Column>";
			vml += "</x:ClientData>";
			node.InnerXml = vml;
		}

		private static void RemoveLegacyDrawingRel(ExcelWorksheet sheet)
		{
			var vmlNode = sheet.WorksheetXml.DocumentElement.SelectSingleNode("d:legacyDrawing/@r:id", sheet.NameSpaceManager);
			if (sheet.Part.RelationshipExists(vmlNode.Value))
			{
				var rel = sheet.Part.GetRelationship(vmlNode.Value);
				var n = sheet.WorksheetXml.DocumentElement.SelectSingleNode($"d:legacyDrawing[@r:id=\"{rel.Id}\"]", sheet.NameSpaceManager);
				if (n != null)
					n.ParentNode.RemoveChild(n);
			}
		}
		#endregion
	}
}
