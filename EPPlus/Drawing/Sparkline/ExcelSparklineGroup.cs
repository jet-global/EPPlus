/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * ExcelSparklineGroup.cs Copyright (C) 2016 Matt Delaney.
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
 * Author					Change						                Date
 * ******************************************************************************
 * Matt Delaney		        Sparklines                                2016-05-20
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting;

namespace OfficeOpenXml.Drawing.Sparkline
{
	/// <summary>
	/// Represents the CT_SparklineGroup type as defined at https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx
	/// </summary>
	public class ExcelSparklineGroup : XmlHelper
	{
		#region Properties
		#region Subnodes
		/// <summary>
		/// Gets or sets the series color of the sparkline.
		/// </summary>
		public Color ColorSeries { get; set; }

		/// <summary>
		/// Gets or sets the color of negative values in this group's sparklines.
		/// </summary>
		public Color ColorNegative { get; set; }

		/// <summary>
		/// Gets or sets the color of the axes in sparklines that belong to this group.
		/// </summary>
		public Color ColorAxis { get; set; }

		/// <summary>
		/// Gets or sets the color of the markers for sparklines in this group.
		/// </summary>
		public Color ColorMarkers { get; set; }

		/// <summary>
		/// Gets or sets the color of the first value in each sparkline in this group.
		/// </summary>
		public Color ColorFirst { get; set; }

		/// <summary>
		/// Gets or sets the color of the last value in each sparkline in this group.
		/// </summary>
		public Color ColorLast { get; set; }

		/// <summary>
		/// Gets or sets the color of the highest value in each sparkline in this group.
		/// </summary>
		public Color ColorHigh { get; set; }

		/// <summary>
		/// Gets or sets the color of the lowest value in each sparkline in this group.
		/// </summary>
		public Color ColorLow { get; set; }

		/// <summary>
		/// This collection must exist and must have at least one sparkline.
		/// </summary>
		public List<ExcelSparkline> Sparklines { get; private set; } = new List<ExcelSparkline>();
		#endregion

		#region Attributes
		/// <summary>
		/// Gets or sets the maximum axis value. Optional.
		/// </summary>
		public double? ManualMax { get; set; }

		/// <summary>
		/// Gets or sets the minimum axis value. Optional.
		/// </summary>
		public double? ManualMin { get; set; }

		/// <summary>
		/// Gets or sets the line weight of the sparklines. Defaults to 0.75. Optional.
		/// </summary>
		public double? LineWeight { get; set; }

		/// <summary>
		/// Gets or sets the <see cref="SparklineType"/> of the sparklines in the group. 
		/// Defaults to <see cref="SparklineType.Line"/>.
		/// </summary>
		public SparklineType? Type { get; set; } = SparklineType.Line;

		/// <summary>
		/// Gets or sets a value indicating if the axis represents a date.
		/// </summary>
		public bool DateAxis { get; set; }

		/// <summary>
		/// Gets or sets a value indicating how blank cells should be displayed.
		/// Defaults to <see cref="DispBlanksAs.Zero"/>.
		/// </summary>
		public DispBlanksAs? DisplayEmptyCellsAs { get; set; } = DispBlanksAs.Zero;

		/// <summary>
		/// Gets or sets a value indicating if value markers should be displayed.
		/// </summary>
		public bool Markers { get; set; }

		/// <summary>
		/// Gets or sets a value that represents if the highest value in a sparkline should be indicated.
		/// </summary>
		public bool High { get; set; }

		/// <summary>
		/// Gets or sets a value that represents if the lowest value in a sparkline should be indicated.
		/// </summary>
		public bool Low { get; set; }

		/// <summary>
		/// Gets or sets a value that represents if the first value in a sparkline should be highlighted.
		/// </summary>
		public bool First { get; set; }

		/// <summary>
		/// Gets or sets a value that represents if the last value in a sparkline should be highlighted.
		/// </summary>
		public bool Last { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates if negative values should be recolored.
		/// </summary>
		public bool Negative { get; set; }

		/// <summary>
		///  Gets or sets a value that indicates if the X Axis should be displayed.
		/// </summary>
		public bool DisplayXAxis { get; set; }

		/// <summary>
		/// Gets or sets a value indicating if hidden cells in the range should be included in the sparkline.
		/// </summary>
		public bool DisplayHidden { get; set; }

		/// <summary>
		/// Gets or sets a value specifying the minimum axis grouping type <see cref="SparklineAxisMinMax"/>.
		/// </summary>
		public SparklineAxisMinMax MinAxisType { get; set; } = SparklineAxisMinMax.Individual;

		/// <summary>
		/// Gets or sets a value specifying the maximum axis grouping type <see cref="SparklineAxisMinMax"/>.
		/// </summary>
		public SparklineAxisMinMax MaxAxisType { get; set; } = SparklineAxisMinMax.Individual;

		/// <summary>
		/// Gets or sets if the values should be graphed right-to-left.
		/// </summary>
		public bool RightToLeft { get; set; }
		#endregion

		/// <summary>
		/// Gets or sets the <see cref="ExcelWorksheet"/> that this group of sparklines lives on.
		/// </summary>
		public ExcelWorksheet Worksheet { get; internal set; }
		#endregion

		#region Public Methods
		/// <summary>
		/// Save the SparklineGroup's properties and attributes to the TopNode's XML.
		/// </summary>
		public void Save()
		{
			this.SaveAttributes();
			this.SaveColors();
			this.SaveSparklines();
		}

		/// <summary>
		/// Copies the color and style contents of the given node to the current <see cref="ExcelSparklineGroup"/>, 
		/// excluding the node's contained sparklines.
		/// </summary>
		/// <param name="topNode">The node containing the color and style elements to copy.</param>
		public void CopyNodeStyle(XmlNode topNode)
		{
			// Parse the color nodes.
			var node = topNode.SelectSingleNode("x14:colorSeries", base.NameSpaceManager);
			if (node != null)
				this.ColorSeries = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorNegative", base.NameSpaceManager);
			if (node != null)
				this.ColorNegative = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorAxis", base.NameSpaceManager);
			if (node != null)
				this.ColorAxis = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorMarkers", base.NameSpaceManager);
			if (node != null)
				this.ColorMarkers = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorFirst", base.NameSpaceManager);
			if (node != null)
				this.ColorFirst = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorLast", base.NameSpaceManager);
			if (node != null)
				this.ColorLast = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorHigh", base.NameSpaceManager);
			if (node != null)
				this.ColorHigh = this.GetColorFromColorNode(node);
			node = topNode.SelectSingleNode("x14:colorLow", base.NameSpaceManager);
			if (node != null)
				this.ColorLow = this.GetColorFromColorNode(node);
			// Parse the attributes.
			var attribute = topNode.Attributes["type"];
			if (attribute != null)
				this.Type = ParseSparklineType(attribute.InnerText);
			attribute = topNode.Attributes["displayEmptyCellsAs"];
			if (attribute != null)
				this.DisplayEmptyCellsAs = ParseDisplayEmptyCellsAs(attribute.InnerText);
			this.Negative = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "negative");
			this.ManualMin = GetAttributeDouble(topNode, "manualMin");
			this.ManualMax = GetAttributeDouble(topNode, "manualMax");
			this.LineWeight = GetAttributeDouble(topNode, "lineWeight");
			this.DateAxis = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "dateAxis");
			this.Markers = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "markers");
			this.High = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "high");
			this.Low = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "low");
			this.First = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "first");
			this.Last = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "last");
			this.DisplayXAxis = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "displayXAxis");
			this.DisplayHidden = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "displayHidden");
			attribute = topNode.Attributes["minAxisType"];
			if (attribute != null)
				this.MinAxisType = ParseSparklineAxisMinMax(attribute.InnerText);
			attribute = topNode.Attributes["maxAxisType"];
			if (attribute != null)
				this.MaxAxisType = ParseSparklineAxisMinMax(attribute.InnerText);
			this.RightToLeft = ExcelConditionalFormattingHelper.GetAttributeBool(topNode, "rightToLeft");
		}
		#endregion

		#region Private Methods
		private void SaveSparklines()
		{
			if (this.Sparklines.Count == 0)
			{
				throw new InvalidOperationException("At least one sparkline must exist to save sparklines.");
			}
			var sparklinesNode = this.CreateNode("x14:sparklines");
			sparklinesNode.RemoveAll();
			foreach (var sparkline in this.Sparklines)
			{
				sparkline.Save();
				sparklinesNode.AppendChild(sparkline.TopNode);
			}
		}

		private void SaveColors()
		{
			var node = this.CreateNode("x14:colorSeries");
			this.UpdateColorNode(node, this.ColorSeries);
			this.UpdateColorNode(this.CreateNode("x14:colorNegative"), this.ColorNegative);
			this.UpdateColorNode(this.CreateNode("x14:colorAxis"), this.ColorAxis);
			this.UpdateColorNode(this.CreateNode("x14:colorMarkers"), this.ColorMarkers);
			this.UpdateColorNode(this.CreateNode("x14:colorFirst"), this.ColorFirst);
			this.UpdateColorNode(this.CreateNode("x14:colorLast"), this.ColorLast);
			this.UpdateColorNode(this.CreateNode("x14:colorHigh"), this.ColorHigh);
			this.UpdateColorNode(this.CreateNode("x14:colorLow"), this.ColorLow);
		}

		private void UpdateColorNode(XmlNode node, Color color)
		{
			if (!color.IsEmpty)
			{
				ExcelSparklineGroup.SetAttribute(node, "rgb", ExcelSparklineGroup.XmlColor(color));
			}
			else if (node.Attributes.Count == 0)
			{
				this.TopNode.RemoveChild(node);
			}
			// Else: the node has an unsupported color definition (such as a theme) and should be left alone.
		}

		private static string XmlColor(Color color)
		{
			return string.Format("{0:X2}{1:X2}{2:X2}{3:X2}", color.A, color.R, color.G, color.B);
		}

		private void SaveAttributes()
		{
			if (this.ManualMax != null)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "manualMax", this.ManualMax.ToString());
			else
				this.ClearAttribute(this.TopNode, "manualMax");
			if (this.ManualMin != null)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "manualMin", this.ManualMin.ToString());
			else
				this.ClearAttribute(this.TopNode, "manualMin");
			if (this.LineWeight != null && this.LineWeight != 0.75)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "lineWeight", this.LineWeight.ToString());
			else
				this.ClearAttribute(this.TopNode, "lineWeight");
			if (this.Type != null && this.Type != SparklineType.Line)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "type", ExcelSparklineGroup.SparklineTypeToString(this.Type.Value));
			else
				this.ClearAttribute(this.TopNode, "type");
			if (this.DateAxis)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "dateAxis", "1");
			else
				this.ClearAttribute(this.TopNode, "dateAxis");
			if (this.DisplayEmptyCellsAs != null && this.DisplayEmptyCellsAs != DispBlanksAs.Zero)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "displayEmptyCellsAs", ExcelSparklineGroup.DisplayBlanksAsToString(this.DisplayEmptyCellsAs.Value));
			else
				this.ClearAttribute(this.TopNode, "displayEmptyCellsAs");
			if (this.Markers)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "markers", "1");
			else
				this.ClearAttribute(this.TopNode, "markers");
			if (this.High)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "high", "1");
			else
				this.ClearAttribute(this.TopNode, "high");
			if (this.Low)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "low", "1");
			else
				this.ClearAttribute(this.TopNode, "low");
			if (this.First)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "first", "1");
			else
				this.ClearAttribute(this.TopNode, "first");
			if (this.Last)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "last", "1");
			else
				this.ClearAttribute(this.TopNode, "last");
			if (this.Negative)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "negative", "1");
			else
				this.ClearAttribute(this.TopNode, "negative");
			if (this.DisplayXAxis)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "displayXAxis", "1");
			else
				this.ClearAttribute(this.TopNode, "displayXAxis");
			if (this.DisplayHidden)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "displayHidden", "1");
			else
				this.ClearAttribute(this.TopNode, "displayHidden");
			if (this.MinAxisType != SparklineAxisMinMax.Individual)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "minAxisType", ExcelSparklineGroup.SparklineAxisMinMaxToString(this.MinAxisType));
			else
				this.ClearAttribute(this.TopNode, "minAxisType");
			if (this.MaxAxisType != SparklineAxisMinMax.Individual)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "maxAxisType", ExcelSparklineGroup.SparklineAxisMinMaxToString(this.MaxAxisType));
			else
				this.ClearAttribute(this.TopNode, "maxAxisType");
			if (this.RightToLeft)
				ExcelSparklineGroup.SetAttribute(this.TopNode, "rightToLeft", "1");
			else
				this.ClearAttribute(this.TopNode, "rightToLeft");
		}

		private void ClearAttribute(XmlNode node, string attribute)
		{
			node.Attributes.RemoveNamedItem(attribute);
		}

		/// <summary>
		/// Variation of <see cref="ExcelConditionalFormattingHelper.GetAttributeDouble(XmlNode, string)"/> that 
		/// returns null if the value cannot be found instead of double.NaN.
		/// </summary>
		/// <param name="topNode"></param>
		/// <param name="attribute"></param>
		/// <returns></returns>
		private static double? GetAttributeDouble(XmlNode topNode, string attribute)
		{
			var result = ExcelConditionalFormattingHelper.GetAttributeDouble(topNode, attribute);
			if (double.IsNaN(result))
				return null;
			else
				return result;
		}
		private static DispBlanksAs ParseDisplayEmptyCellsAs(string type)
		{
			if (type.Equals("gap"))
				return DispBlanksAs.Gap;
			if (type.Equals("span"))
				return DispBlanksAs.Span;
			else
				return DispBlanksAs.Zero;
		}
		private static string DisplayBlanksAsToString(DispBlanksAs type)
		{
			if (DispBlanksAs.Gap == type)
				return "gap";
			else if (DispBlanksAs.Span == type)
				return "span";
			else
				return "zero";
		}

		private static SparklineAxisMinMax ParseSparklineAxisMinMax(string type)
		{
			if (type.Equals("custom"))
				return SparklineAxisMinMax.Custom;
			if (type.Equals("group"))
				return SparklineAxisMinMax.Group;
			else
				return SparklineAxisMinMax.Individual;
		}
		private static string SparklineAxisMinMaxToString(SparklineAxisMinMax type)
		{
			if (SparklineAxisMinMax.Custom == type)
				return "custom";
			else if (SparklineAxisMinMax.Group == type)
				return "group";
			else
				return "individual";
		}

		private static SparklineType ParseSparklineType(string type)
		{
			if (type.Equals("stacked"))
				return SparklineType.Stacked;
			else if (type.Equals("column"))
				return SparklineType.Column;
			else
				return SparklineType.Line;
		}

		private static string SparklineTypeToString(SparklineType type)
		{
			if (type.Equals(SparklineType.Column))
				return "column";
			else if (type.Equals(SparklineType.Stacked))
				return "stacked";
			else
				return "line";
		}

		private Color GetColorFromColorNode(XmlNode node)
		{
			// At the moment, only RGB color codes are supported. 
			// However, other color definition types are possible, such as themes as outlined in the 
			// ExcelDxfColor class.
			if (node == null)
				return Color.Empty;
			var attribute = node.Attributes.GetNamedItem("rgb");
			if (attribute == null)
				return Color.Empty;
			var parsed = ExcelConditionalFormattingHelper.ConvertFromColorCode(attribute.InnerText);
			if (Color.White.Equals(parsed))
				return Color.Empty;
			else
				return parsed;
		}
		#endregion

		#region XmlHelper Overrides
		/// <summary>
		/// Creates a new <see cref="ExcelSparklineGroup"/> based on the specified <see cref="XmlNode"/>.
		/// </summary>
		/// <param name="worksheet">The <see cref="ExcelWorksheet"/> the <see cref="ExcelSparklineGroup"/> is defined on.</param>
		/// <param name="nameSpaceManager">The namespace manager for the object.</param>
		/// <param name="topNode">the x14:sparklineGroup node that defines the <see cref="ExcelSparklineGroup"/>.</param>
		public ExcelSparklineGroup(ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
		{
			if (worksheet == null)
				throw new ArgumentNullException(nameof(worksheet));
			if (topNode == null)
				throw new ArgumentNullException(nameof(topNode));
			this.Worksheet = worksheet;
			this.CopyNodeStyle(topNode);
			var sparklineNodes = topNode.SelectSingleNode("x14:sparklines", base.NameSpaceManager)?.ChildNodes;
			if (sparklineNodes == null)
				return;
			foreach (var sparklineNode in sparklineNodes)
			{
				this.Sparklines.Add(new ExcelSparkline(this, base.NameSpaceManager, (XmlNode)sparklineNode));
			}
		}

		/// <summary>
		/// Creates a new <see cref="ExcelSparklineGroup"/> from scratch (without an XML Node).
		/// </summary>
		/// <param name="worksheet">The worksheet to add the <see cref="ExcelSparklineGroup"/> to.</param>
		/// <param name="nameSpaceManager">The namespace manager for the <see cref="ExcelSparklineGroup"/>.</param>
		public ExcelSparklineGroup(ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
		{
			this.Worksheet = worksheet;
			base.TopNode = worksheet.TopNode.OwnerDocument.CreateElement("x14:sparklineGroup", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		}
		#endregion
	}
}
