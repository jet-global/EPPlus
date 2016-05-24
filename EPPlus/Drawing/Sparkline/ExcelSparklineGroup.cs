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
 * Author					Change						                Date
 * ******************************************************************************
 * emdelaney		        Sparklines                                2016-05-20
 *******************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using System;

namespace OfficeOpenXml.Drawing.Sparkline
{
    public class ExcelSparklineGroup : XmlHelper
    {
        #region Properties
        #region Subnodes
        public Color ColorSeries { get; set; }
        public Color ColorNegative { get; set; }
        public Color ColorAxis { get; set; }
        public Color ColorMarkers { get; set; }
        public Color ColorFirst { get; set; }
        public Color ColorLast { get; set; }
        public Color ColorHigh { get; set; }
        public Color ColorLow { get; set; }

        /// <summary>
        /// This collection must exist and must have at least one sparkline.
        /// </summary>
        public List<ExcelSparkline> Sparklines { get; private set; } = new List<ExcelSparkline>();

        public void Save()
        {
            // TODO: Remove attributes if they don't exist on the model?
            this.SaveAttributes();
            this.SaveColors();
            this.SaveSparklines();
        }

        private void SaveSparklines()
        {
            if (this.Sparklines.Count == 0)
            {
                throw new InvalidOperationException("At least one sparkline must exist to save sparklines.");
            }
            var sparklinesNode = this.CreateNode("x14:sparklines");
            sparklinesNode.RemoveAll();
            foreach(var sparkline in this.Sparklines)
            {
                sparklinesNode.AppendChild(this.CreateSparklineNode(sparkline));
            }
        }

        private XmlNode CreateSparklineNode(ExcelSparkline sparkline)
        {
            var node = this.TopNode.OwnerDocument.CreateElement("x14:sparkline", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var f = this.TopNode.OwnerDocument.CreateElement("xm:f", "http://schemas.microsoft.com/office/excel/2006/main");
            f.InnerText = sparkline.Formula.Address;
            node.AppendChild(f);
            var sqref = this.TopNode.OwnerDocument.CreateElement("xm:sqref", "http://schemas.microsoft.com/office/excel/2006/main");
            sqref.InnerText = sparkline.HostCell.Address;
            node.AppendChild(sqref);
            return node;
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
                this.SetAttribute(node, "rgb", ExcelSparklineGroup.XmlColor(color));
            }
            else
            {
                this.TopNode.RemoveChild(node);
            }
        }

        private static string XmlColor(Color color)
        {
            return string.Format("{0:X2}{1:X2}{2:X2}{3:X2}", color.A, color.R, color.G, color.B);
        }

        private void SaveAttributes()
        {
            if (this.ManualMax != null)
                this.SetAttribute(this.TopNode, "manualMax", this.ManualMax.ToString());
            else
                this.ClearAttribute(this.TopNode, "manualMax");
            if (this.ManualMin != null)
                this.SetAttribute(this.TopNode, "manualMin", this.ManualMin.ToString());
            else
                this.ClearAttribute(this.TopNode, "manualMin");
            if (this.LineWeight != null && this.LineWeight != 0.75)
                this.SetAttribute(this.TopNode, "lineWeight", this.LineWeight.ToString());
            else
                this.ClearAttribute(this.TopNode, "lineWeight");
            if (this.Type != null && this.Type != SparklineType.Line)
                this.SetAttribute(this.TopNode, "type", ExcelSparklineGroup.SparklineTypeToString(this.Type.Value));
            else
                this.ClearAttribute(this.TopNode, "type");
            if (this.DateAxis)
                this.SetAttribute(this.TopNode, "dateAxis", this.DateAxis.ToString());
            else
                this.ClearAttribute(this.TopNode, "dateAxis");
            if (this.DisplayEmptyCellsAs != null && this.DisplayEmptyCellsAs != DispBlanksAs.Zero)
                this.SetAttribute(this.TopNode, "displayEmptyCellsAs", ExcelSparklineGroup.DisplayBlanksAsToString(this.DisplayEmptyCellsAs.Value));
            else
                this.ClearAttribute(this.TopNode, "displayEmptyCellsAs");
            if (this.Markers)
                this.SetAttribute(this.TopNode, "markers", this.Markers.ToString());
            else
                this.ClearAttribute(this.TopNode, "markers");
            if (this.High)
                this.SetAttribute(this.TopNode, "high", this.High.ToString());
            else
                this.ClearAttribute(this.TopNode, "high");
            if (this.Low)
                this.SetAttribute(this.TopNode, "low", this.Low.ToString());
            else
                this.ClearAttribute(this.TopNode, "low");
            if (this.First)
                this.SetAttribute(this.TopNode, "first", this.First.ToString());
            else
                this.ClearAttribute(this.TopNode, "first");
            if (this.Last)
                this.SetAttribute(this.TopNode, "last", this.Last.ToString());
            else
                this.ClearAttribute(this.TopNode, "last");
            if (this.Negative)
                this.SetAttribute(this.TopNode, "negative", this.Negative.ToString());
            else
                this.ClearAttribute(this.TopNode, "negative");
            if (this.DisplayXAxis)
                this.SetAttribute(this.TopNode, "displayXAxis", this.DisplayXAxis.ToString());
            else
                this.ClearAttribute(this.TopNode, "displayXAxis");
            if (this.DisplayHidden)
                this.SetAttribute(this.TopNode, "displayHidden", this.DisplayHidden.ToString());
            else
                this.ClearAttribute(this.TopNode, "displayHidden");
            if (this.MinAxisType != SparklineAxisMinMax.Individual)
                this.SetAttribute(this.TopNode, "minAxisType", ExcelSparklineGroup.SparklineAxisMinMaxToString(this.MinAxisType));
            else
                this.ClearAttribute(this.TopNode, "minAxisType");
            if (this.MaxAxisType != SparklineAxisMinMax.Individual)
                this.SetAttribute(this.TopNode, "maxAxisType", ExcelSparklineGroup.SparklineAxisMinMaxToString(this.MaxAxisType));
            else
                this.ClearAttribute(this.TopNode, "maxAxisType");
            if (this.RightToLeft)
                this.SetAttribute(this.TopNode, "rightToLeft", this.RightToLeft.ToString());
            else
                this.ClearAttribute(this.TopNode, "rightToLeft");
        }

        private void SetAttribute(XmlNode topNode, string attribute, string value)
        {
            var node = topNode.Attributes[attribute];
            if(node == null)
            {
                node = topNode.OwnerDocument.CreateAttribute(attribute);
                topNode.Attributes.Append(node);
            }
            node.Value = value;
        }

        private void ClearAttribute(XmlNode node, string attribute)
        {
            node.Attributes.RemoveNamedItem(attribute);
        }
        #endregion

        #region Attributes
        public double? ManualMax { get; set; }
        public double? ManualMin { get; set; }
        public double? LineWeight { get; set; } // defaults to 0.75, may not be necessary
        public SparklineType? Type { get; set; } = SparklineType.Line;
        public bool DateAxis { get; set; }
        public DispBlanksAs? DisplayEmptyCellsAs { get; set; } = DispBlanksAs.Zero;
        public bool Markers { get; set; }
        public bool High { get; set; }
        public bool Low { get; set; }
        public bool First { get; set; }
        public bool Last { get; set; }
        public bool Negative { get; set; }
        public bool DisplayXAxis { get; set; }
        //public bool DisplayYAxis { get; set; }
        public bool DisplayHidden { get; set; }
        public SparklineAxisMinMax MinAxisType { get; set; } = SparklineAxisMinMax.Individual;
        public SparklineAxisMinMax MaxAxisType { get; set; } = SparklineAxisMinMax.Individual;
        public bool RightToLeft { get; set; }

        #endregion
        #endregion
        #region XmlHelper Overrides
        public ExcelSparklineGroup(XmlNamespaceManager nameSpaceManager, XmlNode topNode): base(nameSpaceManager, topNode)
        {
            // Parse the interesting nodes.
            var node = TopNode.SelectSingleNode("x14:colorSeries", nameSpaceManager);
            if(node != null)
                this.ColorSeries = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorNegative", nameSpaceManager);
            if (node != null)
                this.ColorNegative = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorAxis", nameSpaceManager);
            if (node != null)
                this.ColorAxis = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorMarkers", nameSpaceManager);
            if (node != null)
                this.ColorMarkers = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorFirst", nameSpaceManager);
            if (node != null)
                this.ColorFirst = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorLast", nameSpaceManager);
            if (node != null)
                this.ColorLast = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorHigh", nameSpaceManager);
            if (node != null)
                this.ColorHigh = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            node = TopNode.SelectSingleNode("x14:colorLow", nameSpaceManager);
            if (node != null)
                this.ColorLow = ExcelConditionalFormattingHelper.ConvertFromColorCode(node.Attributes["rgb"].InnerText);
            // Parse the attributes.
            var attribute = topNode.Attributes["type"];
            if (attribute != null)
                this.Type = ParseSparklineType(attribute.InnerText);
            attribute = TopNode.Attributes["displayEmptyCellsAs"];
            if (attribute != null)
                this.DisplayEmptyCellsAs = ParseDisplayEmptyCellsAs(attribute.InnerText);
            this.Negative = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "negative");
            this.ManualMin = GetAttributeDouble(topNode, "manualMin");
            this.ManualMax = GetAttributeDouble(topNode, "manualMax");
            this.LineWeight = GetAttributeDouble(topNode, "lineWeight");
            this.DateAxis = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "dateAxis"); // DateAxis
            this.Markers = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "markers");
            this.High = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "high");
            this.Low = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "low");
            this.First = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "first");
            this.Last = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "last");
            this.DisplayXAxis = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "displayXAxis");
            //this.DisplayYAxis = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "displayYAxis");
            this.DisplayHidden = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "displayHidden");
            attribute = topNode.Attributes["minAxisType"];
            if (attribute != null)
                this.MinAxisType = ParseSparklineAxisMinMax(attribute.InnerText);
            attribute = topNode.Attributes["maxAxisType"];
            if (attribute != null)
                this.MaxAxisType = ParseSparklineAxisMinMax(attribute.InnerText);
            this.RightToLeft = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "rightToLeft");
            // Parse the actual sparklines.
            var sparklineNodes = TopNode.SelectSingleNode("x14:sparklines", nameSpaceManager)?.ChildNodes;
            foreach (var sparklineNode in sparklineNodes)
                this.Sparklines.Add(new ExcelSparkline(nameSpaceManager, (XmlNode) sparklineNode));
        }

        public ExcelSparklineGroup(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {

        }
        #endregion
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
    }
}
