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
        public List<ExcelSparkline> Sparklines { get; private set; }
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
        public bool DisplayYAxis { get; set; }
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
            this.DisplayYAxis = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "displayYAxis");
            this.DisplayHidden = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "displayHidden");
            attribute = topNode.Attributes["minAxisType"];
            if (attribute != null)
                this.MinAxisType = ParseSparklineAxisMinMax(attribute.InnerText);
            attribute = topNode.Attributes["maxAxisType"];
            if (attribute != null)
                this.MaxAxisType = ParseSparklineAxisMinMax(attribute.InnerText);
            this.RightToLeft = ExcelConditionalFormattingHelper.GetAttributeBool(TopNode, "rightToLeft");
            // Parse the actual sparklines.

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
        private static SparklineAxisMinMax ParseSparklineAxisMinMax(string type)
        {
            if (type.Equals("custom"))
                return SparklineAxisMinMax.Custom;
            if (type.Equals("group"))
                return SparklineAxisMinMax.Group;
            else
                return SparklineAxisMinMax.Individual;
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
    }
}
