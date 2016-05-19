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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
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
        public double? ManuaMin { get; set; }
        public double? LineWeight { get; set; } = 0.75;
        public SparklineType? Type { get; set; } = SparklineType.Line;
        public bool? DateAxis { get; set; } = false;
        public DispBlanksAs? DisplayEmptyCellsAs { get; set; } = DispBlanksAs.Zero;
        public bool? Markers { get; set; } = false;
        public bool? High { get; set; } = false;
        public bool? Low { get; set; } = false;
        public bool? First { get; set; } = false;
        public bool? Last { get; set; } = false;
        public bool? Negative { get; set; } = false;
        public bool? DisplayXAxis { get; set; } = false;
        public bool? DisplayYAxis { get; set; } = false;
        public SparklineAxisMinMax? MinAxisType { get; set; } = SparklineAxisMinMax.Individual;
        public SparklineAxisMinMax? MaxAxisType { get; set; } = SparklineAxisMinMax.Individual;
        public bool? RightToLeft { get; set; } = false;

        #endregion
        #endregion
        #region XmlHelper Overrides
        public ExcelSparklineGroup(XmlNamespaceManager nameSpaceManager, XmlNode topNode): base(nameSpaceManager, topNode)
        {

        }

        public ExcelSparklineGroup(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {

        }
        #endregion
    }
}
