﻿/*******************************************************************************
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
using System.Xml;

namespace OfficeOpenXml.Drawing.Sparkline
{
    /// <summary>
    /// Designed to be compliant with the Excel 2009 SparklineGroups schema ( https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx ).
    /// </summary>
    public class ExcelSparklineGroups : XmlHelper
    {
        #region Properties
        private ExcelWorksheet Worksheet;

        public List<ExcelSparklineGroup> SparklineGroups { get; } = new List<ExcelSparklineGroup>();
        #endregion

        #region Public Methods
        public void Save()
        {
            if (this.SparklineGroups.Count == 0 || this.SparklineGroups[0].Sparklines.Count == 0)
            {
                return;
            }
            else if (this.TopNode == null)
            {
                throw new NotImplementedException("Saving new SparkineGroups is currently not supported.");
            }
            else
            {
                foreach (var group in this.SparklineGroups)
                    group.Save();
            }
        }
        #endregion

        #region XmlHelper Overrides
        public ExcelSparklineGroups(ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager, XmlNode topNode): base(nameSpaceManager, topNode)
        {
            this.Worksheet = worksheet;
            foreach(var groupNode in topNode.ChildNodes)
            {
                SparklineGroups.Add(new ExcelSparklineGroup(worksheet, nameSpaceManager, (XmlNode) groupNode));
            }
        }

        public ExcelSparklineGroups(ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            this.Worksheet = worksheet;
        }
        #endregion
    }
}
