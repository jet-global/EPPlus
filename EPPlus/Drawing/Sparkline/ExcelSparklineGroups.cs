/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * ExcelSparklineGroups.cs Copyright (C) 2016 Matt Delaney.
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

		/// <summary>
		/// Gets the <see cref="ExcelSparklineGroup"/>s that exist in this <see cref="ExcelSparklineGroups"/> node.
		/// </summary>
		public List<ExcelSparklineGroup> SparklineGroups { get; } = new List<ExcelSparklineGroup>();
		#endregion

		#region Public Methods
		/// <summary>
		/// Save Sparkline Groups to an existing TopNode.
		/// </summary>
		public void Save()
		{
			if (this.SparklineGroups.Count == 0 || this.SparklineGroups[0].Sparklines.Count == 0)
				return;
			if (base.TopNode == null)
			{
				XmlNode extNode = this.Worksheet.TopNode.SelectSingleNode("d:extLst/d:ext", base.NameSpaceManager);
				XmlNode extLstNode = null;
				if (extNode == null)
				{
					extNode = this.Worksheet.TopNode.OwnerDocument.CreateNode(XmlNodeType.Element, "ext", base.NameSpaceManager.DefaultNamespace);
					extLstNode = this.Worksheet.TopNode.SelectSingleNode("d:extLst", base.NameSpaceManager);
					if (extLstNode == null)
					{
						extLstNode = this.Worksheet.TopNode.OwnerDocument.CreateNode(XmlNodeType.Element, "extLst", base.NameSpaceManager.DefaultNamespace);
						this.Worksheet.TopNode.AppendChild(extLstNode);
					}
					extLstNode.AppendChild(extNode);
				}
				base.TopNode = this.Worksheet.TopNode.OwnerDocument.CreateElement("x14:sparklineGroups", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
				extNode.AppendChild(base.TopNode);
			}
			base.TopNode.RemoveAll();
			foreach (var group in this.SparklineGroups)
			{
				group.Save();
				base.TopNode.AppendChild(group.TopNode);
			}
		}
		#endregion

		#region XmlHelper Overrides
		/// <summary>
		/// Creates a new <see cref="ExcelSparklineGroups"/> based on the specified <see cref="XmlNode"/>.
		/// </summary>
		/// <param name="worksheet">The <see cref="ExcelWorksheet"/> the <see cref="ExcelSparklineGroups"/> node is defined on.</param>
		/// <param name="nameSpaceManager">The namespace manager for the object.</param>
		/// <param name="topNode">the x14:sparklineGroups node that defines the <see cref="ExcelSparklineGroups"/>.</param>
		public ExcelSparklineGroups(ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
		{
			this.Worksheet = worksheet;
			foreach (var groupNode in topNode.ChildNodes)
			{
				this.SparklineGroups.Add(new ExcelSparklineGroup(worksheet, nameSpaceManager, (XmlNode)groupNode));
			}
		}

		/// <summary>
		/// Create a new <see cref="ExcelSparklineGroups"/> from scratch (without an existing XML Node).
		/// </summary>
		/// <param name="worksheet">The worksheet the sparkline groups exist on.</param>
		/// <param name="nameSpaceManager">The namespace manager for the object.</param>
		public ExcelSparklineGroups(ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
		{
			if (worksheet == null)
				throw new ArgumentNullException(nameof(worksheet));
			this.Worksheet = worksheet; 
		}
		#endregion
	}
}
