/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * ExcelSparkline.cs Copyright (C) 2016 Matt Delaney.
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
using System.Xml;

namespace OfficeOpenXml.Drawing.Sparkline
{
	/// <summary>
	/// Represents the CT_Sparkline XML schema element as defined at https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx
	/// </summary> 
	public class ExcelSparkline : XmlHelper
	{
		#region Properties
		/// <summary>
		///  Optional, gets or sets a value that corresponds to the XSD Schema's "F" argument.
		/// </summary>
		public ExcelAddress Formula { get; set; }

		/// <summary>
		/// Required, gets or sets a value that corresponds to the XSD Schema's "SqRef" argument.
		/// </summary>
		public ExcelAddress HostCell { get; set; }

		/// <summary>
		/// Gets the <see cref="ExcelSparklineGroup"/> this <see cref="ExcelSparkline"/> belongs to.
		/// </summary>
		public ExcelSparklineGroup Group { get; private set; }
		#endregion

		#region XmlHelper Overrides
		/// <summary>
		/// Create a new <see cref="ExcelSparkline"/> from an existing XML Node.
		/// </summary>
		/// <param name="group">The <see cref="ExcelSparklineGroup"/> this line will belong to.</param>
		/// <param name="nameSpaceManager">The Namespace Manager for the object.</param>
		/// <param name="topNode">The x14:Sparkline node containing information about the sparkline.</param>
		public ExcelSparkline(ExcelSparklineGroup group, XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
		{
			if (group == null)
				throw new ArgumentNullException(nameof(group));
			if (topNode == null)
				throw new ArgumentNullException(nameof(topNode));
			this.Group = group;
			var formulaNode = topNode.SelectSingleNode("xm:f", nameSpaceManager);
			var hostNode = topNode.SelectSingleNode("xm:sqref", nameSpaceManager);
			Formula = new ExcelAddress(formulaNode.InnerText);
			HostCell = group.Worksheet.Cells[hostNode.InnerText];
			group.Worksheet.Cells[HostCell.Address].Sparklines.Add(this);
		}

		/// <summary>
		/// Create a new <see cref="ExcelSparkline"/> from scratch (Without using an existing XML Node).
		/// </summary>
		/// <param name="group">The <see cref="ExcelSparklineGroup"/> that this line will belong to.</param>
		/// <param name="nameSpaceManager">The namespace manager for the object.</param>
		public ExcelSparkline(ExcelSparklineGroup group, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
		{
			if (group == null)
				throw new ArgumentNullException(nameof(group));
			this.Group = group;
		}
		#endregion
	}
}
