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
			if (topNode == null)
				throw new ArgumentNullException(nameof(topNode));
			if (group == null)
				throw new ArgumentNullException(nameof(group));
			this.Group = group;
			var formulaNode = topNode.SelectSingleNode("xm:f", nameSpaceManager);
			var hostNode = topNode.SelectSingleNode("xm:sqref", nameSpaceManager);
			this.Formula = formulaNode != null ? new ExcelAddress(formulaNode.InnerText) : null;
			this.HostCell = group.Worksheet.Cells[hostNode.InnerText];
		}

		/// <summary>
		/// Create a new <see cref="ExcelSparkline"/> from scratch (without using an existing XML Node).
		/// </summary>
		/// <param name="hostCell">The <see cref="ExcelAddress"/> that hosts the sparkline.</param>
		/// <param name="formula">The <see cref="ExcelAddress"/> that the sparkline references. Can be null.</param>
		/// <param name="group">The <see cref="ExcelSparklineGroup"/> that this line will belong to.</param>
		/// <param name="nameSpaceManager">The namespace manager for the object.</param>
		public ExcelSparkline(ExcelAddress hostCell, ExcelAddress formula, ExcelSparklineGroup group, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
		{
			if (hostCell == null)
				throw new ArgumentNullException(nameof(hostCell));
			if (group == null)
				throw new ArgumentNullException(nameof(group));
			this.HostCell = hostCell;
			this.Group = group;
			this.Formula = formula;
			this.TopNode = group.TopNode.OwnerDocument.CreateElement("x14:sparkline", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Save the Sparkline's properties and attributes to the TopNode's XML. 
		/// </summary>
		public void Save()
		{
			this.TopNode.RemoveAll();
			if (this.Formula != null)
			{
				var formulaNode = this.TopNode.OwnerDocument.CreateElement("xm:f", "http://schemas.microsoft.com/office/excel/2006/main");
				formulaNode.InnerText = this.Formula.FullAddress;
				this.TopNode.AppendChild(formulaNode);
			}
			var hostNode = this.TopNode.OwnerDocument.CreateElement("xm:sqref", "http://schemas.microsoft.com/office/excel/2006/main");
			hostNode.InnerText = this.HostCell.Address;
			this.TopNode.AppendChild(hostNode);
		}
		#endregion
	}
}
