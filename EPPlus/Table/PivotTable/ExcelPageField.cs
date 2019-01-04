/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Evan Schallerer, and others as noted in the source history.
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
* For code change notes, see the source control history.
*******************************************************************************/
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Represents an Excel Pivot Table Page Field XML element.
	/// </summary>
	public class ExcelPageField : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets or sets the index of the field that appears on the page or filter report area of the <see cref="PivotTable"/>.
		/// </summary>
		/// <remarks>Corresponds the "fld" attribute.</remarks>
		public int Field
		{
			get { return base.GetXmlNodeInt("@fld"); }
			set { base.SetXmlNodeString("@fld", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the index of the <see cref="ExcelPivotTableFieldItem"/> that this page field refers to.
		/// </summary>
		public int? Item
		{
			get { return base.GetXmlNodeIntNull("@item"); }
			set { base.SetXmlNodeString("@item", value?.ToString() ?? null, true); }
		}

		/// <summary>
		/// Gets or sets the index of the OLAP hierarchy to which this item belongs.
		/// </summary>
		/// <remarks>Corresponds the "hier" attribute.</remarks>
		public int Hierarchy
		{
			get { return base.GetXmlNodeInt("@hier"); }
			set { base.SetXmlNodeString("@hier", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the unique name of the hierarchy.
		/// </summary>
		public string Name
		{
			get { return base.GetXmlNodeString("@name"); }
			set { base.SetXmlNodeString("@name", value); }
		}

		/// <summary>
		/// Gets or sets the display name of the hierarchy.
		/// </summary>
		/// <remarks>Corresponds to the "cap" attribute.</remarks>
		public string Caption
		{
			get { return base.GetXmlNodeString("@cap"); }
			set { base.SetXmlNodeString("@cap", value); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPageField"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		public ExcelPageField(XmlNamespaceManager namespaceManager, XmlNode node) 
			: base(namespaceManager, node) { }
		#endregion
	}
}
