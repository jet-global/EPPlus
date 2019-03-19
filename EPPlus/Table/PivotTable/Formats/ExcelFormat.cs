/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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
using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable.Formats
{
	/// <summary>
	/// A format item in the <see cref="ExcelFormatsCollection"/>.
	/// </summary>
	public class ExcelFormat : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets the index into the formats collection in Styles.xml.
		/// </summary>
		public int FormatId
		{
			get { return base.GetXmlNodeInt("@dxfId"); }
		}

		/// <summary>
		/// Gets the collection of format references.
		/// </summary>
		public ExcelFormatReferencesCollection References { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new instance of <see cref="ExcelFormat"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public ExcelFormat(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			var referencesNode = node.SelectSingleNode(".//d:references", this.NameSpaceManager);
			if (referencesNode != null)
				this.References = new ExcelFormatReferencesCollection(this.NameSpaceManager, referencesNode);
		}
		#endregion
	}
}
