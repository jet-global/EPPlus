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
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable.Formats
{
	/// <summary>
	/// A collection of <see cref="ExcelFormat"/>s.
	/// </summary>
	public class ExcelFormatsCollection : XmlCollectionBase<ExcelFormat>
	{
		#region Constructors
		/// <summary>
		/// Creates a new instance of an <see cref="ExcelFormatsCollection"/>.
		/// </summary>
		/// <param name="namespaceManager"></param>
		/// <param name="node"></param>
		public ExcelFormatsCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node) { }
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the formats from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="ExcelFormat"/>s.</returns>
		protected override List<ExcelFormat> LoadItems()
		{
			var items = new List<ExcelFormat>();
			foreach (XmlNode format in base.TopNode.SelectNodes("d:format", this.NameSpaceManager))
			{
				items.Add(new ExcelFormat(this.NameSpaceManager, format));
			}
			return items;
		}
		#endregion
	}
}
