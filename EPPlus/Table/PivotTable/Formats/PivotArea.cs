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
	/// The possible values for type of rule used to describe an area of the pivot table.
	/// </summary>
	public enum PivotAreaType
	{
		/// <summary>
		/// Refers to the whole pivot table.
		/// </summary>
		All,
		/// <summary>
		/// Refers to a field button.
		/// </summary>
		Button,
		/// <summary>
		/// Refers to the data area.
		/// </summary>
		Data,
		/// <summary>
		/// Refers to no pivot area.
		/// </summary>
		None,
		/// <summary>
		/// Refers to a header or item.
		/// </summary>
		Normal,
		/// <summary>
		/// Refers to the blank cells at the top left of the pivot table.
		/// </summary>
		Origin,
		/// <summary>
		/// Refers to the blank cells at the top right of the pivot table.
		/// </summary>
		TopRight
	}

	/// <summary>
	/// A rule describing the pivot table formatted area.
	/// </summary>
	public class PivotArea : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets a flag indicating if the collapsed levels are considered subtotals.
		/// </summary>
		public bool CollapsedLevelsAreSubtotals
		{
			get { return base.GetXmlNodeBool("@collapsedLevelsAreSubtotals", false); }
		}

		/// <summary>
		/// Gets a flag indicating if the area is in outline form.
		/// </summary>
		public bool Outline
		{
			get { return base.GetXmlNodeBool("@outline", true); }
		}

		/// <summary>
		/// Gets the type of selection rule.
		/// </summary>
		public PivotAreaType RuleType
		{
			get
			{
				string type = base.GetXmlNodeString("@type");
				return (PivotAreaType)Enum.Parse(typeof(PivotAreaType), type, true);
			}
		}

		/// <summary>
		/// Gets the references collection for this pivot area.
		/// </summary>
		public ExcelFormatReferencesCollection ReferencesCollection { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new instance of a <see cref="PivotArea"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public PivotArea(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			var references = node.SelectSingleNode("d:references", this.NameSpaceManager);
			if (references != null)
				this.ReferencesCollection = new ExcelFormatReferencesCollection(this.NameSpaceManager, references);
		}
		#endregion
	}
}
