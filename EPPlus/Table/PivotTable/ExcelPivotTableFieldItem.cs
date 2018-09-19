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
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A field Item. Used for grouping
	/// </summary>
	public class ExcelPivotTableFieldItem : XmlHelper
	{
		#region Class Variables
		private ExcelPivotTableField myField;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the text with unique values only.
		/// </summary>
		public string Text
		{
			get
			{
				return base.GetXmlNodeString("@n");
			}
			set
			{
				if (string.IsNullOrEmpty(value))
				{
					base.DeleteNode("@n");
					return;
				}
				foreach (var item in myField.Items)
				{
					if (item.Text == value)
						throw (new ArgumentException("Duplicate Text"));
				}
				base.SetXmlNodeString("@n", value);
			}
		}

		/// <summary>
		/// Gets the reference values.
		/// </summary>
		internal int X
		{
			get
			{
				return base.GetXmlNodeInt("@x");
			}
		}

		/// <summary>
		/// Gets the grand total value.
		/// </summary>
		internal string T
		{
			get
			{
				return base.GetXmlNodeString("@t");
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldItem"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node of the xml.</param>
		/// <param name="field">The pivot table field.</param>
		internal ExcelPivotTableFieldItem(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field) :
			 base(ns, topNode)
		{
			myField = field;
		}
		#endregion
	}
}