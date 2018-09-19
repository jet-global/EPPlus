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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A page/report filter field.
	/// </summary>
	public class ExcelPivotTablePageFieldSettings : XmlHelper
	{
		#region Class Variables
		private ExcelPivotTableField myField;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the name of the field.
		/// </summary>
		public string Name
		{
			get
			{
				return base.GetXmlNodeString("@name");
			}
			set
			{
				base.SetXmlNodeString("@name", value);
			}
		}

		/// <summary>
		/// Gets or sets the index of the field.
		/// </summary>
		internal int Index
		{
			get
			{
				return base.GetXmlNodeInt("@fld");
			}
			set
			{
				base.SetXmlNodeString("@fld", value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the number format id of the field.
		/// </summary>
		internal int NumFmtId
		{
			get
			{
				return base.GetXmlNodeInt("@numFmtId");
			}
			set
			{
				base.SetXmlNodeString("@numFmtId", value.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the hier of the field.
		/// </summary>
		internal int Hier
		{
			get
			{
				return base.GetXmlNodeInt("@hier");
			}
			set
			{
				base.SetXmlNodeString("@hier", value.ToString());
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTablePageFieldSettings"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node in the xml.</param>
		/// <param name="field">The pivot table field.</param>
		/// <param name="index">The index of the field.</param>
		internal ExcelPivotTablePageFieldSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field, int index) :
			 base(ns, topNode)
		{
			if (base.GetXmlNodeString("@hier") == "")
				this.Hier = -1;
			myField = field;
		}
		#endregion
	}
}