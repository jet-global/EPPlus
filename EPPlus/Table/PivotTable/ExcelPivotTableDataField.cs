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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A pivot table data field.
	/// </summary>
	public class ExcelPivotTableDataField : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets the pivot table field.
		/// </summary>
		public ExcelPivotTableField Field { get; private set; }
		
		/// <summary>
		/// Gets or sets the index of the data field.
		/// </summary>
		public int Index
		{
			get
			{
				return base.GetXmlNodeInt("@fld");
			}
			internal set
			{
				base.SetXmlNodeString("@fld", value.ToString());
			}
		}
		
		/// <summary>
		/// Gets or sets the name of the data field.
		/// </summary>
		public string Name
		{
			get
			{
				return base.GetXmlNodeString("@name");
			}
			set
			{
				if (this.Field.myTable.DataFields.ExistsDfName(value, this))
					throw (new InvalidOperationException("Duplicate datafield name"));
				base.SetXmlNodeString("@name", value);
			}
		}
		
		/// <summary>
		/// Gets or sets the field index referencing the field collection.
		/// </summary>
		public int BaseField
		{
			get
			{
				return base.GetXmlNodeInt("@baseField");
			}
			set
			{
				base.SetXmlNodeString("@baseField", value.ToString());
			}
		}
		
		/// <summary>
		/// Gets or sets the index to the base item when the ShowDataAs calculation is in use.
		/// </summary>
		public int BaseItem
		{
			get
			{
				return base.GetXmlNodeInt("@baseItem");
			}
			set
			{
				base.SetXmlNodeString("@baseItem", value.ToString());
			}
		}
		
		/// <summary>
		/// Gets or sets the number format id. 
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
		/// Gets or sets the number format for the data column.
		/// </summary>
		public string Format
		{
			get
			{
				foreach (var nf in this.Field.myTable.Worksheet.Workbook.Styles.NumberFormats)
				{
					if (nf.NumFmtId == this.NumFmtId)
						return nf.Format;
				}
				return this.Field.myTable.Worksheet.Workbook.Styles.NumberFormats[0].Format;
			}
			set
			{
				var styles = this.Field.myTable.Worksheet.Workbook.Styles;

				ExcelNumberFormatXml nf = null;
				if (!styles.NumberFormats.FindByID(value, ref nf))
				{
					nf = new ExcelNumberFormatXml(this.NameSpaceManager) { Format = value, NumFmtId = styles.NumberFormats.NextId++ };
					styles.NumberFormats.Add(value, nf);
				}
				this.NumFmtId = nf.NumFmtId;
			}
		}
		
		/// <summary>
		/// Gets or sets the type of aggregate function.
		/// </summary>
		public DataFieldFunctions Function
		{
			get
			{
				string s = base.GetXmlNodeString("@subtotal");
				if (s == "")
					return DataFieldFunctions.None;
				else
					return (DataFieldFunctions)Enum.Parse(typeof(DataFieldFunctions), s, true);
			}
			set
			{
				string v;
				switch (value)
				{
					case DataFieldFunctions.None:
						base.DeleteNode("@subtotal");
						return;
					case DataFieldFunctions.CountNums:
						v = "CountNums";
						break;
					case DataFieldFunctions.StdDev:
						v = "stdDev";
						break;
					case DataFieldFunctions.StdDevP:
						v = "stdDevP";
						break;
					default:
						v = value.ToString().ToLower(CultureInfo.InvariantCulture);
						break;
				}
				base.SetXmlNodeString("@subtotal", v);
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableDataField"/>.
		/// </summary>
		/// <param name="ns">The namespace of the sheet.</param>
		/// <param name="topNode">The top node of the xml.</param>
		/// <param name="field">The pivot table field.</param>
		internal ExcelPivotTableDataField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field) :
			 base(ns, topNode)
		{
			if (topNode.Attributes.Count == 0)
			{
				this.Index = field.Index;
				this.BaseField = 0;
				this.BaseItem = 0;
			}
			this.Field = field;
		}
		#endregion
	}
}