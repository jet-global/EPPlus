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
	/// Collection class for data fields in a pivot table. 
	/// </summary>
	public class ExcelPivotTableDataFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableDataField>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableDataFieldCollection"/>.
		/// </summary>
		/// <param name="table">The pivot table.</param>
		internal ExcelPivotTableDataFieldCollection(ExcelPivotTable table) :
			 base(table)
		{

		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add a new data field.
		/// </summary>
		/// <param name="field">The field being added.</param>
		/// <returns>The new data field.</returns>
		public ExcelPivotTableDataField Add(ExcelPivotTableField field)
		{
			var dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
			if (dataFieldsNode == null)
			{
				myTable.CreateNode("d:dataFields");
				dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
			}

			XmlElement node = myTable.PivotTableXml.CreateElement("dataField", ExcelPackage.schemaMain);
			node.SetAttribute("fld", field.Index.ToString());
			dataFieldsNode.AppendChild(node);

			//XmlElement node = field.AppendField(dataFieldsNode, field.Index, "dataField", "fld");
			field.SetXmlNodeBool("@dataField", true, false);

			var dataField = new ExcelPivotTableDataField(field.NameSpaceManager, node, field);
			ValidateDupName(dataField);

			myList.Add(dataField);
			return dataField;
		}

		/// <summary>
		/// Remove a data field.
		/// </summary>
		/// <param name="dataField">The specified data field.</param>
		public void Remove(ExcelPivotTableDataField dataField)
		{
			if (dataField.Field.TopNode.SelectSingleNode(string.Format("../../d:dataFields/d:dataField[@fld={0}]", dataField.Index), dataField.NameSpaceManager) is XmlElement node)
				node.ParentNode.RemoveChild(node);
			myList.Remove(dataField);
		}

		/// <summary>
		/// Checks if a data field exists.
		/// </summary>
		/// <param name="name">The name of the data field.</param>
		/// <param name="datafield">The pivot table data field.</param>
		/// <returns>True if the data field exists.</returns>
		internal bool ExistsDfName(string name, ExcelPivotTableDataField datafield)
		{
			foreach (var df in myList)
			{
				if (((!string.IsNullOrEmpty(df.Name) && df.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase) ||
					  (string.IsNullOrEmpty(df.Name) && df.Field.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)))) && datafield != df)
					return true;
			}
			return false;
		}
		#endregion

		#region Private Methods
		private void ValidateDupName(ExcelPivotTableDataField dataField)
		{
			if (this.ExistsDfName(dataField.Field.Name, null))
			{
				var index = 2;
				string name;
				do
				{
					name = dataField.Field.Name + "_" + index++.ToString();
				}
				while (this.ExistsDfName(name, null));
				dataField.Name = name;
			}
		}
		#endregion
	}
}
