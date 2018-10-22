/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Jan Källman, Evan Schallerer, and others as noted in the source history.
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
using OfficeOpenXml.Extensions;

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
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		/// <param name="table">The pivot table.</param>
		internal ExcelPivotTableDataFieldCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) 
			: base(namespaceManager, node, table)
		{
			if (node == null)
				base.TopNode = base.PivotTable.CreateNode("d:dataFields");
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add a new data field to this collection.
		/// </summary>
		/// <param name="field">The <see cref="ExcelPivotTableField"/> being added.</param>
		/// <returns>The added field.</returns>
		public ExcelPivotTableDataField Add(ExcelPivotTableField field)
		{
			var dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
			if (dataFieldsNode == null)
			{
				base.PivotTable.CreateNode("d:dataFields");
				dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
			}

			XmlElement node = base.PivotTable.PivotTableXml.CreateElement("dataField", ExcelPackage.schemaMain);
			node.SetAttribute("fld", field.Index.ToString());
			dataFieldsNode.AppendChild(node);

			//XmlElement node = field.AppendField(dataFieldsNode, field.Index, "dataField", "fld");
			field.SetXmlNodeBool("@dataField", true, false);

			var dataField = new ExcelPivotTableDataField(field.NameSpaceManager, node, field);
			this.ValidateDupName(dataField);

			base.AddItem(dataField);
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
			base.RemoveItem(dataField);
		}

		/// <summary>
		/// Checks if a data field exists.
		/// </summary>
		/// <param name="name">The name of the data field.</param>
		/// <param name="datafield">The pivot table data field.</param>
		/// <returns>True if the data field exists.</returns>
		internal bool ExistsDfName(string name, ExcelPivotTableDataField datafield)
		{
			foreach (var dataField in this)
			{
				if (((!string.IsNullOrEmpty(dataField.Name) && dataField.Name.IsEquivalentTo(name) ||
					  (string.IsNullOrEmpty(dataField.Name) && dataField.Field.Name.IsEquivalentTo(name)))) && datafield != dataField)
					return true;
			}
			return false;
		}

		#endregion

		#region ExcelPivotTableFieldCollectionBase Overrides
		protected override List<ExcelPivotTableDataField> LoadItems()
		{
			var collection = new List<ExcelPivotTableDataField>();
			foreach (XmlElement dataElem in base.TopNode.SelectNodes("d:dataField", base.NameSpaceManager))
			{
				if (int.TryParse(dataElem.GetAttribute("fld"), out var fld) && fld >= 0)
				{
					var field = base.PivotTable.Fields[fld];
					var dataField = new ExcelPivotTableDataField(base.NameSpaceManager, dataElem, field);
					collection.Add(dataField);
				}
			}
			return collection;
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
