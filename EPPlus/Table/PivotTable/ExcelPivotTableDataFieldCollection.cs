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
