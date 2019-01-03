using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Represents an Excel Pivot Table Page Field collection XML element.
	/// </summary>
	public class ExcelPageFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPageField>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPageFieldCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		public ExcelPageFieldCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) 
			: base(namespaceManager, node, table) { }
		#endregion

		#region ExcelPivotTableFieldCollectionBase Overrides
		/// <summary>
		/// Loads the initial collection of items from XML.
		/// </summary>
		/// <returns>A new list of page fields from the XML.</returns>
		protected override List<ExcelPageField> LoadItems()
		{
			var collection = new List<ExcelPageField>();
			var fields = base.TopNode.SelectNodes("d:pageField", base.NameSpaceManager);
			foreach (XmlNode xmlNode in fields)
			{
				collection.Add(new ExcelPageField(base.NameSpaceManager, xmlNode));
			}
			return collection;
		}
		#endregion
	}
}
