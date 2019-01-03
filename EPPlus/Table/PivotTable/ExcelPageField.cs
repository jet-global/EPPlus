using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Represents an Excel Pivot Table Page Field XML element.
	/// </summary>
	public class ExcelPageField : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets or sets the index of the field that appears on the page or filter report area of the <see cref="PivotTable"/>.
		/// </summary>
		/// <remarks>Corresponds the "fld" attribute.</remarks>
		public int Field
		{
			get { return base.GetXmlNodeInt("@fld"); }
			set { base.SetXmlNodeString("@fld", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the index of the <see cref="ExcelPivotTableFieldItem"/> that this page field refers to.
		/// </summary>
		public int? Item
		{
			get { return base.GetXmlNodeIntNull("@item"); }
			set { base.SetXmlNodeString("@item", value?.ToString() ?? null, true); }
		}

		/// <summary>
		/// Gets or sets the index of the OLAP hierarchy to which this item belongs.
		/// </summary>
		/// <remarks>Corresponds the "hier" attribute.</remarks>
		public int Hierarchy
		{
			get { return base.GetXmlNodeInt("@hier"); }
			set { base.SetXmlNodeString("@hier", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the unique name of the hierarchy.
		/// </summary>
		public string Name
		{
			get { return base.GetXmlNodeString("@name"); }
			set { base.SetXmlNodeString("@name", value); }
		}

		/// <summary>
		/// Gets or sets the display name of the hierarchy.
		/// </summary>
		/// <remarks>Corresponds to the "cap" attribute.</remarks>
		public string Caption
		{
			get { return base.GetXmlNodeString("@cap"); }
			set { base.SetXmlNodeString("@cap", value); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPageField"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		public ExcelPageField(XmlNamespaceManager namespaceManager, XmlNode node) 
			: base(namespaceManager, node) { }
		#endregion
	}
}
