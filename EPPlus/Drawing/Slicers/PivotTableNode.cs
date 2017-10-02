using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <pivotTable /> in <slicerCacheDefinition-pivotTables />.
	/// </summary>
	public class PivotTableNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the tabId, which identifies which worksheet the pivot table corresponding to this slicerCache exists on.
		/// </summary>
		public string TabId
		{
			get { return this.Node.Attributes["tabId"].Value; }
			set { this.Node.Attributes["tabId"].Value = value; }
		}

		/// <summary>
		/// Gets or sets the name of the PivotTable this slicer cache's slicer is affecting.
		/// </summary>
		public string PivotTableName
		{
			get { return this.Node.Attributes["name"].Value; }
			set { this.Node.Attributes["name"].Value = value; }
		}

		private XmlNode Node { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="PivotTableNode"/>.
		/// This is a wrapper for the <pivotTable /> in <slicerCacheDefinition />
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="PivotTableNode"/>.</param>
		public PivotTableNode(XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Node = node;
		}
		#endregion
	}
}
