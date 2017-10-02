using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <selection /> node in <slicerCacheDefinition-data-olap-selections />.
	/// </summary>
	public class SlicerSelectionNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the tabId, which identifies which worksheet the pivot table corresponding to this slicerCache exists on.
		/// </summary>
		public string Name
		{
			get { return this.Node.Attributes["n"].Value; }
			set { this.Node.Attributes["n"].Value = value; }
		}

		private XmlNode Node { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SlicerSelectionNode"/>.
		/// This is a wrapper for the <selection /> in <slicerCacheDefinition />
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerSelectionNode"/>.</param>
		public SlicerSelectionNode(XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Node = node;
		}
		#endregion
	}
}
