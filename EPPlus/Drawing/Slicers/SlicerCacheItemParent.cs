using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <p /> node in:
	///  <slicerCacheDefinition-data-olap-levels-level-ranges-range-i /> and
	///  <slicerCacheDefinition-data-olap-selections-selection />.
	/// </summary>
	public class SlicerCacheItemParent
	{
		#region Properties
		/// <summary>
		/// Gets or sets the name of this item.
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
		/// Creates an instance of a <see cref="SlicerCacheItemParent"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerCacheItemParent"/>.</param>
		public SlicerCacheItemParent(XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Node = node;
		}
		#endregion
	}
}
