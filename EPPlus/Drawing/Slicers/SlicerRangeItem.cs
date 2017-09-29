using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <i /> node in <slicerCacheDefinition-data-olap-levels-level-ranges-range />.
	/// </summary>
	public class SlicerRangeItem
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

		/// <summary>
		/// Gets or sets the display name of this item.
		/// </summary>
		public string DisplayName
		{
			get { return this.Node.Attributes["c"].Value; }
			set { this.Node.Attributes["c"].Value = value; }
		}

		/// <summary>
		/// Gets or sets whether this item is a display item.
		/// </summary>
		public bool NonDisplay
		{
			get { return this.Node.Attributes["nd"]?.Value == "1"; }
			set
			{
				var attribute = this.Node.Attributes["nd"];
				if (attribute == null)
				{
					attribute = this.Node.OwnerDocument.CreateAttribute("nd");
					this.Node.Attributes.Append(attribute);
				}
				attribute.Value = (value ? "1" : "0");
			}
		}

		private XmlNode Node { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SlicerRangeItem"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerRangeItem"/>.</param>
		public SlicerRangeItem(XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Node = node;
		}
		#endregion
	}
}
