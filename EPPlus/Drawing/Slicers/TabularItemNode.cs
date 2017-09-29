using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <i /> node in <slicerCacheDefinition-data-tabular-items />.
	/// </summary>
	public class TabularItemNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the index of this item in the pivot cache fields.
		/// </summary>
		public int AtomIndex
		{
			get
			{
				if (int.TryParse(this.Node.Attributes["x"]?.Value ?? string.Empty, out int atom))
					return atom;
				return -1;
			}
			set
			{
				var attribute = this.Node.Attributes["x"];
				if (attribute == null)
				{
					attribute = this.Node.OwnerDocument.CreateAttribute("x");
					this.Node.Attributes.Append(attribute);
				}
				attribute.Value = value.ToString();
			}
		}

		/// <summary>
		/// Gets or sets whether or not this item is selected.
		/// </summary>
		public bool IsSelected
		{
			get { return this.Node.Attributes["s"]?.Value == "1"; }
			set
			{
				var attribute = this.Node.Attributes["s"];
				if (attribute == null)
				{
					attribute = this.Node.OwnerDocument.CreateAttribute("s");
					this.Node.Attributes.Append(attribute);
				}
				attribute.Value = (value ? "1" : "0");
			}
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
		/// Creates an instance of a <see cref="TabularItemNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="TabularItemNode"/>.</param>
		public TabularItemNode(XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Node = node;
		}
		#endregion
	}
}
