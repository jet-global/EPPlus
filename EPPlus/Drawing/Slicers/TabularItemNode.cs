using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <i /> node in <slicerCacheDefinition-data-tabular-items />.
	/// </summary>
	public class TabularItemNode : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets or sets the index of this item in the pivot cache fields.
		/// </summary>
		public int AtomIndex
		{
			get { return base.GetXmlNodeInt("@x", -1); }
			set { base.SetXmlNodeString("@x", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets whether or not this item is selected.
		/// </summary>
		public bool IsSelected
		{
			get { return base.GetXmlNodeBool("@s", false); }
			set { base.SetXmlNodeBool("@s", value); }
		}

		/// <summary>
		/// Gets or sets whether this item has data in the pivot table.
		/// </summary>
		public bool NoData
		{
			get { return base.GetXmlNodeBool("@nd", false); }
			set { base.SetXmlNodeBool("@nd", value); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="TabularItemNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="TabularItemNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager.</param>
		public TabularItemNode(XmlNode node, XmlNamespaceManager namespaceManager) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion
	}
}
