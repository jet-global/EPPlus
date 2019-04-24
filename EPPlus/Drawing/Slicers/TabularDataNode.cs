using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	#region Enums
	public enum SortOrder
	{
		Ascending = 0,
		Descending = 1
	};
	#endregion

	/// <summary>
	/// Wraps a <tabular /> node in <slicerCacheDefinition-data />.
	/// </summary>
	public class TabularDataNode : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets or sets the pivot cache ID.
		/// </summary>
		public string PivotCacheId
		{
			get { return base.GetXmlNodeString("@pivotCacheId"); }
			set { base.SetXmlNodeString("@pivotCacheId", value); }
		}

		/// <summary>
		/// Gets or sets the sort order attribute.
		/// </summary>
		public SortOrder SortOrder
		{
			get
			{
				var value = base.GetXmlNodeString("@sortOrder", "ascending");
				if (Enum.TryParse(value, true, out SortOrder sortOrder))
					return sortOrder;
				throw new InvalidOperationException($"Unexpected sortOrder value: '{value}'");
			}
			set
			{
				// The default value is "ascending" so we will leave that blank.
				string stringValue = null;
				if (value == SortOrder.Descending)
					stringValue = value.ToString();
				base.SetXmlNodeString("@sortOrder", stringValue, true);
			}
		}

		/// <summary>
		/// Gets or sets the custom list sort attribute.
		/// </summary>
		public bool CustomListSort
		{
			get { return base.GetXmlNodeBool("@customListSort", true); }
			set { base.SetXmlNodeBool("@customListSort", true); }
		}

		public string CrossFilter
		{
			get { return base.GetXmlNodeString("@crossFilter"); }
			set { base.SetXmlNodeString("@crossFilter", value); }
		}

		/// <summary>
		/// Gets the slicer cache data items.
		/// </summary>
		public SlicerCacheTabularItems Items { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="TabularDataNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="TabularDataNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public TabularDataNode(XmlNode node, XmlNamespaceManager namespaceManager) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			var itemsNode = base.TopNode.SelectSingleNode("default:items", base.NameSpaceManager);
			this.Items = new SlicerCacheTabularItems(itemsNode, base.NameSpaceManager);
		}
		#endregion
	}
}
