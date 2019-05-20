using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	#region Enums
	/// <summary>
	/// Enum representing the sort order of tabular cached slicer data.
	/// </summary>
	public enum SortOrder
	{
		/// <summary>
		/// Values are sorted ascending (A-Z).
		/// </summary>
		Ascending = 0,
		/// <summary>
		/// Values are sorted descending (Z-A).
		/// </summary>
		Descending = 1
	};

	/// <summary>
	/// Enum representing the combination of the tabular slicer data settings
	/// "Visually indicate items with no data" and "Show items with no data last".
	/// </summary>
	public enum CrossFilter
	{
		/// <summary>
		/// Indicates that both settings are selected.
		/// </summary>
		Both = 0,
		/// <summary>
		/// Indicates that "Visually indicate items with no data" is checked and 
		/// "Show items with no data last" is unchecked.
		/// </summary>
		ShowItemsWithNoData = 1,
		/// <summary>
		/// Indicates that "Visually indicates items with no data" is unchecked regardless
		/// of the value of "Show items with no data last".
		/// </summary>
		None = 2
	}
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
					stringValue = value.ToString().ToLower();
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

		/// <summary>
		/// Gets or sets a value indicating the combination of the 
		/// "Visually indicate items with no data" and "Show items with no data last"
		/// setting selections.
		/// </summary>
		public CrossFilter CrossFilter
		{
			get
			{
				var value = base.GetXmlNodeString("@crossFilter");
				if (string.IsNullOrEmpty(value))
					return CrossFilter.Both;
				if (Enum.TryParse(value, true, out CrossFilter result))
					return result;
				throw new InvalidOperationException($"Unexpected sortOrder value: '{value}'");
			}
			set
			{
				if (value == CrossFilter.Both)
					base.SetXmlNodeString("@crossFilter", null, true);
				else
				{
					string stringValue = value.ToString();
					// Lowercase the first letter of the enum string value.
					stringValue = char.ToLower(stringValue[0]) + stringValue.Substring(1);
					base.SetXmlNodeString("@crossFilter", stringValue);
				}
			}
		}

		/// <summary>
		/// Gets or sets a value indicating whether to show items 
		/// deleted from the data source.
		/// </summary>
		public bool ShowMissing
		{
			get { return base.GetXmlNodeBool("@showMissing", true); }
			set { base.SetXmlNodeBool("@showMissing", value, true); }
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
