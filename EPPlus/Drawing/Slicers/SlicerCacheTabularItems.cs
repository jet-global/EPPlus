using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// XML wrapper class for data/tabular/items in slicerCaches.
	/// </summary>
	public class SlicerCacheTabularItems : XmlCollectionBase<TabularItemNode>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SlicerCacheTabularItems"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerCacheTabularItems"/>.</param>
		/// <param name="namespaceManager">The namespace manager.</param>
		public SlicerCacheTabularItems(XmlNode node, XmlNamespaceManager namespaceManager) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new item to the collection with the specified values.
		/// </summary>
		/// <param name="i">The index of the shared item that the added item refers to.</param>
		/// <param name="isSelected">A value indicating whether the item is selected.</param>
		public void Add(int i, bool isSelected)
		{
			var node = base.TopNode.OwnerDocument.CreateElement("i", base.TopNode.NamespaceURI);
			var item = new TabularItemNode(node, base.NameSpaceManager);
			item.AtomIndex = i;
			item.IsSelected = isSelected;
			base.AddItem(item);
		}

		/// <summary>
		/// Adds a list of <see cref="TabularItemNode"/>s to the collection.
		/// </summary>
		/// <param name="items">The items to add.</param>
		public void AddRange(List<TabularItemNode> items)
		{
			foreach (var item in items)
			{
				base.AddItem(item);
			}
		}

		/// <summary>
		/// Clears the collection of items.
		/// </summary>
		public void Clear() => base.ClearItems();
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the tabular items from the xml document.
		/// </summary>
		/// <returns>The collection of tabular items.</returns>
		protected override List<TabularItemNode> LoadItems()
		{
			var items = new List<TabularItemNode>();
			foreach (XmlNode item in base.TopNode.SelectNodes("default:i", base.NameSpaceManager))
			{
				items.Add(new TabularItemNode(item, base.NameSpaceManager));
			}
			return items;
		}
		#endregion
	}
}
