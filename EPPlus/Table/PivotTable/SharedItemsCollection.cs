using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	public class SharedItemsCollection : XmlHelper
	{
		#region Class Variables
		private List<CacheItem> myItems = new List<CacheItem>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the count.
		/// </summary>
		public int Count
		{
			get { return base.GetXmlNodeIntNull("@count") ?? 0; }
			private set { base.SetXmlNodeString("@count", value.ToString()); }
		}

		/// <summary>
		/// Gets a readonly list of the items in this <see cref="CacheFieldNode"/>.
		/// </summary>
		public IReadOnlyList<CacheItem> Items
		{
			get { return myItems; }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SharedItemsCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public SharedItemsCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			// Selects all possible child node types.
			foreach (XmlNode sharedItem in base.TopNode.SelectNodes("d:b | d:d | d:e | d:m | d:n | d:s | d:x", this.NameSpaceManager))
			{
				myItems.Add(new CacheItem(this.NameSpaceManager, sharedItem));
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new field item to the list.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <returns>The index of the new item.</returns>
		public int Add(object value)
		{
			string stringValue = ConvertUtil.ConvertObjectToXmlAttributeString(value);
			myItems.Add(new CacheItem(this.NameSpaceManager, base.TopNode, CacheItem.GetObjectType(value), stringValue));
			this.Count++;
			return myItems.Count - 1;
		}
		#endregion
	}
}