using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection of row or column items.
	/// </summary>
	public class ItemsCollection : XmlHelper
	{
		#region Class Variables
		private List<RowColumnItem> myItems;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the row/column items.
		/// </summary>
		public IReadOnlyList<RowColumnItem> Items
		{
			get
			{
				if (myItems == null)
				{
					myItems = new List<RowColumnItem>();
					var xNodes = base.TopNode.SelectNodes("d:i", base.NameSpaceManager);
					foreach (XmlNode xmlNode in xNodes)
					{
						myItems.Add(new RowColumnItem(base.NameSpaceManager, base.TopNode));
					}
				}
				return myItems;
			}
		}

		/// <summary>
		/// Gets or sets the count.
		/// </summary>
		public int Count
		{
			get { return base.GetXmlNodeIntNull("@count") ?? 0; }
			private set { base.SetXmlNodeString("@count", value.ToString()); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ItemsCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top node.</param>
		public ItemsCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
		}
		#endregion
	}
}