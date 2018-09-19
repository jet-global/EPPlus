using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Wraps a <cacheField/> node in <pivotcachedefinition-cacheFields/>.
	/// </summary>
	public class CacheFieldNode
	{
		#region Class Variables
		private List<CacheFieldItem> myItems = new List<CacheFieldItem>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the name for this <see cref="CacheFieldNode"/>.
		/// </summary>
		public string Name
		{
			get { return this.Node.Attributes["name"].Value; }
			set { this.Node.Attributes["name"].Value = value; }
		}

		/// <summary>
		/// Gets or sets the number format ID for this <see cref="CacheFieldNode"/>.
		/// </summary>
		public string NumFormatId
		{
			get { return this.Node.Attributes["numFmtId"].Value; }
			set { this.Node.Attributes["numFmtId"].Value = value; }
		}

		/// <summary>
		/// Gets a readonly list of the items in this <see cref="CacheFieldNode"/>.
		/// </summary>
		public IReadOnlyList<CacheFieldItem> Items
		{
			get { return myItems; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheFieldNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheFieldNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public CacheFieldNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode cacheFieldItem in this.Node.SelectNodes("d:sharedItems/d:s", this.NameSpaceManager))
			{
				myItems.Add(new CacheFieldItem(cacheFieldItem));
			}
		}
		#endregion
	}
}