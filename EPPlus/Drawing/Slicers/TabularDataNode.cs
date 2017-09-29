using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <tabular /> node in <slicerCacheDefinition-data />.
	/// </summary>
	public class TabularDataNode
	{
		#region Class Variables
		private List<TabularItemNode> myItems = new List<TabularItemNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the pivot cache ID.
		/// </summary>
		public string PivotCacheId
		{
			get { return this.Node.Attributes["pivotCacheId"].Value; }
			set { this.Node.Attributes["pivotCacheId"].Value = value; }
		}

		/// <summary>
		/// Gets a readonly list of <see cref="TabularItemNode"/>.
		/// </summary>
		public IReadOnlyList<TabularItemNode> Items
		{
			get { return myItems; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="TabularDataNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="TabularDataNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public TabularDataNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode item in this.Node.SelectNodes("default:items/default:i", this.NameSpaceManager))
			{
				myItems.Add(new TabularItemNode(item));
			}
		}
		#endregion
	}
}
