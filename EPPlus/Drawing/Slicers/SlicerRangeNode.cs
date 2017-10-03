using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <range /> node in <slicerCacheDefinition-data-olap-levels-level-ranges />.
	/// </summary>
	public class SlicerRangeNode
	{
		#region Class Variables
		private List<SlicerRangeItem> myItems = new List<SlicerRangeItem>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the start item for the range.
		/// </summary>
		public string StartItem
		{
			get { return this.Node.Attributes["startItem"].Value; }
			set { this.Node.Attributes["startItem"].Value = value; }
		}

		/// <summary>
		/// Gets a readonly list of the items in this range.
		/// </summary>
		public IReadOnlyList<SlicerRangeItem> Items
		{
			get { return myItems; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SlicerRangeNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerRangeNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public SlicerRangeNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode slicerRangeItem in this.Node.SelectNodes("default:i", this.NameSpaceManager))
			{
				myItems.Add(new SlicerRangeItem(slicerRangeItem, this.NameSpaceManager));
			}
		}
		#endregion
	}
}
