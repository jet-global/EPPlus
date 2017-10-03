using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <selection /> node in <slicerCacheDefinition-data-olap-selections />.
	/// </summary>
	public class SlicerSelectionNode
	{
		#region Class Variables
		private List<SlicerCacheItemParent> myParents = new List<SlicerCacheItemParent>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the tabId, which identifies which worksheet the pivot table corresponding to this slicerCache exists on.
		/// </summary>
		public string Name
		{
			get { return this.Node.Attributes["n"].Value; }
			set { this.Node.Attributes["n"].Value = value; }
		}

		/// <summary>
		/// Gets a readonly list of the OLAP parents in this range.
		/// </summary>
		public IReadOnlyList<SlicerCacheItemParent> Parents
		{
			get { return myParents; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SlicerSelectionNode"/>.
		/// This is a wrapper for the <selection /> in <slicerCacheDefinition />
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerSelectionNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public SlicerSelectionNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode slicerCacheItemParent in this.Node.SelectNodes("default:p", this.NameSpaceManager))
			{
				myParents.Add(new SlicerCacheItemParent(slicerCacheItemParent));
			}
		}
		#endregion
	}
}
