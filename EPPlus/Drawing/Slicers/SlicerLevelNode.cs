using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps a <level /> node in <slicerCacheDefinition-data-olap-levels />.
	/// </summary>
	public class SlicerLevelNode
	{
		#region Class Variables
		private List<SlicerRangeNode> mySlicerRanges = new List<SlicerRangeNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the unique name for the level.
		/// </summary>
		public string UniqueName
		{
			get { return this.Node.Attributes["uniqueName"].Value; }
			set { this.Node.Attributes["uniqueName"].Value = value; }
		}

		/// <summary>
		/// Gets or sets the source caption for the level.
		/// </summary>
		public string SourceCaption
		{
			get { return this.Node.Attributes["sourceCaption"].Value; }
			set { this.Node.Attributes["sourceCaption"].Value = value; }
		}

		/// <summary>
		/// Gets or sets the count for the level.
		/// </summary>
		public string Count
		{
			get { return this.Node.Attributes["count"].Value; }
			set { this.Node.Attributes["count"].Value = value; }
		}

		/// <summary>
		/// Gets a readonly list of the slicer ranges for the level.
		/// </summary>
		public IReadOnlyList<SlicerRangeNode> SlicerRanges
		{
			get { return mySlicerRanges; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SlicerLevelNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="SlicerLevelNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public SlicerLevelNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode slicerRangeNode in this.Node.SelectNodes("default:ranges/default:range", this.NameSpaceManager))
			{
				mySlicerRanges.Add(new SlicerRangeNode(slicerRangeNode, this.NameSpaceManager));
			}
		}
		#endregion
	}
}
