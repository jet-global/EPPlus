using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Wraps an <olap /> node in <slicerCacheDefinition-data />.
	/// </summary>
	public class OlapDataNode
	{
		#region Class Variables
		private List<SlicerSelectionNode> mySelections = new List<SlicerSelectionNode>();
		private List<SlicerLevelNode> mySlicerLevels = new List<SlicerLevelNode>();
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
		/// Gets a readonly list of <see cref="SlicerLevelNode"/>.
		/// </summary>
		public IReadOnlyList<SlicerLevelNode> SlicerLevels
		{
			get { return mySlicerLevels; }
		}

		/// <summary>
		/// Gets a readonly list of <see cref="SlicerSelectionNode"/>.
		/// </summary>
		public IReadOnlyList<SlicerSelectionNode> Selections
		{
			get { return mySelections; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="OlapDataNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="OlapDataNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public OlapDataNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode slicerLevel in this.Node.SelectNodes("default:levels/default:level", this.NameSpaceManager))
			{
				mySlicerLevels.Add(new SlicerLevelNode(slicerLevel, this.NameSpaceManager));
			}
			foreach (XmlNode slicerSelectionNode in this.Node.SelectNodes("default:selections/default:selection", this.NameSpaceManager))
			{
				mySelections.Add(new SlicerSelectionNode(slicerSelectionNode, this.NameSpaceManager));
			}
		}
		#endregion
	}
}
