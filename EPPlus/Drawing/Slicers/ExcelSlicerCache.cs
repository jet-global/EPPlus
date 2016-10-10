using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Represents an Excel Slicer Cache.
	/// When this part exists, it can be found at /xl/slicerCaches/slicerCacheN.xml.
	/// </summary>
	public class ExcelSlicerCache: XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets or sets the name of this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public string Name
		{
			get
			{
				return this.TopNode.Attributes["name"].Value;
			}
			set
			{
				this.TopNode.Attributes["name"].Value = value;
				if (this.Slicer != null)
					this.Slicer.TopNode.Attributes["cache"].Value = value;
			}
		}

		/// <summary>
		/// Gets or sets the <see cref="ExcelSlicer"/> that uses this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public ExcelSlicer Slicer { get; set; }

		/// <summary>
		/// Gets the Uri of the Slicer Cache's associated XML part.
		/// </summary>
		public Uri SlicerCacheUri { get; private set; }

		/// <summary>
		/// Gets or sets the tabId, which identifies which worksheet the slicer corresponding to this slicerCache exists on.
		/// </summary>
		public string TabId
		{
			get
			{
				return this.TopNode.SelectSingleNode("default:pivotTables/default:pivotTable").Attributes["tabId"].Value;
			}
			set
			{
				this.TopNode.SelectSingleNode("default:pivotTables/default:pivotTable", this.NameSpaceManager).Attributes["tabId"].Value = value;
			}
		}

		private XmlDocument Part { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a new <see cref="ExcelSlicerCache"/> object to represent the slicerCacheN.xml part. 
		/// </summary>
		/// <param name="node">The slicerCacheDefinition node to represent.</param>
		/// <param name="namespaceManager">The namespaceManager to use when parsing nodes (This should usually be based on <see cref="ExcelSlicer.SlicerDocumentNamespaceManager"/>).</param>
		/// <param name="slicerCacheUri">The path to this Slicer Cache's part in the package.</param>
		/// <param name="part">The <see cref="XmlDocument"/> based on the <paramref name="slicerCacheUri"/>.</param>
		internal ExcelSlicerCache(XmlNode node, XmlNamespaceManager namespaceManager, Uri slicerCacheUri, XmlDocument part): base(namespaceManager, node)
		{
			this.SlicerCacheUri = slicerCacheUri;
			this.Part = part;
			this.Name = node.Attributes["name"].Value;
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Save this <see cref="ExcelSlicerCache"/> back into the <paramref name="package"/> at 
		/// </summary>
		/// <param name="package"></param>
		internal void Save(ExcelPackage package)
		{
			package.SavePart(new Uri("/xl/" + this.SlicerCacheUri, UriKind.Relative), this.Part);
		}
		#endregion
	}
}
