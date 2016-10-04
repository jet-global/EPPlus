using OfficeOpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
	/// <summary>
	/// Represents an Excel Slicer.
	/// </summary>
	public class ExcelSlicerDrawing: ExcelDrawing
	{
		#region Properties
		private ExcelWorksheet Worksheet { get; set; }

		public ExcelSlicerCache SlicerCache { get; private set; }

		public XmlDocument SlicerDocument { get; set; }

		private string _name = null;
		public override string Name
		{
			get
			{
				if(_name == null)
					_name = this.TopNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/sle:slicer", this.NameSpaceManager).Attributes["name"].Value;
				return _name;
			}
			set
			{
				this.TopNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/sle:slicer", this.NameSpaceManager).Attributes["name"].Value = value;
				// TODO: Update Slicer Cache's name, sourceName here too.
			}
		}
		#endregion

		private static XmlNamespaceManager GetDefaultSlicerDocumentNamespaceManager()
		{
			var nameTable = new NameTable();
			var namespaceManager = new XmlNamespaceManager(nameTable);
			namespaceManager.AddNamespace(string.Empty, ExcelPackage.schemaMain2009);
			// Hack to work around a bug where SelectSingleNode ignores the default namespace.
			namespaceManager.AddNamespace("default", ExcelPackage.schemaMain2009);
			namespaceManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
			return namespaceManager;
		}

		#region Constructors
		internal ExcelSlicerDrawing(ExcelDrawings drawings, XmlNode node) : base(drawings, node, "xdr:wsDr/mc:AlternateContent/mc:Choice/xdr:GraphicFrame/a:Graphic/a:graphicData/sle:slicer/@name")
		{
			this.Worksheet = drawings.Worksheet;
			// Locate the relevant slicer.xml and slicerCache.xml files by opening the worksheet's slicer links and looking for the slicer whose name matches the name listed in the drawing.
			var slicers = this.Worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaSlicerRelationship);
			XmlNode slicerNode = null;
			var namespaceManager = ExcelSlicerDrawing.GetDefaultSlicerDocumentNamespaceManager();
			foreach (var slicer in slicers)
			{
				var path = slicer.TargetUri.ToString().Replace("..", "/xl");
				var uri = new Uri(path, UriKind.Relative);
				var possiblePart = this.Worksheet._package.GetXmlFromUri(uri);
				slicerNode = possiblePart?.SelectSingleNode("default:slicers/default:slicer", namespaceManager);
				if (slicerNode?.Attributes["name"].Value == this.Name)
				{
					this.SlicerDocument = possiblePart;
					break;
				}
			}
			if (this.SlicerDocument == null)
				throw new NotImplementedException();
			var cacheName = slicerNode.Attributes["cache"].Value;
			var slicerCaches = this.Worksheet.Workbook.Part.GetRelationshipsByType(ExcelPackage.schemaSlicerCache);
			foreach (var cache in slicerCaches)
			{
				var possiblePart = this.Worksheet._package.GetXmlFromUri(new Uri($"/xl/{cache.TargetUri.ToString()}", UriKind.Relative));
				var slicerCacheNode = possiblePart.SelectSingleNode("default:slicerCacheDefinition", namespaceManager);
				if (slicerCacheNode.Attributes["name"].Value == cacheName)
				{
					this.SlicerCache = new ExcelSlicerCache(slicerCacheNode, namespaceManager);
					break;
				}
			}
		}
		#endregion

		#region Public Methods

		#endregion

	}
}
