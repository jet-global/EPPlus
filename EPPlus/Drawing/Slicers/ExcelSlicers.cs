using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Represents a Slicers.xml file, which contains a collection of Excel Slicers.
	/// </summary>
	public class ExcelSlicers : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets the collection of <see cref="ExcelSlicer"/>s that are contained within this <see cref="ExcelSlicers"/> file.
		/// </summary>
		public List<ExcelSlicer> Slicers { get; } = new List<ExcelSlicer>();

		private ExcelWorksheet Worksheet { get; set; }

		private XmlDocument Part { get; set; }

		private Uri SlicersUri { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiate a new <see cref="ExcelSlicers"/> object representing the slicers on a particular <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="worksheet">The <see cref="ExcelWorksheet"/> whose slicers are being represented.</param>
		internal ExcelSlicers(ExcelWorksheet worksheet): base(ExcelSlicer.SlicerDocumentNamespaceManager, null)
		{
			this.Worksheet = worksheet;
			var slicerFiles = this.Worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaSlicerRelationship);
			foreach (var slicerFile in slicerFiles)
			{
				var path = slicerFile.TargetUri.ToString().Replace("..", "/xl");
				var uri = new Uri(path, UriKind.Relative);
				var possiblePart = this.Worksheet._package.GetXmlFromUri(uri);
				XmlNodeList slicerNodes = possiblePart.SelectNodes("default:slicers/default:slicer", this.NameSpaceManager);
				for(int i = 0; i < slicerNodes.Count; i++)
				{
					var slicerNode = slicerNodes[i];
					this.Slicers.Add(new ExcelSlicer(slicerNode, this.NameSpaceManager, this.Worksheet));
				}
				if (this.TopNode == null)
					this.TopNode = possiblePart.DocumentElement;
				if (this.Part == null)
				{
					this.Part = possiblePart;
					this.SlicersUri = uri;
				}
			}
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Serialize changes made to the worksheet's slicers to the slicerN.xml file.
		/// </summary>
		internal void Save()
		{
			this.Worksheet.Workbook._package.SavePart(this.SlicersUri, this.Part);
		}
		#endregion
	}
}
