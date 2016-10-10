using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Represents an Excel Slicer Drawing.
	/// Loosely corresponds to a drawing.xml part.
	/// </summary>
	public class ExcelSlicerDrawing: ExcelDrawing
	{
		#region Class Variables
		private string _name = null;
		#endregion

		#region Properties
		private ExcelWorksheet Worksheet { get; set; }

		/// <summary>
		/// Gets or sets the slicer that this drawing represents on the worksheet.
		/// </summary>
		public ExcelSlicer Slicer { get; set; }

		/// <summary>
		/// Gets or sets the name of this <see cref="ExcelSlicerDrawing"/> object.
		/// This also updates the drawing's XML to reflect the name change.
		/// </summary>
		public override string Name
		{
			get
			{
				if(this._name == null)
					this._name = this.TopNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/sle:slicer", this.NameSpaceManager).Attributes["name"].Value;
				return this._name;
			}
			set
			{
				this._name = value;
				this.TopNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/sle:slicer", this.NameSpaceManager).Attributes["name"].Value = value;
				this.TopNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr", this.NameSpaceManager).Attributes["name"].Value = value;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new instance of the <see cref="ExcelSlicerDrawing"/> class.
		/// </summary>
		/// <param name="drawings">The drawings collection for this drawing's worksheet.</param>
		/// <param name="node">The xdr:twoCellAnchor node from this drawing's xml.</param>
		internal ExcelSlicerDrawing(ExcelDrawings drawings, XmlNode node) : base(drawings, node, "xdr:wsDr/mc:AlternateContent/mc:Choice/xdr:GraphicFrame/a:Graphic/a:graphicData/sle:slicer/@name")
		{
			this.Worksheet = drawings.Worksheet;
			this.Slicer = this.Worksheet.Slicers.Slicers.First(slicer => slicer.Name == this.Name);
		}
		#endregion
	}
}
