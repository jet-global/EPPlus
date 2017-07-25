/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * EPPlus Copyright (C) 2011 Jan Källman.
 * This File Copyright (C) 2016 Matt Delaney.
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author                      Change                            Date
 * ******************************************************************************
 * Matt Delaney                Added support for slicers.        11 October 2016     
 *******************************************************************************/
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Represents an Excel Slicer Drawing.
	/// Loosely corresponds to a drawing.xml part.
	/// </summary>
	public class ExcelSlicerDrawing : ExcelDrawing
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
				if (this._name == null)
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
			this.Slicer = this.Worksheet.Slicers.Slicers.First(slicer => this.CompareSlicerName(slicer.Name, this.Name));
		}
		#endregion

		#region Private Methods
		private bool CompareSlicerName(string name1, string name2)
		{
			// Excel will sometimes encode line feed characters as unicode rather than XML encoding,
			// so known problem unicode characters (such as line feed) are scrubbed.
			return this.StandardizeNewlineFormats(name1) == this.StandardizeNewlineFormats(name2);
		}

		private string StandardizeNewlineFormats(string str)
		{
			return str.Replace("_x000a_", "\n");
		}
		#endregion
	}
}
