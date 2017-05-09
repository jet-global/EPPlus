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
		internal ExcelSlicers(ExcelWorksheet worksheet) : base(ExcelSlicer.SlicerDocumentNamespaceManager, null)
		{
			this.Worksheet = worksheet;
			var slicerFiles = this.Worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaSlicerRelationship);
			foreach (var slicerFile in slicerFiles)
			{
				var path = slicerFile.TargetUri.ToString().Replace("..", "/xl");
				var uri = new Uri(path, UriKind.Relative);
				var possiblePart = this.Worksheet.Package.GetXmlFromUri(uri);
				XmlNodeList slicerNodes = possiblePart.SelectNodes("default:slicers/default:slicer", this.NameSpaceManager);
				for (int i = 0; i < slicerNodes.Count; i++)
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
			this.Worksheet.Workbook.Package.SavePart(this.SlicersUri, this.Part);
		}
		#endregion
	}
}
