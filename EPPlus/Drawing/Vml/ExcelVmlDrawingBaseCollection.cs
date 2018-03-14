/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
	/// <summary>
	/// A base class for VML Drawings.
	/// </summary>
	public class ExcelVmlDrawingBaseCollection
	{
		#region Class Variables
		private static XmlNamespaceManager myNamespaceManager;
		#endregion

		#region Properties
		internal XmlDocument VmlDrawingXml { get; set; }
		internal Uri Uri { get; set; }
		internal string RelId { get; set; }
		internal Packaging.ZipPackagePart Part { get; set; }

		/// <summary>
		///  The <see cref="XmlNamespaceManager"/> for VML Drawings.
		/// </summary>
		internal static XmlNamespaceManager NameSpaceManager
		{
			get
			{
				if (myNamespaceManager == null)
				{
					myNamespaceManager = new XmlNamespaceManager(new NameTable());
					myNamespaceManager.AddNamespace("v", ExcelPackage.schemaMicrosoftVml);
					myNamespaceManager.AddNamespace("o", ExcelPackage.schemaMicrosoftOffice);
					myNamespaceManager.AddNamespace("x", ExcelPackage.schemaMicrosoftExcel);
				}
				return myNamespaceManager;
			}
		}
		#endregion

		#region Constructors
		internal ExcelVmlDrawingBaseCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
		{
			this.VmlDrawingXml = new XmlDocument() { PreserveWhitespace = false };
			this.Uri = uri;
			if (uri == null)
				this.Part = null;
			else
			{
				this.Part = pck.Package.GetPart(uri);
				XmlHelper.LoadXmlSafe(this.VmlDrawingXml, this.Part.GetStream());
			}
		}
		#endregion
	}
}
