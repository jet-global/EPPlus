/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau, Evan Schallerer, and others as noted in the source history.
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
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// An Excel PivotCacheRecords.
	/// </summary>
	public class ExcelPivotCacheRecords : XmlHelper
	{
		#region Constants
		private const string Name = "pivotCacheRecords";
		#endregion

		#region Class Variables
		private List<CacheRecordNode> myRecords = new List<CacheRecordNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets the uri.
		/// </summary>
		public Uri Uri { get; }

		/// <summary>
		/// Gets the cache records xml document.
		/// </summary>
		public XmlDocument CachaRecordsXml { get; }

		/// <summary>
		/// Gets the list of records.
		/// </summary>
		public IReadOnlyList<CacheRecordNode> Records
		{
			get
			{
				if (myRecords == null)
				{
					var cacheRecordNodes = this.TopNode.SelectNodes("d:r", base.NameSpaceManager);
					foreach (XmlNode recordsNode in cacheRecordNodes)
					{
						myRecords.Add(new CacheRecordNode(recordsNode, base.NameSpaceManager));
					}
				}
				return myRecords;
			}
		}

		/// <summary>
		/// Gets the count of total records.
		/// </summary>
		public int Count
		{
			get { return int.Parse(this.TopNode.Attributes["count"].Value); }
		}

		/// <summary>
		/// Gets or sets the reference to the internal package part.
		/// </summary>
		internal Packaging.ZipPackagePart Part { get; set; }

		private ExcelPivotCacheDefinition CacheDefinition { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an existing <see cref="ExcelPivotCacheRecords"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="cacheRecordsXml">The <see cref="ExcelPivotCacheRecords"/> xml document.</param>
		/// <param name="targetUri">The <see cref="ExcelPivotCacheRecords"/> target uri.</param>
		/// <param name="cacheDefinition">The cache definition of the pivot table.</param>
		public ExcelPivotCacheRecords(XmlNamespaceManager ns, XmlDocument cacheRecordsXml, Uri targetUri, ExcelPivotCacheDefinition cacheDefinition) : base(ns, null)
		{
			this.CachaRecordsXml = cacheRecordsXml;
			base.TopNode = cacheRecordsXml.SelectSingleNode($"d:{ExcelPivotCacheRecords.Name}", ns);
			this.Uri = targetUri;
			this.CacheDefinition = cacheDefinition;

			var cacheRecordNodes = this.TopNode.SelectNodes("d:r", base.NameSpaceManager);
			foreach (XmlNode record in cacheRecordNodes)
			{
				myRecords.Add(new CacheRecordNode(record, base.NameSpaceManager));
			}
		}

		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotCacheRecords"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="package">The <see cref="Packaging.ZipPackage"/> of the Excel package.</param>
		/// <param name="tableId">The <see cref="ExcelPivotTable"/>'s ID.</param>
		/// <param name="cacheDefinition">The cache definition of the pivot table.</param>
		public ExcelPivotCacheRecords(XmlNamespaceManager ns, Packaging.ZipPackage package, ref int tableId, ExcelPivotCacheDefinition cacheDefinition) : base(ns, null)
		{
			// CacheRecord. Create an empty one.
			this.Uri = XmlHelper.GetNewUri(package, $"/xl/pivotCache/{ExcelPivotCacheRecords.Name}{{0}}.xml", ref tableId);
			var cacheRecord = new XmlDocument();
			cacheRecord.LoadXml($"<{ExcelPivotCacheRecords.Name} xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
			var recPart = package.CreatePart(this.Uri, ExcelPackage.schemaPivotCacheRecords);
			cacheRecord.Save(recPart.GetStream());

			base.TopNode = cacheRecord.FirstChild;
			this.CacheDefinition = cacheDefinition;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Update the cache items.
		/// </summary>
		/// <param name="row">The row in the source table.</param>
		/// <param name="col">The column in the source table.</param>
		/// <param name="value">The new value.</param>
		public void UpdateRecord(int row, int col, object value)
		{
			var targetRecord = myRecords[row];
			targetRecord.UpdateItem(col, value, this.CacheDefinition.CacheFields[col]);
		}
		#endregion
	}
}