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
	/// Represents an Excel Slicer Cache.
	/// When this part exists, it can be found at /xl/slicerCaches/slicerCacheN.xml.
	/// </summary>
	public class ExcelSlicerCache : XmlHelper
	{
		#region Class Variables
		private List<PivotTableNode> myPivotTables = new List<PivotTableNode>();
		private List<OlapDataNode> myOlapDataSources = new List<OlapDataNode>();
		private List<TabularDataNode> myTabularDataSources = new List<TabularDataNode>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the name of this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public string Name
		{
			get { return this.TopNode.Attributes["name"].Value; }
			set
			{
				this.TopNode.Attributes["name"].Value = value;
				if (this.Slicer != null)
					this.Slicer.TopNode.Attributes["cache"].Value = value;
			}
		}

		/// <summary>
		/// Gets or sets the source name of this <see cref="ExcelSlicerCache"/>.
		/// </summary>
		public string SourceName
		{
			get { return this.TopNode.Attributes["sourceName"]?.Value; }
			set
			{
				var attribute = this.TopNode.Attributes["sourceName"];
				if (attribute == null)
				{
					attribute = this.TopNode.OwnerDocument.CreateAttribute("sourceName");
					this.TopNode.Attributes.Append(attribute);
				}
				attribute.Value = value;
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
		/// Gets a readonly list of <see cref="PivotTableNode"/>s that wrap the <pivotTable /> element.
		/// </summary>
		public IReadOnlyList<PivotTableNode> PivotTables
		{
			get { return myPivotTables; }
		}

		/// <summary>
		/// Gets a readonly list of olap data sources for this slicer cache.
		/// </summary>
		public IReadOnlyList<OlapDataNode> OlapDataSources
		{
			get { return myOlapDataSources; }
		}

		/// <summary>
		/// Gets a readonly list of table data sources for this slicer cache.
		/// </summary>
		public IReadOnlyList<TabularDataNode> TabularDataSources
		{
			get { return myTabularDataSources; }
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
		internal ExcelSlicerCache(XmlNode node, XmlNamespaceManager namespaceManager, Uri slicerCacheUri, XmlDocument part) : base(namespaceManager, node)
		{
			this.SlicerCacheUri = slicerCacheUri;
			this.Part = part;
			this.Name = node.Attributes["name"].Value;
			foreach (XmlNode pivotTableNode in this.TopNode.SelectNodes("default:pivotTables/default:pivotTable", this.NameSpaceManager))
			{
				myPivotTables.Add(new PivotTableNode(pivotTableNode));
			}
			foreach (XmlNode olapDataNode in this.TopNode.SelectNodes("default:data/default:olap", this.NameSpaceManager))
			{
				myOlapDataSources.Add(new OlapDataNode(olapDataNode, this.NameSpaceManager));
			}
			foreach (XmlNode tabularDataNode in this.TopNode.SelectNodes("default:data/default:tabular", this.NameSpaceManager))
			{
				myTabularDataSources.Add(new TabularDataNode(tabularDataNode, this.NameSpaceManager));
			}
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
