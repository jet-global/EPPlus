/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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
using System.Linq;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable.Formats
{
	/// <summary>
	/// A collection of pivot table conditional formats.
	/// </summary>
	public class PivotTableConditionalFormatsCollection : XmlCollectionBase<PivotTableConditionalFormat>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="PivotTableConditionalFormatsCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public PivotTableConditionalFormatsCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Update the values of the shared items in the references collection.
		/// </summary>
		/// <param name="root">The row items tree node that has the shared string values in the given list.</param>
		/// <param name="conditionalFormattingStringValues">The list of cached references shared string values.</param>
		public void UpdateConditionalFormatReferences(PivotItemTreeNode root, List<Tuple<int, int, List<Tuple<int, string>>>> conditionalFormattingStringValues)
		{
			if (conditionalFormattingStringValues.Count > 0)
			{
				foreach (var conditionalFormat in this)
				{
					var matchingPriorityList = conditionalFormattingStringValues.Where(i => i.Item1 == conditionalFormat.Priority);
					int matchingPriorityListIndex = 0;
					if (matchingPriorityList.Count() > 0)
					{
						foreach (var pivotArea in conditionalFormat.PivotAreasCollection)
						{
							int sharedStringIndex = 0;
							foreach (var reference in pivotArea.ReferencesCollection)
							{
								if (reference.FieldIndex < 0)
									continue;
								int temp = matchingPriorityListIndex;
								foreach (var item in reference.SharedItems)
								{
									var cf = matchingPriorityList.ElementAt(temp++);
									var sharedString = cf.Item3[sharedStringIndex].Item2;
									var node = root.GetNode(root, sharedString);
									if (node != null)
										item.Value = node.PivotFieldItemIndex.ToString();
								}
								sharedStringIndex++;
							}
							matchingPriorityListIndex++;
						}
					}
				}
			}
		}
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the <see cref="PivotTableConditionalFormat"/>s from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="PivotTableConditionalFormat"/>s.</returns>
		protected override List<PivotTableConditionalFormat> LoadItems()
		{
			var conditionalFormatCollection = new List<PivotTableConditionalFormat>();
			foreach (XmlNode conditionalFormat in base.TopNode.SelectNodes("d:conditionalFormat", this.NameSpaceManager))
			{
				conditionalFormatCollection.Add(new PivotTableConditionalFormat(this.NameSpaceManager, conditionalFormat));
			}
			return conditionalFormatCollection;
		}
		#endregion
	}
}
