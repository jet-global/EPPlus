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
 *******************************************************************************
 * Jan Källman		Initial Release		     
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents a comment on an excel cell.
	/// </summary>
	public class ExcelComment
	{
		#region Constants
		private const string AUTHORS_PATH = "d:comments/d:authors";
		private const string AUTHOR_PATH = "d:comments/d:authors/d:author";
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the name of the Author of the comment.
		/// </summary>
		public string Author
		{
			get
			{
				int authorRef = this.CommentHelper.GetXmlNodeInt("@authorId");
				return this.CommentHelper.TopNode.OwnerDocument.SelectSingleNode(string.Format("{0}[{1}]", AUTHOR_PATH, authorRef + 1), this.CommentHelper.NameSpaceManager).InnerText;
			}
			set
			{
				int authorRef = GetAuthor(value);
				this.CommentHelper.SetXmlNodeString("@authorId", authorRef.ToString());
			}
		}

		/// <summary>
		/// Gets or sets the text of the comment.
		/// </summary>
		public string Text
		{
			get { return this.RichText.Text ?? string.Empty; }
			set { this.RichText.Text = value; }
		}

		/// <summary>
		/// Gets the font of the first richtext item.
		/// </summary>
		public ExcelRichText Font
		{
			get
			{
				if (this.RichText.Count > 0)
					return this.RichText[0];
				return null;
			}
		}

		/// <summary>
		/// Gets or sets the Rich Text Collection.
		/// </summary>
		public ExcelRichTextCollection RichText { get; set; }

		/// <summary>
		/// Gets or sets the cell reference that specifies which cell this comment is associated with.
		/// </summary>
		internal string Reference
		{
			get { return this.CommentHelper.GetXmlNodeString("@ref"); }
			set
			{
				this.Range.Address = value;
				this.CommentHelper.SetXmlNodeString("@ref", value);
			}
		}

		/// <summary>
		/// Gets or sets the <see cref="ExcelRangeBase"/> that the comment refers to.
		/// </summary>
		public ExcelRangeBase Range { get; set; }

		/// <summary>
		/// Gets the <see cref="XmlHelper"/> for this <see cref="ExcelComment"/>.
		/// </summary>
		internal XmlHelper CommentHelper { get; }
		#endregion

		#region Constructors
		internal ExcelComment(XmlNamespaceManager ns, XmlNode commentTopNode, ExcelRangeBase cell)
		{
			this.CommentHelper = XmlHelperFactory.Create(ns, commentTopNode);
			var textElem = commentTopNode.SelectSingleNode("d:text", ns);
			if (textElem == null)
			{
				textElem = commentTopNode.OwnerDocument.CreateElement("text", ExcelPackage.schemaMain);
				commentTopNode.AppendChild(textElem);
			}
			this.Range = cell;
			this.RichText = new ExcelRichTextCollection(ns, textElem);
		}
		#endregion

		#region Private Methods
		private int GetAuthor(string value)
		{
			int authorRef = 0;
			bool found = false;
			foreach (XmlElement node in this.CommentHelper.TopNode.OwnerDocument.SelectNodes(AUTHOR_PATH, this.CommentHelper.NameSpaceManager))
			{
				if (node.InnerText == value)
				{
					found = true;
					break;
				}
				authorRef++;
			}
			if (!found)
			{
				var elem = this.CommentHelper.TopNode.OwnerDocument.CreateElement("d", "author", ExcelPackage.schemaMain);
				this.CommentHelper.TopNode.OwnerDocument.SelectSingleNode(AUTHORS_PATH, this.CommentHelper.NameSpaceManager).AppendChild(elem);
				elem.InnerText = value;
			}
			return authorRef;
		}
		#endregion
		
	}
}
