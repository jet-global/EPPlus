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
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml
{
	/// <summary>
	/// Collection of Excelcomment objects
	/// </summary>  
	public class ExcelCommentCollection : IEnumerable, IDisposable
	{
		private List<ExcelComment> myComments;

		List<ExcelComment> Comments
		{
			get
			{
				if (myComments == null)
					myComments = new List<ExcelComment>();
				return myComments;
			}
		}

		internal ExcelCommentCollection(ExcelPackage pck, ExcelWorksheet ws, XmlNamespaceManager ns)
		{
			this.CommentXml = new XmlDocument();
			this.CommentXml.PreserveWhitespace = false;
			if (!ns.HasNamespace("sme"))
				ns.AddNamespace("sme", ExcelPackage.schemaMicrosoftExcel);
			if (!ns.HasNamespace("smo"))
				ns.AddNamespace("smo", ExcelPackage.schemaMicrosoftOffice);
			this.NameSpaceManager = ns;
			this.Worksheet = ws;
			this.CreateXml(pck);
			this.AddCommentsFromXml();
		}

		private void CreateXml(ExcelPackage package)
		{
			var commentParts = Worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaComment);
			bool isLoaded = false;
			this.CommentXml = new XmlDocument();
			foreach (var commentPart in commentParts)
			{
				this.Uri = UriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
				this.Part = package.Package.GetPart(Uri);
				XmlHelper.LoadXmlSafe(this.CommentXml, this.Part.GetStream());
				this.RelId = commentPart.Id;
				isLoaded = true;
			}
			//Create a new document
			if (!isLoaded)
			{
				this.CommentXml.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors /><commentList /></comments>");
				Uri = null;
			}
		}

		private void AddCommentsFromXml()
		{
			foreach (XmlElement node in this.CommentXml.SelectNodes("//d:commentList/d:comment", NameSpaceManager))
			{
				var comment = new ExcelComment(NameSpaceManager, node, new ExcelRangeBase(this.Worksheet, node.GetAttribute("ref")));
				this.Comments.Add(comment);
				this.Worksheet._commentsStore.SetValue(comment.Range._fromRow, comment.Range._fromCol, this.Comments.Count - 1);
			}
		}

		/// <summary>
		/// Access to the comment xml document
		/// </summary>
		public XmlDocument CommentXml { get; set; }
		internal Uri Uri { get; set; }
		internal string RelId { get; set; }
		internal XmlNamespaceManager NameSpaceManager { get; set; }
		internal Packaging.ZipPackagePart Part
		{
			get;
			set;
		}

		/// <summary>
		/// A reference to the worksheet object
		/// </summary>
		public ExcelWorksheet Worksheet
		{
			get;
			set;
		}

		/// <summary>
		/// Number of comments in the collection
		/// </summary>
		public int Count
		{
			get
			{
				return this.Comments.Count;
			}
		}

		/// <summary>
		/// Indexer for the comments collection
		/// </summary>
		/// <param name="index">The index</param>
		/// <returns>The comment</returns>
		public ExcelComment this[int index]
		{
			get
			{
				if (index < 0 || index >= this.Comments.Count)
					throw (new ArgumentOutOfRangeException("Comment index out of range"));
				return this.Comments[index] as ExcelComment;
			}
		}

		/// <summary>
		/// Indexer for the comments collection
		/// </summary>
		/// <param name="cell">The comment cell.</param>
		/// <returns>The comment</returns>
		public ExcelComment this[ExcelCellAddress cell]
		{
			get
			{
				int i = -1;
				if (this.Worksheet._commentsStore.Exists(cell.Row, cell.Column, out i))
					return this.Comments[i];
				else
					return null;
			}
		}

		/// <summary>
		/// Adds a comment.
		/// </summary>
		/// <param name="cell">The cell to which the comment is added.</param>
		/// <param name="text">The text of the comment.</param>
		/// <param name="author">The author of the comment.</param>
		public ExcelComment Add(ExcelRangeBase cell, string text, string author)
		{
			var element = this.CommentXml.CreateElement("comment", ExcelPackage.schemaMain);
			// Make sure the nodes are sorted, by column and then by row.
			int nextCommentRow = cell._fromRow;
			int nextCommentColumn = cell._fromCol;
			if (this.Comments.Count == 0 || !this.Worksheet._commentsStore.NextCell(ref nextCommentRow, ref nextCommentColumn))
				this.CommentXml.SelectSingleNode("d:comments/d:commentList", this.NameSpaceManager).AppendChild(element);
			else
			{
				ExcelComment nextComment = this.Comments[Worksheet._commentsStore.GetValue(nextCommentRow, nextCommentColumn)];
				nextComment.CommentHelper.TopNode.ParentNode.InsertBefore(element, nextComment.CommentHelper.TopNode);
			}
			ExcelComment comment = new ExcelComment(this.NameSpaceManager, element, cell);
			comment.Reference = new ExcelAddress(cell._fromRow, cell._fromCol, cell._fromRow, cell._fromCol).Address;
			comment.RichText.Add(text);
			if (author != string.Empty)
				comment.Author = author;
			this.Comments.Add(comment);
			this.Worksheet._commentsStore.SetValue(cell.Start.Row, cell.Start.Column, this.Comments.Count - 1);
			if (!this.Worksheet.ExistsValueInner(cell._fromRow, cell._fromCol))
				this.Worksheet.SetValueInner(cell._fromRow, cell._fromCol, null);
			return comment;
		}

		/// <summary>
		/// Adds a comment that is styled the same as the specified <paramref name="copyComment"/>.
		/// </summary>
		/// <param name="cell">The cell to which the comment is added.</param>
		/// <param name="copyComment">The comment to copy.</param>
		public void Add(ExcelRangeBase cell, ExcelComment copyComment)
		{
			var element = this.CommentXml.CreateElement("comment", ExcelPackage.schemaMain);
			// Make sure the nodes come in order.
			int nextCommentRow = cell._fromRow;
			int nextCommentColumn = cell._fromCol;
			if (this.Comments.Count == 0 || !this.Worksheet._commentsStore.NextCell(ref nextCommentRow, ref nextCommentColumn))
				this.CommentXml.SelectSingleNode("d:comments/d:commentList", this.NameSpaceManager).AppendChild(element);
			else
			{
				ExcelComment nextComment = this.Comments[Worksheet._commentsStore.GetValue(nextCommentRow, nextCommentColumn)];
				nextComment.CommentHelper.TopNode.ParentNode.InsertBefore(element, nextComment.CommentHelper.TopNode);
			}
			ExcelComment comment = new ExcelComment(this.NameSpaceManager, element, cell);
			comment.RichText = copyComment.RichText;
			// Copy text styling.
			comment.CommentHelper.TopNode.SelectSingleNode(".//d:text", this.NameSpaceManager).InnerXml = copyComment.CommentHelper.TopNode.SelectSingleNode(".//d:text", this.NameSpaceManager).InnerXml;
			string author = copyComment.Author;
			if (string.IsNullOrEmpty(author))
				author = Thread.CurrentPrincipal.Identity.Name;
			comment.Reference = new ExcelAddress(cell._fromRow, cell._fromCol, cell._fromRow, cell._fromCol).Address;
			comment.Author = author;
			float rowMarginOffset = 0, columnMarginOffset = 0;
			int rowDirection = comment.Range._fromRow.CompareTo(copyComment.Range._fromRow);
			var fromRow = Math.Min(comment.Range._fromRow, copyComment.Range._fromRow);
			var toRow = Math.Max(comment.Range._fromRow, copyComment.Range._fromRow);
			for (int i = fromRow; i < toRow; i++)
			{
				rowMarginOffset += (float)this.Worksheet.Row(i).Height;
			}
			int columnDirection = comment.Range._fromCol.CompareTo(copyComment.Range._fromCol);
			var fromColumn = Math.Min(comment.Range._fromCol, copyComment.Range._fromCol);
			var toColumn = Math.Max(comment.Range._fromCol, copyComment.Range._fromCol);
			for (int i = fromColumn; i < toColumn; i++)
			{
				columnMarginOffset += (float)this.Worksheet.Column(i).Width;
			}
			this.Comments.Add(comment);
			this.Worksheet._commentsStore.SetValue(cell._fromRow, cell._fromCol, this.Comments.Count - 1);
			// Check if a value exists otherwise add one so it is saved when the cells collection is iterated.
			if (!this.Worksheet.ExistsValueInner(cell._fromRow, cell._fromCol))
				this.Worksheet.SetValueInner(cell._fromRow, cell._fromCol, null);
		}

		/// <summary>
		/// Removes the comment
		/// </summary>
		/// <param name="comment">The comment to remove</param>
		public void Remove(ExcelComment comment)
		{
			if (comment != null && this.Comments.Contains(comment))
			{
				comment.CommentHelper.TopNode.ParentNode.RemoveChild(comment.CommentHelper.TopNode); //Remove Comment
				this.Comments.Remove(comment);
			}
			else
				throw (new ArgumentException("Comment does not exist in the worksheet"));
		}

		/// <summary>
		/// Shifts all comments based on their address and the location of deleted rows and columns.
		/// </summary>
		/// <param name="fromRow">The start row.</param>
		/// <param name="fromCol">The start column.</param>
		/// <param name="rows">The number of rows to deleted.</param>
		/// <param name="columns">The number of columns to deleted.</param>
		internal void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			List<ExcelComment> deletedComments = new List<ExcelComment>();
			ExcelAddress address = null;
			foreach (ExcelComment comment in this.Comments)
			{
				address = new ExcelAddress(comment.Range);
				if (fromCol > 0 && address._fromCol >= fromCol)
				{
					address = address.DeleteColumn(fromCol, columns);
				}
				if (fromRow > 0 && address._fromRow >= fromRow)
				{
					address = address.DeleteRow(fromRow, rows);
				}
				if (address == null || address.Address == "#REF!")
				{
					deletedComments.Add(comment);
				}
				else
				{
					comment.Reference = address.Address;
				}
			}
			int i = -1;
			List<int> deletedIndices = new List<int>();
			foreach (var comment in deletedComments)
			{
				if (this.Worksheet._commentsStore.Exists(comment.Range._fromRow, comment.Range._fromCol, out i))
				{
					this.Remove(comment);
					deletedIndices.Add(i);
				}
			}
			this.Worksheet._commentsStore.Delete(fromRow, fromCol, rows, columns);
			var commentEnumerator = this.Worksheet._commentsStore.GetEnumerator();
			while (commentEnumerator.MoveNext())
			{
				int offset = deletedIndices.Count(di => commentEnumerator.Value > di);
				commentEnumerator.Value -= offset;
			}
		}

		/// <summary>
		/// Shifts all comments based on their address and the location of inserted rows and columns.
		/// </summary>
		/// <param name="fromRow">The start row.</param>
		/// <param name="fromCol">The start column.</param>
		/// <param name="rows">The number of rows to insert.</param>
		/// <param name="columns">The number of columns to insert.</param>
		public void Insert(int fromRow, int fromCol, int rows, int columns)
		{
			foreach (ExcelComment comment in this.Comments)
			{
				var address = new ExcelAddress(comment.Range);
				if (rows > 0 && address._fromRow >= fromRow)
				{
					comment.Reference = comment.Range.AddRow(fromRow, rows).Address;
				}
				if (columns > 0 && address._fromCol >= fromCol)
				{
					comment.Reference = comment.Range.AddColumn(fromCol, columns).Address;
				}
			}
			this.Worksheet._commentsStore.Insert(fromRow, fromCol, rows, columns);
		}

		void IDisposable.Dispose()
		{
		}

		/// <summary>
		/// Removes the comment at the specified position
		/// </summary>
		/// <param name="Index">The index</param>
		public void RemoveAt(int Index)
		{
			Remove(this[Index]);
		}

		#region IEnumerable Members
		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.Comments.GetEnumerator();
		}
		#endregion

		internal void Clear()
		{
			while (Count > 0)
			{
				RemoveAt(0);
			}
		}
	}
}
