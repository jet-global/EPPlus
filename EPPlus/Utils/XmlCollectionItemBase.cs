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
using System.Xml;

namespace OfficeOpenXml.Utils
{
	/// <summary>
	/// The base class for xml collection items.
	/// </summary>
	public abstract class XmlCollectionItemBase : XmlHelper
	{
		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml node.</param>
		public XmlCollectionItemBase(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
		}
		#endregion

		#region Public Virtual Methods
		/// <summary>
		/// Adds this <see cref="XmlCollectionItemBase"/>'s <see cref="XmlNode"/> to the specified <paramref name="parentNode"/>.
		/// </summary>
		/// <param name="parentNode">The parent xml node to append to.</param>
		public virtual void AddSelf(XmlNode parentNode)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			parentNode.AppendChild(base.TopNode);
		}

		/// <summary>
		/// Inserts this <see cref="XmlCollectionItemBase"/>'s <see cref="XmlNode"/> into the specified <paramref name="index"/> of the <paramref name="parentNode"/>.
		/// </summary>
		/// <param name="index">The specified position to insert into.</param>
		/// <param name="parentNode">The parent node to append to.</param>
		public virtual void InsertSelf(int index, XmlNode parentNode)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			var childNodes = parentNode.ChildNodes;
			var childrenCount = childNodes.Count;
			if (index >= childrenCount)
			{
				parentNode.AppendChild(base.TopNode);
				return;
			}
			var followingNode = childNodes[index];
			parentNode.InsertBefore(base.TopNode, followingNode);
		}

		/// <summary>
		/// Removes this <see cref="XmlCollectionItemBase"/>'s <see cref="XmlNode"/> from the specified <paramref name="parentNode"/>.
		/// </summary>
		/// <param name="parentNode">The parent node to remove from.</param>
		public virtual void RemoveSelf(XmlNode parentNode)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			parentNode.RemoveChild(base.TopNode);
		}
		#endregion
	}
}