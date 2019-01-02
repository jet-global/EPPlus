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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Utils
{
	/// <summary>
	/// The base class for xml collections.
	/// </summary>
	/// <typeparam name="T">The type of the collection.</typeparam>
	public abstract class XmlCollectionBase<T> : XmlHelper, IEnumerable<T> where T : XmlCollectionItemBase
	{
		#region Class Variables
		private List<T> myCollection;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the count of items in the collection.
		/// </summary>
		public int Count
		{
			get { return base.GetXmlNodeIntNull("@count") ?? 0; }
			private set { base.SetXmlNodeString("@count", value.ToString()); }
		}


		/// <summary>
		/// Gets the item at the given index.
		/// </summary>
		/// <param name="index">The position of the item in the collection.</param>
		/// <returns>The item at the specified index.</returns>
		public T this[int index]
		{
			get
			{
				if (index < 0 || index >= this.Collection.Count)
					throw new IndexOutOfRangeException($"Index out of range: {index}");
				return this.Collection[index];
			}
		}

		private List<T> Collection
		{
			get
			{
				if (myCollection == null)
				{
					myCollection = this.LoadItems();
					this.Count = myCollection.Count();
				}
				return myCollection;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The root XML node of the collection.</param>
		internal XmlCollectionBase(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
		}
		#endregion

		#region Abstract Methods
		/// <summary>
		/// Loads the initial collection of items from XML.
		/// </summary>
		/// <returns>A new list of items from the XML.</returns>
		protected abstract List<T> LoadItems();
		#endregion

		#region Protected Methods
		/// <summary>
		/// Adds an item to the collection.
		/// </summary>
		/// <param name="item">The item to add.</param>
		protected void AddItem(T item)
		{
			this.Collection.Add(item);
			item.AddSelf(base.TopNode);
			this.Count++;
		}

		/// <summary>
		/// Inserts an item to the collection at the specified index.
		/// </summary>
		/// <param name="index">The index to insert at.</param>
		/// <param name="item">The item to add.</param>
		protected void InsertItem(int index, T item)
		{
			this.Collection.Insert(index, item);
			item.InsertSelf(index, base.TopNode);
			this.Count++;
		}

		/// <summary>
		/// Determines if an item is in the collection.
		/// </summary>
		/// <param name="item">The specified item.</param>
		/// <returns>True if the item is in the collection. Otherwise, false.</returns>
		protected bool ContainsItem(T item)
		{
			return this.Collection.Contains(item);
		}

		/// <summary>
		/// Removes an item from the collection.
		/// </summary>
		/// <param name="item">The item to remove.</param>
		protected void RemoveItem(T item)
		{
			this.Collection.Remove(item);
			item.RemoveSelf(base.TopNode);
			this.Count--;
		}

		/// <summary>
		/// Clears all the items in the collection.
		/// </summary>
		protected void ClearItems()
		{
			this.Collection.Clear();
			base.TopNode.RemoveAll();
			this.Count = 0;
		}
		#endregion

		#region IEnumerable Overrides
		/// <summary>
		/// Gets the generic enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator<T> GetEnumerator()
		{
			return this.Collection.GetEnumerator();
		}

		/// <summary>
		/// Gets the specified type enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.Collection.GetEnumerator();
		}
		#endregion
	}
}