using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Wraps a <s /> node in <pivotcachedefinition-cacheFields-cacheField-sharedItems />.
	/// </summary>
	public class CacheFieldItem
	{
		#region Properties
		/// <summary>
		/// Gets or sets the value of this item.
		/// </summary>
		public string Value
		{
			get { return this.Node.Attributes["v"].Value; }
			set { this.Node.Attributes["v"].Value = value; }
		}

		private XmlNode Node { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheFieldItem"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheFieldItem"/>.</param>
		public CacheFieldItem(XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Node = node;
		}
		#endregion
	}
}
