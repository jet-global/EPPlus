using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A reference item used for custom sorting.
	/// </summary>
	public class AutoSortScopeReference : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets the type of this reference.
		/// </summary>
		public PivotCacheRecordType Type { get; private set; }

		/// <summary>
		/// Gets the value of this reference.
		/// </summary>
		public string Value
		{
			get
			{
				if (this.Type == PivotCacheRecordType.m)
					return null;
				return base.GetXmlNodeString("@v");
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="AutoSortScopeReference"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public AutoSortScopeReference(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Type = (PivotCacheRecordType)Enum.Parse(typeof(PivotCacheRecordType), node.Name);
		}
		#endregion
	}
}
