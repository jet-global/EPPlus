using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A row or column item object.
	/// </summary>
	public class RowColumnItem : XmlHelper
	{
		#region Class Variables
		private List<int> myMemberPropertyIndexes;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the list of member property indexes.
		/// </summary>
		public IReadOnlyList<int> MemberPropertyIndex
		{
			get
			{
				if (myMemberPropertyIndexes == null)
				{
					myMemberPropertyIndexes = new List<int>();
					var xNodes = base.TopNode.SelectNodes("d:x", base.NameSpaceManager);
					foreach (XmlNode xmlNode in xNodes)
					{
						var value = xmlNode.Attributes["v"]?.Value;
						int index = value == null ? 0 : int.Parse(value);
						myMemberPropertyIndexes.Add(index);
					}
				}
				return myMemberPropertyIndexes;
			}
		}

		/// <summary>
		/// Gets or sets the data field index.
		/// </summary>
		public int DataFieldIndex
		{
			get { return base.GetXmlNodeIntNull("@i") ?? 0; }
			set { base.SetXmlNodeString("@i", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the repeated items count.
		/// </summary>
		public int RepeatedItemsCount
		{
			get { return base.GetXmlNodeIntNull("@r") ?? 0; }
			set { base.SetXmlNodeString("@r", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the item type.
		/// </summary>
		public string ItemType
		{
			get { return base.GetXmlNodeString("@t"); }
			set { base.SetXmlNodeString("@t", value, true); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new <see cref="RowColumnItem"/> object.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The item <see cref="XmlNode"/>.</param>
		public RowColumnItem(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
		}
		#endregion
	}
}