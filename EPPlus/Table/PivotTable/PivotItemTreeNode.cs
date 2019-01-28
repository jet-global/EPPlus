using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Maintains state for a tree data structure that represents a pivot table.
	/// </summary>
	internal class PivotItemTreeNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the cache record item node "v" index.
		/// </summary>
		public int Value { get; set; }
		
		/// <summary>
		/// Gets or sets the index of the cache record.
		/// </summary>
		public int CacheRecordIndex { get; set; }

		/// <summary>
		/// Gets or sets the index of the datafield referenced.
		/// </summary>
		public int DataFieldIndex { get; set; }

		/// <summary>
		/// Gets or sets the index of the pivot field referenced.
		/// </summary>
		public int PivotFieldIndex { get; set; } = -2;

		/// <summary>
		/// Gets or sets the index of the referenced pivot field item.
		/// </summary>
		public int PivotFieldItemIndex { get; set; } = -2;

		/// <summary>
		/// Gets or sets whether or not subtotal top is enabled.
		/// </summary>
		public bool SubtotalTop { get; set; } = true; // Excel defaults this to true, so we will too.

		/// <summary>
		/// Gets whether or not this node represents a datafield.
		/// </summary>
		public bool IsDataField => this.CacheRecordIndex == -2 && this.Value == -2;

		/// <summary>
		/// Gets or sets the list of children that this node parents.
		/// </summary>
		public List<PivotItemTreeNode> Children { get; set; } = new List<PivotItemTreeNode>();
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="value">The cache record item node "v" index.</param>
		/// <param name="cacheRecordIndex">The index of the cache record.</param>
		public PivotItemTreeNode(int value, int cacheRecordIndex)
		{
			this.Value = value;
			this.CacheRecordIndex = cacheRecordIndex;
		}
		#endregion

		#region Public Methods
		public void AddChild(PivotItemTreeNode child)
		{
			this.Children.Add(child);
		}

		public bool HasChild(int value)
		{
			return this.Children?.Any(i => i.Value == value) ?? false;
		}

		public PivotItemTreeNode GetChildNode(int value)
		{
			return this.Children.Find(i => i.Value == value);
		}

		public PivotItemTreeNode Clone()
		{
			var clone = new PivotItemTreeNode(this.Value, this.CacheRecordIndex);
			clone.DataFieldIndex = this.DataFieldIndex;
			clone.PivotFieldIndex = this.PivotFieldIndex;
			clone.PivotFieldItemIndex = this.PivotFieldItemIndex;
			clone.SubtotalTop = this.SubtotalTop;

			foreach (var child in this.Children)
			{
				clone.AddChild(child.Clone());
			}
			return clone;
		}

		public void RecursivelySetDataFieldIndex(int index)
		{
			this.DataFieldIndex = index;
			foreach (var child in this.Children)
			{
				child.RecursivelySetDataFieldIndex(index);
			}
		}
		#endregion

	}
}
