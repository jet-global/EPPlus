using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Maintains state for a tree data structure that represents a pivot table.
	/// </summary>
	public class PivotItemTreeNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the cache record item node "v" index.
		/// </summary>
		public int Value { get; set; }
		
		/// <summary>
		/// Gets or sets the list of cache record indices.
		/// </summary>
		public List<int> CacheRecordIndices { get; set; } = new List<int>();

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
		public bool IsDataField => this.Value == -2;

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
		public PivotItemTreeNode(int value)
		{
			this.Value = value;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add the given node as a child of this node.
		/// </summary>
		/// <param name="child">The node to add.</param>
		public void AddChild(PivotItemTreeNode child)
		{
			this.Children.Add(child);
		}

		/// <summary>
		/// Checks whether or not a child already exists with the specified value.
		/// </summary>
		/// <param name="value">The value to look for in the children list.</param>
		/// <returns>True if the child exists, otherwise false.</returns>
		public bool HasChild(int value)
		{
			return this.Children?.Any(i => i.Value == value) ?? false;
		}

		/// <summary>
		/// Gets the child node that has the specified value.
		/// </summary>
		/// <param name="value">The value to look for in the children list.</param>
		/// <returns>The child node if it exists.</returns>
		public PivotItemTreeNode GetChildNode(int value)
		{
			return this.Children.Find(i => i.Value == value);
		}

		/// <summary>
		/// Creates a deep copy of this node.
		/// </summary>
		/// <returns>The newly created node.</returns>
		public PivotItemTreeNode Clone()
		{
			var clone = new PivotItemTreeNode(this.Value);
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

		/// <summary>
		/// Sets the data field index for this node and all it's children.
		/// </summary>
		/// <param name="index">The specified data field index.</param>
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
