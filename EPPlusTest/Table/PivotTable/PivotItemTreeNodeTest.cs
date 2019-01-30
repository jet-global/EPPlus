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
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class PivotItemTreeNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ConstructPivotItemTreeNodeTest()
		{
			var node = new PivotItemTreeNode(5);
			Assert.AreEqual(-2, node.PivotFieldIndex);
			Assert.AreEqual(-2, node.PivotFieldItemIndex);
			Assert.AreEqual(5, node.Value);
			Assert.IsTrue(node.SubtotalTop);
			Assert.IsFalse(node.IsDataField);
			Assert.IsFalse(node.HasChildren);
		}

		[TestMethod]
		public void ConstructPivotItemTreeNodeDataFieldTest()
		{
			var node = new PivotItemTreeNode(-2);
			Assert.AreEqual(-2, node.PivotFieldIndex);
			Assert.AreEqual(-2, node.PivotFieldItemIndex);
			Assert.AreEqual(-2, node.Value);
			Assert.IsTrue(node.SubtotalTop);
			Assert.IsTrue(node.IsDataField);
			Assert.IsFalse(node.HasChildren);
		}
		#endregion

		#region AddChild Tests
		[TestMethod]
		public void AddChildWithDefaultValues()
		{
			var node = new PivotItemTreeNode(-1);
			var child = node.AddChild(4);
			Assert.AreEqual(child, node.Children.Single());
			Assert.AreEqual(-2, child.PivotFieldIndex);
			Assert.AreEqual(-2, child.PivotFieldItemIndex);
			Assert.AreEqual(4, child.Value);
			Assert.IsTrue(child.SubtotalTop);
			Assert.IsFalse(child.IsDataField);
			Assert.IsFalse(child.HasChildren);
		}

		[TestMethod]
		public void AddChildWithNonDefaultValues()
		{
			var node = new PivotItemTreeNode(-1);
			var child = node.AddChild(4, 2, 3);
			Assert.AreEqual(child, node.Children.Single());
			Assert.AreEqual(2, child.PivotFieldIndex);
			Assert.AreEqual(3, child.PivotFieldItemIndex);
			Assert.AreEqual(4, child.Value);
			Assert.IsTrue(child.SubtotalTop);
			Assert.IsFalse(child.IsDataField);
			Assert.IsFalse(child.HasChildren);
		}
		#endregion

		#region HasChild Tests
		[TestMethod]
		public void HasChildTrue()
		{
			var node = new PivotItemTreeNode(-1);
			var child = node.AddChild(4);
			Assert.IsTrue(node.HasChild(4));
		}

		[TestMethod]
		public void HasChildFalse()
		{
			var node = new PivotItemTreeNode(-1);
			var child = node.AddChild(4);
			Assert.IsFalse(node.HasChild(3));
		}
		#endregion

		#region GetChildNode Tests
		[TestMethod]
		public void GetChildNodeHasChild()
		{
			var node = new PivotItemTreeNode(-1);
			var child = node.AddChild(4);
			Assert.AreEqual(child, node.GetChildNode(4));
		}

		[TestMethod]
		public void GetChildNodeDoesNotHaveChild()
		{
			var node = new PivotItemTreeNode(-1);
			var child = node.AddChild(4);
			Assert.IsNull(node.GetChildNode(5));
		}
		#endregion

		#region Clone Tests
		[TestMethod]
		public void CloneTest()
		{
			var node = new PivotItemTreeNode(3);
			node.DataFieldIndex = 4;
			node.PivotFieldIndex = 5;
			node.PivotFieldItemIndex = 6;
			node.SubtotalTop = false;
			var child1 = node.AddChild(2);
			var child1Child = child1.AddChild(8);

			var clone = node.Clone();
			var child1Clone = clone.Children.Single();
			var child1ChildClone = child1Clone.Children.Single();
			Assert.AreEqual(child1.Value, child1Clone.Value);
			Assert.AreEqual(child1Child.Value, child1ChildClone.Value);
			Assert.AreNotEqual(child1, child1Clone);
			Assert.AreNotEqual(child1Child, child1ChildClone);
		}
		#endregion

		#region RecursivelySetDataFieldIndex Tests
		[TestMethod]
		public void RecursivelySetDataFieldIndex()
		{
			var node = new PivotItemTreeNode(3);
			node.DataFieldIndex = 4;
			var child1 = node.AddChild(2);
			child1.DataFieldIndex = 5;
			var child2 = node.AddChild(3);
			child2.DataFieldIndex = 7;
			var child1Child = child1.AddChild(8);
			child1Child.DataFieldIndex = 6;
			node.RecursivelySetDataFieldIndex(2);
			Assert.AreEqual(2, node.DataFieldIndex);
			Assert.AreEqual(2, child1.DataFieldIndex);
			Assert.AreEqual(2, child2.DataFieldIndex);
			Assert.AreEqual(2, child1Child.DataFieldIndex);
		}
		#endregion

		#region ExpandIfDataFieldNode Tests
		[TestMethod]
		public void ExpandIfDataFieldNodeTest()
		{
			var node = new PivotItemTreeNode(3);
			var child1 = node.AddChild(-2); // Data field
			var child1Child1 = child1.AddChild(8);
			var child1Child2 = child1.AddChild(9);

			Assert.AreEqual(1, node.Children.Count);
			node.ExpandIfDataFieldNode(3);
			Assert.AreEqual(3, node.Children.Count);

			foreach (var child in node.Children)
			{
				Assert.AreEqual(-2, child.Value);
				Assert.AreEqual(2, child.Children.Count);
				Assert.AreEqual(8, child.Children[0].Value);
				Assert.AreEqual(9, child.Children[1].Value);
			}
		}

		[TestMethod]
		public void ExpandIfDataFieldNodeNotDataFieldNode()
		{
			var node = new PivotItemTreeNode(3);
			var child1 = node.AddChild(2); // Not data field
			var child1Child1 = child1.AddChild(8);
			var child1Child2 = child1.AddChild(9);

			Assert.AreEqual(1, node.Children.Count);
			node.ExpandIfDataFieldNode(3);
			Assert.AreEqual(1, node.Children.Count);
			Assert.AreEqual(8, node.Children[0].Children[0].Value);
			Assert.AreEqual(9, node.Children[0].Children[1].Value);
		}

		[TestMethod]
		public void ExpandIfDataFieldNodeHasMultipleChildren()
		{
			var node = new PivotItemTreeNode(3);
			var child1 = node.AddChild(-2); // Data field
			var child2 = node.AddChild(-2); // Data field
			var child1Child1 = child1.AddChild(8);
			var child1Child2 = child1.AddChild(9);

			// The children are data fields, but are not expanded because there 
			// are multiple. This is an error state.
			Assert.AreEqual(2, node.Children.Count);
			node.ExpandIfDataFieldNode(3);
			Assert.AreEqual(2, node.Children.Count);
			Assert.AreEqual(8, node.Children[0].Children[0].Value);
			Assert.AreEqual(9, node.Children[0].Children[1].Value);
		}
		#endregion
	}
}
