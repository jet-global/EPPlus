using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class StructuredReferenceTest
	{
		#region Constructor Tests
		[TestMethod]
		public void StructuredReferenceWithTableAndColumnAll()
		{
			var structuredReference = new StructuredReference("MyTable[[#All],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.All, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnAllIgnoreCase()
		{
			var structuredReference = new StructuredReference("MyTable[[#ALL],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.All, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnData()
		{
			var structuredReference = new StructuredReference("MyTable[[#Data],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnDataIgnoreCase()
		{
			var structuredReference = new StructuredReference("MyTable[[#DATA],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnHeaders()
		{
			var structuredReference = new StructuredReference("MyTable[[#Headers],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Headers, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnHeadersIgnoreCase()
		{
			var structuredReference = new StructuredReference("MyTable[[#HEADERS],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Headers, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnTotals()
		{
			var structuredReference = new StructuredReference("MyTable[[#Totals],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Totals, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnTotalsIgnoreCase()
		{
			var structuredReference = new StructuredReference("MyTable[[#TOTALS],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Totals, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnThisRow()
		{
			var structuredReference = new StructuredReference("MyTable[[#This Row],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.ThisRow, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnThisRowIgnoreCase()
		{
			var structuredReference = new StructuredReference("MyTable[[#THIS ROW],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.ThisRow, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithMultipleSpecifiers()
		{
			var structuredReference = new StructuredReference("MyTable[[#Headers],[#Data],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Headers | ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithEverySpecifier()
		{
			var structuredReference = new StructuredReference("MyTable[[#All],[#Totals],[#This Row],[#Headers],[#Data],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.All | ItemSpecifiers.Totals | ItemSpecifiers.ThisRow | ItemSpecifiers.Headers | ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithColumnRange()
		{
			var structuredReference = new StructuredReference("MyTable[[#Headers],[MyStartColumn]:[MyEndColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyStartColumn", structuredReference.StartColumn);
			Assert.AreEqual("MyEndColumn", structuredReference.EndColumn);
			Assert.AreEqual(ItemSpecifiers.Headers, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceNoItemSpecifierDefaultsToThisRow()
		{
			var structuredReference = new StructuredReference("MyTable[MyColumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void StructuredReferenceNullReferenceStringThrowsException()
		{
			new StructuredReference(null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void StructuredReferenceEmptyReferenceStringThrowsException()
		{
			new StructuredReference(string.Empty);
		}
		#endregion

		#region HasValidItemSpecifier Tests
		[TestMethod]
		public void HasValidItemSpecifiersValidTests()
		{
			var structuredReference = new StructuredReference("MyTable[#Data]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[#Headers]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[#Totals]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[#This row]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[#All]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Data],[#Headers]]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Data],[#Totals]]");
			Assert.IsTrue(structuredReference.HasValidItemSpecifiers());
		}

		[TestMethod]
		public void HasValidItemSpecifiersInvalidTests()
		{
			var structuredReference = new StructuredReference("MyTable[[#Data],[#This row]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Data],[#All]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Headers],[#Totals]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Headers],[#This row]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Headers],[#All]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Totals],[#This row]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Totals],[#All]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#This row],[#All]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
			structuredReference = new StructuredReference("MyTable[[#Data],[#Headers],[#Totals]]");
			Assert.IsFalse(structuredReference.HasValidItemSpecifiers());
		}
		#endregion
	}
}
