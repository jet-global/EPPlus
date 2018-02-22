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
	}
}
