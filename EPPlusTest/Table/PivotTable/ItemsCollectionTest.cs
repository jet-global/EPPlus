using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ItemsCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ItemsCollection()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ItemsCollection(TestUtility.CreateDefaultNSM(), node);
			Assert.AreEqual(3, itemsCollection.Items.Count);
			Assert.AreEqual(itemsCollection.Count, itemsCollection.Items.Count);
		}
		#endregion
	}
}