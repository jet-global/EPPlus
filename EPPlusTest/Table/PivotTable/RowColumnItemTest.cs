using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class RowColumnItemTest
	{
		#region Constructor Tests
		[TestMethod]
		public void RowColumnItem()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<i xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" r=""1"" i=""1"" t=""grand"">
					<x v=""1""/>
					<x v=""1048832""/>
					<x/>
				</i>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new RowColumnItem(TestUtility.CreateDefaultNSM(), node);
			Assert.AreEqual(3, itemsCollection.MemberPropertyIndex.Count);
			Assert.AreEqual(1, itemsCollection.RepeatedItemsCount);
			Assert.AreEqual(1, itemsCollection.DataFieldIndex);
			Assert.AreEqual("grand", itemsCollection.ItemType);
		}
		#endregion
	}
}
