using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Formats;

namespace EPPlusTest.Table.PivotTable.Formats
{
	[TestClass]
	public class PivotTableConditionalFormatsCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void PivotTableConditionalFormatsCollectionConstructorTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
				<conditionalFormats count=""2"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
					<conditionalFormat priority=""2"">
						<pivotAreas count=""1"">
							<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
								<references count=""2"">
									<reference field=""4294967294"" count=""1"" selected=""0"">
										<x v=""1""/>
									</reference>
									<reference field=""1"" count=""1"" selected=""0"">
										<x v=""0""/>
									</reference>
								</references>
							</pivotArea>
					</pivotAreas>
				</conditionalFormat>
				<conditionalFormat priority=""1"">
					<pivotAreas count=""1"">
						<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
							<references count=""4"">
								<reference field=""4294967294"" count=""1"" selected=""0"">
									<x v=""1""/>
								</reference>
								<reference field=""1"" count=""1"" selected=""0"">
									<x v=""1""/>
								</reference>
								<reference field=""2"" count=""1"" selected=""0"">
									<x v=""2""/>
								</reference>
								<reference field=""3"" count=""1"" selected=""0"">
									<x v=""0""/>
								</reference>
							</references>
						</pivotArea>
					</pivotAreas>
				</conditionalFormat>
			</conditionalFormats>");
			var collection = new PivotTableConditionalFormatsCollection(TestUtility.CreateDefaultNSM(), xmlDoc.FirstChild);
			Assert.IsNotNull(collection);
			Assert.AreEqual(2, collection.Count);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotTableConditionalFormatsCollectioNullNamespaceManagerTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
				<conditionalFormats count=""2"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
					<conditionalFormat priority=""2"">
						<pivotAreas count=""1"">
							<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
								<references count=""2"">
									<reference field=""4294967294"" count=""1"" selected=""0"">
										<x v=""1""/>
									</reference>
									<reference field=""1"" count=""1"" selected=""0"">
										<x v=""0""/>
									</reference>
								</references>
							</pivotArea>
					</pivotAreas>
				</conditionalFormat>
				<conditionalFormat priority=""1"">
					<pivotAreas count=""1"">
						<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
							<references count=""4"">
								<reference field=""4294967294"" count=""1"" selected=""0"">
									<x v=""1""/>
								</reference>
								<reference field=""1"" count=""1"" selected=""0"">
									<x v=""1""/>
								</reference>
								<reference field=""2"" count=""1"" selected=""0"">
									<x v=""2""/>
								</reference>
								<reference field=""3"" count=""1"" selected=""0"">
									<x v=""0""/>
								</reference>
							</references>
						</pivotArea>
					</pivotAreas>
				</conditionalFormat>
			</conditionalFormats>");
			new PivotTableConditionalFormatsCollection(null, xmlDoc.FirstChild);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotTableConditionalFormatsCollectionNullNodeTest()
		{
			new PivotTableConditionalFormatsCollection(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion
	}
}
