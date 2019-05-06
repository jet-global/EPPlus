using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Formats;

namespace EPPlusTest.Table.PivotTable.Formats
{
	[TestClass]
	public class PivotTableConditionalFormatTest
	{
		#region Constructor Tests
		[TestMethod]
		public void PivotTableConditionalFormatsConstructorTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
					<conditionalFormat priority=""2"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
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
					</conditionalFormat>");
			var conditionalFormat = new PivotTableConditionalFormat(TestUtility.CreateDefaultNSM(), xmlDoc.FirstChild);
			Assert.IsNotNull(conditionalFormat);
			Assert.IsNotNull(conditionalFormat.PivotAreasCollection);
			Assert.AreEqual(2, conditionalFormat.Priority);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotTableConditionalFormatsNullNamespaceManagerTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
					<conditionalFormat priority=""2"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
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
					</conditionalFormat>");
			new PivotTableConditionalFormat(null, xmlDoc.FirstChild);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotTableConditionalFormatsNullNodeTest()
		{
			new PivotTableConditionalFormat(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion
	}
}
