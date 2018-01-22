using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlusTest.ConditionalFormatting.Rules
{
	[TestClass]
	public class ExcelConditionalFormattingRuleTest
	{
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void AddressPropertyTest()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				var rule = formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.DataBar);
				Assert.AreEqual("G5:G16,S6:S8", rule.Address.Address);
				Assert.AreEqual("G5:G16 S6:S8", rule.Node.ParentNode.Attributes["sqref"].Value);

				rule.Address = new ExcelAddress("A1:B5,B3:F7,Z8,W2");

				Assert.AreEqual("A1:B5,B3:F7,Z8,W2", rule.Address.Address);
				Assert.AreEqual("A1:B5 B3:F7 Z8 W2", rule.Node.ParentNode.Attributes["sqref"].Value);
			}
		}
	}
}
