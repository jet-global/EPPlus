using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.ConditionalFormatting.Rules
{
	[TestClass]
	public class X14ConditionalFormattingRuleTest
	{
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void X14AddressPropertyTest()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var rule = worksheet.X14ConditionalFormatting.X14Rules.First(r => r.TopNode.ChildNodes[0].Attributes["type"].Value == "dataBar");

				Assert.AreEqual("G5:G16,S6:S8", rule.Address);
				Assert.AreEqual("G5:G16 S6:S8", rule.GetXmlNodeString("xm:sqref"));

				rule.Address = "A1:B5,B3:F7,Z8,W2";

				Assert.AreEqual("A1:B5,B3:F7,Z8,W2", rule.Address);
				Assert.AreEqual("A1:B5 B3:F7 Z8 W2", rule.GetXmlNodeString("xm:sqref"));
			}
		}
	}
}
