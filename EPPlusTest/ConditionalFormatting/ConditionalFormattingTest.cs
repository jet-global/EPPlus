using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace EPPlusTest
{
	/// <summary>
	/// Tests the Conditional Formatting feature.
	/// </summary>
	[TestClass]
	public class ConditionalFormattingTest
	{
		#region Properties
		private static ExcelPackage ExcelPackage { get; set; }
		#endregion

		#region Setup
		[ClassInitialize()]
		public static void MyClassInitialize(TestContext testContext)
		{
			if (!Directory.Exists("Test"))
				Directory.CreateDirectory(string.Format("Test"));
			ExcelPackage = new ExcelPackage(new FileInfo(@"Test\ConditionalFormatting.xlsx"));
		}

		[ClassCleanup()]
		public static void MyClassCleanup()
		{
			ExcelPackage = null;
		}
		#endregion

		#region Tests Methods
		[TestMethod]
		public void TwoColorScale()
		{
			var ws = ExcelPackage.Workbook.Worksheets.Add("ColorScale");
			ws.ConditionalFormatting.AddTwoColorScale(ws.Cells["A1:A5"]);
			ws.SetValue(1, 1, 1);
			ws.SetValue(2, 1, 2);
			ws.SetValue(3, 1, 3);
			ws.SetValue(4, 1, 4);
			ws.SetValue(5, 1, 5);
		}

		[TestMethod]
		public void TwoBackColor()
		{
			var ws = ExcelPackage.Workbook.Worksheets.Add("TwoBackColor");
			IExcelConditionalFormattingEqual condition1 = ws.ConditionalFormatting.AddEqual(ws.Cells["A1"]);
			condition1.StopIfTrue = true;
			condition1.Priority = 1;
			condition1.Formula = "TRUE";
			condition1.Style.Fill.BackgroundColor.Color = Color.Green;
			IExcelConditionalFormattingEqual condition2 = ws.ConditionalFormatting.AddEqual(ws.Cells["A2"]);
			condition2.StopIfTrue = true;
			condition2.Priority = 2;
			condition2.Formula = "FALSE";
			condition2.Style.Fill.BackgroundColor.Color = Color.Red;
		}

		[TestMethod]
		public void Databar()
		{
			var ws = ExcelPackage.Workbook.Worksheets.Add("Databar");
			var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
			ws.SetValue(1, 1, 1);
			ws.SetValue(2, 1, 2);
			ws.SetValue(3, 1, 3);
			ws.SetValue(4, 1, 4);
			ws.SetValue(5, 1, 5);
		}

		[TestMethod]
		public void DatabarChangingAddressAddsConditionalFormatNodeInSchemaOrder()
		{
			var ws = ExcelPackage.Workbook.Worksheets.Add("DatabarAddressing");
			// Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
			ws.HeaderFooter.AlignWithMargins = true;
			var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
			Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
			Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
			cf.Address = new ExcelAddress("C3");
			Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
			Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
		}

		[TestMethod]
		public void IconSet()
		{
			var ws = ExcelPackage.Workbook.Worksheets.Add("IconSet");
			var cf = ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A1:A3"], eExcelconditionalFormatting3IconsSetType.Symbols);
			ws.SetValue(1, 1, 1);
			ws.SetValue(2, 1, 2);
			ws.SetValue(3, 1, 3);

			var cf4 = ws.ConditionalFormatting.AddFourIconSet(ws.Cells["B1:B4"], eExcelconditionalFormatting4IconsSetType.Rating);
			cf4.Icon1.Type = eExcelConditionalFormattingValueObjectType.Formula;
			cf4.Icon1.Formula = "0";
			cf4.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula;
			cf4.Icon2.Formula = "1/3";
			cf4.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula;
			cf4.Icon3.Formula = "2/3";
			ws.SetValue(1, 2, 1);
			ws.SetValue(2, 2, 2);
			ws.SetValue(3, 2, 3);
			ws.SetValue(4, 2, 4);

			var cf5 = ws.ConditionalFormatting.AddFiveIconSet(ws.Cells["C1:C5"], eExcelconditionalFormatting5IconsSetType.Quarters);
			cf5.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
			cf5.Icon1.Value = 1;
			cf5.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
			cf5.Icon2.Value = 2;
			cf5.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
			cf5.Icon3.Value = 3;
			cf5.Icon4.Type = eExcelConditionalFormattingValueObjectType.Num;
			cf5.Icon4.Value = 4;
			cf5.Icon5.Type = eExcelConditionalFormattingValueObjectType.Num;
			cf5.Icon5.Value = 5;
			cf5.ShowValue = false;
			cf5.Reverse = true;

			ws.SetValue(1, 3, 1);
			ws.SetValue(2, 3, 2);
			ws.SetValue(3, 3, 3);
			ws.SetValue(4, 3, 4);
			ws.SetValue(5, 3, 5);
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\MultiColorConditionalFormatting.xlsx")]
		public void TwoAndThreeColorConditionalFormattingFromFileDoesNotGetOverwrittenWithDefaultValues()
		{
			var file = new FileInfo(@"MultiColorConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(2, sheet.ConditionalFormatting.Count);
				var twoColor = (ExcelConditionalFormattingTwoColorScale)sheet.ConditionalFormatting.First(cf => cf is ExcelConditionalFormattingTwoColorScale);
				var threeColor = (ExcelConditionalFormattingThreeColorScale)sheet.ConditionalFormatting.First(cf => cf is ExcelConditionalFormattingThreeColorScale);

				var defaultTwoColorScale = new ExcelConditionalFormattingTwoColorScale(new ExcelAddress("A1"), 2, sheet);
				var defaultThreeColorScale = new ExcelConditionalFormattingThreeColorScale(new ExcelAddress("A1"), 2, sheet);

				Assert.IsNull(twoColor.HighValue);
				Assert.IsNull(twoColor.LowValue);
				Assert.IsNotNull(defaultTwoColorScale.HighValue);
				Assert.IsNotNull(defaultTwoColorScale.LowValue);
				Assert.IsNull(threeColor.HighValue);
				Assert.IsNull(threeColor.LowValue);
				Assert.IsNotNull(defaultThreeColorScale.HighValue);
				Assert.IsNotNull(defaultThreeColorScale.LowValue);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\NAV024 - Item Sales and Profit.xlsx")]
		public void ConditionalFormattingsAndRulesWithoutSqrefsAreIgnored()
		{
			var file = new FileInfo(@"NAV024 - Item Sales and Profit.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(3, sheet.ConditionalFormatting.Count);
				Assert.IsTrue(sheet.ConditionalFormatting.All(cf => !string.IsNullOrEmpty(cf.Address.ToString())));
				Assert.AreEqual(2, sheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsTrue(sheet.X14ConditionalFormatting.X14Rules.All(cfr => !string.IsNullOrEmpty(cfr.Address.ToString())));
			}
		}

		[TestMethod]
		public void SetAddressUpdatesParentAddress()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("sheet1");
				worksheet.ConditionalFormatting.AddGreaterThan(worksheet.Cells["A1:A5"]);
				var format = worksheet.ConditionalFormatting.ElementAt(0);
				Assert.AreEqual("A1:A5", format.Address.Address);
				format.Address = new ExcelAddress("C2");
				Assert.AreEqual("C2", format.Address.Address);
			}
		}

		[TestMethod]
		public void SetAddressDoesNotGroupRules()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("sheet1");
				worksheet.ConditionalFormatting.AddGreaterThan(worksheet.Cells["A1:A5"]);
				worksheet.ConditionalFormatting.AddLessThan(worksheet.Cells["A1:A5"]);
				worksheet.ConditionalFormatting.AddBeginsWith(worksheet.Cells["A2"]);
				var format1 = worksheet.ConditionalFormatting.ElementAt(0);
				var format2 = worksheet.ConditionalFormatting.ElementAt(1);
				var format3 = worksheet.ConditionalFormatting.ElementAt(2);
				Assert.AreEqual(format1.Node.ParentNode, format2.Node.ParentNode);
				Assert.AreNotEqual(format2.Node.ParentNode, format3.Node.ParentNode);
				Assert.AreEqual(format1.Address.Address, format2.Address.Address);
				Assert.AreNotEqual(format2.Address, format3.Address);
				Assert.AreEqual(2, format1.Node.ParentNode.ChildNodes.Count);
				format2.Address = new ExcelAddress("A2");
				Assert.AreNotEqual(format1.Node.ParentNode, format2.Node.ParentNode);
				Assert.AreNotEqual(format2.Node.ParentNode, format3.Node.ParentNode);
				Assert.AreNotEqual(format1.Address.Address, format2.Address.Address);
				Assert.AreEqual(format2.Address.Address, format3.Address.Address);
				Assert.AreEqual("A2", format2.Address.Address);
				Assert.AreEqual(1, format1.Node.ParentNode.ChildNodes.Count);
				Assert.AreEqual(1, format2.Node.ParentNode.ChildNodes.Count);
				Assert.AreEqual(1, format3.Node.ParentNode.ChildNodes.Count);
			}
		}
		#endregion

		#region TransformFormulaReferences Test
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void TransformFormulaReferencesValAttributeTest()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				var twoColorScale = (ExcelConditionalFormattingTwoColorScale)formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.TwoColorScale);
				Assert.AreEqual("colorScale", twoColorScale.Node.Attributes["type"].Value);
				var innerNodes = twoColorScale.Node.ChildNodes[0].ChildNodes;
				var minNode = innerNodes[0];
				var maxNode = innerNodes[1];
				Assert.AreEqual("$K$6", minNode.Attributes["val"].Value);
				Assert.AreEqual("$K$7", maxNode.Attributes["val"].Value);

				formattings.TransformFormulaReferences(s => { try { return new ExcelAddress(s).AddColumn(1, 2).Address; } catch { return s; } });

				formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				twoColorScale = (ExcelConditionalFormattingTwoColorScale)formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.TwoColorScale);
				Assert.AreEqual("colorScale", twoColorScale.Node.Attributes["type"].Value);
				innerNodes = twoColorScale.Node.ChildNodes[0].ChildNodes;
				minNode = innerNodes[0];
				maxNode = innerNodes[1];
				Assert.AreEqual("$M$6", minNode.Attributes["val"].Value);
				Assert.AreEqual("$M$7", maxNode.Attributes["val"].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void TransformFormulaReferencesFiveIconListTest()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				var iconSet = (ExcelConditionalFormattingFiveIconSet)formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.FiveIconSet);
				Assert.AreEqual("iconSet", iconSet.Node.Attributes["type"].Value);
				Assert.AreEqual("0", iconSet.Icon1.Formula);
				Assert.AreEqual("$Q$7", iconSet.Icon2.Formula);
				Assert.AreEqual("$Q$8", iconSet.Icon3.Formula);
				Assert.AreEqual("$Q$9", iconSet.Icon4.Formula);
				Assert.AreEqual("$Q$10", iconSet.Icon5.Formula);

				formattings.TransformFormulaReferences(s => { try { return new ExcelAddress(s).AddColumn(1, 2).AddRow(1, 3).Address; } catch { return s; } });

				formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				iconSet = (ExcelConditionalFormattingFiveIconSet)formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.FiveIconSet);
				Assert.AreEqual("iconSet", iconSet.Node.Attributes["type"].Value);
				Assert.AreEqual("0", iconSet.Icon1.Formula);
				Assert.AreEqual("$S$10", iconSet.Icon2.Formula);
				Assert.AreEqual("$S$11", iconSet.Icon3.Formula);
				Assert.AreEqual("$S$12", iconSet.Icon4.Formula);
				Assert.AreEqual("$S$13", iconSet.Icon5.Formula);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void TransformFormulaReferencesFormulaTagTest()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				var between = (ExcelConditionalFormattingBetween)formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("cellIs", between.Node.Attributes["type"].Value);
				Assert.AreEqual(between.Type, eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("$Q$22", between.Formula);
				Assert.AreEqual("$Q$23", between.Formula2);
				
				formattings.TransformFormulaReferences(s => $"IF({s},1,0)");

				formattings = package.Workbook.Worksheets.First().ConditionalFormatting;
				between = (ExcelConditionalFormattingBetween)formattings.First(x => x.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("cellIs", between.Node.Attributes["type"].Value);
				Assert.AreEqual(between.Type, eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("IF($Q$22,1,0)", between.Formula);
				Assert.AreEqual("IF($Q$23,1,0)", between.Formula2);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void TransformFormulaReferencesX14Test()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var formattings = package.Workbook.Worksheets.First().X14ConditionalFormatting;
				Assert.AreEqual(2, formattings.X14Rules.Count);
				var dataBarRule = formattings.X14Rules.First();
				Assert.AreEqual(1, dataBarRule.Formulae.Count);
				Assert.AreEqual("$O$6", dataBarRule.Formulae.First().Formula);
				var containsTextRule = formattings.X14Rules[1];
				Assert.AreEqual(2, containsTextRule.Formulae.Count);
				Assert.AreEqual("NOT(ISERROR(SEARCH($S$22,E22)))", containsTextRule.Formulae.First().Formula);
				Assert.AreEqual("$S$22", containsTextRule.Formulae[1].Formula);

				formattings.TransformFormulaReferences(s => $"IF({s},1,0)");

				formattings = package.Workbook.Worksheets.First().X14ConditionalFormatting;
				dataBarRule = formattings.X14Rules.First();
				Assert.AreEqual(1, dataBarRule.Formulae.Count);
				Assert.AreEqual("IF($O$6,1,0)", dataBarRule.Formulae.First().Formula);
				Assert.AreEqual("IF(NOT(ISERROR(SEARCH($S$22,E22))),1,0)", containsTextRule.Formulae.First().Formula);
				Assert.AreEqual("IF($S$22,1,0)", containsTextRule.Formulae[1].Formula);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\MultipleConditionalFormattingAtSameAddress.xlsx")]
		public void ExcelConditionalFormattingRuleTransformTwoFormulasWithSameAddressInSameXmlConditionalFormattingNode()
		{
			FileInfo tempFile = new FileInfo(Path.GetTempFileName());
			ExcelAddress originalAddress = new ExcelAddress("C4");
			ExcelAddress newAddress = new ExcelAddress("D8");
			FileInfo file = new FileInfo(@"MultipleConditionalFormattingAtSameAddress.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets["Sheet1"];
				var conditionalFormattingNodes = sheet.ConditionalFormatting.TopNode.SelectNodes(".//d:conditionalFormatting", sheet.NameSpaceManager);
				Assert.AreEqual(1, conditionalFormattingNodes.Count);
				Assert.AreEqual(2, conditionalFormattingNodes.Item(0).ChildNodes.Count);
				var format1 = sheet.ConditionalFormatting.First();
				var format2 = sheet.ConditionalFormatting[1];
				Assert.IsTrue(format1 is ExcelConditionalFormattingEqual);
				Assert.IsTrue(format2 is ExcelConditionalFormattingEqual);
				Assert.IsTrue(format1 is ExcelConditionalFormattingRule);
				Assert.IsTrue(format2 is ExcelConditionalFormattingRule);
				Assert.AreEqual(originalAddress.ToString(), format1.Address.ToString());
				Assert.AreEqual(originalAddress.ToString(), format2.Address.ToString());

				format1.Address = newAddress;
				conditionalFormattingNodes = sheet.ConditionalFormatting.TopNode.SelectNodes(".//d:conditionalFormatting", sheet.NameSpaceManager);
				Assert.AreEqual(2, conditionalFormattingNodes.Count);
				Assert.AreEqual(1, conditionalFormattingNodes.Item(0).ChildNodes.Count);
				Assert.AreEqual(1, conditionalFormattingNodes.Item(1).ChildNodes.Count);
				Assert.IsTrue(format1 is ExcelConditionalFormattingEqual);
				Assert.IsTrue(format2 is ExcelConditionalFormattingEqual);
				Assert.IsTrue(format1 is ExcelConditionalFormattingRule);
				Assert.IsTrue(format2 is ExcelConditionalFormattingRule);
				Assert.AreEqual(newAddress.ToString(), format1.Address.ToString());
				Assert.AreEqual(originalAddress.ToString(), format2.Address.ToString());

				format2.Address = newAddress;
				Assert.AreEqual(2, sheet.ConditionalFormatting.Count);
				Assert.IsTrue(format1 is ExcelConditionalFormattingEqual);
				Assert.IsTrue(format2 is ExcelConditionalFormattingEqual);
				Assert.IsTrue(format1 is ExcelConditionalFormattingRule);
				Assert.IsTrue(format2 is ExcelConditionalFormattingRule);
				Assert.AreEqual(newAddress.ToString(), format1.Address.ToString());
				Assert.AreEqual(newAddress.ToString(), format2.Address.ToString());
			}
			File.Delete(tempFile.ToString());
		}
		#endregion
	}
}