using System;
using EPPlusTest.FormulaParsing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class StructuredReferenceTest : DataProviderTestBase
	{
		#region Constructor Tests
		[TestMethod]
		public void StructuredReferenceWithAll()
		{
			var structuredReference = new StructuredReference("MyTable[#All]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.All, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithData()
		{
			var structuredReference = new StructuredReference("MyTable[#Data]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithThisRow()
		{
			var structuredReference = new StructuredReference("MyTable[#This row]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.ThisRow, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithHeaders()
		{
			var structuredReference = new StructuredReference("MyTable[#Headers]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Headers, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTotals()
		{
			var structuredReference = new StructuredReference("MyTable[#Totals]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Totals, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceNoArguments()
		{
			var structuredReference = new StructuredReference("MyTable[]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsTrue(string.IsNullOrEmpty(structuredReference.StartColumn));
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceAtArgument()
		{
			var structuredReference = new StructuredReference("MyTable[#this row]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.IsNull(structuredReference.EndColumn);
			Assert.AreEqual(ItemSpecifiers.ThisRow, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceNestedAtArgument()
		{
			var structuredReference = new StructuredReference("MyTable[[#this row]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.IsNull(structuredReference.StartColumn);
			Assert.IsNull(structuredReference.EndColumn);
			Assert.AreEqual(ItemSpecifiers.ThisRow, structuredReference.ItemSpecifiers);
		}
		
		[TestMethod]
		public void StructuredReferenceWithTableAndColumnAll()
		{
			var structuredReference = new StructuredReference("MyTable[[#all],[MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.All, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnNoBrackets()
		{
			var structuredReference = new StructuredReference("MyTable[[#all],MyColumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.All, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithTableAndColumnNoBracketsFirst()
		{
			var structuredReference = new StructuredReference("MyTable[MyColumn,[#data]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
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
		public void StructuredReferenceColumnNameWithHashtagEscaped()
		{
			var structuredReference = new StructuredReference("MyTable['#NotAnItemSpecifier]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("#NotAnItemSpecifier", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceColumnNameWithHashtagEscapedNested()
		{
			var structuredReference = new StructuredReference("MyTable[['#NotAnItemSpecifier]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("#NotAnItemSpecifier", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceEscapedLeftBracketInColumn()
		{
			var structuredReference = new StructuredReference("MyTable[My'[olumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("My[olumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceEscapedRightBracketInColumn()
		{
			var structuredReference = new StructuredReference("MyTable[My']olumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("My]olumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceEscapedQuotationInColumn()
		{
			var structuredReference = new StructuredReference("MyTable[My''olumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("My'olumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceEscapedThreeQuotationInColumn()
		{
			var structuredReference = new StructuredReference("MyTable[My'''[olumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("My'[olumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceEscapedFourQuotationInColumn()
		{
			var structuredReference = new StructuredReference("MyTable[My''''olumn]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("My''olumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceSpecialCharactersInColumn()
		{
			var structuredReference = new StructuredReference("MyTable[!|\\;`']  ,;:.ss'[]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("!|\\;`]  ,;:.ss[", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceSpecialCharactersInColumnNested()
		{
			var structuredReference = new StructuredReference("MyTable[[!|\\;`']  ,;:.ss'[]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("!|\\;`]  ,;:.ss[", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceSpecialCharactersInColumnWithWhitespace()
		{
			var structuredReference = new StructuredReference("MyTable[ \t [#data]  \t  , \t [!|\\;`']  ,;:.ss'[]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("!|\\;`]  ,;:.ss[", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceWithMultipleSpecifiersWithWhitespace()
		{
			var structuredReference = new StructuredReference("MyTable[\t[#Headers]   ,\t [#Data] ,\t     [MyColumn]]");
			Assert.AreEqual("MyTable", structuredReference.TableName);
			Assert.AreEqual("MyColumn", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Headers | ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		public void StructuredReferenceSpecialCharactersInTableName()
		{
			var structuredReference = new StructuredReference("My\\}+'Table[column]");
			Assert.AreEqual("My\\}+'Table", structuredReference.TableName);
			Assert.AreEqual("column", structuredReference.StartColumn);
			Assert.AreEqual(ItemSpecifiers.Data, structuredReference.ItemSpecifiers);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void StructuredReferenceNoBrackets()
		{
			var structuredReference = new StructuredReference("MyTable");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void StructuredReferenceNoTable()
		{
			var structuredReference = new StructuredReference("[#this row]");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void StructuredReferenceTrailingCharacters()
		{
			var structuredReference = new StructuredReference("MyTable[#this row]sdfsdf");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void StructuredReferenceInvalidArgument()
		{
			var structuredReference = new StructuredReference("MyTable[[#All],fsdfs[[]dfsd]]]]]");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void StructuredReferenceNoEndBracket()
		{
			var structuredReference = new StructuredReference("MyTable[#this row");
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

		#region Integration Tests
		[TestMethod]
		public void StructuredReferenceTests()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				this.BuildTableHeaders(worksheet);
				this.BuildTableData(worksheet, true);
				var table = worksheet.Tables.Add(new ExcelRange(worksheet, 3, 3, 9, 6), DataProviderTestBase.TableName);
				worksheet.Cells[6, 20].Formula = $"{DataProviderTestBase.TableName}[[#this row],[{DataProviderTestBase.Header2}]]";
				worksheet.Cells[6, 21].Formula = $"=SUM({DataProviderTestBase.TableName}[[#data],[{DataProviderTestBase.Header2}]])";
				worksheet.Cells[6, 22].Formula = $"=SUM({DataProviderTestBase.TableName}[[#all],[{DataProviderTestBase.Header2}]])";
				worksheet.Cells[6, 23].Formula = $"{DataProviderTestBase.TableName}[[#headers],[{DataProviderTestBase.Header2}]]";
				worksheet.Cells[6, 24].Formula = $"{DataProviderTestBase.TableName}[[#totals],[{DataProviderTestBase.Header3}]]";
				worksheet.Cells[6, 25].Formula = $"{DataProviderTestBase.TableName}[[#totals]]";
				worksheet.Cells[6, 26].Formula = $"SUM({DataProviderTestBase.TableName}[[#this row],[{DataProviderTestBase.Header2}]:[{DataProviderTestBase.Header4}]])";
				worksheet.Calculate();
				Assert.AreEqual(32, worksheet.Cells[6, 20].Value);
				Assert.AreEqual(12+22+32+42+52+62d, worksheet.Cells[6, 21].Value);
				Assert.AreEqual(12+22+32+42+52+62d, worksheet.Cells[6, 22].Value);
				Assert.AreEqual(DataProviderTestBase.Header2, worksheet.Cells[6, 23].Value);
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)worksheet.Cells[6, 24].Value).Type);
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)worksheet.Cells[6, 25].Value).Type);
				Assert.AreEqual(32+33+34d, worksheet.Cells[6, 26].Value);
			}
		}

		[TestMethod]
		public void StructuredReferenceTestsWithHeaderSpecialCharacters()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[3, 3].Value = DataProviderTestBase.Header1;
				worksheet.Cells[3, 4].Value = "#!'][%:$";
				worksheet.Cells[3, 5].Value = DataProviderTestBase.Header3;
				worksheet.Cells[3, 6].Value = DataProviderTestBase.Header4;
				this.BuildTableData(worksheet, true);
				var table = worksheet.Tables.Add(new ExcelRange(worksheet, 3, 3, 9, 6), DataProviderTestBase.TableName);
				worksheet.Cells[6, 20].Formula = $"SUM({DataProviderTestBase.TableName}[[#this row],['#!''']'[%:$]:[{DataProviderTestBase.Header3}]])";
				worksheet.Calculate();
				Assert.AreEqual(32+33d, worksheet.Cells[6, 20].Value);
			}
		}
		#endregion
	}
}
