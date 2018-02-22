using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing
{
	[TestClass]
	public class EpplusExcelDataProviderTest
	{
		#region Constants
		private const string TableName = "MyTable";
		private const string Header1 = "Header1";
		private const string Header2 = "Header2";
		private const string Header3 = "Header3";
		private const string Header4 = "Header4";
		#endregion

		#region Class Variables
		private ExcelPackage myPackage;
		private EpplusExcelDataProvider myDataProvider;
		#endregion

		#region Properties
		private ExcelPackage Package
		{
			get
			{
				if (myPackage == null)
				{
					myPackage = new ExcelPackage();
					var worksheet = myPackage.Workbook.Worksheets.Add("Sheet1");
					worksheet.Cells[3, 3].Value = EpplusExcelDataProviderTest.Header1;
					worksheet.Cells[3, 4].Value = EpplusExcelDataProviderTest.Header2;
					worksheet.Cells[3, 5].Value = EpplusExcelDataProviderTest.Header3;
					worksheet.Cells[3, 6].Value = EpplusExcelDataProviderTest.Header4;
					worksheet.Cells[4, 3].Value = "h1_r1";
					worksheet.Cells[4, 4].Value = "h2_r1";
					worksheet.Cells[4, 5].Value = "h3_r1";
					worksheet.Cells[4, 6].Value = "h4_r1";
					worksheet.Cells[5, 3].Value = "h1_r2";
					worksheet.Cells[5, 4].Value = "h2_r2";
					worksheet.Cells[5, 5].Value = "h3_r2";
					worksheet.Cells[5, 6].Value = "h4_r2";
					worksheet.Cells[6, 3].Value = "h1_r3";
					worksheet.Cells[6, 4].Value = "h2_r3";
					worksheet.Cells[6, 5].Value = "h3_r3";
					worksheet.Cells[6, 6].Value = "h4_r3";
					worksheet.Cells[7, 3].Value = "h1_r4";
					worksheet.Cells[7, 4].Value = "h2_r4";
					worksheet.Cells[7, 5].Value = "h3_r4";
					worksheet.Cells[7, 6].Value = "h4_r4";
					worksheet.Cells[8, 3].Value = "h1_r5";
					worksheet.Cells[8, 4].Value = "h2_r5";
					worksheet.Cells[8, 5].Value = "h3_r5";
					worksheet.Cells[8, 6].Value = "h4_r5";
					worksheet.Cells[9, 3].Value = "h1_r6";
					worksheet.Cells[9, 4].Value = "h2_r6";
					worksheet.Cells[9, 5].Value = "h3_r6";
					worksheet.Cells[9, 6].Value = "h4_r6";
					var table = worksheet.Tables.Add(new ExcelAddress("Sheet1", 3, 3, 9, 6), EpplusExcelDataProviderTest.TableName);
					// TODO :: add variations where total row is or is not shown, same with headers.
				}
				return myPackage;
			}
		}

		private EpplusExcelDataProvider DataProvider
		{
			get
			{
				if (myDataProvider == null)
					myDataProvider = new EpplusExcelDataProvider(this.Package);
				return myDataProvider;
			}
		}
		#endregion

		#region ResolveStructuredReference Tests
		#region #ThisRow Tests
		[TestMethod]
		public void ResolveStructuredReferenceThisRowRightOfTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 7, 7);
			Assert.AreEqual(7, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(7, result.Address.End.Row);
			Assert.AreEqual(3, result.Address.End.Column);
			Assert.AreEqual("h1_r4", result.GetOffset(0, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowRange()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]:[{EpplusExcelDataProviderTest.Header3}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 7, 7);
			Assert.AreEqual(7, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(7, result.Address.End.Row);
			Assert.AreEqual(5, result.Address.End.Column);
			Assert.AreEqual("h1_r4", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(0, 1));
			Assert.AreEqual("h3_r4", result.GetOffset(0, 2));
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowLeftOfTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 7, 1);
			Assert.AreEqual(7, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(7, result.Address.End.Row);
			Assert.AreEqual(3, result.Address.End.Column);
			Assert.AreEqual("h1_r4", result.GetOffset(0, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowAboveTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 2, 4);
			Assert.IsNull(result);
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowBelowTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 12, 4);
			Assert.IsNull(result);
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowHeaderRow()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 3, 7);
			Assert.IsNull(result);
		}
		#endregion

		#region #Data Tests
		[TestMethod]
		public void ResolveStructuredReferenceDataWithoutRowReturnsFullRange()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[{EpplusExcelDataProviderTest.Header4}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(6, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h4_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h4_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h4_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h4_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h4_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithoutRowReturnsFullMultiDimensionalRange()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[{EpplusExcelDataProviderTest.Header3}]:[{EpplusExcelDataProviderTest.Header4}]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(5, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h3_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h3_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h3_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h3_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h3_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h3_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h4_r1", result.GetOffset(0, 1));
			Assert.AreEqual("h4_r2", result.GetOffset(1, 1));
			Assert.AreEqual("h4_r3", result.GetOffset(2, 1));
			Assert.AreEqual("h4_r4", result.GetOffset(3, 1));
			Assert.AreEqual("h4_r5", result.GetOffset(4, 1));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 1));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeaders()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header1, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 2));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 3));
			Assert.AreEqual("h1_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(6, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersNoHeaders()
		{
			var table = this.Package.Workbook.GetTable(EpplusExcelDataProviderTest.TableName);
			this.Package.Workbook.Worksheets["Sheet1"].Tables.Delete(table);
			this.Package.Workbook.Worksheets["Sheet1"].Tables.Add(new ExcelAddress(4, 3, 9, 6), EpplusExcelDataProviderTest.TableName);
			try
			{
				var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers]]");
				var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
				Assert.AreEqual(4, result.Address.Start.Row);
				Assert.AreEqual(3, result.Address.Start.Column);
				Assert.AreEqual(9, result.Address.End.Row);
				Assert.AreEqual(6, result.Address.End.Column);
				Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
				Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
				Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
				Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
				Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
				Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
			}
			finally
			{
				this.Package.Workbook.GetTable(EpplusExcelDataProviderTest.TableName).ShowHeader = true;
			}
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersColumn()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers],[Header2]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual("h2_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(6, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersColumns()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers],[Header2]:[Header4]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual("h2_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(6, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(6, 2));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsNoTotals()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsColumnNoTotals()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals],[Header2]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual("h2_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsColumnsNoTotals()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals],[Header2]:[Header4]]");
			var result = this.DataProvider.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h2_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 2));
		}
		#endregion
		#endregion
	}
}
