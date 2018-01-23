using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.X14DataValidation;

namespace EPPlusTest.DataValidation
{
	[TestClass]
	public class DataValidationTests : ValidationTestBase
	{
		#region Test Setup/Teardown
		[TestInitialize]
		public void Setup()
		{
			SetupTestData();
		}

		[TestCleanup]
		public void Cleanup()
		{
			CleanupTestData();
		}
		#endregion

		#region Data Validations
		[TestMethod]
		public void DataValidations_ShouldSetOperatorFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "greaterThanOrEqual", "1");
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.AreEqual(ExcelDataValidationOperator.greaterThanOrEqual, validation.Operator);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
		{
			var validations = _sheet.DataValidations.AddIntegerValidation("A1");
			validations.Operator = ExcelDataValidationOperator.equal;
			validations.Validate();
		}

		[TestMethod]
		public void DataValidations_ShouldSetShowErrorMessageFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "1", true, false);
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.IsTrue(validation.ShowErrorMessage ?? false);
		}

		[TestMethod]
		public void DataValidations_ShouldSetShowInputMessageFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "1", false, true);
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.IsTrue(validation.ShowInputMessage ?? false);
		}

		[TestMethod]
		public void DataValidations_ShouldSetPromptFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.AreEqual("Prompt", validation.Prompt);
		}

		[TestMethod]
		public void DataValidations_ShouldSetPromptTitleFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.AreEqual("PromptTitle", validation.PromptTitle);
		}

		[TestMethod]
		public void DataValidations_ShouldSetErrorFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.AreEqual("Error", validation.Error);
		}

		[TestMethod]
		public void DataValidations_ShouldSetErrorTitleFromExistingXml()
		{
			// Arrange
			LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
			// Act
			var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
			// Assert
			Assert.AreEqual("ErrorTitle", validation.ErrorTitle);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
		{
			var validation = _sheet.DataValidations.AddIntegerValidation("A1");
			validation.Formula.Value = 1;
			validation.Operator = ExcelDataValidationOperator.between;
			validation.Validate();
		}

		[TestMethod]
		public void DataValidations_ShouldAcceptOneItemOnly()
		{
			var validation = _sheet.DataValidations.AddListValidation("A1");
			validation.Formula.Values.Add("1");
			validation.Validate();
		}

		[TestMethod]
		public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsOneColumn()
		{
			// Act
			var validation = _sheet.DataValidations.AddIntegerValidation("A:A");

			// Assert
			Assert.AreEqual("A1:A" + ExcelPackage.MaxRows.ToString(), validation.Address.Address);
		}

		[TestMethod]
		public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsDifferentColumns()
		{
			// Act
			var validation = _sheet.DataValidations.AddIntegerValidation("A:B");

			// Assert
			Assert.AreEqual(string.Format("A1:B{0}", ExcelPackage.MaxRows), validation.Address.Address);
		}
		#endregion

		#region X14 Data Validations
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\X14DataValidations.xlsx")]
		public void ReadX14DataValidations()
		{
			var file = new FileInfo("X14DataValidations.xlsx");
			Assert.IsTrue(file.Exists);
			using (var excelPackage = new ExcelPackage(file))
			{
				var worksheet = excelPackage.Workbook.Worksheets["Data Validation"];
				Assert.AreEqual(2, worksheet.X14DataValidations.Count);
				Assert.AreEqual("2", worksheet.X14DataValidations.TopNode.Attributes["count"].Value);
				ExcelX14DataValidation firstListValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.List);
				ExcelX14DataValidation firstWholeValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.Whole);
				Assert.IsNotNull(firstListValidation);
				Assert.IsNotNull(firstWholeValidation);
				Assert.AreEqual("D7,J7", firstListValidation.Address.Address);
				Assert.AreEqual("F7", firstWholeValidation.Address.Address);
				Assert.AreEqual("'Source data'!$E$8:$E$9", firstListValidation.Formula);
				Assert.AreEqual("'Source data'!$E$8", firstWholeValidation.Formula);
				Assert.AreEqual("'Source data'!$E$9", firstWholeValidation.Formula2);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\X14DataValidations.xlsx")]
		public void TranslateX14DataValidationFormulas()
		{
			var file = new FileInfo("X14DataValidations.xlsx");
			Assert.IsTrue(file.Exists);
			var tempFile = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var excelPackage = new ExcelPackage(file))
				{
					var sourceWorksheet = excelPackage.Workbook.Worksheets["Source data"];
					sourceWorksheet.InsertColumn(5, 3);
					sourceWorksheet.InsertRow(7, 1);

					var worksheet = excelPackage.Workbook.Worksheets["Data Validation"];
					ExcelX14DataValidation firstListValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.List);
					ExcelX14DataValidation firstWholeValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.Whole);
					Assert.AreEqual("D7,J7", firstListValidation.Address.Address);
					Assert.AreEqual("F7", firstWholeValidation.Address.Address);
					Assert.AreEqual("'Source data'!$H$9:$H$10", firstListValidation.Formula);
					Assert.AreEqual("'Source data'!$H$9", firstWholeValidation.Formula);
					Assert.AreEqual("'Source data'!$H$10", firstWholeValidation.Formula2);
					excelPackage.SaveAs(tempFile);
				}

				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var worksheet = excelPackage.Workbook.Worksheets["Data Validation"];
					ExcelX14DataValidation firstListValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.List);
					ExcelX14DataValidation firstWholeValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.Whole);
					Assert.AreEqual("D7,J7", firstListValidation.Address.Address);
					Assert.AreEqual("F7", firstWholeValidation.Address.Address);
					Assert.AreEqual("'Source data'!$H$9:$H$10", firstListValidation.Formula);
					Assert.AreEqual("'Source data'!$H$9", firstWholeValidation.Formula);
					Assert.AreEqual("'Source data'!$H$10", firstWholeValidation.Formula2);
					var fNode = firstListValidation.TopNode.SelectSingleNode(".//x14:formula1", worksheet.NameSpaceManager).SelectSingleNode(".//xm:f", worksheet.NameSpaceManager);
					Assert.AreEqual("'Source data'!$H$9:$H$10", fNode.InnerText);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\X14DataValidations.xlsx")]
		public void TranslateX14DataValidationAddresses()
		{
			var file = new FileInfo("X14DataValidations.xlsx");
			Assert.IsTrue(file.Exists);
			var tempFile = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var excelPackage = new ExcelPackage(file))
				{
					var worksheet = excelPackage.Workbook.Worksheets["Data Validation"];
					worksheet.InsertColumn(5, 3);
					worksheet.InsertRow(7, 1);
					ExcelX14DataValidation firstListValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.List);
					ExcelX14DataValidation firstWholeValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.Whole);
					Assert.AreEqual("D8,M8", firstListValidation.Address.Address);
					Assert.AreEqual("I8", firstWholeValidation.Address.Address);
					Assert.AreEqual("'Source data'!$E$8:$E$9", firstListValidation.Formula);
					Assert.AreEqual("'Source data'!$E$8", firstWholeValidation.Formula);
					Assert.AreEqual("'Source data'!$E$9", firstWholeValidation.Formula2);
					excelPackage.SaveAs(tempFile);
				}

				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var worksheet = excelPackage.Workbook.Worksheets["Data Validation"];
					ExcelX14DataValidation firstListValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.List);
					ExcelX14DataValidation firstWholeValidation = (ExcelX14DataValidation)worksheet.X14DataValidations.First(d => d.ValidationType.Type == eDataValidationType.Whole);
					Assert.AreEqual("D8,M8", firstListValidation.Address.Address);
					Assert.AreEqual("I8", firstWholeValidation.Address.Address);
					Assert.AreEqual("'Source data'!$E$8:$E$9", firstListValidation.Formula);
					Assert.AreEqual("'Source data'!$E$8", firstWholeValidation.Formula);
					Assert.AreEqual("'Source data'!$E$9", firstWholeValidation.Formula2);
					var fNode = firstListValidation.TopNode.SelectSingleNode(".//x14:formula1", worksheet.NameSpaceManager).SelectSingleNode(".//xm:f", worksheet.NameSpaceManager);
					Assert.AreEqual("'Source data'!$E$8:$E$9", fNode.InnerText);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}
		#endregion
	}
}
