using System;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation.X14DataValidation
{
	public class ExcelX14DataValidation : XmlHelper, IExcelDataValidation //WithFormula2<IExcelDataValidationFormula>
	{
		#region Constants
		public const string DataValidationPath = "//x14:dataValidations";
		public const string DataValidationItemsPath = "//x14:dataValidations/x14:dataValidation";
		public const string SqrefLocalPath = ".//xm:sqref";
		public const string Formula1LocalPath = ".//x14:formula1";
		public const string Formula2LocalPath = ".//x14:formula2";
		#endregion

		#region Properties
		public ExcelDataValidationType ValidationType { get; }
		public bool AllowsOperator => ValidationType.AllowOperator;

		public ExcelAddress Address
		{
			get
			{
				var sqref = base.GetXmlNodeString(ExcelX14DataValidation.SqrefLocalPath);
				sqref = sqref?.Replace(' ', ',');
				return new ExcelAddress(sqref);
			}
			set
			{
				var address = AddressUtility.ParseEntireColumnSelections(value.Address.Replace(',', ' '));
				base.SetXmlNodeString(ExcelX14DataValidation.SqrefLocalPath, address);
			}
		}
		public string Formula
		{
			get
			{
				return base.GetXmlNodeString(ExcelX14DataValidation.Formula1LocalPath);
			}
			set
			{
				if (base.ExistNode(ExcelX14DataValidation.Formula1LocalPath))
					base.SetXmlNodeString(ExcelX14DataValidation.Formula1LocalPath, value);
			}
		}
		public string Formula2
		{
			get
			{
				return base.GetXmlNodeString(ExcelX14DataValidation.Formula2LocalPath);
			}
			set
			{
				if (base.ExistNode(ExcelX14DataValidation.Formula2LocalPath))
					base.SetXmlNodeString(ExcelX14DataValidation.Formula2LocalPath, value);
			}
		}

		// TODO: back these with xml
		public ExcelDataValidationWarningStyle ErrorStyle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool? AllowBlank { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool? ShowInputMessage { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool? ShowErrorMessage { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string ErrorTitle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string Error { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string PromptTitle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string Prompt { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }


		#endregion

		// todo refactor constructor
		/// <summary>
		/// Creates a new ExcelX14DataValidation and adds the node to the xml
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="address"></param>
		/// <param name="validationType"></param>
		//internal ExcelX14DataValidation(ExcelWorksheet worksheet, ExcelAddress address, eDataValidationType validationType)
		//	: base(worksheet.NameSpaceManager)
		//{
		//	if (worksheet == null)
		//		throw new ArgumentNullException(nameof(worksheet));
		//	if (address == null)
		//		throw new ArgumentNullException(nameof(address));
		//	this.ValidationType = ExcelDataValidationType.GetByValidationType(validationType);
		//	this.Address = address;
		//	// TODO: create node and potentially parent node
		//	//var datavalidationsNode = worksheet.WorksheetXml.SelectSingleNode("//d:dataValidations", worksheet.NameSpaceManager);
		//	//var itemNode = this.TopNode.OwnerDocument.CreateElement("d:dataValidation");
		//	// set TopNode
		//}

		/// <summary>
		/// Creates ExcelX14DataValidation from existing xml node
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="address"></param>
		/// <param name="validationType"></param>
		/// <param name="itemNode"></param>
		internal ExcelX14DataValidation(ExcelWorksheet worksheet, string address, string validationType, XmlNode itemNode)
			: base(worksheet.NameSpaceManager)
		{
			if (string.IsNullOrEmpty(address))
				throw new ArgumentException(address, nameof(address));
			if (itemNode == null)
				throw new ArgumentNullException(nameof(itemNode));
			this.ValidationType = ExcelDataValidationType.GetBySchemaName(validationType);
			this.TopNode = itemNode;
		}

		public void Validate()
		{
			// implement based on type
			throw new NotImplementedException();
		}
	}
}
