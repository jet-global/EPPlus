using System;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation.X14DataValidation
{
	/// <summary>
	/// Represents an X14 Data Validation
	/// </summary>
	public class ExcelX14DataValidation : XmlHelper, IExcelDataValidation //WithFormula2<IExcelDataValidationFormula>
	{
		#region Constants
		public const string DataValidationPath = "//x14:dataValidations";
		public const string DataValidationItemsPath = "//x14:dataValidations/x14:dataValidation";
		public const string SqrefLocalPath = ".//xm:sqref";
		public const string Formula1LocalPath = ".//x14:formula1/xm:f";
		public const string Formula2LocalPath = ".//x14:formula2/xm:f";
		#endregion

		#region Properties
		/// <summary>
		/// The <see cref="ExcelDataValidationType"/> for this Data Validation.
		/// </summary>
		public ExcelDataValidationType ValidationType { get; }
		/// <summary>
		/// True if the current validation type allows operator.
		/// </summary>
		public bool AllowsOperator => ValidationType.AllowOperator;

		/// <summary>
		/// The address of the <see cref="ExcelX14DataValidation"/>.
		/// </summary>
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

		/// <summary>
		/// The first formula for the <see cref="ExcelX14DataValidation"/>
		/// </summary>
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

		/// <summary>
		/// The second formula for the <see cref="ExcelX14DataValidation"/>
		/// </summary>
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

		// TODO: Implement Xml backing for these properties.
		public ExcelDataValidationWarningStyle ErrorStyle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool? AllowBlank { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool? ShowInputMessage { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool? ShowErrorMessage { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string ErrorTitle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string Error { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string PromptTitle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public string Prompt { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates ExcelX14DataValidation from existing xml node
		/// </summary>
		/// <param name="worksheet">The <see cref="ExcelWorksheet"/> of the <see cref="ExcelX14DataValidation"/>.</param>
		/// <param name="address">The address of the <see cref="ExcelX14DataValidation"/></param>
		/// <param name="validationType">They data validation type.</param>
		/// <param name="itemNode">The <see cref="XmlNode"/> of the <see cref="ExcelX14DataValidation"/>.</param>
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

		// TODO: Create constructor for creating a new (non-pre-existing) ExcelX14DataValidation.
		#endregion

		#region Public Methods
		/// <summary>
		/// This method will validate the state of the validation
		/// </summary>
		/// <exception cref="InvalidOperationException">If the state breaks the rules of the validation</exception>
		public void Validate()
		{
			if (string.IsNullOrEmpty(this.Address?.Address))
				throw new InvalidOperationException("Address cannot be empty");
			// TODO: More validation based on type.
		}
		#endregion
	}
}
