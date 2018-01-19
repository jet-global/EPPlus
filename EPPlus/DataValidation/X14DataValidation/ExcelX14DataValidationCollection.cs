using System.Linq;
using System.Xml;
using OfficeOpenXml.DataValidation.X14DataValidation;

namespace OfficeOpenXml.DataValidation
{
	public class ExcelX14DataValidationCollection : ExcelDataValidationCollectionBase
	{
		#region Constants
		protected override string DataValidationPath => ExcelX14DataValidation.DataValidationPath;
		protected override string DataValidationItemsPath => ExcelX14DataValidation.DataValidationItemsPath;
		#endregion

		#region Constructors
		internal ExcelX14DataValidationCollection(ExcelWorksheet worksheet)
			: base(worksheet)
		{
			// check existing nodes and load them
			var dataValidationNodes = worksheet.WorksheetXml.SelectNodes(DataValidationItemsPath, worksheet.NameSpaceManager);
			if (dataValidationNodes != null && dataValidationNodes.Count > 0)
			{
				foreach (XmlNode node in dataValidationNodes)
				{
					var address = base.GetXmlNodeString(node, "xm:sqref");
					if (!string.IsNullOrEmpty(address))
					{
						var type = node.Attributes["type"]?.Value ?? string.Empty;
						_validations.Add(new ExcelX14DataValidation(worksheet, address, type, node));
					}
				}
			}
			if (_validations.Any())
				base.OnValidationCountChanged();
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a <see cref="ExcelDataValidationAny"/> to the worksheet.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns></returns>
		/// // TODO: Add ability to add an x14 validation
		//public ExcelX14DataValidation AddValidation(string address, eDataValidationType type)
		//{
		//	this.ValidateAddress(address);
		//	this.EnsureRootElementExists();
		//	var item = new ExcelX14DataValidation(_worksheet, address, ExcelDataValidationType.GetByValidationType(type));
		//	this._validations.Add(item);
		//	this.OnValidationCountChanged();
		//	return item;
		//}
		#endregion
	}
}
