using System.Linq;
using System.Xml;
using OfficeOpenXml.DataValidation.X14DataValidation;

namespace OfficeOpenXml.DataValidation
{
	/// <summary>
	/// Represenst a collection of X14 Data Validations on a worksheet.
	/// </summary>
	public class ExcelX14DataValidationCollection : ExcelDataValidationCollectionBase
	{
		#region Properties
		protected override string DataValidationPath => ExcelX14DataValidation.DataValidationPath;
		protected override string DataValidationItemsPath => ExcelX14DataValidation.DataValidationItemsPath;
		#endregion

		#region Constructors
		/// <summary>
		/// Construcs an <see cref="ExcelX14DataValidationCollection"/> and loads data validations from exising <paramref name="worksheet"/> xml.
		/// </summary>
		/// <param name="worksheet"></param>
		internal ExcelX14DataValidationCollection(ExcelWorksheet worksheet)
			: base(worksheet)
		{
			// check existing nodes and load them
			var dataValidationNodes = worksheet.WorksheetXml.SelectNodes(DataValidationItemsPath, worksheet.NameSpaceManager);
			if (dataValidationNodes != null && dataValidationNodes.Count > 0)
			{
				foreach (XmlNode node in dataValidationNodes)
				{
					var address = base.GetXmlNodeString(node, ExcelX14DataValidation.SqrefLocalPath);
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

		#region ExcelDataValidationCollectionBase Overrides
		protected override void ClearWorksheetValidations()
		{
			_worksheet.ClearX14Validations();
		}
		#endregion
	}
}
