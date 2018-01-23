/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-01-01
 * Mats Alm                         Applying patch submitted    2011-11-14
 *                                  by Ted Heatherington
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 * Raziq York		                Added support for Any type  2014-08-08
*******************************************************************************/
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{


	/// <summary>
	/// <para>
	/// Collection of <see cref="ExcelDataValidation"/>. This class is providing the API for EPPlus data validation.
	/// </para>
	/// <para>
	/// The public methods of this class (Add[...]Validation) will create a datavalidation entry in the worksheet. When this
	/// validation has been created changes to the properties will affect the workbook immediately.
	/// </para>
	/// <para>
	/// Each type of validation has either a formula or a typed value/values, except for custom validation which has a formula only.
	/// </para>
	/// <code>
	/// // Add a date time validation
	/// var validation = worksheet.DataValidation.AddDateTimeValidation("A1");
	/// // set validation properties
	/// validation.ShowErrorMessage = true;
	/// validation.ErrorTitle = "An invalid date was entered";
	/// validation.Error = "The date must be between 2011-01-31 and 2011-12-31";
	/// validation.Prompt = "Enter date here";
	/// validation.Formula.Value = DateTime.Parse("2011-01-01");
	/// validation.Formula2.Value = DateTime.Parse("2011-12-31");
	/// validation.Operator = ExcelDataValidationOperator.between;
	/// </code>
	/// </summary>
	public class ExcelDataValidationCollection : ExcelDataValidationCollectionBase
	{
		#region Properties
		protected override string DataValidationPath => "//d:dataValidations";
		protected override string DataValidationItemsPath => string.Format("{0}/d:dataValidation", DataValidationPath);
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an <see cref="ExcelDataValidationCollection"/>. Loads any existing data validations from the <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="worksheet">The worksheet of the <see cref="ExcelDataValidationCollection"/>.</param>
		internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
			 : base(worksheet)
		{
			// check existing nodes and load them
			var dataValidationNodes = worksheet.WorksheetXml.SelectNodes(DataValidationItemsPath, worksheet.NameSpaceManager);
			if (dataValidationNodes != null && dataValidationNodes.Count > 0)
			{
				foreach (XmlNode node in dataValidationNodes)
				{
					if (node.Attributes["sqref"] == null) continue;

					var addr = node.Attributes["sqref"].Value;

					var typeSchema = node.Attributes["type"] != null ? node.Attributes["type"].Value : "";

					var type = ExcelDataValidationType.GetBySchemaName(typeSchema);
					_validations.Add(ExcelDataValidationFactory.Create(type, worksheet, addr, node));
				}
			}
			if (_validations.Count > 0)
				base.OnValidationCountChanged();
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a <see cref="ExcelDataValidationAny"/> to the worksheet.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationAny"/> that was added.</returns>
		public IExcelDataValidationAny AddAnyValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationAny(_worksheet, address, ExcelDataValidationType.Any);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Adds an <see cref="IExcelDataValidationInt"/> to the worksheet. Whole means that the only accepted values
		/// are integer values.
		/// </summary>
		/// <param name="address">the range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationInt"/> that was added.</returns>
		public IExcelDataValidationInt AddIntegerValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationInt(_worksheet, address, ExcelDataValidationType.Whole);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Addes an <see cref="IExcelDataValidationDecimal"/> to the worksheet. The only accepted values are
		/// decimal values.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationDecimal"/> that was added.</returns>
		public IExcelDataValidationDecimal AddDecimalValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationDecimal(_worksheet, address, ExcelDataValidationType.Decimal);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Adds an <see cref="IExcelDataValidationList"/> to the worksheet. The accepted values are defined
		/// in a list.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationList"/> that was added.</returns>
		public IExcelDataValidationList AddListValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationList(_worksheet, address, ExcelDataValidationType.List);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Adds an <see cref="IExcelDataValidationInt"/> regarding text length to the worksheet.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationInt"/> that was added.</returns>
		public IExcelDataValidationInt AddTextLengthValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationInt(_worksheet, address, ExcelDataValidationType.TextLength);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationDateTime"/> that was added.</returns>
		public IExcelDataValidationDateTime AddDateTimeValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationDateTime(_worksheet, address, ExcelDataValidationType.DateTime);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Adds an <see cref="IExcelDataValidationTime"/> to the worksheet.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationTime"/> that was added.</returns>
		public IExcelDataValidationTime AddTimeValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationTime(_worksheet, address, ExcelDataValidationType.Time);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}

		/// <summary>
		/// Adds a <see cref="IExcelDataValidationCustom"/> to the worksheet.
		/// </summary>
		/// <param name="address">The range/address to validate</param>
		/// <returns>The <see cref="IExcelDataValidationCustom"/> that was added.</returns>
		public IExcelDataValidationCustom AddCustomValidation(string address)
		{
			this.ValidateAddress(address);
			this.EnsureRootElementExists();
			var item = new ExcelDataValidationCustom(_worksheet, address, ExcelDataValidationType.Custom);
			this._validations.Add(item);
			this.OnValidationCountChanged();
			return item;
		}
		#endregion

		#region ExcelDataValidationCollectionBase Overrides
		protected override void ClearWorksheetValidations()
		{
			_worksheet.ClearValidations();
		}
		#endregion
	}
}
