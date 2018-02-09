using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{
	/// <summary>
	/// Base class for original-style and X14 style data validations.
	/// </summary>
	public abstract class ExcelDataValidationCollectionBase : XmlHelper, IEnumerable<IExcelDataValidation>
	{
		#region Properties
		protected abstract string DataValidationPath { get; }
		protected abstract string DataValidationItemsPath { get; }
		#endregion

		#region Class Variables
		protected List<IExcelDataValidation> _validations = new List<IExcelDataValidation>();
		protected ExcelWorksheet _worksheet = null;
		#endregion

		#region Constructors
		protected ExcelDataValidationCollectionBase(ExcelWorksheet worksheet)
			 : base(worksheet?.NameSpaceManager, worksheet?.WorksheetXml?.DocumentElement)
		{
			if (worksheet == null)
				throw new ArgumentNullException(nameof(worksheet));
			this._worksheet = worksheet;
			this.SchemaNodeOrder = worksheet.SchemaNodeOrder;
		}
		#endregion

		#region Protected Helper Methods
		protected void EnsureRootElementExists()
		{
			var node = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
			if (node == null)
				base.CreateNode(DataValidationPath.TrimStart('/'));
		}

		protected XmlNode GetRootNode()
		{
			this.EnsureRootElementExists();
			this.TopNode = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
			return this.TopNode;
		}

		protected void OnValidationCountChanged()
		{
			var dvNode = this.GetRootNode();
			if (_validations.Any())
			{
				var attr = _worksheet.WorksheetXml.DocumentElement.SelectSingleNode(DataValidationPath + "[@count]", _worksheet.NameSpaceManager);
				if (attr == null)
					dvNode.Attributes.Append(_worksheet.WorksheetXml.CreateAttribute("count"));
				dvNode.Attributes["count"].Value = _validations.Count.ToString(CultureInfo.InvariantCulture);
			}
			else
			{
				if (dvNode != null)
					_worksheet.WorksheetXml.DocumentElement.RemoveChild(dvNode);
				this.ClearWorksheetValidations();
			}
		}

		protected void ValidateAddress(string address)
		{
			this.ValidateAddress(address, null);
		}

		protected abstract void ClearWorksheetValidations();
		#endregion

		#region Internal Methods
		/// <summary>
		/// Validates all data validations.
		/// </summary>
		internal void ValidateAll()
		{
			foreach (var validation in _validations)
			{
				validation.Validate();
				this.ValidateAddress(validation.Address.Address, validation);
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Number of validations
		/// </summary>
		public int Count
		{
			get { return _validations.Count; }
		}

		/// <summary>
		/// Removes all validations from the collection.
		/// </summary>
		public void Clear()
		{
			DeleteAllNode(DataValidationItemsPath.TrimStart('/'));
			_validations.Clear();
		}

		/// <summary>
		/// Returns all validations that matches the supplied predicate <paramref name="match"/>.
		/// </summary>
		/// <param name="match">Predicate to filter out matching validations.</param>
		/// <returns>List of <see cref="IExcelDataValidation"/> the match the <paramref name="match"/>.</returns>
		public IEnumerable<IExcelDataValidation> FindAll(Predicate<IExcelDataValidation> match)
		{
			return _validations.FindAll(match);
		}

		/// <summary>
		/// Returns the first matching validation.
		/// </summary>
		/// <param name="match">Predicate to filter out matching validations.</param>
		/// <returns>First <see cref="IExcelDataValidation"/> that matches the <paramref name="match"/>.</returns>
		public IExcelDataValidation Find(Predicate<IExcelDataValidation> match)
		{
			return _validations.Find(match);
		}

		/// <summary>
		/// Removes an <see cref="ExcelDataValidation"/> from the collection.
		/// </summary>
		/// <param name="item">The item to remove.</param>
		/// <returns>True if remove succeeds, otherwise false.</returns>
		/// <exception cref="ArgumentNullException">If <paramref name="item"/> is null</exception>
		public bool Remove(IExcelDataValidation item)
		{
			if (!(item is ExcelDataValidation))
				throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
			if (item == null)
				throw new ArgumentNullException(nameof(item));
			var dvNode = _worksheet.WorksheetXml.DocumentElement.SelectSingleNode(DataValidationPath.TrimStart('/'), NameSpaceManager);
			dvNode?.RemoveChild(((ExcelDataValidation)item).TopNode);
			var retVal = _validations.Remove(item);
			if (retVal)
				this.OnValidationCountChanged();
			return retVal;
		}

		/// <summary>
		/// Removes the validations that matches the predicate
		/// </summary>
		/// <param name="match">Predicate to filter out matching validations.</param>
		public void RemoveAll(Predicate<IExcelDataValidation> match)
		{
			var matches = _validations.FindAll(match);
			foreach (var m in matches)
			{
				if (!(m is ExcelDataValidation))
					throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
				TopNode.RemoveChild(((ExcelDataValidation)m).TopNode);
			}
			_validations.RemoveAll(match);
			this.OnValidationCountChanged();
		}
		#endregion

		#region Private Methods
		private void ValidateAddress(string address, IExcelDataValidation validatingValidation)
		{
			if (string.IsNullOrEmpty(address))
				throw new ArgumentNullException(nameof(address));
			// ensure that the new address does not collide with an existing validation.
			var newAddress = new ExcelAddress(address);
			if (_validations.Count > 0)
			{
				foreach (var validation in _validations)
				{
					if (validatingValidation != null && validatingValidation == validation)
						continue;
					var result = validation.Address.Collide(newAddress);
					if (result != ExcelAddress.eAddressCollition.No)
						throw new InvalidOperationException(string.Format("The address ({0}) collides with an existing validation ({1})", address, validation.Address.Address));
				}
			}
		}
		#endregion

		#region Indexing Overrides
		/// <summary>
		/// Index operator, returns by 0-based index.
		/// </summary>
		/// <param name="index">The 0-based index.</param>
		/// <returns>The <see cref="IExcelDataValidation"/> at the <paramref name="index"/>.</returns>
		public IExcelDataValidation this[int index]
		{
			get { return _validations[index]; }
			set { _validations[index] = value; }
		}

		/// <summary>
		/// Index operator, returns a data validation which address partly or exactly matches the searched address.
		/// </summary>
		/// <param name="address">A cell address or range</param>
		/// <returns>A <see cref="ExcelDataValidation"/> or null if no match</returns>
		public IExcelDataValidation this[string address]
		{
			get
			{
				var searchedAddress = new ExcelAddress(address);
				return _validations.Find(x => x.Address.Collide(searchedAddress) != ExcelAddress.eAddressCollition.No);
			}
		}
		#endregion

		#region IEnumerable Implementation
		IEnumerator<IExcelDataValidation> IEnumerable<IExcelDataValidation>.GetEnumerator()
		{
			return _validations.GetEnumerator();
		}

		IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return _validations.GetEnumerator();
		}
		#endregion
	}
}
