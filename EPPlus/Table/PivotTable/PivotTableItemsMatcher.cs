using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Table.PivotTable.Filters;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Matches pivot table items with page field items and hidden items.
	/// </summary>
	public class PivotTableItemsMatcher
	{
		#region Class Variables
		private readonly ExcelPivotCacheDefinition myCacheDefinition;
		private readonly ExcelPivotTableFieldCollection myFields;
		private readonly ExcelPivotFieldFiltersCollection myFilters;
		#endregion

		#region Properties
		private Dictionary<int, List<int>> PageFieldsItemsToInclude { get; } = new Dictionary<int, List<int>>();
		private Dictionary<int, List<int>> HiddenFieldItems { get; } = new Dictionary<int, List<int>>();
		#endregion

		#region Constructors
		/// <summary>
		/// Constructs a <see cref="PivotTableItemsMatcher"/>.
		/// </summary>
		/// <param name="pivotFields">The <see cref="ExcelPivotTableFieldCollection"/>.</param>
		/// <param name="pageFields">The <see cref="ExcelPageFieldCollection"/>.</param>
		/// <param name="cacheDefinition">The <see cref="ExcelPivotCacheDefinition"/>.</param>
		/// <param name="filters">The <see cref="ExcelPivotFieldFiltersCollection"/>.</param>
		public PivotTableItemsMatcher(ExcelPivotTableFieldCollection pivotFields, ExcelPageFieldCollection pageFields, ExcelPivotCacheDefinition cacheDefinition, ExcelPivotFieldFiltersCollection filters)
		{
			myCacheDefinition = cacheDefinition;
			myFields = pivotFields;
			myFilters = filters;

			if (pageFields != null)
			{
				foreach (var pageField in pageFields)
				{
					if (pageField.Item != null)
					{
						if (!this.PageFieldsItemsToInclude.ContainsKey(pageField.Field))
							this.PageFieldsItemsToInclude.Add(pageField.Field, new List<int> { });
						var pivotFieldItem = pivotFields[pageField.Field].Items[pageField.Item.Value];
						this.PageFieldsItemsToInclude[pageField.Field].Add(pivotFieldItem.X);
					}
					else
					{
						var pageFieldItems = pivotFields[pageField.Field].Items;
						if (pageFieldItems != null)
						{
							foreach (var item in pageFieldItems.Where(i => !i.Hidden))
							{
								if (!this.PageFieldsItemsToInclude.ContainsKey(pageField.Field))
									this.PageFieldsItemsToInclude.Add(pageField.Field, new List<int> { });
								this.PageFieldsItemsToInclude[pageField.Field].Add(item.X);
							}
						}
					}
				}
			}

			foreach (var field in pivotFields.Where(f => f.HasItems))
			{
				var values = new List<int>();
				foreach (var fieldItem in field.Items)
				{
					if (fieldItem.Hidden && fieldItem.X >= 0)
						values.Add(fieldItem.X);
				}
				this.HiddenFieldItems.Add(field.Index, values);
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Determines whether the cache record should be included as data 
		/// or if it should be filtered out based on a page field or a hidden item.
		/// </summary>
		/// <param name="cacheRecord">The <see cref="CacheRecordNode"/>.</param>
		/// <returns>True if the record should be included, or false if the record should be filtered out.</returns>
		public bool ShouldInclude(CacheRecordNode cacheRecord)
		{
			return this.ShouldIncludeBasedOnPageFields(cacheRecord);
		}
		#endregion

		#region Private Methods
		private bool ShouldIncludeBasedOnPageFields(CacheRecordNode cacheRecord)
		{
			// If the report does not have page fields then do not filter items.
			if (!this.PageFieldsItemsToInclude.Any())
				return true;
			// Map of indices into a cacheRecord to cacheRecordItem values.
			foreach (var keyValue in this.PageFieldsItemsToInclude)
			{
				int cacheRecordValue = int.Parse(cacheRecord.Items[keyValue.Key].Value);
				if (!keyValue.Value.Any(v => v == cacheRecordValue))
					return false;
			}
			return true;
		}
		#endregion
	}
}