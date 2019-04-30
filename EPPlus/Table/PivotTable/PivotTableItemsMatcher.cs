using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeOpenXml.Extensions;
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

			foreach (var field in pivotFields.Where(f => f.HasItems && (f.IsRowField || f.IsColumnField)))
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
			return this.ShouldIncludeBasedOnPageFields(cacheRecord)
				&& this.ShouldIncludeBasedOnHiddenFields(cacheRecord);
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

		private bool ShouldIncludeBasedOnHiddenFields(CacheRecordNode cacheRecord)
		{
			foreach (var entry in this.HiddenFieldItems)
			{
				int fieldIndex = entry.Key;
				var hiddenFieldItemIndicies = entry.Value;
				// Ignore data field tuples, group pivot field tuples and custom field subtotal settings.
				if (fieldIndex == -2 || fieldIndex == 1048832)
					continue;

				var cacheField = myCacheDefinition.CacheFields[fieldIndex];
				if (cacheField.IsGroupField)
				{
					//bool groupMatch = this.FindGroupingRecordValueAndTupleMatch(cacheField, cacheRecord, tuple);
					//if (!groupMatch)
					//	return false;
				}
				else
				{
					//if (myFilters != null)
					//{
					//	foreach (var filter in myFilters)
					//	{
					//		int filterFieldIndex = filter.Field;
					//		int recordReferenceIndex = int.Parse(cacheRecord.Items[filterFieldIndex].Value);
					//		string recordReferenceString = myCacheDefinition.CacheFields[filterFieldIndex].SharedItems[recordReferenceIndex].Value;
					//		bool isNumeric = myCacheDefinition.CacheFields[filterFieldIndex].SharedItems[recordReferenceIndex].Type == PivotCacheRecordType.n;
					//		bool isMatch = filter.MatchesFilterCriteriaResult(recordReferenceString, isNumeric);
					//		if (!isMatch)
					//			return false;
					//	}
					//}

					var sharedItemsCollection = myCacheDefinition.CacheFields[fieldIndex].SharedItems;
					int cacheRecordSharedItemIndex = int.Parse(cacheRecord.Items[fieldIndex].Value);
					var cacheRecordValue = sharedItemsCollection[cacheRecordSharedItemIndex].Value;
					var hiddenSharedCacheItems = entry.Value.Select(i => sharedItemsCollection[i]);
					if (hiddenSharedCacheItems.Any(i => i.Value == cacheRecordValue))
						return false;
				}
			}
			return true;
		}

		//private bool FindGroupingRecordValueAndTupleMatch(CacheFieldNode cacheField, CacheRecordNode record, Tuple<int, int> tuple)
		//{
		//	if (cacheField.IsDateGrouping)
		//	{
		//		// Find record indices for date groupings fields.
		//		var recordIndices = this.DateGroupingRecordValueTupleMatch(cacheField, tuple.Item2);
		//		// If the record value is in the list, then the record value and tuple are a match.
		//		int index = tuple.Item1 < record.Items.Count ? tuple.Item1 : cacheField.FieldGroup.BaseField;
		//		var itemValue = record.Items[index].Value;
		//		if (recordIndices.All(i => i != int.Parse(itemValue)))
		//			return false;
		//	}
		//	else
		//	{
		//		// Use discrete grouping property collection to determine match.
		//		int baseIndex = cacheField.FieldGroup.BaseField;
		//		// Get the pivot field item's x value and the record item's v value.
		//		int pivotFieldValue = myFields[tuple.Item1].Items[tuple.Item2].X;
		//		var recordValue = int.Parse(record.Items[baseIndex].Value);
		//		var fieldGroup = cacheField.FieldGroup;
		//		// Get the shared item string the pivot field item and record item is pointing to.
		//		var pivotFieldPtrValue = fieldGroup.GroupItems[pivotFieldValue].Value;
		//		var recordDiscretePtrValue = int.Parse(fieldGroup.DiscreteGroupingProperties[recordValue].Value);
		//		var recordPtrValue = fieldGroup.GroupItems[recordDiscretePtrValue].Value;
		//		// Check if the pivot field item and record item is pointing to the same shared string.
		//		if (!pivotFieldPtrValue.IsEquivalentTo(recordPtrValue))
		//			return false;
		//	}
		//	return true;
		//}

		//private List<int> DateGroupingRecordValueTupleMatch(CacheFieldNode cacheField, int tupleItem2)
		//{
		//	var recordIndices = new List<int>();
		//	// Go through all the shared items and if the item is the targeted value (tuple.Item2 or groupItems[tuple.Item2]), 
		//	// add the index of the shared item to the list.
		//	var groupByType = cacheField.FieldGroup.GroupBy;
		//	string groupFieldItemsValue = cacheField.FieldGroup.GroupItems[tupleItem2].Value;

		//	int baseFieldIndex = cacheField.FieldGroup.BaseField;
		//	var sharedItems = myCacheDefinition.CacheFields[baseFieldIndex].SharedItems;

		//	for (int i = 0; i < sharedItems.Count; i++)
		//	{
		//		var dateTime = DateTime.Parse(sharedItems[i].Value);
		//		var groupByValue = string.Empty;

		//		// Get the sharedItem's groupBy value, unless the groupBy value is months.
		//		if (groupByType == PivotFieldDateGrouping.Months && tupleItem2 == dateTime.Month)
		//			recordIndices.Add(i);
		//		else if (groupByType == PivotFieldDateGrouping.Years)
		//			groupByValue = dateTime.Year.ToString();
		//		else if (groupByType == PivotFieldDateGrouping.Quarters)
		//			groupByValue = "Qtr" + ((dateTime.Month - 1) / 3 + 1);
		//		else if (groupByType == PivotFieldDateGrouping.Days)
		//			groupByValue = dateTime.Day + "-" + dateTime.ToString("MMM");
		//		else if (groupByType == PivotFieldDateGrouping.Minutes)
		//			groupByValue = ":" + dateTime.ToString("mm");
		//		else if (groupByType == PivotFieldDateGrouping.Seconds)
		//			groupByValue = ":" + dateTime.ToString("ss");
		//		else if (groupByType == PivotFieldDateGrouping.Hours)
		//		{
		//			int hour = dateTime.Hour == 00 ? 12 : dateTime.Hour;
		//			groupByValue = hour + " " + dateTime.ToString("tt", Thread.CurrentThread.CurrentCulture);
		//		}

		//		// Check if the sharedItem's groupBy value matches the groupFieldItem's value in the cacheField.
		//		if (groupFieldItemsValue.IsEquivalentTo(groupByValue))
		//			recordIndices.Add(i);
		//	}
		//	return recordIndices;
		//}
		#endregion
	}
}