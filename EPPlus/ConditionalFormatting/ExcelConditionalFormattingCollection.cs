﻿/*******************************************************************************
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
 * Author					Change						                Date
 * ******************************************************************************
 * Eyal Seagull		Conditional Formatting            2012-04-03
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using static OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingConstants;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// Collection of <see cref="ExcelConditionalFormattingRule"/>.
	/// This class is providing the API for EPPlus conditional formatting.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The public methods of this class (Add[...]ConditionalFormatting) will create a ConditionalFormatting/CfRule entry in the worksheet. When this
	/// Conditional Formatting has been created changes to the properties will affect the workbook immediately.
	/// </para>
	/// <para>
	/// Each type of Conditional Formatting Rule has diferente set of properties.
	/// </para>
	/// <code>
	/// // Add a Three Color Scale conditional formatting
	/// var cf = worksheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:C10"));
	/// // Set the conditional formatting properties
	/// cf.LowValue.Type = ExcelConditionalFormattingValueObjectType.Min;
	/// cf.LowValue.Color = Color.White;
	/// cf.MiddleValue.Type = ExcelConditionalFormattingValueObjectType.Percent;
	/// cf.MiddleValue.Value = 50;
	/// cf.MiddleValue.Color = Color.Blue;
	/// cf.HighValue.Type = ExcelConditionalFormattingValueObjectType.Max;
	/// cf.HighValue.Color = Color.Black;
	/// </code>
	/// </remarks>
	public class ExcelConditionalFormattingCollection
		: XmlHelper,
		IEnumerable<IExcelConditionalFormattingRule>
	{
		#region Properties 
		private List<IExcelConditionalFormattingRule> ConditionalFormattingRules { get; }
		private ExcelWorksheet ExcelWorksheet { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingCollection"/>
		/// </summary>
		/// <param name="worksheet">The worksheet from which to construct the ConditionalFormattings.</param>
		internal ExcelConditionalFormattingCollection(
		  ExcelWorksheet worksheet)
		  : base(
			 worksheet.NameSpaceManager,
			 worksheet.WorksheetXml.DocumentElement)
		{
			Require.Argument(worksheet).IsNotNull("worksheet");

			this.ExcelWorksheet = worksheet;
			this.SchemaNodeOrder = this.ExcelWorksheet.SchemaNodeOrder;
			this.ConditionalFormattingRules = new List<IExcelConditionalFormattingRule>();

			// Look for all the <conditionalFormatting> nodes.
			var conditionalFormattingNodes = this.TopNode.SelectNodes(
			  "//" + ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
				this.ExcelWorksheet.NameSpaceManager);

			if ((conditionalFormattingNodes != null) && (conditionalFormattingNodes.Count > 0))
			{
				foreach (XmlNode conditionalFormattingNode in conditionalFormattingNodes)
				{
					// Try to get the @sqref attribute. If it is missing, do not add a cf rule for this node.
					string sqref = conditionalFormattingNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref]?.Value;
					if (string.IsNullOrEmpty(sqref))
						continue;
					ExcelAddressBase address = new ExcelAddressBase(sqref);

					// Check for all the <cfRules> nodes and load them.
					var cfRuleNodes = conditionalFormattingNode.SelectNodes(
					  ExcelConditionalFormattingConstants.Paths.CfRule,
						this.ExcelWorksheet.NameSpaceManager);

					// Foreach <cfRule> inside the current <conditionalFormatting>.
					foreach (XmlNode cfRuleNode in cfRuleNodes)
					{
						// Check if @type attribute exists.
						if (cfRuleNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Type] == null)
							throw new Exception(ExcelConditionalFormattingConstants.Errors.MissingTypeAttribute);

						// Check if @priority attribute exists.
						if (cfRuleNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Priority] == null)
							throw new Exception(ExcelConditionalFormattingConstants.Errors.MissingPriorityAttribute);

						// Get the <cfRule> main attributes.
						string typeAttribute = ExcelConditionalFormattingHelper.GetAttributeString(
						  cfRuleNode,
						  ExcelConditionalFormattingConstants.Attributes.Type);

						int priority = ExcelConditionalFormattingHelper.GetAttributeInt(
						  cfRuleNode,
						  ExcelConditionalFormattingConstants.Attributes.Priority);

						// Transform the @type attribute to EPPlus Rule Type (slighty different).
						var type = ExcelConditionalFormattingRuleType.GetTypeByAttrbiute(
						  typeAttribute,
						  cfRuleNode,
							this.ExcelWorksheet.NameSpaceManager);

						var cfRule = ExcelConditionalFormattingRuleFactory.Create(
						  type,
						  address,
						  priority,
							this.ExcelWorksheet,
						  cfRuleNode);

						if (cfRule != null)
							this.ConditionalFormattingRules.Add(cfRule);
					}
				}
			}
		}
		#endregion Constructors

		#region Private Methods
		/// <summary>
		/// Throws an exception if the &lt;worksheet&gt; node is missing.
		/// </summary>
		private void EnsureRootElementExists()
		{
			// Find the <worksheet> node
			if (this.ExcelWorksheet.WorksheetXml.DocumentElement == null)
				throw new Exception(ExcelConditionalFormattingConstants.Errors.MissingWorksheetNode);
		}

		/// <summary>
		/// Gets the root node.
		/// </summary>
		/// <returns>Returns an <see cref="XmlNode"/> of the root node.</returns>
		private XmlNode GetRootNode()
		{
			EnsureRootElementExists();
			return this.ExcelWorksheet.WorksheetXml.DocumentElement;
		}

		/// <summary>
		/// Validates address that the given address is not null.
		/// </summary>
		/// <param name="address">The address to validate.</param>
		/// <returns>Returns the given <see cref="ExcelAddressBase"/> if it is valid.</returns>
		private ExcelAddressBase ValidateAddress(ExcelAddressBase address)
		{
			Require.Argument(address).IsNotNull("address");
			//TODO: Are there any other validation we need to do?
			return address;
		}

		/// <summary>
		/// Gets the next priority sequential number.
		/// </summary>
		/// <returns>Returns the next priority sequential integer.</returns>
		private int GetNextPriority()
		{
			// Consider zero as the last priority when we have no CF rules.
			int lastPriority = 0;
			foreach (var cfRule in this.ConditionalFormattingRules)
			{
				if (cfRule.Priority > lastPriority)
					lastPriority = cfRule.Priority;
			}
			return lastPriority + 1;
		}
		#endregion

		#region IEnumerable<IExcelConditionalFormatting>
		/// <summary>
		/// Counts the number of validations.
		/// </summary>
		/// <returns>Returns the number of validations in the collection.</returns>
		public int Count
		{
			get { return this.ConditionalFormattingRules.Count; }
		}

		/// <summary>
		/// Index operator, returns by 0-based index.
		/// </summary>
		/// <param name="index"></param>
		/// <returns>Returns the <see cref="IExcelConditionalFormattingRule"/> at the given index.</returns>
		public IExcelConditionalFormattingRule this[int index]
		{
			get { return this.ConditionalFormattingRules[index]; }
			set { this.ConditionalFormattingRules[index] = value; }
		}

		/// <summary>
		/// Gets the 'cfRule' enumerator.
		/// </summary>
		/// <returns>Returns an IEnumerator object that can be used to iterate through the collection.</returns>
		IEnumerator<IExcelConditionalFormattingRule> IEnumerable<IExcelConditionalFormattingRule>.GetEnumerator()
		{
			return this.ConditionalFormattingRules.GetEnumerator();
		}

		/// <summary>
		/// Get the 'cfRule' enumerator.
		/// </summary>
		/// <returns></returns>
		IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.ConditionalFormattingRules.GetEnumerator();
		}

		/// <summary>
		/// Removes all 'cfRule' from the collection and from the XML.
		/// <remarks>
		/// This is the same as removing all the 'conditionalFormatting' nodes.
		/// </remarks>
		/// </summary>
		public void RemoveAll()
		{
			// Look for all the <conditionalFormatting> nodes
			var conditionalFormattingNodes = this.TopNode.SelectNodes(
			  "//" + ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
				this.ExcelWorksheet.NameSpaceManager);

			// Remove all the <conditionalFormatting> nodes one by one
			foreach (XmlNode conditionalFormattingNode in conditionalFormattingNodes)
			{
				conditionalFormattingNode.ParentNode.RemoveChild(conditionalFormattingNode);
			}

			// Clear the <cfRule> item list
			this.ConditionalFormattingRules.Clear();
		}

		/// <summary>
		/// Remove a Conditional Formatting Rule by its object.
		/// </summary>
		/// <param name="item">The item to remove.</param>
		public void Remove(IExcelConditionalFormattingRule item)
		{
			Require.Argument(item).IsNotNull("item");
			try
			{
				// Point to the parent node
				var oldParentNode = item.Node.ParentNode;
				// Remove the <cfRule> from the old <conditionalFormatting> parent node
				oldParentNode.RemoveChild(item.Node);
				// Check if the old <conditionalFormatting> parent node has <cfRule> node inside it
				if (!oldParentNode.HasChildNodes)
					oldParentNode.ParentNode.RemoveChild(oldParentNode);
				this.ConditionalFormattingRules.Remove(item);
			}
			catch
			{
				throw new Exception(ExcelConditionalFormattingConstants.Errors.InvalidRemoveRuleOperation);
			}
		}

		/// <summary>
		/// Remove a Conditional Formatting Rule by its 0-based index.
		/// </summary>
		/// <param name="index">The index of the item to remove.</param>
		public void RemoveAt(int index)
		{
			Require.Argument(index).IsInRange(0, this.Count - 1, "index");
			Remove(this[index]);
		}

		/// <summary>
		/// Remove a Conditional Formatting Rule by its priority.
		/// </summary>
		/// <param name="priority">The priority of the rules to be removed.</param>
		public void RemoveByPriority(int priority)
		{
			try
			{
				Remove(RulesByPriority(priority));
			}
			catch { }
		}

		/// <summary>
		/// Get a rule by its priority.
		/// </summary>
		/// <param name="priority">The priority of the rule to get.</param>
		/// <returns>Returns the <see cref="IExcelConditionalFormattingRule"/> with the given priority.</returns>
		public IExcelConditionalFormattingRule RulesByPriority(int priority)
		{
			return this.ConditionalFormattingRules.Find(x => x.Priority == priority);
		}
		#endregion IEnumerable<IExcelConditionalFormatting>

		#region Conditional Formatting Rule Methods
		/// <summary>
		/// Add rule (internal)
		/// </summary>
		/// <param name="type">The <see cref="eExcelConditionalFormattingRuleType"/> of the rule to add.</param>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added to the collection.</returns>
		internal IExcelConditionalFormattingRule AddRule(eExcelConditionalFormattingRuleType type, ExcelAddressBase address)
		{
			Require.Argument(address).IsNotNull("address");
			address = ValidateAddress(address);
			EnsureRootElementExists();
			IExcelConditionalFormattingRule cfRule = ExcelConditionalFormattingRuleFactory.Create(
			  type,
			  address,
			  GetNextPriority(),
				this.ExcelWorksheet,
			  null);
			this.ConditionalFormattingRules.Add(cfRule);
			return cfRule;
		}

		/// <summary>
		/// Add AboveAverage Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingAverageGroup AddAboveAverage(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingAverageGroup)AddRule(
			  eExcelConditionalFormattingRuleType.AboveAverage,
			  address);
		}

		/// <summary>
		/// Add AboveOrEqualAverage Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage(
		  ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingAverageGroup)AddRule(
			  eExcelConditionalFormattingRuleType.AboveOrEqualAverage,
			  address);
		}

		/// <summary>
		/// Add BelowAverage Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingAverageGroup AddBelowAverage(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingAverageGroup)AddRule(
			  eExcelConditionalFormattingRuleType.BelowAverage,
			  address);
		}

		/// <summary>
		/// Add BelowOrEqualAverage Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingAverageGroup)AddRule(
			  eExcelConditionalFormattingRuleType.BelowOrEqualAverage,
			  address);
		}

		/// <summary>
		/// Add AboveStdDev Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingStdDevGroup AddAboveStdDev(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingStdDevGroup)AddRule(
			  eExcelConditionalFormattingRuleType.AboveStdDev,
			  address);
		}

		/// <summary>
		/// Add BelowStdDev Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingStdDevGroup AddBelowStdDev(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingStdDevGroup)AddRule(
			  eExcelConditionalFormattingRuleType.BelowStdDev,
			  address);
		}

		/// <summary>
		/// Add Bottom Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTopBottomGroup AddBottom(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTopBottomGroup)AddRule(
			  eExcelConditionalFormattingRuleType.Bottom,
			  address);
		}

		/// <summary>
		/// Add BottomPercent Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTopBottomGroup AddBottomPercent(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTopBottomGroup)AddRule(
			  eExcelConditionalFormattingRuleType.BottomPercent,
			  address);
		}

		/// <summary>
		/// Add Top Rule
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTopBottomGroup AddTop(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTopBottomGroup)AddRule(
			  eExcelConditionalFormattingRuleType.Top,
			  address);
		}

		/// <summary>
		/// Add TopPercent Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTopBottomGroup AddTopPercent(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTopBottomGroup)AddRule(
			  eExcelConditionalFormattingRuleType.TopPercent,
			  address);
		}

		/// <summary>
		/// Add Last7Days Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.Last7Days,
			  address);
		}

		/// <summary>
		/// Add LastMonth Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddLastMonth(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.LastMonth,
			  address);
		}

		/// <summary>
		/// Add LastWeek Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.LastWeek,
			  address);
		}

		/// <summary>
		/// Add NextMonth Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddNextMonth(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.NextMonth,
			  address);
		}

		/// <summary>
		/// Add NextWeek Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.NextWeek,
			  address);
		}

		/// <summary>
		/// Add ThisMonth Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddThisMonth(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.ThisMonth,
			  address);
		}

		/// <summary>
		/// Add ThisWeek Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.ThisWeek,
			  address);
		}

		/// <summary>
		/// Add Today Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddToday(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.Today,
			  address);
		}

		/// <summary>
		/// Add Tomorrow Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.Tomorrow,
			  address);
		}

		/// <summary>
		/// Add Yesterday Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTimePeriodGroup AddYesterday(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
			  eExcelConditionalFormattingRuleType.Yesterday,
			  address);
		}

		/// <summary>
		/// Add BeginsWith Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingBeginsWith AddBeginsWith(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingBeginsWith)AddRule(
			  eExcelConditionalFormattingRuleType.BeginsWith,
			  address);
		}

		/// <summary>
		/// Add Between Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingBetween AddBetween(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingBetween)AddRule(
			  eExcelConditionalFormattingRuleType.Between,
			  address);
		}

		/// <summary>
		/// Add ContainsBlanks Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingContainsBlanks AddContainsBlanks(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingContainsBlanks)AddRule(
			  eExcelConditionalFormattingRuleType.ContainsBlanks,
			  address);
		}

		/// <summary>
		/// Add ContainsErrors Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingContainsErrors AddContainsErrors(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingContainsErrors)AddRule(
			  eExcelConditionalFormattingRuleType.ContainsErrors,
			  address);
		}

		/// <summary>
		/// Add ContainsText Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingContainsText AddContainsText(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingContainsText)AddRule(
			  eExcelConditionalFormattingRuleType.ContainsText,
			  address);
		}

		/// <summary>
		/// Add DuplicateValues Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingDuplicateValues AddDuplicateValues(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingDuplicateValues)AddRule(
			  eExcelConditionalFormattingRuleType.DuplicateValues,
			  address);
		}

		/// <summary>
		/// Add EndsWith Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingEndsWith AddEndsWith(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingEndsWith)AddRule(
			  eExcelConditionalFormattingRuleType.EndsWith,
			  address);
		}

		/// <summary>
		/// Add Equal Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingEqual AddEqual(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingEqual)AddRule(
			  eExcelConditionalFormattingRuleType.Equal,
			  address);
		}

		/// <summary>
		/// Add Expression Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingExpression AddExpression(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingExpression)AddRule(
			  eExcelConditionalFormattingRuleType.Expression,
			  address);
		}

		/// <summary>
		/// Add GreaterThan Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingGreaterThan AddGreaterThan(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingGreaterThan)AddRule(
			  eExcelConditionalFormattingRuleType.GreaterThan,
			  address);
		}

		/// <summary>
		/// Add GreaterThanOrEqual Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingGreaterThanOrEqual)AddRule(
			  eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
			  address);
		}

		/// <summary>
		/// Add LessThan Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingLessThan AddLessThan(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingLessThan)AddRule(
			  eExcelConditionalFormattingRuleType.LessThan,
			  address);
		}

		/// <summary>
		/// Add LessThanOrEqual Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingLessThanOrEqual)AddRule(
			  eExcelConditionalFormattingRuleType.LessThanOrEqual,
			  address);
		}

		/// <summary>
		/// Add NotBetween Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingNotBetween AddNotBetween(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingNotBetween)AddRule(
			  eExcelConditionalFormattingRuleType.NotBetween,
			  address);
		}

		/// <summary>
		/// Add NotContainsBlanks Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingNotContainsBlanks)AddRule(
			  eExcelConditionalFormattingRuleType.NotContainsBlanks,
			  address);
		}

		/// <summary>
		/// Add NotContainsErrors Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingNotContainsErrors)AddRule(
			  eExcelConditionalFormattingRuleType.NotContainsErrors,
			  address);
		}

		/// <summary>
		/// Add NotContainsText Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingNotContainsText AddNotContainsText(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingNotContainsText)AddRule(
			  eExcelConditionalFormattingRuleType.NotContainsText,
			  address);
		}

		/// <summary>
		/// Add NotEqual Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingNotEqual AddNotEqual(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingNotEqual)AddRule(
			  eExcelConditionalFormattingRuleType.NotEqual,
			  address);
		}

		/// <summary>
		/// Add Unique Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingUniqueValues AddUniqueValues(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingUniqueValues)AddRule(
			  eExcelConditionalFormattingRuleType.UniqueValues,
			  address);
		}

		/// <summary>
		/// Add ThreeColorScale Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingThreeColorScale AddThreeColorScale(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingThreeColorScale)AddRule(
			  eExcelConditionalFormattingRuleType.ThreeColorScale,
			  address);
		}

		/// <summary>
		/// Add TwoColorScale Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingTwoColorScale AddTwoColorScale(ExcelAddressBase address)
		{
			return (IExcelConditionalFormattingTwoColorScale)AddRule(
			  eExcelConditionalFormattingRuleType.TwoColorScale,
			  address);
		}

		/// <summary>
		/// Add ThreeIconSet Rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <param name="iconSet">Type of iconset</param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(ExcelAddressBase address, eExcelconditionalFormatting3IconsSetType iconSet)
		{
			var icon = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)AddRule(
				 eExcelConditionalFormattingRuleType.ThreeIconSet,
				 address);
			icon.IconSet = iconSet;
			return icon;
		}

		/// <summary>
		/// Adds a FourIconSet rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <param name="iconSet"></param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(ExcelAddressBase address, eExcelconditionalFormatting4IconsSetType iconSet)
		{
			var icon = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)AddRule(
				 eExcelConditionalFormattingRuleType.FourIconSet,
				 address);
			icon.IconSet = iconSet;
			return icon;
		}

		/// <summary>
		/// Adds a FiveIconSet rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <param name="iconSet"></param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(ExcelAddressBase address, eExcelconditionalFormatting5IconsSetType iconSet)
		{
			var icon = (IExcelConditionalFormattingFiveIconSet)AddRule(
				 eExcelConditionalFormattingRuleType.FiveIconSet,
				 address);
			icon.IconSet = iconSet;
			return icon;
		}

		/// <summary>
		/// Adds a databar rule.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> of the rule to add.</param>
		/// <param name="color"></param>
		/// <returns>Returns the rule added.</returns>
		public IExcelConditionalFormattingDataBarGroup AddDatabar(ExcelAddressBase address, Color color)
		{
			var dataBar = (IExcelConditionalFormattingDataBarGroup)AddRule(
				 eExcelConditionalFormattingRuleType.DataBar,
				 address);
			dataBar.Color = color;
			return dataBar;
		}
		#endregion Conditional Formatting Rules

		#region Virtual Methods
		/// <summary>
		/// Applies the <paramref name="transformer"/> to all formulas in the <see cref="ExcelConditionalFormattingCollection"/>.
		/// </summary>
		/// <param name="transformer">The transformation to apply.</param>
		public virtual void TransformFormulaReferences(Func<string, string> transformer)
		{
			XmlHelper.TransformValuesInNode(this.TopNode, this.NameSpaceManager, transformer, ".//d:conditionalFormatting//d:formula");
			XmlHelper.TransformAttributesInNode(this.TopNode, this.NameSpaceManager, transformer, ".//d:conditionalFormatting//d:cfvo", Attributes.Val);
		}
		#endregion
	}
}